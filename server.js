require("dotenv").config();

const express = require("express");
const multer = require("multer");
const pdfParse = require("pdf-parse");
const mammoth = require("mammoth");
const WordExtractor = require("word-extractor");
const OpenAI = require("openai");
const fs = require("fs/promises");
const os = require("os");
const path = require("path");

const app = express();
const port = Number(process.env.PORT) || 3000;
const provider = String(process.env.LLM_PROVIDER || "openai").toLowerCase();
const openaiModel = process.env.OPENAI_MODEL || "gpt-4.1-mini";
const geminiModel = process.env.GEMINI_MODEL || "gemini-2.5-flash";
const dataDir = path.join(__dirname, "data");
const customPromptsPath = path.join(dataDir, "custom-prompts.json");

if (provider === "openai" && !process.env.OPENAI_API_KEY) {
  console.warn("OPENAI_API_KEY is not set. Requests to /api/grade will fail until it is configured.");
}
if (provider === "gemini" && !process.env.GEMINI_API_KEY) {
  console.warn("GEMINI_API_KEY is not set. Requests to /api/grade will fail until it is configured.");
}

const openai = process.env.OPENAI_API_KEY ? new OpenAI({ apiKey: process.env.OPENAI_API_KEY }) : null;

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 12 * 1024 * 1024
  }
});

const PROMPT_PROFILES = {
  graduate_professor:
    "You are a rigorous graduate-level professor. Evaluate with high standards for depth, originality, precision, and scholarly quality.",
  high_school_teacher:
    "You are a high school teacher. Evaluate for clarity, structure, factual correctness, and age-appropriate expectations.",
  supportive_tutor:
    "You are a supportive tutor. Grade fairly and provide highly actionable feedback that helps the student improve quickly.",
  strict_examiner:
    "You are a strict examiner. Prioritize adherence to rubric criteria and assignment requirements over stylistic generosity."
};

app.use(express.static("public"));
app.use(express.json());

function getExtension(name) {
  return path.extname(name || "").toLowerCase();
}

async function extractDocText(buffer) {
  const tempPath = path.join(os.tmpdir(), `upload-${Date.now()}-${Math.random().toString(16).slice(2)}.doc`);

  try {
    await fs.writeFile(tempPath, buffer);
    const extractor = new WordExtractor();
    const extracted = await extractor.extract(tempPath);
    return extracted.getBody();
  } finally {
    await fs.unlink(tempPath).catch(() => {});
  }
}

function htmlToSignalText(html) {
  let value = String(html || "");
  value = value.replace(/<\s*(em|i)\b[^>]*>/gi, "[ITALIC]");
  value = value.replace(/<\s*\/\s*(em|i)\s*>/gi, "[/ITALIC]");
  value = value.replace(/<\s*br\s*\/?\s*>/gi, "\n");
  value = value.replace(/<\s*\/\s*p\s*>/gi, "\n");
  value = value.replace(/<[^>]+>/g, " ");
  value = value
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
  return value;
}

async function extractTextFromFile(file, options = {}) {
  const preserveDocxFormatting = Boolean(options.preserveDocxFormatting);
  const ext = getExtension(file.originalname);

  if (ext === ".pdf") {
    const parsed = await pdfParse(file.buffer);
    return parsed.text;
  }

  if (ext === ".docx" && preserveDocxFormatting) {
    const parsedHtml = await mammoth.convertToHtml({ buffer: file.buffer });
    return htmlToSignalText(parsedHtml.value);
  }

  if (ext === ".docx") {
    const parsed = await mammoth.extractRawText({ buffer: file.buffer });
    return parsed.value;
  }

  if (ext === ".doc") {
    return extractDocText(file.buffer);
  }

  throw new Error(`Unsupported file type: ${ext || "unknown"}. Please upload PDF, DOC, or DOCX files.`);
}

function normalizeText(text) {
  return String(text || "")
    .replace(/\r/g, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function normalizeSubmissionType(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "discussion") {
    return "discussion";
  }
  if (normalized === "discussion_reply") {
    return "discussion_reply";
  }
  return "none";
}

function isReplyCriterionText(text) {
  const value = String(text || "").toLowerCase();
  return (
    value.includes("reply") ||
    value.includes("replies") ||
    value.includes("peer response") ||
    value.includes("response to peer") ||
    value.includes("respond to peer") ||
    value.includes("classmate response") ||
    value.includes("discussion response")
  );
}

function isItalicIssueText(text) {
  const value = String(text || "").toLowerCase();
  return value.includes("italic") || value.includes("italics") || value.includes("non-italic");
}

function filterBySubmissionType(items, submissionType, textAccessor) {
  if (!Array.isArray(items) || submissionType === "none") {
    return items || [];
  }

  return items.filter((item) => {
    const isReply = isReplyCriterionText(textAccessor(item));
    if (submissionType === "discussion") {
      return !isReply;
    }
    if (submissionType === "discussion_reply") {
      return isReply;
    }
    return true;
  });
}

function calculateWordCount(text) {
  const normalized = String(text || "").replace(/\[\/?ITALIC\]/g, " ");
  const matches = normalized.match(/[A-Za-z0-9]+(?:[’'-][A-Za-z0-9]+)*/g);
  return matches ? matches.length : 0;
}

async function ensureCustomPromptStore() {
  await fs.mkdir(dataDir, { recursive: true });
  try {
    await fs.access(customPromptsPath);
  } catch {
    await fs.writeFile(customPromptsPath, "[]", "utf8");
  }
}

async function readCustomPrompts() {
  await ensureCustomPromptStore();
  const raw = await fs.readFile(customPromptsPath, "utf8");
  const parsed = JSON.parse(raw);
  if (!Array.isArray(parsed)) {
    return [];
  }
  return parsed;
}

async function writeCustomPrompts(prompts) {
  await ensureCustomPromptStore();
  await fs.writeFile(customPromptsPath, JSON.stringify(prompts, null, 2), "utf8");
}

function buildMessages({
  rubric,
  instructions,
  submission,
  selectedPrompt,
  otherContext,
  submissionType,
  submissionWordCount
}) {
  const modeRules =
    submissionType === "discussion"
      ? [
          "Submission type is Discussion.",
          "Ignore ALL reply/peer-response requirements in rubric and instructions.",
          "Do not deduct for missing replies.",
          "Set pointsPossible using only non-reply criteria."
        ]
      : submissionType === "discussion_reply"
        ? [
            "Submission type is Discussion Reply.",
            "Grade ONLY reply/peer-response requirements and criteria.",
            "Ignore non-reply criteria entirely.",
            "Set pointsPossible using only reply-related criteria."
          ]
        : ["Submission type is Standard. Grade all criteria as provided."];

  const system = [
    "You are an expert academic evaluator.",
    "Grade the student submission using ONLY the provided rubric and assignment instructions.",
    "If rubric criteria are unclear, make minimal assumptions and explain them.",
    "For every deduction, cite specific evidence from the submission using exact snippets from the text.",
    "Evidence must be concrete, actionable, and tied to a correction.",
    "Formatting penalties must be conservative: do not penalize italicization unless missing italics are clearly evidenced.",
    "Output valid JSON only with this exact schema:",
    "{",
    '  "pointsEarned": number,',
    '  "pointsPossible": number,',
    '  "letterGrade": string,',
    '  "rubricBreakdown": [{ "criterion": string, "score": number, "maxScore": number, "rationale": string }],',
    '  "deductions": [{',
    '    "criterion": string,',
    '    "pointsLost": number,',
    '    "reason": string,',
    '    "actionableFix": string,',
    '    "evidence": [{ "location": string, "snippet": string, "issue": string, "suggestion": string }]',
    "  }],",
    '  "citationsNeeded": [{',
    '    "location": string,',
    '    "claimText": string,',
    '    "whyCitationNeeded": string,',
    '    "suggestedCitationType": string',
    "  }]",
    "}"
  ].join(" ");

  const user = [
    `Evaluation mode: ${selectedPrompt}`,
    "",
    "[Grading Rubric]",
    rubric,
    "",
    "[Assignment Instructions]",
    instructions,
    "",
    "[Student Submission - Original]",
    submission,
    "",
    `[Authoritative Submission Word Count] ${submissionWordCount}`,
    "",
    "[Additional Context]",
    otherContext || "None provided.",
    "",
    "Rules:",
    "1) Grade against rubric and instructions.",
    "2) Give concise but specific feedback with exact evidence locations.",
    "3) pointsEarned must be between 0 and pointsPossible.",
    "4) Ensure rubricBreakdown totals reasonably align with pointsEarned/pointsPossible.",
    "5) Include at least one evidence item for each deduction.",
    "5a) For grammar deductions, include exact problematic snippet(s) and clear correction guidance.",
    "6) If citations are missing, include citationsNeeded entries showing exactly where citation is required.",
    "7) Use the authoritative submission word count provided above. Do not re-count words yourself.",
    "8) The submission may include [ITALIC]...[/ITALIC] markers from DOCX. Treat these as explicit italic formatting evidence.",
    "9) Do not use line-number labels (L1, L2, etc.) in evidence locations."
  ]
    .concat(modeRules)
    .join("\n");

  return [
    { role: "system", content: system },
    { role: "user", content: user }
  ];
}

app.get("/api/custom-prompts", async (_req, res) => {
  try {
    const prompts = await readCustomPrompts();
    res.json({
      prompts: prompts.map((item) => ({
        id: String(item.id || ""),
        name: String(item.name || "Unnamed saved prompt"),
        text: String(item.text || "")
      }))
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to load saved prompts." });
  }
});

app.post("/api/custom-prompts", async (req, res) => {
  try {
    const name = normalizeText(req.body?.name);
    const text = normalizeText(req.body?.text);

    if (!name || !text) {
      return res.status(400).json({ error: "Prompt name and prompt text are required." });
    }

    const prompts = await readCustomPrompts();
    const id = `saved_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`;
    const next = [...prompts, { id, name, text }];
    await writeCustomPrompts(next);

    res.status(201).json({ prompt: { id, name, text } });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to save custom prompt." });
  }
});

app.delete("/api/custom-prompts/:id", async (req, res) => {
  try {
    const promptId = normalizeText(req.params.id);
    if (!promptId) {
      return res.status(400).json({ error: "Prompt id is required." });
    }

    const prompts = await readCustomPrompts();
    const existing = prompts.find((item) => String(item.id) === promptId);
    if (!existing) {
      return res.status(404).json({ error: "Saved prompt not found." });
    }

    const next = prompts.filter((item) => String(item.id) !== promptId);
    await writeCustomPrompts(next);

    res.json({ success: true });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to delete saved prompt." });
  }
});

function parseModelJson(content) {
  try {
    return JSON.parse(content);
  } catch {
    const match = content.match(/\{[\s\S]*\}$/);
    if (!match) {
      throw new Error("Model response was not valid JSON.");
    }
    return JSON.parse(match[0]);
  }
}

function hasMissingEvidenceForDeductions(parsed) {
  const deductions = Array.isArray(parsed?.deductions) ? parsed.deductions : [];
  if (deductions.length === 0) {
    return false;
  }

  return deductions.some((deduction) => {
    const evidence = Array.isArray(deduction?.evidence) ? deduction.evidence : [];
    if (evidence.length === 0) {
      return true;
    }
    return evidence.some((item) => !String(item?.snippet || "").trim());
  });
}

function roundTo2(value) {
  return Math.round(value * 100) / 100;
}

function normalizePointsResult(parsed, submissionType, submission) {
  const hasItalicMarkers = String(submission || "").includes("[ITALIC]");

  let rubric = Array.isArray(parsed.rubricBreakdown)
    ? parsed.rubricBreakdown.map((item) => ({
        criterion: String(item.criterion || "Unnamed criterion"),
        score: Number(item.score) || 0,
        maxScore: Number(item.maxScore) || 0,
        rationale: String(item.rationale || "")
      }))
    : [];

  rubric = filterBySubmissionType(rubric, submissionType, (item) => `${item.criterion} ${item.rationale}`);
  rubric = rubric.map((item) => {
    if (hasItalicMarkers && isItalicIssueText(`${item.criterion} ${item.rationale}`)) {
      return {
        ...item,
        score: item.maxScore,
        rationale: item.rationale
          ? `${item.rationale} Italic formatting markers were detected; italicization is treated as satisfied unless explicit contrary evidence exists.`
          : "Italic formatting markers were detected; italicization is treated as satisfied unless explicit contrary evidence exists."
      };
    }
    return item;
  });

  const rubricPointsEarned = rubric.reduce((sum, item) => sum + (Number.isFinite(item.score) ? item.score : 0), 0);
  const rubricPointsPossible = rubric.reduce((sum, item) => sum + (Number.isFinite(item.maxScore) ? item.maxScore : 0), 0);

  let pointsPossible = rubricPointsPossible > 0 ? rubricPointsPossible : Number(parsed.pointsPossible);
  if (!Number.isFinite(pointsPossible) || pointsPossible <= 0) {
    pointsPossible = 100;
  }

  let pointsEarned = rubricPointsPossible > 0 ? rubricPointsEarned : Number(parsed.pointsEarned);
  if (!Number.isFinite(pointsEarned)) {
    const overallScore = Number(parsed.overallScore);
    if (Number.isFinite(overallScore)) {
      pointsEarned = (overallScore / 100) * pointsPossible;
    } else {
      pointsEarned = rubricPointsEarned;
    }
  }

  pointsEarned = Math.max(0, Math.min(pointsEarned, pointsPossible));

  let deductions = Array.isArray(parsed.deductions)
    ? parsed.deductions.map((item) => ({
        criterion: String(item.criterion || "General"),
        pointsLost: Number(item.pointsLost) || 0,
        reason: String(item.reason || ""),
        actionableFix: String(item.actionableFix || ""),
        evidence: Array.isArray(item.evidence)
          ? item.evidence.map((e) => ({
              location: String(e.location || ""),
              snippet: String(e.snippet || ""),
              issue: String(e.issue || ""),
              suggestion: String(e.suggestion || "")
            }))
          : []
      }))
    : [];
  deductions = filterBySubmissionType(deductions, submissionType, (item) => `${item.criterion} ${item.reason}`);
  if (hasItalicMarkers) {
    deductions = deductions.filter((item) => !isItalicIssueText(`${item.criterion} ${item.reason}`));
  }
  deductions = deductions.filter((item) =>
    Array.isArray(item.evidence) && item.evidence.some((evidence) => String(evidence.snippet || "").trim())
  );

  const citationsNeeded = Array.isArray(parsed.citationsNeeded)
    ? parsed.citationsNeeded.map((item) => ({
        location: String(item.location || ""),
        claimText: String(item.claimText || ""),
        whyCitationNeeded: String(item.whyCitationNeeded || ""),
        suggestedCitationType: String(item.suggestedCitationType || "")
      }))
    : [];

  return {
    pointsEarned: roundTo2(pointsEarned),
    pointsPossible: roundTo2(pointsPossible),
    letterGrade: String(parsed.letterGrade || "N/A"),
    summary: String(parsed.summary || "No summary returned."),
    strengths: Array.isArray(parsed.strengths) ? parsed.strengths.map(String) : [],
    improvements: Array.isArray(parsed.improvements) ? parsed.improvements.map(String) : [],
    rubricBreakdown: rubric,
    deductions,
    citationsNeeded
  };
}

async function generateWithOpenAI(messages) {
  if (!openai) {
    throw new Error("Server is missing OPENAI_API_KEY.");
  }

  const completion = await openai.chat.completions.create({
    model: openaiModel,
    response_format: { type: "json_object" },
    messages,
    temperature: 0.2
  });

  return completion.choices?.[0]?.message?.content || "{}";
}

async function generateWithGemini(messages) {
  if (!process.env.GEMINI_API_KEY) {
    throw new Error("Server is missing GEMINI_API_KEY.");
  }

  const systemText = messages.find((message) => message.role === "system")?.content || "";
  const userText = messages
    .filter((message) => message.role === "user")
    .map((message) => message.content)
    .join("\n\n");

  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${geminiModel}:generateContent`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-goog-api-key": process.env.GEMINI_API_KEY
      },
      body: JSON.stringify({
        systemInstruction: { parts: [{ text: systemText }] },
        contents: [{ role: "user", parts: [{ text: userText }] }],
        generationConfig: {
          temperature: 0.2,
          responseMimeType: "application/json"
        }
      })
    }
  );

  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    const detail = data?.error?.message || `HTTP ${response.status}`;
    throw new Error(`Gemini API error: ${detail}`);
  }

  const text = data?.candidates?.[0]?.content?.parts?.map((part) => part.text || "").join("\n").trim();
  if (!text) {
    throw new Error("Gemini returned an empty response.");
  }

  return text;
}

async function generateGradeOutput(messages) {
  if (provider === "openai") {
    return generateWithOpenAI(messages);
  }
  if (provider === "gemini") {
    return generateWithGemini(messages);
  }
  throw new Error(`Invalid LLM_PROVIDER '${provider}'. Use 'openai' or 'gemini'.`);
}

app.post(
  "/api/grade",
  upload.fields([
    { name: "rubric", maxCount: 1 },
    { name: "instructions", maxCount: 1 },
    { name: "submission", maxCount: 1 },
    { name: "other", maxCount: 1 }
  ]),
  async (req, res) => {
    try {
      if (provider === "openai" && !process.env.OPENAI_API_KEY) {
        return res.status(500).json({ error: "Server is missing OPENAI_API_KEY." });
      }
      if (provider === "gemini" && !process.env.GEMINI_API_KEY) {
        return res.status(500).json({ error: "Server is missing GEMINI_API_KEY." });
      }

      const rubricFile = req.files?.rubric?.[0];
      const instructionsFile = req.files?.instructions?.[0];
      const submissionFile = req.files?.submission?.[0];
      const otherFile = req.files?.other?.[0];

      if (!rubricFile || !instructionsFile || !submissionFile) {
        return res.status(400).json({
          error: "Please upload all required files: rubric, assignment instructions, and student submission."
        });
      }

      const profile = String(req.body.promptProfile || "");
      const promptInstructions = normalizeText(req.body.promptInstructions);
      const submissionType = normalizeSubmissionType(req.body.submissionType);
      const otherRelevance = normalizeText(req.body.otherRelevance);

      let selectedPrompt;
      if (profile === "custom") {
        selectedPrompt = promptInstructions;
      } else if (profile.startsWith("saved:")) {
        const savedId = profile.slice("saved:".length);
        const prompts = await readCustomPrompts();
        const saved = prompts.find((item) => String(item.id) === savedId);
        if (!saved) {
          return res.status(400).json({ error: "Selected saved prompt was not found." });
        }
        selectedPrompt = normalizeText(saved.text);
      } else if (PROMPT_PROFILES[profile]) {
        selectedPrompt = PROMPT_PROFILES[profile];
      } else {
        return res.status(400).json({ error: "Invalid prompt profile selected." });
      }

      const finalPrompt = promptInstructions || selectedPrompt;
      if (!finalPrompt) {
        return res.status(400).json({ error: "Prompt Instructions is required." });
      }

      if (otherFile && !otherRelevance) {
        return res.status(400).json({ error: "Please explain why the 'Other' file is relevant." });
      }

      const [rubricText, instructionsText, submissionText, otherTextRaw] = await Promise.all([
        extractTextFromFile(rubricFile),
        extractTextFromFile(instructionsFile),
        extractTextFromFile(submissionFile, { preserveDocxFormatting: true }),
        otherFile ? extractTextFromFile(otherFile) : Promise.resolve("")
      ]);

      const rubric = normalizeText(rubricText);
      const instructions = normalizeText(instructionsText);
      const submission = normalizeText(submissionText);
      const submissionWordCount = calculateWordCount(submission);
      const otherText = normalizeText(otherTextRaw);

      if (!rubric || !instructions || !submission) {
        return res.status(400).json({
          error: "One or more uploaded files had no readable text. Please verify the document contents."
        });
      }

      const otherContext = otherFile
        ? `Reason this file is relevant: ${otherRelevance}\n\n[Other Supporting File]\n${otherText || "No readable text found in other file."}`
        : "";

      const messages = buildMessages({
        rubric,
        instructions,
        submission,
        selectedPrompt: finalPrompt,
        otherContext,
        submissionType,
        submissionWordCount
      });
      let raw = await generateGradeOutput(messages);
      let parsed = parseModelJson(raw);

      if (hasMissingEvidenceForDeductions(parsed)) {
        const retryMessages = [
          ...messages,
          {
            role: "user",
            content:
              "Regenerate the JSON. Every deduction must include at least one evidence item with a non-empty exact snippet from the submission."
          }
        ];
        raw = await generateGradeOutput(retryMessages);
        parsed = parseModelJson(raw);
      }

      const normalized = normalizePointsResult(parsed, submissionType, submission);

      res.json({
        result: {
          ...normalized,
          submissionWordCount
        }
      });
    } catch (error) {
      console.error(error);
      res.status(500).json({
        error: error.message || "Failed to grade submission."
      });
    }
  }
);

app.listen(port, () => {
  console.log(`Assignment grader is running on http://localhost:${port}`);
});
