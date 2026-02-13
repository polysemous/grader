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

async function extractTextFromFile(file) {
  const ext = getExtension(file.originalname);

  if (ext === ".pdf") {
    const parsed = await pdfParse(file.buffer);
    return parsed.text;
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

function buildMessages({ rubric, instructions, submission, selectedPrompt, otherContext }) {
  const system = [
    "You are an expert academic evaluator.",
    "Grade the student submission using ONLY the provided rubric and assignment instructions.",
    "If rubric criteria are unclear, make minimal assumptions and explain them.",
    "Output valid JSON only with this exact schema:",
    "{",
    '  "pointsEarned": number,',
    '  "pointsPossible": number,',
    '  "letterGrade": string,',
    '  "summary": string,',
    '  "strengths": string[],',
    '  "improvements": string[],',
    '  "rubricBreakdown": [{ "criterion": string, "score": number, "maxScore": number, "rationale": string }]',
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
    "[Student Submission]",
    submission,
    "",
    "[Additional Context]",
    otherContext || "None provided.",
    "",
    "Rules:",
    "1) Grade against rubric and instructions.",
    "2) Give concise but specific feedback.",
    "3) pointsEarned must be between 0 and pointsPossible.",
    "4) Ensure rubricBreakdown totals reasonably align with pointsEarned/pointsPossible."
  ].join("\n");

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

function roundTo2(value) {
  return Math.round(value * 100) / 100;
}

function normalizePointsResult(parsed) {
  const rubric = Array.isArray(parsed.rubricBreakdown)
    ? parsed.rubricBreakdown.map((item) => ({
        criterion: String(item.criterion || "Unnamed criterion"),
        score: Number(item.score) || 0,
        maxScore: Number(item.maxScore) || 0,
        rationale: String(item.rationale || "")
      }))
    : [];

  const rubricPointsEarned = rubric.reduce((sum, item) => sum + (Number.isFinite(item.score) ? item.score : 0), 0);
  const rubricPointsPossible = rubric.reduce((sum, item) => sum + (Number.isFinite(item.maxScore) ? item.maxScore : 0), 0);

  let pointsPossible = Number(parsed.pointsPossible);
  if (!Number.isFinite(pointsPossible) || pointsPossible <= 0) {
    pointsPossible = rubricPointsPossible > 0 ? rubricPointsPossible : 100;
  }

  let pointsEarned = Number(parsed.pointsEarned);
  if (!Number.isFinite(pointsEarned)) {
    const overallScore = Number(parsed.overallScore);
    if (Number.isFinite(overallScore)) {
      pointsEarned = (overallScore / 100) * pointsPossible;
    } else {
      pointsEarned = rubricPointsEarned;
    }
  }

  pointsEarned = Math.max(0, Math.min(pointsEarned, pointsPossible));

  return {
    pointsEarned: roundTo2(pointsEarned),
    pointsPossible: roundTo2(pointsPossible),
    letterGrade: String(parsed.letterGrade || "N/A"),
    summary: String(parsed.summary || "No summary returned."),
    strengths: Array.isArray(parsed.strengths) ? parsed.strengths.map(String) : [],
    improvements: Array.isArray(parsed.improvements) ? parsed.improvements.map(String) : [],
    rubricBreakdown: rubric
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
        extractTextFromFile(submissionFile),
        otherFile ? extractTextFromFile(otherFile) : Promise.resolve("")
      ]);

      const rubric = normalizeText(rubricText);
      const instructions = normalizeText(instructionsText);
      const submission = normalizeText(submissionText);
      const otherText = normalizeText(otherTextRaw);

      if (!rubric || !instructions || !submission) {
        return res.status(400).json({
          error: "One or more uploaded files had no readable text. Please verify the document contents."
        });
      }

      const otherContext = otherFile
        ? `Reason this file is relevant: ${otherRelevance}\n\n[Other Supporting File]\n${otherText || "No readable text found in other file."}`
        : "";

      const messages = buildMessages({ rubric, instructions, submission, selectedPrompt: finalPrompt, otherContext });
      const raw = await generateGradeOutput(messages);
      const parsed = parseModelJson(raw);
      const normalized = normalizePointsResult(parsed);

      res.json({
        result: normalized
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
