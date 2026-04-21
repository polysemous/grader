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
const envProvider = String(process.env.LLM_PROVIDER || "openai").toLowerCase();
const envOpenaiModel = process.env.OPENAI_MODEL || "gpt-4.1-mini";
const envGeminiModel = process.env.GEMINI_MODEL || "gemini-2.5-flash";
const envOpenrouterModel = process.env.OPENROUTER_MODEL || "meta-llama/llama-3.1-8b-instruct:free";
const dataDir = path.join(__dirname, "data");
const customPromptsPath = path.join(dataDir, "custom-prompts.json");
const reportsPath = path.join(dataDir, "reports.json");
const profilesPath = path.join(dataDir, "profiles.json");
const llmSettingsPath = path.join(dataDir, "llm-settings.json");
const DEFAULT_PROFILE_ID = "default";
const OPENROUTER_MODELS_URL = "https://openrouter.ai/api/v1/models";
const BASE_SUPPORTED_MODELS = [
  { provider: "gemini", model: "gemini-2.5-flash", label: "Gemini 2.5 Flash" },
];
const OPENROUTER_GRADING_INCLUDE_TOKENS = [
  "gemini",
  "qwen",
  "llama",
  "mistral",
  "deepseek",
  "command",
  "claude",
  "gpt"
];
const OPENROUTER_GRADING_EXCLUDE_TOKENS = [
  "vision",
  "image",
  "audio",
  "speech",
  "tts",
  "whisper",
  "embedding",
  "embed",
  "rerank",
  "search",
  "guard",
  "moderation"
];
const MIN_GRADING_CONTEXT_LENGTH = 32000;
const OPENROUTER_MODEL_ALLOWLIST_TOKENS = ["aurora-alpha", "aurora alpha"];
const DEFAULT_MODEL_CONFIG = (() => {
  if (envProvider === "openrouter") {
    return { provider: "openrouter", model: String(envOpenrouterModel || "").trim() || "meta-llama/llama-3.1-8b-instruct:free" };
  }
  if (envProvider === "openai") {
    return { provider: "openai", model: String(envOpenaiModel || "").trim() || "gpt-4.1-mini" };
  }
  return { provider: "gemini", model: String(envGeminiModel || "").trim() || "gemini-2.5-flash" };
})();

if (envProvider === "openai" && !process.env.OPENAI_API_KEY) {
  console.warn("OPENAI_API_KEY is not set. Requests to /api/grade will fail until it is configured.");
}
if (envProvider === "gemini" && !process.env.GEMINI_API_KEY) {
  console.warn("GEMINI_API_KEY is not set. Requests to /api/grade will fail until it is configured.");
}
if (envProvider === "openrouter" && !process.env.OPENROUTER_API_KEY) {
  console.warn("OPENROUTER_API_KEY is not set. Requests to /api/grade will fail until it is configured.");
}

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 12 * 1024 * 1024
  }
});

const PROMPT_PROFILES = {
  graduate_professor:
    "You are a rigorous graduate-level professor. Evaluate with high standards for depth, originality, precision, and scholarly quality.",
  professor_from_hell:
    "You are an extremely strict professor with very high standards. Reward only precise, well-supported, high-quality work and provide direct, specific correction guidance for every flaw.",
  strict_examiner:
    "You are a strict examiner. Prioritize adherence to rubric criteria and assignment requirements over stylistic generosity.",
  grammar_spelling_apa7:
    "Focus only on grammar, spelling, and APA 7 writing/formatting quality. Return actionable feedback with exact examples and corrections. Do not assign a grade or score."
};

app.use(express.static("public"));
app.use(express.json());

app.get("/api/profiles", async (_req, res) => {
  try {
    const profiles = await readProfiles();
    res.json({
      profiles: profiles.map((item) => ({
        id: String(item.id || ""),
        name: String(item.name || "Unnamed Profile")
      }))
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to load student profiles." });
  }
});

app.post("/api/profiles", async (req, res) => {
  try {
    const name = normalizeProfileName(req.body?.name);
    if (!name) {
      return res.status(400).json({ error: "Profile name is required." });
    }

    const profiles = await readProfiles();
    const existingByName = profiles.find((item) => normalizeProfileName(item.name).toLowerCase() === name.toLowerCase());
    if (existingByName) {
      return res.json({
        profile: {
          id: String(existingByName.id || ""),
          name: String(existingByName.name || "")
        },
        created: false
      });
    }

    let baseId = sanitizeProfileId(name);
    let candidate = baseId;
    let i = 2;
    while (profiles.some((item) => String(item.id || "") === candidate)) {
      candidate = `${baseId}-${i}`;
      i += 1;
    }

    const profile = { id: candidate, name };
    await writeProfiles([...profiles, profile]);
    res.status(201).json({ profile, created: true });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to create student profile." });
  }
});

app.patch("/api/profiles/:id", async (req, res) => {
  try {
    const requestedId = sanitizeProfileId(req.params?.id);
    const name = normalizeProfileName(req.body?.name);
    if (!name) {
      return res.status(400).json({ error: "Profile name is required." });
    }

    const profiles = await readProfiles();
    const index = profiles.findIndex((item) => String(item.id || "") === requestedId);
    if (index < 0) {
      return res.status(404).json({ error: "Student profile not found." });
    }

    const duplicate = profiles.find(
      (item, i) => i !== index && normalizeProfileName(item.name).toLowerCase() === name.toLowerCase()
    );
    if (duplicate) {
      return res.status(400).json({ error: "A profile with that name already exists." });
    }

    const updated = {
      ...profiles[index],
      name
    };
    profiles[index] = updated;
    await writeProfiles(profiles);
    res.json({ profile: { id: String(updated.id || ""), name: String(updated.name || "") } });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to rename student profile." });
  }
});

app.get("/api/settings/llm", async (_req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(_req);
    const settings = await readLlmSettings();
    const profileSettings = getLlmSettingsForProfile(settings, profileId);
    const supportedModels = await getSupportedModels();
    const hasApiKey = Boolean(getActiveProviderApiKey(profileSettings));
    const openaiHasApiKey = Boolean(String(settings.openaiApiKey || "").trim());
    const geminiHasApiKey = Boolean(String(settings.geminiApiKey || "").trim());
    const openrouterHasApiKey = Boolean(String(settings.openrouterApiKey || "").trim());
    res.json({
      settings: {
        modelId: buildModelId(profileSettings.provider, profileSettings.model),
        provider: profileSettings.provider,
        model: profileSettings.model,
        hasApiKey
      },
      supportedModels: supportedModels.map((item) => ({
        id: buildModelId(item.provider, item.model),
        provider: item.provider,
        model: item.model,
        label: item.label,
        hasApiKey:
          item.provider === "openai"
            ? openaiHasApiKey
            : item.provider === "gemini"
              ? geminiHasApiKey
              : openrouterHasApiKey
      }))
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to load LLM settings." });
  }
});

app.post("/api/settings/llm", async (req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(req);
    const modelId = String(req.body?.modelId || "").trim();
    const apiKey = String(req.body?.apiKey || "").trim();
    const selected = await getSupportedModelById(modelId);
    if (!selected) {
      return res.status(400).json({ error: "Selected model is not supported." });
    }

    const current = await readLlmSettings();
    const next = {
      ...current
    };
    next.profileModelSelections = {
      ...(next.profileModelSelections && typeof next.profileModelSelections === "object" ? next.profileModelSelections : {}),
      [profileId]: {
        provider: selected.provider,
        model: selected.model
      }
    };
    if (apiKey) {
      if (selected.provider === "openai") {
        next.openaiApiKey = apiKey;
      } else if (selected.provider === "gemini") {
        next.geminiApiKey = apiKey;
      } else if (selected.provider === "openrouter") {
        next.openrouterApiKey = apiKey;
      }
    }
    const saved = await writeLlmSettings(next);
    const savedForProfile = getLlmSettingsForProfile(saved, profileId);
    const hasApiKey = Boolean(getActiveProviderApiKey(savedForProfile));
    res.json({
      settings: {
        modelId: buildModelId(savedForProfile.provider, savedForProfile.model),
        provider: savedForProfile.provider,
        model: savedForProfile.model,
        hasApiKey
      }
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to save LLM settings." });
  }
});

app.post("/api/settings/llm/test", async (req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(req);
    const modelId = String(req.body?.modelId || "").trim();
    const apiKeyInput = String(req.body?.apiKey || "").trim();
    if (!modelId) {
      return res.status(400).json({ error: "Select a model before running the test." });
    }

    const selected = await getSupportedModelById(modelId);
    if (!selected) {
      return res.status(400).json({ error: "Selected model is not supported." });
    }

    const current = await readLlmSettings();
    const currentForProfile = getLlmSettingsForProfile(current, profileId);
    const provider = String(selected.provider || "");
    const model = String(selected.model || "");
    const providerSavedKey =
      provider === "openai"
        ? String(currentForProfile.openaiApiKey || "")
        : provider === "gemini"
          ? String(currentForProfile.geminiApiKey || "")
          : String(currentForProfile.openrouterApiKey || "");
    const apiKey = apiKeyInput || providerSavedKey.trim();
    if (!apiKey) {
      const providerLabel = provider === "openai" ? "OpenAI" : provider === "openrouter" ? "OpenRouter" : "Gemini";
      return res.status(400).json({ error: `No ${providerLabel} API key is available for this model.` });
    }

    const messages = [
      {
        role: "system",
        content: "You are a model availability test. Return a minimal JSON object."
      },
      {
        role: "user",
        content: 'Return: {"ok": true}'
      }
    ];

    try {
      const raw = await generateOutputForSettings(messages, {
        provider,
        model,
        openaiApiKey: provider === "openai" ? apiKey : "",
        geminiApiKey: provider === "gemini" ? apiKey : "",
        openrouterApiKey: provider === "openrouter" ? apiKey : ""
      });
      if (!String(raw || "").trim()) {
        throw new Error("Model returned an empty response.");
      }
      res.json({
        usable: true,
        status: "ok",
        provider,
        model,
        modelId: buildModelId(provider, model),
        checkedAt: new Date().toISOString(),
        message: "Model responded successfully and appears usable right now."
      });
    } catch (error) {
      const message = String(error?.message || "Model test failed.");
      res.json({
        usable: false,
        status: classifyModelTestError(message),
        provider,
        model,
        modelId: buildModelId(provider, model),
        checkedAt: new Date().toISOString(),
        message
      });
    }
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to test selected model." });
  }
});

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

function sanitizeProfileId(value) {
  const normalized = String(value || "")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9_-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "");
  return normalized || DEFAULT_PROFILE_ID;
}

function normalizeProfileName(value) {
  return String(value || "").trim().replace(/\s+/g, " ");
}

function resolveRequestedProfileId(req) {
  const headerValue = req.get("x-profile-id");
  const bodyValue = req.body?.profileId;
  const queryValue = req.query?.profileId;
  return sanitizeProfileId(headerValue || bodyValue || queryValue || DEFAULT_PROFILE_ID);
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

function normalizeGradingOption(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "grammar_only") {
    return "grammar_only";
  }
  return "standard";
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

function buildModelId(provider, model) {
  return `${String(provider || "").trim()}:${String(model || "").trim()}`;
}

function parsePricingNumber(value) {
  const num = Number.parseFloat(String(value ?? "").trim());
  return Number.isFinite(num) ? num : null;
}

function isOpenRouterModelFree(item) {
  const id = String(item?.id || "").trim();
  if (!id) {
    return false;
  }
  if (id.endsWith(":free")) {
    return true;
  }
  const pricing = item?.pricing && typeof item.pricing === "object" ? item.pricing : {};
  const keys = Object.keys(pricing);
  if (keys.length === 0) {
    return false;
  }
  const numeric = keys
    .map((key) => parsePricingNumber(pricing[key]))
    .filter((value) => value !== null);
  if (numeric.length === 0) {
    return false;
  }
  return numeric.every((value) => value <= 0);
}

function isOpenRouterModelGoodForGrading(item) {
  const id = String(item?.id || "").toLowerCase();
  const name = String(item?.name || "").toLowerCase();
  const merged = `${id} ${name}`;
  if (!merged.trim()) {
    return false;
  }
  const inAllowlist = OPENROUTER_MODEL_ALLOWLIST_TOKENS.some((token) => merged.includes(token));
  if (inAllowlist) {
    return true;
  }

  const hasExcludedToken = OPENROUTER_GRADING_EXCLUDE_TOKENS.some((token) => merged.includes(token));
  if (hasExcludedToken) {
    return false;
  }

  const hasIncludedToken = OPENROUTER_GRADING_INCLUDE_TOKENS.some((token) => merged.includes(token));
  if (!hasIncludedToken) {
    return false;
  }

  const contextLength = Number(item?.context_length);
  if (Number.isFinite(contextLength) && contextLength > 0) {
    return contextLength >= MIN_GRADING_CONTEXT_LENGTH;
  }
  return true;
}

async function fetchOpenRouterFreeModels() {
  try {
    const response = await fetch(OPENROUTER_MODELS_URL, {
      method: "GET",
      headers: {
        "Content-Type": "application/json"
      }
    });
    if (!response.ok) {
      return [];
    }
    const data = await response.json().catch(() => ({}));
    const models = Array.isArray(data?.data) ? data.data : [];
    const freeModels = models.filter(isOpenRouterModelFree);
    const gradingModels = freeModels.filter(isOpenRouterModelGoodForGrading);
    const source = gradingModels.length > 0 ? gradingModels : freeModels;
    return source
      .map((item) => {
        const model = String(item.id || "").trim();
        const name = String(item.name || model).trim();
        return {
          provider: "openrouter",
          model,
          label: `OpenRouter ${name}`
        };
      })
      .filter((item) => Boolean(item.model))
      .sort((a, b) => a.label.localeCompare(b.label));
  } catch {
    return [];
  }
}

async function getSupportedModels() {
  const openrouterModels = await fetchOpenRouterFreeModels();
  return [...BASE_SUPPORTED_MODELS, ...openrouterModels];
}

async function getSupportedModelById(modelId) {
  const supported = await getSupportedModels();
  return supported.find((item) => buildModelId(item.provider, item.model) === modelId) || null;
}

function normalizeLlmSettings(input) {
  const normalized = input && typeof input === "object" ? input : {};
  const candidate = normalizeModelSelection({
    provider: normalized.provider || DEFAULT_MODEL_CONFIG.provider || "gemini",
    model: normalized.model || ""
  });
  const rawSelections =
    normalized.profileModelSelections && typeof normalized.profileModelSelections === "object"
      ? normalized.profileModelSelections
      : {};
  const profileModelSelections = Object.entries(rawSelections).reduce((acc, [key, value]) => {
    const profileId = sanitizeProfileId(key);
    if (!profileId) {
      return acc;
    }
    acc[profileId] = normalizeModelSelection(value);
    return acc;
  }, {});
  return {
    provider: candidate.provider,
    model: candidate.model,
    profileModelSelections,
    openaiApiKey: String(normalized.openaiApiKey || process.env.OPENAI_API_KEY || "").trim(),
    geminiApiKey: String(normalized.geminiApiKey || process.env.GEMINI_API_KEY || "").trim(),
    openrouterApiKey: String(normalized.openrouterApiKey || process.env.OPENROUTER_API_KEY || "").trim()
  };
}

function normalizeModelSelection(input) {
  const value = input && typeof input === "object" ? input : {};
  const provider = String(value.provider || DEFAULT_MODEL_CONFIG.provider || "gemini").toLowerCase();
  const model = String(value.model || "").trim();
  const isKnownProvider = provider === "openai" || provider === "gemini" || provider === "openrouter";
  const baseCandidate = BASE_SUPPORTED_MODELS.find((item) => item.provider === provider && item.model === model) || null;
  const candidate =
    baseCandidate ||
    (isKnownProvider && model ? { provider, model } : DEFAULT_MODEL_CONFIG);
  return {
    provider: candidate.provider,
    model: candidate.model
  };
}

function getLlmSettingsForProfile(settings, profileId) {
  const normalizedSettings = normalizeLlmSettings(settings);
  const id = sanitizeProfileId(profileId || DEFAULT_PROFILE_ID);
  const selection =
    normalizedSettings.profileModelSelections && normalizedSettings.profileModelSelections[id]
      ? normalizeModelSelection(normalizedSettings.profileModelSelections[id])
      : normalizeModelSelection({
          provider: normalizedSettings.provider,
          model: normalizedSettings.model
        });
  return {
    ...normalizedSettings,
    provider: selection.provider,
    model: selection.model
  };
}

async function ensureLlmSettingsStore() {
  await fs.mkdir(dataDir, { recursive: true });
  try {
    await fs.access(llmSettingsPath);
  } catch {
    const initial = normalizeLlmSettings({});
    await fs.writeFile(llmSettingsPath, JSON.stringify(initial, null, 2), "utf8");
  }
}

async function readLlmSettings() {
  await ensureLlmSettingsStore();
  const raw = await fs.readFile(llmSettingsPath, "utf8");
  const parsed = JSON.parse(raw);
  return normalizeLlmSettings(parsed);
}

async function writeLlmSettings(settings) {
  await ensureLlmSettingsStore();
  const next = normalizeLlmSettings(settings);
  await fs.writeFile(llmSettingsPath, JSON.stringify(next, null, 2), "utf8");
  return next;
}

function getActiveProviderApiKey(settings) {
  if (settings.provider === "openai") {
    return String(settings.openaiApiKey || "");
  }
  if (settings.provider === "gemini") {
    return String(settings.geminiApiKey || "");
  }
  if (settings.provider === "openrouter") {
    return String(settings.openrouterApiKey || "");
  }
  return "";
}

async function ensureCustomPromptStore() {
  await fs.mkdir(dataDir, { recursive: true });
  try {
    await fs.access(customPromptsPath);
  } catch {
    await fs.writeFile(customPromptsPath, "[]", "utf8");
  }
}

async function ensureReportsStore() {
  await fs.mkdir(dataDir, { recursive: true });
  try {
    await fs.access(reportsPath);
  } catch {
    await fs.writeFile(reportsPath, "[]", "utf8");
  }
}

async function ensureProfilesStore() {
  await fs.mkdir(dataDir, { recursive: true });
  try {
    await fs.access(profilesPath);
  } catch {
    await fs.writeFile(
      profilesPath,
      JSON.stringify([{ id: DEFAULT_PROFILE_ID, name: "KC" }], null, 2),
      "utf8"
    );
  }
}

async function readProfiles() {
  await ensureProfilesStore();
  const raw = await fs.readFile(profilesPath, "utf8");
  const parsed = JSON.parse(raw);
  const list = Array.isArray(parsed) ? parsed : [];
  if (!list.some((item) => String(item.id || "") === DEFAULT_PROFILE_ID)) {
    return [{ id: DEFAULT_PROFILE_ID, name: "KC" }, ...list];
  }
  return list;
}

async function writeProfiles(profiles) {
  await ensureProfilesStore();
  const next = Array.isArray(profiles) ? profiles : [];
  if (!next.some((item) => String(item.id || "") === DEFAULT_PROFILE_ID)) {
    next.unshift({ id: DEFAULT_PROFILE_ID, name: "KC" });
  }
  await fs.writeFile(profilesPath, JSON.stringify(next, null, 2), "utf8");
}

async function resolveProfileIdOrThrow(req) {
  const profileId = resolveRequestedProfileId(req);
  const profiles = await readProfiles();
  const exists = profiles.some((item) => String(item.id || "") === profileId);
  if (!exists) {
    return DEFAULT_PROFILE_ID;
  }
  return profileId;
}

async function readCustomPrompts(profileId = DEFAULT_PROFILE_ID) {
  await ensureCustomPromptStore();
  const raw = await fs.readFile(customPromptsPath, "utf8");
  const parsed = JSON.parse(raw);
  const all = Array.isArray(parsed) ? parsed : [];
  const scoped = all.filter((item) => sanitizeProfileId(item?.profileId || DEFAULT_PROFILE_ID) === profileId);
  if (!Array.isArray(scoped)) {
    return [];
  }
  return scoped;
}

async function writeCustomPrompts(profileId = DEFAULT_PROFILE_ID, prompts) {
  await ensureCustomPromptStore();
  const raw = await fs.readFile(customPromptsPath, "utf8");
  const parsed = JSON.parse(raw);
  const all = Array.isArray(parsed) ? parsed : [];
  const others = all.filter((item) => sanitizeProfileId(item?.profileId || DEFAULT_PROFILE_ID) !== profileId);
  const scoped = (Array.isArray(prompts) ? prompts : []).map((item) => ({
    ...item,
    profileId
  }));
  await fs.writeFile(customPromptsPath, JSON.stringify([...others, ...scoped], null, 2), "utf8");
}

function buildProfessorPromptKey(professorName) {
  return normalizeText(professorName)
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function buildProfessorPromptName(professorName) {
  return `Professor Calibrated - ${normalizeText(professorName)}`;
}

function getLegacyProfessorPromptKey(item) {
  const name = String(item?.name || "");
  const prefix = "Professor Calibrated - ";
  if (!name.startsWith(prefix)) {
    return "";
  }
  return buildProfessorPromptKey(name.slice(prefix.length));
}

async function findProfessorCustomPrompt(professorName, profileId = DEFAULT_PROFILE_ID) {
  const prompts = await readCustomPrompts(profileId);
  const professorKey = buildProfessorPromptKey(professorName);
  const targetName = buildProfessorPromptName(professorName);
  return (
    prompts.find(
      (item) =>
        String(item.promptType || "") === "professor_calibrated" &&
        String(item.professorKey || "") === professorKey
    ) ||
    prompts.find((item) => getLegacyProfessorPromptKey(item) === professorKey) ||
    prompts.find((item) => String(item.name || "") === targetName) ||
    null
  );
}

async function readReports(profileId = DEFAULT_PROFILE_ID) {
  await ensureReportsStore();
  const raw = await fs.readFile(reportsPath, "utf8");
  const parsed = JSON.parse(raw);
  const all = Array.isArray(parsed) ? parsed : [];
  return all.filter((item) => sanitizeProfileId(item?.profileId || DEFAULT_PROFILE_ID) === profileId);
}

async function writeReports(profileId = DEFAULT_PROFILE_ID, reports) {
  await ensureReportsStore();
  const raw = await fs.readFile(reportsPath, "utf8");
  const parsed = JSON.parse(raw);
  const all = Array.isArray(parsed) ? parsed : [];
  const others = all.filter((item) => sanitizeProfileId(item?.profileId || DEFAULT_PROFILE_ID) !== profileId);
  const scoped = (Array.isArray(reports) ? reports : []).map((item) => ({
    ...item,
    profileId
  }));
  await fs.writeFile(reportsPath, JSON.stringify([...others, ...scoped], null, 2), "utf8");
}

function normalizeGradingContext(value) {
  const context = value && typeof value === "object" ? value : {};
  return {
    rubric: normalizeText(context.rubric),
    instructions: normalizeText(context.instructions),
    promptInstructions: normalizeText(context.promptInstructions),
    otherContext: normalizeText(context.otherContext),
    submissionType: normalizeSubmissionType(context.submissionType),
    gradingOption: normalizeGradingOption(context.gradingOption)
  };
}

function normalizeActualOutcome(value) {
  const input = value && typeof value === "object" ? value : {};
  return {
    trueScore: normalizeText(input.trueScore),
    trueFeedback: normalizeText(input.trueFeedback),
    recordedAt: normalizeText(input.recordedAt) || new Date().toISOString()
  };
}

function normalizePromptCalibration(value) {
  const input = value && typeof value === "object" ? value : {};
  const confidenceRaw = String(input.confidence || "").toLowerCase();
  const confidence = ["low", "medium", "high"].includes(confidenceRaw) ? confidenceRaw : "low";
  return {
    sampleCount: Math.max(1, Number(input.sampleCount) || 1),
    recommendedPromptName: normalizeText(input.recommendedPromptName) || "Calibrated Professor Prompt",
    suggestedPromptText: normalizeText(input.suggestedPromptText),
    improvementNotes: Array.isArray(input.improvementNotes)
      ? input.improvementNotes.map((item) => normalizeText(item)).filter(Boolean)
      : [],
    confidence,
    generatedAt: normalizeText(input.generatedAt) || new Date().toISOString()
  };
}

function buildCalibrationMessages({ professorName, currentPrompt, samples }) {
  const system = [
    "You are an expert prompt engineer calibrating an academic grading prompt.",
    "Use the historical mismatch between AI grading and professor grading feedback to improve the prompt.",
    "Return valid JSON only using this exact schema:",
    "{",
    '  "recommendedPromptName": string,',
    '  "suggestedPromptText": string,',
    '  "improvementNotes": string[],',
    '  "confidence": "low" | "medium" | "high"',
    "}",
    "Focus on preserving useful rigor while matching this professor's likely expectations."
  ].join(" ");

  const user = [
    `Professor: ${professorName}`,
    "",
    "[Current Prompt]",
    currentPrompt || "None",
    "",
    "[Historical Calibration Samples]",
    JSON.stringify(samples, null, 2),
    "",
    "Rules:",
    "1) suggestedPromptText must be complete and ready to use.",
    "2) improvementNotes must be concise and actionable.",
    "3) If sample count is less than 3, confidence should be low."
  ].join("\n");

  return [
    { role: "system", content: system },
    { role: "user", content: user }
  ];
}

function buildPromptFromFeedbackMessages({ professorName, originalPrompt, buildInstructions, selectedFeedback }) {
  const system = [
    "You are an expert prompt engineer for academic assignment grading.",
    "You will revise an existing grading prompt to better align with a professor's actual feedback patterns.",
    "Return a concise, high-level evaluation prompt only.",
    "Do NOT include rubric tables, point scales, category lists, scoring procedures, or feedback examples.",
    "Do NOT restate assignment-specific constraints that the app already enforces programmatically.",
    "Target length: 80-220 words.",
    "Return valid JSON only using this exact schema:",
    "{",
    '  "recommendedPromptName": string,',
    '  "promptText": string,',
    '  "notes": string[]',
    "}",
    "promptText must be complete and directly usable as an evaluation prompt."
  ].join(" ");

  const user = [
    `Professor: ${professorName}`,
    "",
    "[Original Prompt]",
    originalPrompt,
    "",
    "[Build Instructions]",
    buildInstructions,
    "",
    "[Selected Historical Professor Feedback]",
    JSON.stringify(selectedFeedback, null, 2),
    "",
    "Rules:",
    "1) Keep rubric-based grading and evidence requirements intact.",
    "2) Incorporate recurring professor preferences from selected feedback.",
    "3) Avoid overfitting to one sample; synthesize patterns.",
    "4) Keep the rewritten prompt high-level and concise.",
    "5) Do not output any rubric/category/points breakdown.",
    "6) Return JSON only with the schema above."
  ].join("\n");

  return [
    { role: "system", content: system },
    { role: "user", content: user }
  ];
}

function isPromptBuilderOutputTooSpecific(promptText) {
  const text = normalizeText(promptText);
  if (!text) {
    return true;
  }
  if (text.length > 1400) {
    return true;
  }
  const lower = text.toLowerCase();
  const blockedPatterns = [
    "rubric",
    "scoring procedure",
    "feedback example",
    "out of",
    "points possible",
    "category",
    "criterion"
  ];
  const matches = blockedPatterns.filter((pattern) => lower.includes(pattern));
  return matches.length >= 2;
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

function buildGrammarOnlyMessages({ submission, selectedPrompt, submissionWordCount }) {
  const system = [
    "You are an expert writing reviewer.",
    "Review ONLY grammar, spelling, and APA 7 writing/formatting quality.",
    "Do not grade rubric criteria and do not assign scores, points, or letter grades.",
    "Return valid JSON only with this exact schema:",
    "{",
    '  "deductions": [{',
    '    "criterion": string,',
    '    "pointsLost": number,',
    '    "reason": string,',
    '    "actionableFix": string,',
    '    "evidence": [{ "location": string, "snippet": string, "issue": string, "suggestion": string }]',
    "  }],",
    '  "citationsNeeded": []',
    "}",
    "For each issue, include exact evidence snippets and concrete corrections."
  ].join(" ");

  const user = [
    `Evaluation mode: ${selectedPrompt}`,
    "",
    "[Student Submission - Original]",
    submission,
    "",
    `[Authoritative Submission Word Count] ${submissionWordCount}`,
    "",
    "Rules:",
    "1) Focus strictly on grammar, spelling, and APA 7 quality.",
    "2) Do not include numeric grading, score, points, or letter-grade judgments.",
    "3) Return only actionable issue findings with evidence and fixes.",
    "4) Do not use line-number labels like L1/L2."
  ].join("\n");

  return [
    { role: "system", content: system },
    { role: "user", content: user }
  ];
}

app.get("/api/custom-prompts", async (_req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(_req);
    const prompts = await readCustomPrompts(profileId);
    res.json({
      prompts: prompts.map((item) => ({
        id: String(item.id || ""),
        name: String(item.name || "Unnamed saved prompt"),
        text: String(item.text || ""),
        promptType: String(item.promptType || "custom"),
        professorName: String(item.professorName || "")
      }))
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to load saved prompts." });
  }
});

app.get("/api/prompt-builder/professors", async (req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(req);
    const reports = await readReports(profileId);
    const counts = reports.reduce((acc, report) => {
      const professor = normalizeText(report?.professorName);
      const trueFeedback = normalizeText(report?.actualOutcome?.trueFeedback);
      if (!professor || !trueFeedback) {
        return acc;
      }
      acc[professor] = (acc[professor] || 0) + 1;
      return acc;
    }, {});
    const professors = Object.entries(counts)
      .map(([name, feedbackCount]) => ({ name, feedbackCount }))
      .sort((a, b) => a.name.localeCompare(b.name));
    res.json({ professors });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to load professors with true feedback." });
  }
});

app.get("/api/prompt-builder/feedback", async (req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(req);
    const professorName = normalizeText(req.query?.professorName);
    if (!professorName) {
      return res.status(400).json({ error: "Professor name is required." });
    }
    const reports = await readReports(profileId);
    const feedbackRows = reports
      .filter(
        (report) =>
          normalizeText(report?.professorName).toLowerCase() === professorName.toLowerCase() &&
          normalizeText(report?.actualOutcome?.trueFeedback)
      )
      .map((report) => ({
        reportId: String(report?.id || ""),
        reportName: String(report?.name || ""),
        className: String(report?.className || ""),
        runDate: String(report?.runDate || ""),
        trueScore: String(report?.actualOutcome?.trueScore || ""),
        trueFeedback: String(report?.actualOutcome?.trueFeedback || ""),
        projectedScore: `${report?.result?.pointsEarned ?? "N/A"}/${report?.result?.pointsPossible ?? "N/A"}`
      }))
      .sort((a, b) => new Date(b.runDate).getTime() - new Date(a.runDate).getTime());

    res.json({ feedbackRows });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to load true feedback rows." });
  }
});

app.post("/api/prompt-builder/build", async (req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(req);
    const professorName = normalizeText(req.body?.professorName);
    const originalPromptId = normalizeText(req.body?.originalPromptId);
    const originalPromptTextFallback = normalizeText(req.body?.originalPromptText);
    const buildInstructions = normalizeText(req.body?.buildInstructions);
    const selectedReportIdsRaw = Array.isArray(req.body?.selectedReportIds) ? req.body.selectedReportIds : [];
    const selectedReportIds = selectedReportIdsRaw.map((value) => normalizeText(value)).filter(Boolean);

    if (!professorName) {
      return res.status(400).json({ error: "Select a professor first." });
    }
    if (!originalPromptId) {
      return res.status(400).json({ error: "Select the original saved prompt." });
    }
    if (!buildInstructions) {
      return res.status(400).json({ error: "Build prompt instructions are required." });
    }
    if (selectedReportIds.length === 0) {
      return res.status(400).json({ error: "Select at least one feedback item to incorporate." });
    }

    const resolveOriginalPrompt = async () => {
      if (originalPromptId.startsWith("system:")) {
        const key = originalPromptId.slice("system:".length);
        const text = normalizeText(PROMPT_PROFILES[key]);
        return text ? { text, source: "system", label: key } : null;
      }
      if (originalPromptId.startsWith("saved:")) {
        const savedId = originalPromptId.slice("saved:".length);
        const prompts = await readCustomPrompts(profileId);
        const found = prompts.find((item) => String(item?.id || "") === savedId);
        if (!found) {
          return null;
        }
        return { text: normalizeText(found.text), source: "saved", label: String(found.name || savedId) };
      }
      if (PROMPT_PROFILES[originalPromptId]) {
        return { text: normalizeText(PROMPT_PROFILES[originalPromptId]), source: "system", label: originalPromptId };
      }
      const prompts = await readCustomPrompts(profileId);
      const found = prompts.find((item) => String(item?.id || "") === originalPromptId);
      if (!found) {
        return null;
      }
      return { text: normalizeText(found.text), source: "saved", label: String(found.name || originalPromptId) };
    };

    let originalPrompt = await resolveOriginalPrompt();
    if ((!originalPrompt || !originalPrompt.text) && originalPromptTextFallback) {
      originalPrompt = {
        text: originalPromptTextFallback,
        source: "client_fallback",
        label: "provided_text"
      };
    }
    if (!originalPrompt?.text) {
      return res.status(404).json({ error: "Original prompt not found." });
    }

    const reports = await readReports(profileId);
    const selectedFeedback = reports
      .filter((report) => selectedReportIds.includes(String(report?.id || "")))
      .filter(
        (report) =>
          normalizeText(report?.professorName).toLowerCase() === professorName.toLowerCase() &&
          normalizeText(report?.actualOutcome?.trueFeedback)
      )
      .map((report) => ({
        reportName: String(report?.name || ""),
        className: String(report?.className || ""),
        runDate: String(report?.runDate || ""),
        trueScore: String(report?.actualOutcome?.trueScore || ""),
        trueFeedback: String(report?.actualOutcome?.trueFeedback || ""),
        projectedScore: `${report?.result?.pointsEarned ?? "N/A"}/${report?.result?.pointsPossible ?? "N/A"}`
      }));

    if (selectedFeedback.length === 0) {
      return res.status(400).json({ error: "Selected feedback could not be matched to this professor." });
    }

    const llmSettings = getLlmSettingsForProfile(await readLlmSettings(), profileId);
    const llmApiKey = getActiveProviderApiKey(llmSettings);
    const modelUsed = String(llmSettings.model || "gemini-2.5-flash");
    if (llmSettings.provider === "openai" && !llmApiKey) {
      return res.status(500).json({ error: "OpenAI API key is not configured in Settings." });
    }
    if (llmSettings.provider === "gemini" && !llmApiKey) {
      return res.status(500).json({ error: "Gemini API key is not configured in Settings." });
    }
    if (llmSettings.provider === "openrouter" && !llmApiKey) {
      return res.status(500).json({ error: "OpenRouter API key is not configured in Settings." });
    }

    const messages = buildPromptFromFeedbackMessages({
      professorName,
      originalPrompt: String(originalPrompt.text || ""),
      buildInstructions,
      selectedFeedback
    });
    const raw = await generateOutputForSettings(messages, llmSettings);
    let parsed = parseModelJson(raw);
    let recommendedPromptName = normalizeText(parsed?.recommendedPromptName) || `${professorName} Refined Prompt`;
    let promptText = normalizeText(parsed?.promptText);
    if (isPromptBuilderOutputTooSpecific(promptText)) {
      const simplifyMessages = [
        ...messages,
        {
          role: "user",
          content:
            "Your prior prompt was too detailed. Regenerate a simpler high-level prompt (80-220 words). " +
            "Do not include any rubric sections, category lists, scoring steps, or point values."
        }
      ];
      const simplifiedRaw = await generateOutputForSettings(simplifyMessages, llmSettings);
      const simplifiedParsed = parseModelJson(simplifiedRaw);
      const simplifiedPrompt = normalizeText(simplifiedParsed?.promptText);
      if (simplifiedPrompt) {
        parsed = simplifiedParsed;
        recommendedPromptName = normalizeText(parsed?.recommendedPromptName) || recommendedPromptName;
        promptText = simplifiedPrompt;
      }
    }
    const notes = Array.isArray(parsed?.notes) ? parsed.notes.map((item) => normalizeText(item)).filter(Boolean) : [];

    if (!promptText) {
      return res.status(500).json({ error: "Prompt build failed: model returned empty prompt text." });
    }

    res.json({
      generated: {
        recommendedPromptName,
        promptText,
        notes
      },
      modelUsed
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({
      error: error.message || "Failed to build prompt from feedback."
    });
  }
});

app.post("/api/custom-prompts", async (req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(req);
    const name = normalizeText(req.body?.name);
    const text = normalizeText(req.body?.text);

    if (!name || !text) {
      return res.status(400).json({ error: "Prompt name and prompt text are required." });
    }

    const prompts = await readCustomPrompts(profileId);
    const id = `saved_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`;
    const next = [...prompts, { id, name, text, promptType: "custom" }];
    await writeCustomPrompts(profileId, next);

    res.status(201).json({ prompt: { id, name, text, promptType: "custom" } });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to save custom prompt." });
  }
});

app.delete("/api/custom-prompts/:id", async (req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(req);
    const promptId = normalizeText(req.params.id);
    if (!promptId) {
      return res.status(400).json({ error: "Prompt id is required." });
    }

    const prompts = await readCustomPrompts(profileId);
    const existing = prompts.find((item) => String(item.id) === promptId);
    if (!existing) {
      return res.status(404).json({ error: "Saved prompt not found." });
    }

    const next = prompts.filter((item) => String(item.id) !== promptId);
    await writeCustomPrompts(profileId, next);

    res.json({ success: true });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to delete saved prompt." });
  }
});

app.put("/api/custom-prompts/:id", async (req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(req);
    const promptId = normalizeText(req.params.id);
    const text = normalizeText(req.body?.text);
    const name = normalizeText(req.body?.name);
    if (!promptId) {
      return res.status(400).json({ error: "Prompt id is required." });
    }
    if (!text) {
      return res.status(400).json({ error: "Prompt text is required." });
    }

    const prompts = await readCustomPrompts(profileId);
    const index = prompts.findIndex((item) => String(item.id || "") === promptId);
    if (index < 0) {
      return res.status(404).json({ error: "Saved prompt not found." });
    }

    const existing = prompts[index];
    prompts[index] = {
      ...existing,
      name: name || String(existing.name || "Unnamed saved prompt"),
      text
    };
    await writeCustomPrompts(profileId, prompts);

    res.json({
      prompt: {
        id: String(prompts[index].id || ""),
        name: String(prompts[index].name || ""),
        text: String(prompts[index].text || ""),
        promptType: String(prompts[index].promptType || "custom"),
        professorName: String(prompts[index].professorName || "")
      }
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to update saved prompt." });
  }
});

app.get("/api/reports", async (_req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(_req);
    const reports = await readReports(profileId);
    const summaries = reports.map((report) => ({
      id: String(report.id || ""),
      name: String(report.name || "Untitled report"),
      className: String(report.className || ""),
      professorName: String(report.professorName || ""),
      university: String(report.university || ""),
      runDate: String(report.runDate || ""),
      canRegrade:
        String(report.gradingContext?.gradingOption || "standard") === "grammar_only"
          ? Boolean(report.gradingContext?.promptInstructions)
          : Boolean(report.gradingContext?.rubric && report.gradingContext?.instructions && report.gradingContext?.promptInstructions),
      hasActualOutcome: Boolean(report.actualOutcome?.trueScore || report.actualOutcome?.trueFeedback),
      modelUsed: String(report?.result?.modelUsed || report?.modelUsed || "gemini-2.5-flash"),
      projectedScore: `${report?.result?.pointsEarned ?? "N/A"}/${report?.result?.pointsPossible ?? "N/A"}`,
      trueScore: String(report?.actualOutcome?.trueScore || "")
    }));
    res.json({ reports: summaries });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to load saved reports." });
  }
});

app.get("/api/reports/:id", async (req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(req);
    const reportId = normalizeText(req.params.id);
    const reports = await readReports(profileId);
    const found = reports.find((item) => String(item.id) === reportId);
    if (!found) {
      return res.status(404).json({ error: "Report not found." });
    }
    res.json({ report: found });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to load report." });
  }
});

app.post("/api/reports", async (req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(req);
    const name = normalizeText(req.body?.name);
    const className = normalizeText(req.body?.className);
    const professorName = normalizeText(req.body?.professorName);
    const university = normalizeText(req.body?.university);
    const runDate = normalizeText(req.body?.runDate) || new Date().toISOString();
    const result = req.body?.result;
    const gradingContext = normalizeGradingContext(req.body?.gradingContext);

    if (!name || !className || !professorName || !university) {
      return res.status(400).json({
        error: "Report name, class name, professor name, and university are required."
      });
    }
    if (!result || typeof result !== "object") {
      return res.status(400).json({ error: "A grading result is required to save a report." });
    }
    const isGrammarOnly = String(gradingContext.gradingOption || "standard") === "grammar_only";
    if (!gradingContext.promptInstructions || (!isGrammarOnly && (!gradingContext.rubric || !gradingContext.instructions))) {
      return res.status(400).json({
        error:
          "This report is missing required grading context. Please grade again and then save so rubric, instructions, and prompt are stored for regrading."
      });
    }

    const reports = await readReports(profileId);
    const id = `report_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`;
    const entry = {
      id,
      name,
      className,
      professorName,
      university,
      runDate,
      result,
      gradingContext
    };
    const next = [entry, ...reports];
    await writeReports(profileId, next);

    res.status(201).json({
      report: {
        id,
        name,
        className,
        professorName,
        university,
        runDate
      }
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to save report." });
  }
});

app.delete("/api/reports/:id", async (req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(req);
    const reportId = normalizeText(req.params.id);
    if (!reportId) {
      return res.status(400).json({ error: "Report id is required." });
    }

    const reports = await readReports(profileId);
    const found = reports.find((item) => String(item.id) === reportId);
    if (!found) {
      return res.status(404).json({ error: "Report not found." });
    }

    const next = reports.filter((item) => String(item.id) !== reportId);
    await writeReports(profileId, next);
    res.json({ success: true });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to delete report." });
  }
});

app.post(
  "/api/reports/:id/regrade",
  upload.fields([{ name: "submission", maxCount: 1 }]),
  async (req, res) => {
    try {
      const profileId = await resolveProfileIdOrThrow(req);
      const llmSettings = getLlmSettingsForProfile(await readLlmSettings(), profileId);
      const llmApiKey = getActiveProviderApiKey(llmSettings);
      const modelUsed = String(llmSettings.model || "gemini-2.5-flash");
      if (llmSettings.provider === "openai" && !llmApiKey) {
        return res.status(500).json({ error: "OpenAI API key is not configured in Settings." });
      }
      if (llmSettings.provider === "gemini" && !llmApiKey) {
        return res.status(500).json({ error: "Gemini API key is not configured in Settings." });
      }
      if (llmSettings.provider === "openrouter" && !llmApiKey) {
        return res.status(500).json({ error: "OpenRouter API key is not configured in Settings." });
      }

      const reportId = normalizeText(req.params.id);
      const submissionFile = req.files?.submission?.[0];
      if (!reportId) {
        return res.status(400).json({ error: "Report id is required." });
      }
      if (!submissionFile) {
        return res.status(400).json({ error: "Please upload the updated student submission file." });
      }

      const reports = await readReports(profileId);
      const index = reports.findIndex((item) => String(item.id) === reportId);
      if (index < 0) {
        return res.status(404).json({ error: "Report not found." });
      }

      const existing = reports[index];
      const gradingContext = normalizeGradingContext(existing?.gradingContext);
      const isGrammarOnly = String(gradingContext.gradingOption || "standard") === "grammar_only";
      if (!gradingContext.promptInstructions || (!isGrammarOnly && (!gradingContext.rubric || !gradingContext.instructions))) {
        return res.status(400).json({
          error:
            "This saved report does not include rubric/instructions context. Create a new saved report from a fresh grade first."
        });
      }

      const submissionText = await extractTextFromFile(submissionFile, { preserveDocxFormatting: true });
      const submission = normalizeText(submissionText);
      if (!submission) {
        return res.status(400).json({
          error: "The updated submission had no readable text. Please verify the document contents."
        });
      }

      const submissionWordCount = calculateWordCount(submission);
      const result = await gradeSubmissionFromInputs({
        rubric: gradingContext.rubric,
        instructions: gradingContext.instructions,
        submission,
        selectedPrompt: gradingContext.promptInstructions,
        otherContext: gradingContext.otherContext || "",
        submissionType: gradingContext.submissionType,
        gradingOption: gradingContext.gradingOption,
        submissionWordCount,
        modelUsed,
        llmSettings
      });

      const runDate = new Date().toISOString();
      const updated = {
        ...existing,
        runDate,
        result,
        gradingContext
      };
      reports[index] = updated;
      await writeReports(profileId, reports);

      res.json({
        report: {
          id: String(updated.id || ""),
          name: String(updated.name || "Untitled report"),
          className: String(updated.className || ""),
          professorName: String(updated.professorName || ""),
          university: String(updated.university || ""),
          runDate
        },
        result
      });
    } catch (error) {
      console.error(error);
      res.status(500).json({
        error: error.message || "Failed to regrade report."
      });
    }
  }
);

app.post("/api/reports/:id/actual-outcome", async (req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(req);
    const reportId = normalizeText(req.params.id);
    const trueScore = normalizeText(req.body?.trueScore);
    const trueFeedback = normalizeText(req.body?.trueFeedback);
    const clearing = !trueScore && !trueFeedback;

    if (!reportId) {
      return res.status(400).json({ error: "Report id is required." });
    }
    if (!clearing && (!trueScore || !trueFeedback)) {
      return res.status(400).json({
        error: "Provide both true score and professor feedback, or clear both fields to remove true outcome."
      });
    }

    const reports = await readReports(profileId);
    const index = reports.findIndex((item) => String(item.id) === reportId);
    if (index < 0) {
      return res.status(404).json({ error: "Report not found." });
    }

    const existing = reports[index];
    const actualOutcome = clearing
      ? normalizeActualOutcome({ trueScore: "", trueFeedback: "" })
      : normalizeActualOutcome({ trueScore, trueFeedback });
    const updated = {
      ...existing,
      actualOutcome
    };
    reports[index] = updated;
    await writeReports(profileId, reports);

    res.json({ report: updated, cleared: clearing });
  } catch (error) {
    console.error(error);
    res.status(500).json({
      error: error.message || "Failed to save true outcome."
    });
  }
});

app.post("/api/professor-prompts/upsert", async (req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(req);
    const professorName = normalizeText(req.body?.professorName);
    const promptText = normalizeText(req.body?.promptText);

    if (!professorName || !promptText) {
      return res.status(400).json({ error: "Professor name and prompt text are required." });
    }

    const prompts = await readCustomPrompts(profileId);
    const professorKey = buildProfessorPromptKey(professorName);
    const promptName = buildProfessorPromptName(professorName);
    const existingIndex = prompts.findIndex(
      (item) =>
        (String(item.promptType || "") === "professor_calibrated" &&
          String(item.professorKey || "") === professorKey) ||
        getLegacyProfessorPromptKey(item) === professorKey ||
        String(item.name || "") === promptName
    );

    let promptId;
    if (existingIndex >= 0) {
      promptId = String(prompts[existingIndex].id || "");
      prompts[existingIndex] = {
        ...prompts[existingIndex],
        name: promptName,
        text: promptText,
        promptType: "professor_calibrated",
        professorName,
        professorKey
      };
    } else {
      promptId = `saved_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`;
      prompts.push({
        id: promptId,
        name: promptName,
        text: promptText,
        promptType: "professor_calibrated",
        professorName,
        professorKey
      });
    }

    const savedPrompt = prompts.find((item) => String(item.id || "") === promptId);
    const deduped = prompts.filter((item) => {
      const isProfessorPromptForSameProfessor =
        (String(item.promptType || "") === "professor_calibrated" && String(item.professorKey || "") === professorKey) ||
        getLegacyProfessorPromptKey(item) === professorKey;
      if (!isProfessorPromptForSameProfessor) {
        return true;
      }
      return String(item.id || "") === promptId;
    });
    if (savedPrompt && !deduped.some((item) => String(item.id || "") === promptId)) {
      deduped.push(savedPrompt);
    }

    await writeCustomPrompts(profileId, deduped);
    res.json({
      prompt: {
        id: promptId,
        name: promptName,
        text: promptText,
        promptType: "professor_calibrated",
        professorName
      },
      overwrote: existingIndex >= 0
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to save professor prompt." });
  }
});

app.get("/api/professor-prompts/current", async (req, res) => {
  try {
    const profileId = await resolveProfileIdOrThrow(req);
    const professorName = normalizeText(req.query?.professorName);
    if (!professorName) {
      return res.status(400).json({ error: "Professor name is required." });
    }

    const found = await findProfessorCustomPrompt(professorName, profileId);
    if (!found) {
      return res.json({ exists: false, prompt: null });
    }

    res.json({
      exists: true,
      prompt: {
        id: String(found.id || ""),
        name: String(found.name || ""),
        text: String(found.text || "")
      }
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to load professor prompt." });
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

function normalizeGrammarOnlyResult(parsed, submission) {
  const hasItalicMarkers = String(submission || "").includes("[ITALIC]");
  let deductions = Array.isArray(parsed?.deductions)
    ? parsed.deductions.map((item) => ({
        criterion: String(item.criterion || "Writing Quality"),
        pointsLost: 0,
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
  if (hasItalicMarkers) {
    deductions = deductions.filter((item) => !isItalicIssueText(`${item.criterion} ${item.reason}`));
  }
  deductions = deductions.filter((item) =>
    Array.isArray(item.evidence) && item.evidence.some((evidence) => String(evidence.snippet || "").trim())
  );
  return {
    grammarOnlyMode: true,
    pointsEarned: "N/A",
    pointsPossible: "N/A",
    letterGrade: "N/A",
    rubricBreakdown: [],
    deductions,
    citationsNeeded: []
  };
}

async function generateWithOpenAI(messages, model, apiKey) {
  if (!apiKey) {
    throw new Error("Server is missing OPENAI_API_KEY.");
  }
  const client = new OpenAI({ apiKey });

  const completion = await client.chat.completions.create({
    model,
    response_format: { type: "json_object" },
    messages,
    temperature: 0.2
  });

  return completion.choices?.[0]?.message?.content || "{}";
}

async function generateWithGemini(messages, model, apiKey) {
  if (!apiKey) {
    throw new Error("Server is missing GEMINI_API_KEY.");
  }

  const systemText = messages.find((message) => message.role === "system")?.content || "";
  const userText = messages
    .filter((message) => message.role === "user")
    .map((message) => message.content)
    .join("\n\n");

  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-goog-api-key": apiKey
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
    throw new Error(formatGeminiApiError(detail, model, response.status));
  }

  const text = data?.candidates?.[0]?.content?.parts?.map((part) => part.text || "").join("\n").trim();
  if (!text) {
    throw new Error("Gemini returned an empty response.");
  }

  return text;
}

async function generateWithOpenRouter(messages, model, apiKey) {
  if (!apiKey) {
    throw new Error("Server is missing OPENROUTER_API_KEY.");
  }

  const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
      "HTTP-Referer": process.env.APP_BASE_URL || `http://localhost:${port}`,
      "X-Title": "Assignment Grader"
    },
    body: JSON.stringify({
      model,
      messages,
      response_format: { type: "json_object" },
      temperature: 0.2
    })
  });

  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    const detail = {
      message: String(data?.error?.message || data?.message || "").trim(),
      raw:
        String(
          data?.error?.metadata?.raw ||
            data?.error?.metadata?.reason ||
            data?.error?.details ||
            data?.details ||
            ""
        ).trim(),
      provider: String(data?.error?.metadata?.provider_name || data?.provider || "").trim()
    };
    throw new Error(formatOpenRouterApiError(detail, model, response.status));
  }

  const text = data?.choices?.[0]?.message?.content;
  if (!text) {
    throw new Error("OpenRouter returned an empty response.");
  }
  return text;
}

function formatGeminiApiError(detail, model, statusCode) {
  const message = String(detail || "").trim();
  const lower = message.toLowerCase();
  const modelLabel = String(model || "selected model");

  if (lower.includes("not found") || lower.includes("not supported for generatecontent")) {
    return (
      `The selected Gemini model (${modelLabel}) is not available for grading with this API endpoint. ` +
      "Open Settings and choose the available model (Gemini 2.5 Flash), then try again."
    );
  }

  if (lower.includes("quota") || lower.includes("rate") || lower.includes("resource exhausted")) {
    return (
      "Gemini API quota/rate limit reached. Wait briefly and retry, or switch to another model/provider in Settings."
    );
  }

  if (statusCode === 401 || lower.includes("api key") || lower.includes("permission")) {
    return "Gemini authentication failed. Verify your Gemini API key in Settings and try again.";
  }

  return `Gemini request failed for model ${modelLabel}. ${message || `HTTP ${statusCode}`}`;
}

function formatOpenRouterApiError(detail, model, statusCode) {
  const message = String(detail?.message || detail || "").trim();
  const raw = String(detail?.raw || "").trim();
  const provider = String(detail?.provider || "").trim();
  const lower = message.toLowerCase();
  const rawLower = raw.toLowerCase();
  const modelLabel = String(model || "selected model");
  const providerLabel = provider ? ` (provider: ${provider})` : "";

  if (
    lower.includes("model") && (lower.includes("not found") || lower.includes("not available")) ||
    rawLower.includes("model") && (rawLower.includes("not found") || rawLower.includes("not available"))
  ) {
    return (
      `The selected OpenRouter model (${modelLabel}) is currently unavailable${providerLabel}. ` +
      "Open Settings and choose a different OpenRouter free model, then try again."
    );
  }
  if (
    lower.includes("quota") ||
    lower.includes("rate limit") ||
    lower.includes("credits") ||
    rawLower.includes("quota") ||
    rawLower.includes("rate limit") ||
    rawLower.includes("credits")
  ) {
    return `OpenRouter quota/rate limit reached${providerLabel}. Try again shortly or switch models/providers in Settings.`;
  }
  if (
    statusCode === 401 ||
    lower.includes("api key") ||
    lower.includes("unauthorized") ||
    rawLower.includes("api key") ||
    rawLower.includes("unauthorized")
  ) {
    return `OpenRouter authentication failed${providerLabel}. Verify your OpenRouter API key in Settings and try again.`;
  }
  if (lower.includes("provider returned error")) {
    if (raw) {
      return `OpenRouter provider error for model ${modelLabel}${providerLabel}: ${raw}`;
    }
    return (
      `OpenRouter provider error for model ${modelLabel}${providerLabel}. ` +
      "The upstream provider failed the request. Try again, or switch to another model in Settings."
    );
  }
  const combined = [message, raw].filter(Boolean).join(" | ");
  return `OpenRouter request failed for model ${modelLabel}${providerLabel}. ${combined || `HTTP ${statusCode}`}`;
}

function classifyModelTestError(message) {
  const lower = String(message || "").toLowerCase();
  if (lower.includes("quota") || lower.includes("rate limit") || lower.includes("rate-limited")) {
    return "rate_limited";
  }
  if (lower.includes("authentication") || lower.includes("api key") || lower.includes("unauthorized")) {
    return "auth_error";
  }
  if (
    (lower.includes("model") && lower.includes("not")) ||
    lower.includes("not supported") ||
    lower.includes("unavailable")
  ) {
    return "model_unavailable";
  }
  return "api_error";
}

async function generateOutputForSettings(messages, settings) {
  const provider = String(settings?.provider || "").trim();
  const model = String(settings?.model || "").trim();
  const apiKey = getActiveProviderApiKey(settings).trim();
  if (provider === "openai") {
    return generateWithOpenAI(messages, model, apiKey);
  }
  if (provider === "gemini") {
    return generateWithGemini(messages, model, apiKey);
  }
  if (provider === "openrouter") {
    return generateWithOpenRouter(messages, model, apiKey);
  }
  throw new Error(`Invalid LLM provider '${provider}'. Use 'openai', 'gemini', or 'openrouter'.`);
}

async function generateGradeOutput(messages, settings = null) {
  const effectiveSettings = settings ? normalizeLlmSettings(settings) : await readLlmSettings();
  return generateOutputForSettings(messages, effectiveSettings);
}

async function gradeSubmissionFromInputs({
  rubric,
  instructions,
  submission,
  selectedPrompt,
  otherContext,
  submissionType,
  gradingOption,
  submissionWordCount,
  modelUsed,
  llmSettings
}) {
  const mode = normalizeGradingOption(gradingOption);
  const messages =
    mode === "grammar_only"
      ? buildGrammarOnlyMessages({
          submission,
          selectedPrompt,
          submissionWordCount
        })
      : buildMessages({
          rubric,
          instructions,
          submission,
          selectedPrompt,
          otherContext,
          submissionType,
          submissionWordCount
        });
  let raw = await generateGradeOutput(messages, llmSettings);
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
    raw = await generateGradeOutput(retryMessages, llmSettings);
    parsed = parseModelJson(raw);
  }

  return {
    ...(mode === "grammar_only"
      ? normalizeGrammarOnlyResult(parsed, submission)
      : normalizePointsResult(parsed, submissionType, submission)),
    submissionWordCount,
    gradingOption: mode,
    modelUsed: String(modelUsed || "gemini-2.5-flash")
  };
}

async function suggestPromptCalibration({ professorName, currentPrompt, reports }) {
  const normalizedProfessor = normalizeText(professorName);
  const relevant = reports
    .filter(
      (item) =>
        normalizeText(item.professorName).toLowerCase() === normalizedProfessor.toLowerCase() &&
        item?.actualOutcome?.trueFeedback
    )
    .slice(0, 25);

  const samples = relevant.map((item) => ({
    reportName: String(item.name || ""),
    className: String(item.className || ""),
    runDate: String(item.runDate || ""),
    aiScore: `${item?.result?.pointsEarned ?? "N/A"}/${item?.result?.pointsPossible ?? "N/A"}`,
    trueScore: String(item?.actualOutcome?.trueScore || ""),
    trueFeedback: String(item?.actualOutcome?.trueFeedback || ""),
    aiDeductions: Array.isArray(item?.result?.deductions)
      ? item.result.deductions.slice(0, 4).map((d) => ({
          criterion: String(d.criterion || ""),
          reason: String(d.reason || "")
        }))
      : []
  }));

  const safeCurrentPrompt =
    normalizeText(currentPrompt) ||
    "You are a rigorous professor. Grade using the rubric and assignment instructions with strict evidence-based deductions.";

  const fallback = normalizePromptCalibration({
    sampleCount: Math.max(1, samples.length),
    recommendedPromptName: `${normalizedProfessor || "Professor"} Calibrated Prompt`,
    suggestedPromptText: safeCurrentPrompt,
    improvementNotes: [
      "Add stronger weighting for patterns repeatedly mentioned in professor feedback.",
      "Use this as a baseline and refine after more true-grade samples are recorded."
    ],
    confidence: samples.length >= 6 ? "medium" : "low"
  });

  try {
    const messages = buildCalibrationMessages({
      professorName: normalizedProfessor || "Unknown Professor",
      currentPrompt: safeCurrentPrompt,
      samples
    });
    const raw = await generateGradeOutput(messages);
    const parsed = parseModelJson(raw);
    const suggestion = normalizePromptCalibration({
      sampleCount: Math.max(1, samples.length),
      recommendedPromptName: parsed.recommendedPromptName,
      suggestedPromptText: parsed.suggestedPromptText,
      improvementNotes: parsed.improvementNotes,
      confidence: parsed.confidence
    });
    if (!suggestion.suggestedPromptText) {
      return fallback;
    }
    return suggestion;
  } catch {
    return fallback;
  }
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
      const profileId = await resolveProfileIdOrThrow(req);
      const llmSettings = getLlmSettingsForProfile(await readLlmSettings(), profileId);
      const llmApiKey = getActiveProviderApiKey(llmSettings);
      const modelUsed = String(llmSettings.model || "gemini-2.5-flash");
      if (llmSettings.provider === "openai" && !llmApiKey) {
        return res.status(500).json({ error: "OpenAI API key is not configured in Settings." });
      }
      if (llmSettings.provider === "gemini" && !llmApiKey) {
        return res.status(500).json({ error: "Gemini API key is not configured in Settings." });
      }
      if (llmSettings.provider === "openrouter" && !llmApiKey) {
        return res.status(500).json({ error: "OpenRouter API key is not configured in Settings." });
      }

      const rubricFile = req.files?.rubric?.[0];
      const instructionsFile = req.files?.instructions?.[0];
      const submissionFile = req.files?.submission?.[0];
      const otherFile = req.files?.other?.[0];

      const profile = String(req.body.promptProfile || "");
      const promptInstructions = normalizeText(req.body.promptInstructions);
      const submissionType = normalizeSubmissionType(req.body.submissionType);
      const gradingOption = normalizeGradingOption(req.body.gradingOption);
      const otherRelevance = normalizeText(req.body.otherRelevance);

      if (!submissionFile) {
        return res.status(400).json({ error: "Please upload student submission." });
      }
      if (gradingOption !== "grammar_only" && (!rubricFile || !instructionsFile)) {
        return res.status(400).json({
          error: "Please upload rubric and assignment instructions, or use Spelling and Grammar Only mode."
        });
      }

      let selectedPrompt;
      if (profile === "custom") {
        selectedPrompt = promptInstructions;
      } else if (profile.startsWith("saved:")) {
        const savedId = profile.slice("saved:".length);
        const prompts = await readCustomPrompts(profileId);
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

      const forcedPrompt =
        gradingOption === "grammar_only" ? String(PROMPT_PROFILES.grammar_spelling_apa7 || selectedPrompt || "") : "";
      const finalPrompt = forcedPrompt || promptInstructions || selectedPrompt;
      if (!finalPrompt) {
        return res.status(400).json({ error: "Prompt Instructions is required." });
      }

      if (otherFile && !otherRelevance) {
        return res.status(400).json({ error: "Please explain why the 'Other' file is relevant." });
      }

      const [rubricText, instructionsText, submissionText, otherTextRaw] = await Promise.all([
        rubricFile ? extractTextFromFile(rubricFile) : Promise.resolve(""),
        instructionsFile ? extractTextFromFile(instructionsFile) : Promise.resolve(""),
        extractTextFromFile(submissionFile, { preserveDocxFormatting: true }),
        otherFile ? extractTextFromFile(otherFile) : Promise.resolve("")
      ]);

      const rubric = normalizeText(rubricText);
      const instructions = normalizeText(instructionsText);
      const submission = normalizeText(submissionText);
      const submissionWordCount = calculateWordCount(submission);
      const otherText = normalizeText(otherTextRaw);

      if (!submission || (gradingOption !== "grammar_only" && (!rubric || !instructions))) {
        return res.status(400).json({
          error: "One or more uploaded files had no readable text. Please verify the document contents."
        });
      }

      const otherContext = otherFile
        ? `Reason this file is relevant: ${otherRelevance}\n\n[Other Supporting File]\n${otherText || "No readable text found in other file."}`
        : "";

      const result = await gradeSubmissionFromInputs({
        rubric,
        instructions,
        submission,
        selectedPrompt: finalPrompt,
        otherContext,
        submissionType,
        gradingOption,
        submissionWordCount,
        modelUsed,
        llmSettings
      });

      res.json({
        result,
        gradingContext: {
          rubric,
          instructions,
          promptInstructions: finalPrompt,
          otherContext,
          submissionType,
          gradingOption
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
