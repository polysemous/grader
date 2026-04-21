const form = document.getElementById("grade-form");
const studentProfileSelect = document.getElementById("student-profile-select");
const promptProfile = document.getElementById("prompt-profile");
const promptInstructions = document.getElementById("prompt-instructions");
const savePromptWrap = document.getElementById("save-prompt-wrap");
const savedPromptName = document.getElementById("saved-prompt-name");
const savePromptBtn = document.getElementById("save-prompt-btn");
const deletePromptWrap = document.getElementById("delete-prompt-wrap");
const deletePromptBtn = document.getElementById("delete-prompt-btn");
const savePromptEditsWrap = document.getElementById("save-prompt-edits-wrap");
const savePromptEditsBtn = document.getElementById("save-prompt-edits-btn");
const otherFile = document.getElementById("other-file");
const otherRelevanceWrap = document.getElementById("other-relevance-wrap");
const otherRelevance = document.getElementById("other-relevance");
const submissionFileInput = document.getElementById("submission-file");
const rubricFileInput = document.getElementById("rubric-file");
const instructionsFileInput = document.getElementById("instructions-file");
const submissionTypeDiscussion = document.getElementById("submission-type-discussion");
const submissionTypeDiscussionReply = document.getElementById("submission-type-discussion-reply");
const gradingOptionGrammarOnly = document.getElementById("grading-option-grammar-only");

const submitBtn = document.getElementById("submit-btn");
const statusEl = document.getElementById("status");

const resultsSection = document.getElementById("results");
const scorePill = document.getElementById("score-pill");
const resultModelUsed = document.getElementById("result-model-used");
const wordCountDisplay = document.getElementById("word-count-display");
const deductionsWrap = document.getElementById("deductions-wrap");
const rubricWrap = document.getElementById("rubric-table-wrap");

const reportNameInput = document.getElementById("report-name");
const classNameInput = document.getElementById("class-name");
const professorNameInput = document.getElementById("professor-name");
const universityNameInput = document.getElementById("university-name");
const classNameOptions = document.getElementById("class-name-options");
const professorNameOptions = document.getElementById("professor-name-options");
const universityNameOptions = document.getElementById("university-name-options");
const saveReportBtn = document.getElementById("save-report-btn");

const reportSortHeaders = document.querySelectorAll(".sort-header-btn[data-sort-key]");
const reportsTableBody = document.getElementById("reports-table-body");
const tabGraderBtn = document.getElementById("tab-grader");
const tabReportsBtn = document.getElementById("tab-reports");
const tabSettingsBtn = document.getElementById("tab-settings");
const panelGrader = document.getElementById("panel-grader");
const panelReports = document.getElementById("panel-reports");
const panelSettings = document.getElementById("panel-settings");
const llmModelSelect = document.getElementById("llm-model-select");
const llmApiKeyInput = document.getElementById("llm-api-key-input");
const llmKeyHint = document.getElementById("llm-key-hint");
const testLlmModelBtn = document.getElementById("test-llm-model-btn");
const settingsTestStatus = document.getElementById("settings-test-status");
const saveLlmSettingsBtn = document.getElementById("save-llm-settings-btn");
const settingsStatus = document.getElementById("settings-status");
const activeLlmModel = document.getElementById("active-llm-model");
const graderCurrentModel = document.getElementById("grader-current-model");
const graderTimeoutSecondsInput = document.getElementById("grader-timeout-seconds-input");
const saveGraderTimeoutBtn = document.getElementById("save-grader-timeout-btn");
const graderTimeoutStatus = document.getElementById("grader-timeout-status");

const viewerSection = document.getElementById("report-viewer");
const viewerTitle = document.getElementById("viewer-title");
const viewerMeta = document.getElementById("viewer-meta");
const outcomeComparison = document.getElementById("outcome-comparison");
const trueScoreDisplay = document.getElementById("true-score-display");
const trueFeedbackDisplay = document.getElementById("true-feedback-display");
const predictedScoreDisplay = document.getElementById("predicted-score-display");
const predictedFeedbackDisplay = document.getElementById("predicted-feedback-display");
const viewerScore = document.getElementById("viewer-score");
const viewerDeductions = document.getElementById("viewer-deductions");
const viewerRubric = document.getElementById("viewer-rubric");
const closeViewerBtn = document.getElementById("close-viewer-btn");
const regradeDialog = document.getElementById("regrade-dialog");
const closeRegradeDialogBtn = document.getElementById("close-regrade-dialog-btn");
const actualOutcomeDialog = document.getElementById("actual-outcome-dialog");
const closeActualOutcomeDialogBtn = document.getElementById("close-actual-outcome-dialog-btn");
const regradeSubmissionFile = document.getElementById("regrade-submission-file");
const regradeBtn = document.getElementById("regrade-btn");
const actualScoreInput = document.getElementById("actual-score-input");
const actualFeedbackInput = document.getElementById("actual-feedback-input");
const saveActualBtn = document.getElementById("save-actual-btn");
const promptBuilderDialog = document.getElementById("prompt-builder-dialog");
const closePromptBuilderDialogBtn = document.getElementById("close-prompt-builder-dialog-btn");
const promptBuilderModel = document.getElementById("prompt-builder-model");
const promptBuilderProfessorSelect = document.getElementById("prompt-builder-professor-select");
const promptBuilderFeedbackBody = document.getElementById("prompt-builder-feedback-body");
const promptBuilderSelectAllFeedback = document.getElementById("prompt-builder-select-all-feedback");
const promptBuilderOriginalPromptSelect = document.getElementById("prompt-builder-original-prompt-select");
const promptBuilderOriginalPromptText = document.getElementById("prompt-builder-original-prompt-text");
const promptBuilderInstructions = document.getElementById("prompt-builder-instructions");
const promptBuilderBuildBtn = document.getElementById("prompt-builder-build-btn");
const promptBuilderStatus = document.getElementById("prompt-builder-status");
const promptBuilderOutput = document.getElementById("prompt-builder-output");
const promptBuilderSaveName = document.getElementById("prompt-builder-save-name");
const promptBuilderSaveBtn = document.getElementById("prompt-builder-save-btn");

const PROMPT_PREVIEWS = {
  graduate_professor:
    "You are a rigorous graduate-level professor. Evaluate with high standards for depth, originality, precision, and scholarly quality.",
  professor_from_hell:
    "You are an extremely strict professor with very high standards. Reward only precise, well-supported, high-quality work and provide direct, specific correction guidance for every flaw.",
  strict_examiner:
    "You are a strict examiner. Prioritize adherence to rubric criteria and assignment requirements over stylistic generosity.",
  grammar_spelling_apa7:
    "Focus only on grammar, spelling, and APA 7 writing/formatting quality. Return actionable feedback with exact examples and corrections. Do not assign a grade or score."
};
const SYSTEM_PROMPT_LABELS = {
  graduate_professor: "Graduate Professor",
  professor_from_hell: "Professor from Hell",
  strict_examiner: "Strict Examiner",
  grammar_spelling_apa7: "Grammar/Spelling/APA 7"
};

let savedPromptMap = {};
let savedPromptNameMap = {};
let savedPromptTypeMap = {};
let latestResult = null;
let latestRunDate = null;
let latestGradingContext = null;
let reportsCache = [];
let activeReport = null;
let activeTab = "grader";
let profilesCache = [];
let activeProfileId = "default";
let activeProfileName = "KC";
let collapsedClassGroups = {};
let reportTableSortKey = "name";
let reportTableSortDirection = "asc";
let llmModelOptions = [];
let activeModelId = "";
let lastStandardPromptProfile = "graduate_professor";
let lastNonGrammarPromptProfile = "graduate_professor";
let promptBuilderOriginalPromptMap = {};
const DEFAULT_GRADE_REQUEST_TIMEOUT_SECONDS = 30;
const MIN_GRADE_REQUEST_TIMEOUT_SECONDS = 5;
const MAX_GRADE_REQUEST_TIMEOUT_SECONDS = 300;
let gradeRequestTimeoutMs = DEFAULT_GRADE_REQUEST_TIMEOUT_SECONDS * 1000;
const PROMPT_BUILD_TIMEOUT_MS = 45000;
const DEFAULT_PROMPT_BUILD_INSTRUCTIONS =
  "Revise the original grading prompt to better match this professor's feedback patterns. Keep it concise and high-level. Do not include rubric tables, point scales, scoring steps, or long category lists.";

function setStatus(message, isError = false, forceShow = false) {
  const text = String(message || "").trim();
  if (isError || forceShow) {
    statusEl.textContent = text;
    statusEl.className = isError ? "status-error" : "field-hint";
    return;
  }
  statusEl.textContent = "";
  statusEl.className = "";
}

function setSaveReportButtonSaved(isSaved) {
  saveReportBtn.textContent = isSaved ? "Saved" : "Save Feedback Report";
}

function setSaveActualButtonSaved(isSaved) {
  saveActualBtn.textContent = isSaved ? "Saved" : "Save True Outcome";
}

function setSettingsStatus(message, isError = false) {
  settingsStatus.textContent = String(message || "").trim();
  settingsStatus.className = isError ? "status-error" : "field-hint";
}

function setSettingsTestStatus(message, isError = false) {
  settingsTestStatus.textContent = String(message || "").trim();
  settingsTestStatus.className = isError ? "status-error" : "field-hint";
}

function setPromptBuilderStatus(message, isError = false) {
  promptBuilderStatus.textContent = String(message || "").trim();
  promptBuilderStatus.className = isError ? "status-error" : "field-hint";
}

function setGraderTimeoutStatus(message, isError = false) {
  graderTimeoutStatus.textContent = String(message || "").trim();
  graderTimeoutStatus.className = isError ? "status-error" : "field-hint";
}

function getGraderTimeoutStorageKey() {
  return "graderTimeoutSeconds";
}

function normalizeGraderTimeoutSeconds(value) {
  const parsed = Number.parseInt(String(value ?? "").trim(), 10);
  if (!Number.isFinite(parsed)) {
    return DEFAULT_GRADE_REQUEST_TIMEOUT_SECONDS;
  }
  return Math.max(MIN_GRADE_REQUEST_TIMEOUT_SECONDS, Math.min(MAX_GRADE_REQUEST_TIMEOUT_SECONDS, parsed));
}

function applyGraderTimeoutSeconds(seconds) {
  const normalized = normalizeGraderTimeoutSeconds(seconds);
  gradeRequestTimeoutMs = normalized * 1000;
  graderTimeoutSecondsInput.value = String(normalized);
  return normalized;
}

function loadGraderTimeoutSetting() {
  const stored = localStorage.getItem(getGraderTimeoutStorageKey());
  applyGraderTimeoutSeconds(stored || DEFAULT_GRADE_REQUEST_TIMEOUT_SECONDS);
}

function saveGraderTimeoutSetting() {
  const raw = String(graderTimeoutSecondsInput.value || "").trim();
  if (!raw) {
    setGraderTimeoutStatus(
      `Enter a timeout between ${MIN_GRADE_REQUEST_TIMEOUT_SECONDS} and ${MAX_GRADE_REQUEST_TIMEOUT_SECONDS} seconds.`,
      true
    );
    return;
  }
  const normalized = normalizeGraderTimeoutSeconds(raw);
  applyGraderTimeoutSeconds(normalized);
  localStorage.setItem(getGraderTimeoutStorageKey(), String(normalized));
  setGraderTimeoutStatus(`Grader timeout saved: ${normalized}s.`, false);
}

function getCollapsedGroupsStorageKey() {
  return `collapsedReportGroups:${activeProfileId}`;
}

function loadCollapsedGroupsState() {
  try {
    const raw = localStorage.getItem(getCollapsedGroupsStorageKey());
    const parsed = raw ? JSON.parse(raw) : {};
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
      return {};
    }
    return parsed;
  } catch {
    return {};
  }
}

function saveCollapsedGroupsState() {
  localStorage.setItem(getCollapsedGroupsStorageKey(), JSON.stringify(collapsedClassGroups));
}

function updateFileUploadDisplay(input) {
  const displayId = input?.dataset?.fileDisplay;
  if (!displayId) {
    return;
  }
  const label = document.getElementById(displayId);
  if (!label) {
    return;
  }
  const fileName = input.files?.[0]?.name;
  label.textContent = fileName || "No file selected";
}

function initializeFileUploadControls() {
  const uploadButtons = document.querySelectorAll(".file-upload-btn[data-file-input]");
  uploadButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const inputId = button.getAttribute("data-file-input");
      const input = document.getElementById(inputId);
      if (input) {
        input.click();
      }
    });
  });

  const fileInputs = document.querySelectorAll(".file-upload-input[data-file-display]");
  fileInputs.forEach((input) => {
    input.addEventListener("change", () => updateFileUploadDisplay(input));
    updateFileUploadDisplay(input);
  });
}

async function apiFetch(url, options = {}) {
  const headers = new Headers(options.headers || {});
  headers.set("X-Profile-Id", activeProfileId);
  return fetch(url, { ...options, headers });
}

async function apiFetchWithTimeout(url, options = {}, timeoutMs = 0) {
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) {
    return apiFetch(url, options);
  }
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await apiFetch(url, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(timeoutId);
  }
}

async function readApiJson(response) {
  const text = await response.text();
  let data = {};
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      const snippet = text.trim().slice(0, 160);
      const looksLikeHtml =
        snippet.startsWith("<!DOCTYPE") || snippet.startsWith("<html") || snippet.startsWith("<");
      const message = looksLikeHtml
        ? "Server returned HTML instead of API JSON. Restart the app server so the latest API routes are loaded."
        : `API returned non-JSON response: ${snippet}`;
      throw new Error(message);
    }
  }

  if (!response.ok) {
    throw new Error(data.error || `Request failed with status ${response.status}.`);
  }
  return data;
}

function renderProfileOptions() {
  const options = profilesCache
    .map(
      (profile) =>
        `<option value="${escapeHtml(profile.id)}"${profile.id === activeProfileId ? " selected" : ""}>${escapeHtml(
          profile.name
        )}</option>`
    )
    .join("");
  studentProfileSelect.innerHTML = `${options}<option value="__add_new__">+ Add New Profile...</option><option value="__rename_profile__">Rename Current Profile...</option>`;
  fitProfileSelectWidth();
}

function fitProfileSelectWidth() {
  const optionTexts = Array.from(studentProfileSelect.options).map((item) => String(item.text || "").trim());
  const longest = optionTexts.reduce((max, text) => (text.length > max.length ? text : max), "");
  const chWidth = Math.max(longest.length + 4, 12);
  studentProfileSelect.style.width = `${chWidth}ch`;
}

async function loadProfiles() {
  const response = await apiFetch("/api/profiles");
  const data = await readApiJson(response);
  profilesCache = Array.isArray(data.profiles) ? data.profiles : [];
  if (profilesCache.length === 0) {
    profilesCache = [{ id: "default", name: "KC" }];
  }

  const stored = String(localStorage.getItem("activeProfileId") || "").trim();
  const canUseStored = profilesCache.some((item) => item.id === stored);
  if (canUseStored) {
    activeProfileId = stored;
  } else if (!profilesCache.some((item) => item.id === activeProfileId)) {
    activeProfileId = profilesCache[0].id;
  }

  const active = profilesCache.find((item) => item.id === activeProfileId) || profilesCache[0];
  activeProfileId = active.id;
  activeProfileName = active.name;
  localStorage.setItem("activeProfileId", activeProfileId);

  renderProfileOptions();
}

async function reloadProfileData() {
  collapsedClassGroups = loadCollapsedGroupsState();
  viewerSection.classList.add("hidden");
  regradeDialog.classList.add("hidden");
  actualOutcomeDialog.classList.add("hidden");
  closePromptBuilderDialog();
  await loadSavedPrompts();
  await loadPromptBuilderProfessors();
  renderPromptBuilderFeedbackRows([]);
  await loadReports();
  await loadLlmSettings();
  setStatus(`Loaded profile: ${activeProfileName}`);
}

async function handleProfileSelectionChange() {
  const value = String(studentProfileSelect.value || "").trim();
  if (value === "__add_new__") {
    const name = String(window.prompt("Enter a name for the new student profile:") || "").trim();
    if (!name) {
      renderProfileOptions();
      return;
    }
    try {
      const response = await apiFetch("/api/profiles", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name })
      });
      const data = await readApiJson(response);
      const created = data.profile;
      if (!created?.id) {
        throw new Error("Profile creation failed.");
      }
      activeProfileId = created.id;
      activeProfileName = created.name || created.id;
      localStorage.setItem("activeProfileId", activeProfileId);
      await loadProfiles();
      await reloadProfileData();
    } catch (error) {
      setStatus(error.message || "Failed to create profile.", true);
      renderProfileOptions();
    }
    return;
  }

  if (value === "__rename_profile__") {
    const selected = profilesCache.find((item) => item.id === activeProfileId);
    if (!selected) {
      setStatus("No active profile selected to rename.", true);
      renderProfileOptions();
      return;
    }
    const nextName = String(window.prompt("Rename current profile:", selected.name) || "").trim();
    if (!nextName || nextName === selected.name) {
      renderProfileOptions();
      return;
    }
    try {
      const response = await apiFetch(`/api/profiles/${encodeURIComponent(selected.id)}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: nextName })
      });
      const data = await readApiJson(response);
      const updated = data.profile || {};
      activeProfileName = String(updated.name || nextName);
      await loadProfiles();
      await reloadProfileData();
      setStatus(`Renamed profile to: ${activeProfileName}`);
    } catch (error) {
      setStatus(error.message || "Failed to rename profile.", true);
      renderProfileOptions();
    }
    return;
  }

  const selected = profilesCache.find((item) => item.id === value);
  if (!selected) {
    renderProfileOptions();
    return;
  }
  activeProfileId = selected.id;
  activeProfileName = selected.name;
  localStorage.setItem("activeProfileId", activeProfileId);
  renderProfileOptions();
  await reloadProfileData();
}

function switchTab(tab) {
  if (tab === "reports" || tab === "settings" || tab === "grader") {
    activeTab = tab;
  } else {
    activeTab = "grader";
  }
  const isGrader = activeTab === "grader";
  const isReports = activeTab === "reports";
  const isSettings = activeTab === "settings";

  panelGrader.classList.toggle("hidden", !isGrader);
  panelReports.classList.toggle("hidden", !isReports);
  panelSettings.classList.toggle("hidden", !isSettings);

  tabGraderBtn.classList.toggle("active", isGrader);
  tabReportsBtn.classList.toggle("active", isReports);
  tabSettingsBtn.classList.toggle("active", isSettings);

  tabGraderBtn.setAttribute("aria-selected", isGrader ? "true" : "false");
  tabReportsBtn.setAttribute("aria-selected", isReports ? "true" : "false");
  tabSettingsBtn.setAttribute("aria-selected", isSettings ? "true" : "false");
  setStatus("", false);
  setSettingsStatus("", false);
  setSettingsTestStatus("", false);
  setGraderTimeoutStatus("", false);

  if (isSettings) {
    void loadLlmSettings();
  }
}

function applyLlmKeyHint() {
  const selected = llmModelOptions.find((item) => item.id === llmModelSelect.value);
  const provider = String(selected?.provider || "");
  if (provider === "openai") {
    llmApiKeyInput.placeholder = "sk-...";
  } else if (provider === "openrouter") {
    llmApiKeyInput.placeholder = "sk-or-v1-...";
  } else {
    llmApiKeyInput.placeholder = "AIza...";
  }
  const hasKeyForSelection = Boolean(selected?.hasApiKey);
  const providerLabel = provider === "openrouter" ? "OpenRouter" : provider === "openai" ? "OpenAI" : "Gemini";
  llmKeyHint.textContent = hasKeyForSelection
    ? `${providerLabel} API key is saved. Enter a new key only if you want to replace it.`
    : `No ${providerLabel} API key is saved yet.`;
}

async function loadLlmSettings() {
  try {
    const response = await apiFetch("/api/settings/llm");
    const data = await readApiJson(response);
    llmModelOptions = Array.isArray(data.supportedModels) ? data.supportedModels : [];
    const settings = data.settings || {};
    const modelId = String(settings.modelId || "").trim();
    const savedProvider = String(settings.provider || "").trim();
    const savedModel = String(settings.model || "").trim();

    const hasSavedModelInList = llmModelOptions.some((item) => item.id === modelId);
    const optionHtml = llmModelOptions
      .map(
        (item) =>
          `<option value="${escapeHtml(item.id)}"${item.id === modelId ? " selected" : ""}>${escapeHtml(item.label)}${
            item.hasApiKey ? " (key saved)" : ""
          }</option>`
      )
      .join("");
    const currentUnavailableOption =
      modelId && !hasSavedModelInList && savedProvider && savedModel
        ? `<option value="${escapeHtml(modelId)}" selected>${escapeHtml(
            `${savedProvider}:${savedModel} (active, not in current free list)`
          )}</option>`
        : "";
    llmModelSelect.innerHTML = `${currentUnavailableOption}${optionHtml}`;

    if (!llmModelSelect.value && llmModelOptions.length > 0) {
      llmModelSelect.value = llmModelOptions[0].id;
    }
    activeModelId = llmModelSelect.value;
    const activeModel = llmModelOptions.find((item) => item.id === activeModelId);
    if (activeModel) {
      activeLlmModel.textContent = `Active Model: ${activeModel.label}`;
      graderCurrentModel.textContent = `Current model: ${activeModel.label}`;
    } else if (modelId && savedProvider && savedModel) {
      activeLlmModel.textContent = `Active Model: ${savedProvider}:${savedModel}`;
      graderCurrentModel.textContent = `Current model: ${savedProvider}:${savedModel}`;
    } else {
      activeLlmModel.textContent = "Active Model: Not configured";
      graderCurrentModel.textContent = "Current model: Not configured";
    }
    applyLlmKeyHint();
    llmApiKeyInput.value = "";
    setSettingsStatus("", false);
    setSettingsTestStatus("", false);
  } catch (error) {
    setSettingsStatus(error.message || "Failed to load LLM settings.", true);
    setSettingsTestStatus("", false);
  }
}

async function saveLlmSettings() {
  const modelId = String(llmModelSelect.value || "").trim();
  const apiKey = String(llmApiKeyInput.value || "").trim();
  if (!modelId) {
    setSettingsStatus("Select a model before saving.", true);
    return;
  }
  saveLlmSettingsBtn.disabled = true;
  try {
    const response = await apiFetch("/api/settings/llm", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ modelId, apiKey })
    });
    await readApiJson(response);
    await loadLlmSettings();
    activeModelId = llmModelSelect.value;
    setSettingsStatus("Settings saved.", false);
  } catch (error) {
    setSettingsStatus(error.message || "Failed to save LLM settings.", true);
  } finally {
    saveLlmSettingsBtn.disabled = false;
  }
}

async function testSelectedLlmModel() {
  const modelId = String(llmModelSelect.value || "").trim();
  const apiKey = String(llmApiKeyInput.value || "").trim();
  if (!modelId) {
    setSettingsTestStatus("Select a model before running the test.", true);
    return;
  }
  testLlmModelBtn.disabled = true;
  setSettingsTestStatus("Testing selected model...", false);
  try {
    const response = await apiFetch("/api/settings/llm/test", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ modelId, apiKey })
    });
    const data = await readApiJson(response);
    if (data.usable) {
      setSettingsTestStatus(`Usable: ${data.message || "Model test passed."}`, false);
      return;
    }
    setSettingsTestStatus(`Not usable right now: ${data.message || "Model test failed."}`, true);
  } catch (error) {
    setSettingsTestStatus(error.message || "Failed to test selected model.", true);
  } finally {
    testLlmModelBtn.disabled = false;
  }
}

function closeReportViewer() {
  viewerSection.classList.add("hidden");
  closeRegradeDialog();
  closeActualOutcomeDialog();
  setStatus("", false);
}

function openRegradeDialog() {
  regradeDialog.classList.remove("hidden");
}

function closeRegradeDialog() {
  regradeDialog.classList.add("hidden");
  regradeSubmissionFile.value = "";
  updateFileUploadDisplay(regradeSubmissionFile);
  setStatus("", false);
}

function openActualOutcomeDialog() {
  actualOutcomeDialog.classList.remove("hidden");
  setSaveActualButtonSaved(false);
}

function closeActualOutcomeDialog() {
  actualOutcomeDialog.classList.add("hidden");
  setStatus("", false);
}

function getPromptBuilderModelLabel() {
  const selected = llmModelOptions.find((item) => item.id === activeModelId);
  if (selected && selected.label) {
    return selected.label;
  }
  const model = getActiveConfiguredModelId();
  return model || "Unknown";
}

function openPromptBuilderDialog() {
  promptBuilderModel.textContent = `Model used for build: ${getPromptBuilderModelLabel()}`;
  if (!String(promptBuilderInstructions.value || "").trim()) {
    promptBuilderInstructions.value = DEFAULT_PROMPT_BUILD_INSTRUCTIONS;
  }
  promptBuilderDialog.classList.remove("hidden");
}

function closePromptBuilderDialog() {
  promptBuilderDialog.classList.add("hidden");
  setPromptBuilderStatus("", false);
}

function hasActiveTextSelection() {
  const selection = window.getSelection?.();
  return Boolean(selection && String(selection).trim().length > 0);
}

function attachSafeBackdropClose(modal, closeFn) {
  let startedOnBackdrop = false;
  modal.addEventListener("mousedown", (event) => {
    startedOnBackdrop = event.target === modal;
  });
  modal.addEventListener("mouseup", (event) => {
    const endedOnBackdrop = event.target === modal;
    if (startedOnBackdrop && endedOnBackdrop && !hasActiveTextSelection()) {
      closeFn();
    }
    startedOnBackdrop = false;
  });
  modal.addEventListener("mouseleave", () => {
    startedOnBackdrop = false;
  });
}

function renderPredictedFeedbackSummary(result) {
  const deductions = Array.isArray(result?.deductions) ? result.deductions : [];
  if (deductions.length === 0) {
    predictedFeedbackDisplay.innerHTML = "<p>No deduction feedback. Predicted submission received full credit.</p>";
    return;
  }
  const items = deductions
    .map((item) => `<li><strong>${escapeHtml(item.criterion || "General")}:</strong> ${escapeHtml(item.reason || "")}</li>`)
    .join("");
  predictedFeedbackDisplay.innerHTML = `<ul>${items}</ul>`;
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function renderSavedPromptOptions(prompts) {
  const oldSavedGroup = document.getElementById("saved-prompt-group");
  if (oldSavedGroup) {
    oldSavedGroup.remove();
  }
  const oldProfessorGroup = document.getElementById("professor-prompt-group");
  if (oldProfessorGroup) {
    oldProfessorGroup.remove();
  }

  if (!Array.isArray(prompts) || prompts.length === 0) {
    return;
  }

  const professorPrompts = prompts.filter((prompt) => String(prompt.promptType || "") === "professor_calibrated");
  const customPrompts = prompts.filter((prompt) => String(prompt.promptType || "") !== "professor_calibrated");

  if (professorPrompts.length > 0) {
    const professorGroup = document.createElement("optgroup");
    professorGroup.id = "professor-prompt-group";
    professorGroup.label = "Professor Prompts";
    professorPrompts.forEach((prompt) => {
      const option = document.createElement("option");
      option.value = `saved:${prompt.id}`;
      option.textContent = `Professor: ${prompt.professorName || prompt.name}`;
      professorGroup.appendChild(option);
    });
    promptProfile.insertBefore(professorGroup, promptProfile.querySelector('option[value="custom"]'));
  }

  if (customPrompts.length === 0) {
    return;
  }

  const customGroup = document.createElement("optgroup");
  customGroup.id = "saved-prompt-group";
  customGroup.label = "Saved Custom Prompts";

  customPrompts.forEach((prompt) => {
    const option = document.createElement("option");
    option.value = `saved:${prompt.id}`;
    option.textContent = `Saved: ${prompt.name}`;
    customGroup.appendChild(option);
  });

  promptProfile.insertBefore(customGroup, promptProfile.querySelector('option[value="custom"]'));
}

function renderPromptBuilderOriginalPromptOptions() {
  const systemItems = Object.keys(PROMPT_PREVIEWS).map((key) => ({
    id: `system:${key}`,
    name: String(SYSTEM_PROMPT_LABELS[key] || key),
    text: String(PROMPT_PREVIEWS[key] || ""),
    group: "System Prompts"
  }));
  const savedItems = Object.keys(savedPromptMap).map((id) => ({
    id: `saved:${id}`,
    name: String(savedPromptNameMap[id] || "Unnamed saved prompt"),
    text: String(savedPromptMap[id] || ""),
    group: "Saved Prompts"
  }));

  promptBuilderOriginalPromptMap = {};
  [...systemItems, ...savedItems].forEach((item) => {
    promptBuilderOriginalPromptMap[item.id] = item.text;
  });

  const buildGroupOptions = (group, items) => {
    if (!items.length) {
      return "";
    }
    const optionHtml = items
      .sort((a, b) => a.name.localeCompare(b.name))
      .map((item) => `<option value="${escapeHtml(item.id)}">${escapeHtml(item.name)}</option>`)
      .join("");
    return `<optgroup label="${escapeHtml(group)}">${optionHtml}</optgroup>`;
  };

  promptBuilderOriginalPromptSelect.innerHTML = [
    `<option value="">Select original prompt...</option>`,
    buildGroupOptions(
      "System Prompts",
      systemItems
    ),
    buildGroupOptions(
      "Saved Prompts",
      savedItems
    )
  ].join("");
  updatePromptBuilderOriginalPromptText();
}

function updatePromptBuilderOriginalPromptText() {
  const selectedId = String(promptBuilderOriginalPromptSelect.value || "").trim();
  if (!selectedId) {
    promptBuilderOriginalPromptText.value = "";
    return;
  }
  promptBuilderOriginalPromptText.value = String(promptBuilderOriginalPromptMap[selectedId] || "").trim();
}

function isSavedPromptSelection() {
  return promptProfile.value.startsWith("saved:");
}

function getSelectedSavedPromptText() {
  if (!isSavedPromptSelection()) {
    return "";
  }
  const savedId = promptProfile.value.slice("saved:".length);
  return String(savedPromptMap[savedId] || "");
}

function getBasePromptForSelection() {
  if (promptProfile.value === "custom") {
    return "";
  }

  if (promptProfile.value.startsWith("saved:")) {
    const savedId = promptProfile.value.slice("saved:".length);
    return savedPromptMap[savedId] || "";
  }

  return PROMPT_PREVIEWS[promptProfile.value] || "";
}

function syncPromptEditorFromSelection() {
  promptInstructions.value = getBasePromptForSelection();
}

function updateSavePromptEditsButtonState() {
  if (!isSavedPromptSelection()) {
    savePromptEditsBtn.textContent = "Save Prompt Changes";
    savePromptEditsBtn.disabled = false;
    return;
  }
  const currentText = String(promptInstructions.value || "").trim();
  const savedText = String(getSelectedSavedPromptText() || "").trim();
  const isDirty = currentText !== savedText;
  savePromptEditsBtn.textContent = isDirty ? "Save Prompt Changes" : "Saved";
  savePromptEditsBtn.disabled = !isDirty;
}

async function loadSavedPrompts() {
  try {
    const currentValue = promptProfile.value;
    const response = await apiFetch("/api/custom-prompts");
    const data = await readApiJson(response);

    const prompts = Array.isArray(data.prompts) ? data.prompts : [];
    const filteredPrompts = prompts.filter(
      (item) => String(item?.name || "").trim().toLowerCase() !== "professor from hell"
    );
    savedPromptMap = {};
    savedPromptNameMap = {};
    savedPromptTypeMap = {};
    filteredPrompts.forEach((item) => {
      savedPromptMap[item.id] = item.text;
      savedPromptNameMap[item.id] = item.name;
      savedPromptTypeMap[item.id] = item.promptType || "custom";
    });

    renderSavedPromptOptions(filteredPrompts);
    const canKeepCurrent =
      currentValue === "custom" ||
      currentValue in PROMPT_PREVIEWS ||
      currentValue === "create_from_feedback" ||
      (currentValue.startsWith("saved:") && savedPromptMap[currentValue.slice("saved:".length)]);
    promptProfile.value = canKeepCurrent ? currentValue : "graduate_professor";
    if (promptProfile.value !== "create_from_feedback") {
      lastStandardPromptProfile = promptProfile.value;
    }
    if (promptProfile.value !== "grammar_spelling_apa7" && promptProfile.value !== "create_from_feedback") {
      lastNonGrammarPromptProfile = promptProfile.value;
    }
    renderPromptBuilderOriginalPromptOptions();
    syncPromptEditorFromSelection();
    togglePromptControls();
  } catch (error) {
    setStatus(error.message || "Failed to load saved prompts.", true);
  }
}

function togglePromptControls() {
  const showCustomTools = promptProfile.value === "custom";
  savePromptWrap.classList.toggle("hidden", !showCustomTools);
  const showSavedTools = isSavedPromptSelection();
  deletePromptWrap.classList.toggle("hidden", !showSavedTools);
  savePromptEditsWrap.classList.toggle("hidden", !showSavedTools);
  syncPromptEditorFromSelection();
  updateSavePromptEditsButtonState();
}

function handleGrammarOnlyToggle() {
  if (gradingOptionGrammarOnly.checked) {
    if (promptProfile.value !== "grammar_spelling_apa7" && promptProfile.value !== "create_from_feedback") {
      lastNonGrammarPromptProfile = promptProfile.value;
    }
    promptProfile.value = "grammar_spelling_apa7";
    handlePromptProfileChange();
    return;
  }
  if (promptProfile.value === "grammar_spelling_apa7") {
    promptProfile.value = lastNonGrammarPromptProfile || "graduate_professor";
    handlePromptProfileChange();
  }
}

function handlePromptProfileChange() {
  const value = String(promptProfile.value || "");
  if (gradingOptionGrammarOnly.checked && value !== "grammar_spelling_apa7" && value !== "create_from_feedback") {
    promptProfile.value = "grammar_spelling_apa7";
    togglePromptControls();
    return;
  }
  if (value === "create_from_feedback") {
    promptProfile.value = lastStandardPromptProfile;
    togglePromptControls();
    openPromptBuilderDialog();
    return;
  }
  if (value !== "grammar_spelling_apa7") {
    lastNonGrammarPromptProfile = value;
  }
  lastStandardPromptProfile = value;
  togglePromptControls();
}

function toggleOtherRelevance() {
  const hasOtherFile = Boolean(otherFile.files && otherFile.files.length > 0);
  otherRelevanceWrap.classList.toggle("hidden", !hasOtherFile);
  otherRelevance.required = hasOtherFile;
}

function truncateText(text, maxLen = 200) {
  const value = String(text || "").trim();
  if (value.length <= maxLen) {
    return value;
  }
  return `${value.slice(0, maxLen)}...`;
}

function renderPromptBuilderFeedbackRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) {
    promptBuilderFeedbackBody.innerHTML = "<tr><td colspan='5'>No true feedback found for this professor.</td></tr>";
    promptBuilderSelectAllFeedback.checked = false;
    promptBuilderSelectAllFeedback.indeterminate = false;
    promptBuilderSelectAllFeedback.disabled = true;
    return;
  }
  const html = rows
    .map(
      (row) => `
        <tr>
          <td><input type="checkbox" class="prompt-builder-feedback-check" value="${escapeHtml(row.reportId)}" /></td>
          <td>${escapeHtml(row.reportName || "")}</td>
          <td>${escapeHtml(row.className || "")}</td>
          <td>${escapeHtml(row.trueScore || "")}</td>
          <td title="${escapeHtml(row.trueFeedback || "")}">${escapeHtml(truncateText(row.trueFeedback, 180))}</td>
        </tr>
      `
    )
    .join("");
  promptBuilderFeedbackBody.innerHTML = html;
  promptBuilderSelectAllFeedback.checked = false;
  promptBuilderSelectAllFeedback.indeterminate = false;
  promptBuilderSelectAllFeedback.disabled = false;
}

function updatePromptBuilderSelectAllState() {
  const checks = Array.from(promptBuilderDialog.querySelectorAll(".prompt-builder-feedback-check"));
  if (checks.length === 0) {
    promptBuilderSelectAllFeedback.checked = false;
    promptBuilderSelectAllFeedback.indeterminate = false;
    promptBuilderSelectAllFeedback.disabled = true;
    return;
  }
  const checkedCount = checks.filter((item) => item.checked).length;
  promptBuilderSelectAllFeedback.disabled = false;
  if (checkedCount === 0) {
    promptBuilderSelectAllFeedback.checked = false;
    promptBuilderSelectAllFeedback.indeterminate = false;
    return;
  }
  if (checkedCount === checks.length) {
    promptBuilderSelectAllFeedback.checked = true;
    promptBuilderSelectAllFeedback.indeterminate = false;
    return;
  }
  promptBuilderSelectAllFeedback.checked = false;
  promptBuilderSelectAllFeedback.indeterminate = true;
}

async function loadPromptBuilderProfessors() {
  const response = await apiFetch("/api/prompt-builder/professors");
  const data = await readApiJson(response);
  const professors = Array.isArray(data.professors) ? data.professors : [];
  const options = professors
    .map(
      (item) =>
        `<option value="${escapeHtml(item.name)}">${escapeHtml(item.name)} (${Number(item.feedbackCount) || 0})</option>`
    )
    .join("");
  promptBuilderProfessorSelect.innerHTML = `<option value="">Select professor...</option>${options}`;
}

async function loadPromptBuilderFeedbackForProfessor() {
  const professorName = String(promptBuilderProfessorSelect.value || "").trim();
  if (!professorName) {
    renderPromptBuilderFeedbackRows([]);
    return;
  }
  const response = await apiFetch(`/api/prompt-builder/feedback?professorName=${encodeURIComponent(professorName)}`);
  const data = await readApiJson(response);
  const rows = Array.isArray(data.feedbackRows) ? data.feedbackRows : [];
  renderPromptBuilderFeedbackRows(rows);
}

function getSelectedPromptBuilderReportIds() {
  const checks = promptBuilderDialog.querySelectorAll(".prompt-builder-feedback-check:checked");
  return Array.from(checks)
    .map((item) => String(item.value || "").trim())
    .filter(Boolean);
}

async function buildPromptFromFeedback() {
  const professorName = String(promptBuilderProfessorSelect.value || "").trim();
  const originalPromptId = String(promptBuilderOriginalPromptSelect.value || "").trim();
  const originalPromptText = String(promptBuilderOriginalPromptText.value || "").trim();
  const buildInstructions = String(promptBuilderInstructions.value || "").trim();
  const selectedReportIds = getSelectedPromptBuilderReportIds();
  if (!professorName) {
    setPromptBuilderStatus("Select a professor first.", true);
    return;
  }
  if (!originalPromptId) {
    setPromptBuilderStatus("Select the original saved prompt.", true);
    return;
  }
  if (!originalPromptText) {
    setPromptBuilderStatus("The selected original prompt has no text.", true);
    return;
  }
  if (!buildInstructions) {
    setPromptBuilderStatus("Build instructions are required.", true);
    return;
  }
  if (selectedReportIds.length === 0) {
    setPromptBuilderStatus("Select at least one feedback item to incorporate.", true);
    return;
  }

  const originalLabel = String(promptBuilderBuildBtn.textContent || "Build Prompt");
  promptBuilderBuildBtn.disabled = true;
  promptBuilderBuildBtn.textContent = "Building...";
  promptBuilderSaveBtn.disabled = true;
  setPromptBuilderStatus("Sending request to model. Waiting for response...", false);
  try {
    const response = await apiFetchWithTimeout(
      "/api/prompt-builder/build",
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          professorName,
          originalPromptId,
          originalPromptText,
          buildInstructions,
          selectedReportIds
        })
      },
      PROMPT_BUILD_TIMEOUT_MS
    );
    const data = await readApiJson(response);
    const generated = data.generated || {};
    const promptText = String(generated.promptText || "").trim();
    if (!promptText) {
      throw new Error("Model did not return a prompt.");
    }
    promptBuilderOutput.value = promptText;
    promptBuilderSaveName.value = String(generated.recommendedPromptName || `${professorName} Refined Prompt`);
    promptBuilderSaveBtn.disabled = false;
    promptBuilderModel.textContent = `Model used for build: ${String(data.modelUsed || getPromptBuilderModelLabel())}`;
    setPromptBuilderStatus("Prompt generated. Review and save if you want to keep it.", false);
  } catch (error) {
    if (error?.name === "AbortError") {
      setPromptBuilderStatus(
        `Build timed out after ${Math.floor(PROMPT_BUILD_TIMEOUT_MS / 1000)}s. Try again or switch models in Settings.`,
        true
      );
    } else {
      setPromptBuilderStatus(error.message || "Failed to build prompt from feedback.", true);
    }
  } finally {
    promptBuilderBuildBtn.disabled = false;
    promptBuilderBuildBtn.textContent = originalLabel;
  }
}

async function saveGeneratedPromptFromBuilder() {
  const name = String(promptBuilderSaveName.value || "").trim();
  const text = String(promptBuilderOutput.value || "").trim();
  if (!name || !text) {
    setPromptBuilderStatus("Generated prompt and save name are required.", true);
    return;
  }
  promptBuilderSaveBtn.disabled = true;
  try {
    const response = await apiFetch("/api/custom-prompts", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name, text })
    });
    const data = await readApiJson(response);
    const newPromptId = String(data?.prompt?.id || "").trim();
    await loadSavedPrompts();
    if (newPromptId) {
      promptProfile.value = `saved:${newPromptId}`;
      lastStandardPromptProfile = promptProfile.value;
      togglePromptControls();
    }
    setPromptBuilderStatus("Generated prompt saved as a new prompt.", false);
  } catch (error) {
    setPromptBuilderStatus(error.message || "Failed to save generated prompt.", true);
  } finally {
    promptBuilderSaveBtn.disabled = false;
  }
}

function getSubmissionType() {
  const isDiscussion = Boolean(submissionTypeDiscussion.checked);
  const isDiscussionReply = Boolean(submissionTypeDiscussionReply.checked);

  if (isDiscussion && isDiscussionReply) {
    return "invalid";
  }
  if (isDiscussion) {
    return "discussion";
  }
  if (isDiscussionReply) {
    return "discussion_reply";
  }
  return "none";
}

function getGradingOption() {
  return gradingOptionGrammarOnly.checked ? "grammar_only" : "standard";
}

function extractBaseName(fileName) {
  const full = String(fileName || "").trim();
  if (!full) {
    return "Submission";
  }
  const withoutExt = full.replace(/\.[^.]+$/, "");
  return withoutExt.replace(/[._-]+/g, " ").trim() || "Submission";
}

function suggestReportName() {
  const fileName = submissionFileInput?.files?.[0]?.name || "";
  const base = extractBaseName(fileName);
  const className = String(classNameInput.value || "").trim();
  const parts = [base];
  if (className) {
    parts.push(className);
  }
  return `${parts.join(" - ")} Feedback`;
}

function renderRubricTableInto(target, breakdown) {
  if (!Array.isArray(breakdown) || breakdown.length === 0) {
    target.innerHTML = "<p>No rubric breakdown was returned.</p>";
    return;
  }

  const rows = breakdown
    .map(
      (item) => `
      <tr>
        <td>${escapeHtml(item.criterion)}</td>
        <td>${escapeHtml(item.score)}</td>
        <td>${escapeHtml(item.maxScore)}</td>
        <td>${escapeHtml(item.rationale)}</td>
      </tr>
    `
    )
    .join("");

  target.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Criterion</th>
          <th>Score</th>
          <th>Max</th>
          <th>Rationale</th>
        </tr>
      </thead>
      <tbody>
        ${rows}
      </tbody>
    </table>
  `;
}

function renderDeductionsInto(target, deductions) {
  if (!Array.isArray(deductions) || deductions.length === 0) {
    target.innerHTML = "<p>No specific deductions were reported.</p>";
    return;
  }

  const cards = deductions
    .map((item) => {
      const evidenceItems = Array.isArray(item.evidence) && item.evidence.length > 0
        ? item.evidence
            .map(
              (evidence, index) => `
                <li>
                  <strong>Issue #${index + 1}</strong>: "${escapeHtml(evidence.snippet)}"
                  <br />
                  <span>Issue: ${escapeHtml(evidence.issue)}</span>
                  <br />
                  <span>Fix: ${escapeHtml(evidence.suggestion)}</span>
                </li>
              `
            )
            .join("")
        : "<li>No evidence provided.</li>";

      return `
        <article class="feedback-card">
          <h4>${escapeHtml(item.criterion || "General")} (${Number(item.pointsLost) || 0} points lost)</h4>
          <p><strong>Why:</strong> ${escapeHtml(item.reason || "")}</p>
          <p><strong>Actionable fix:</strong> ${escapeHtml(item.actionableFix || "")}</p>
          <p><strong>Specific incidents:</strong></p>
          <ul>${evidenceItems}</ul>
        </article>
      `;
    })
    .join("");

  target.innerHTML = `<div class="feedback-stack">${cards}</div>`;
}

function getActiveConfiguredModelId() {
  const value = String(activeModelId || "").trim();
  const parts = value.split(":");
  if (parts.length >= 2) {
    return parts.slice(1).join(":");
  }
  return "";
}

function ensureResultModelUsed(result) {
  const next = result && typeof result === "object" ? { ...result } : {};
  const modelUsed = String(next.modelUsed || "").trim();
  if (!modelUsed) {
    next.modelUsed = getActiveConfiguredModelId() || "gemini-2.5-flash";
  }
  return next;
}

function renderCurrentResult(result) {
  const normalized = ensureResultModelUsed(result);
  if (normalized.grammarOnlyMode) {
    scorePill.textContent = "Feedback Only";
  } else {
    scorePill.textContent = `${normalized.pointsEarned ?? "N/A"}/${normalized.pointsPossible ?? "N/A"} (${normalized.letterGrade ?? "N/A"})`;
  }
  const usedModel = String(normalized?.modelUsed || "").trim() || getActiveConfiguredModelId() || "Unknown";
  const activeConfigured = getActiveConfiguredModelId();
  if (activeConfigured && usedModel !== activeConfigured) {
    resultModelUsed.textContent = `Model used: ${usedModel} (current setting: ${activeConfigured})`;
  } else {
    resultModelUsed.textContent = `Model used: ${usedModel}`;
  }
  wordCountDisplay.textContent = normalized.grammarOnlyMode
    ? `Grammar/Spelling/APA 7 feedback mode | Submission word count: ${normalized.submissionWordCount ?? "N/A"}`
    : `Authoritative submission word count: ${normalized.submissionWordCount ?? "N/A"}`;
  renderDeductionsInto(deductionsWrap, normalized.deductions);
  if (normalized.grammarOnlyMode) {
    rubricWrap.innerHTML = "<p>Rubric breakdown is not generated in grammar-only mode.</p>";
  } else {
    renderRubricTableInto(rubricWrap, normalized.rubricBreakdown);
  }
}

function formatDate(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "";
  }
  return date.toLocaleString();
}

function formatDateOnly(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "";
  }
  return date.toLocaleDateString();
}

function parsePointsScore(value) {
  const text = String(value || "").trim();
  const match = text.match(/^(-?\d+(?:\.\d+)?)\s*\/\s*(-?\d+(?:\.\d+)?)$/);
  if (!match) {
    return null;
  }
  return {
    earned: Number(match[1]),
    possible: Number(match[2])
  };
}

function getScoreGapMeta(projectedScore, trueScore) {
  const projected = parsePointsScore(projectedScore);
  const actual = parsePointsScore(trueScore);
  if (!projected || !actual || !Number.isFinite(projected.earned) || !Number.isFinite(actual.earned)) {
    return null;
  }

  const delta = projected.earned - actual.earned;
  const absDelta = Math.abs(delta);
  const rounded = Math.round(delta * 100) / 100;
  const deltaText = `${rounded > 0 ? "+" : ""}${rounded}`;

  if (absDelta < 0.01) {
    return { label: "Exact", deltaText, tone: "exact" };
  }
  if (absDelta <= 1) {
    return { label: "Close", deltaText, tone: "close" };
  }
  return { label: "Far", deltaText, tone: "far" };
}

function getSortableScoreValue(value) {
  const parsed = parsePointsScore(value);
  if (!parsed || !Number.isFinite(parsed.earned) || !Number.isFinite(parsed.possible)) {
    return Number.NEGATIVE_INFINITY;
  }
  return parsed.earned + parsed.possible / 1000;
}

function getReportModelUsed(report) {
  return String(report?.result?.modelUsed || report?.modelUsed || "gemini-2.5-flash");
}

function compareReportsByActiveSort(a, b) {
  const direction = reportTableSortDirection === "asc" ? 1 : -1;
  let diff = 0;

  if (reportTableSortKey === "name") {
    diff = String(a.name || "").localeCompare(String(b.name || ""));
  } else if (reportTableSortKey === "modelUsed") {
    diff = getReportModelUsed(a).localeCompare(getReportModelUsed(b));
  } else if (reportTableSortKey === "projectedScore") {
    diff = getSortableScoreValue(a.projectedScore) - getSortableScoreValue(b.projectedScore);
  } else if (reportTableSortKey === "trueScore") {
    diff = getSortableScoreValue(a.trueScore) - getSortableScoreValue(b.trueScore);
  } else if (reportTableSortKey === "professorName") {
    diff = String(a.professorName || "").localeCompare(String(b.professorName || ""));
  }

  if (diff === 0) {
    return String(a.name || "").localeCompare(String(b.name || ""));
  }
  return diff * direction;
}

function updateSortHeaderUi() {
  reportSortHeaders.forEach((button) => {
    const key = button.getAttribute("data-sort-key");
    const indicator = button.querySelector(".sort-indicator");
    const isActive = key === reportTableSortKey;
    button.classList.toggle("active", isActive);
    if (indicator) {
      indicator.textContent = isActive ? (reportTableSortDirection === "asc" ? "▲" : "▼") : "↕";
    }
  });
}

function handleSortHeaderClick(sortKey) {
  if (!sortKey) {
    return;
  }
  if (reportTableSortKey === sortKey) {
    reportTableSortDirection = reportTableSortDirection === "asc" ? "desc" : "asc";
  } else {
    reportTableSortKey = sortKey;
    reportTableSortDirection = "asc";
  }
  renderReportsTable();
}

function renderReportsTable() {
  const rows = Array.isArray(reportsCache) ? reportsCache : [];

  if (rows.length === 0) {
    reportsTableBody.innerHTML = "<tr><td colspan='7'>No saved reports found.</td></tr>";
    updateSortHeaderUi();
    return;
  }

  const groups = rows.reduce((acc, report) => {
    const key = String(report.className || "Uncategorized");
    if (!acc[key]) {
      acc[key] = [];
    }
    acc[key].push(report);
    return acc;
  }, {});

  const groupKeys = Object.keys(groups).sort((a, b) => a.localeCompare(b));
  const html = groupKeys
    .map((groupName) => {
      const collapsed = Boolean(collapsedClassGroups[groupName]);
      const groupCount = groups[groupName].length;
      const groupRows = [...groups[groupName]]
        .sort(compareReportsByActiveSort)
        .map(
          (report) => {
            const gap = getScoreGapMeta(report.projectedScore, report.trueScore);
            const gapCell = gap
              ? `<span class="gap-pill gap-${gap.tone}">${gap.label} (${gap.deltaText})</span>`
              : `<span class="gap-pill gap-na">N/A</span>`;
            const hasActualOutcome = Boolean(String(report.trueScore || "").trim() && String(report.trueScore) !== "Not recorded");
            const trueGradeTitle = hasActualOutcome ? "Update True Grade" : "Record True Grade";
            const regradeDisabled = report.canRegrade ? "" : " disabled";
            const regradeTitle = report.canRegrade
              ? "Re-Grade"
              : "Re-Grade unavailable for this report (missing saved rubric/instructions context)";
            return `
            <tr>
              <td>${escapeHtml(report.name)}</td>
              <td>${escapeHtml(getReportModelUsed(report))}</td>
              <td>${escapeHtml(report.projectedScore || "")}</td>
              <td>${escapeHtml(report.trueScore || "Not recorded")}</td>
              <td>${gapCell}</td>
              <td>${escapeHtml(report.professorName)}</td>
              <td class="table-actions">
                <details class="row-actions-menu">
                  <summary class="secondary-btn row-actions-trigger" title="Actions" aria-label="Actions">•••</summary>
                  <div class="row-actions-panel">
                    <button type="button" class="row-action-item" data-action="open" data-report-id="${escapeHtml(
                      report.id
                    )}" title="Open report">Open report</button>
                    <button type="button" class="row-action-item" data-action="regrade" data-report-id="${escapeHtml(
                      report.id
                    )}" title="${escapeHtml(regradeTitle)}"${regradeDisabled}>Re-grade</button>
                    <button type="button" class="row-action-item" data-action="actual" data-report-id="${escapeHtml(
                      report.id
                    )}" title="${escapeHtml(trueGradeTitle)}">${escapeHtml(trueGradeTitle)}</button>
                    <button type="button" class="row-action-item row-action-delete" data-action="delete" data-report-id="${escapeHtml(
                      report.id
                    )}" data-report-name="${escapeHtml(report.name)}" title="Delete report">Delete report</button>
                  </div>
                </details>
              </td>
            </tr>
          `;
          }
        )
        .join("");
      return `
        <tr class="class-group-row">
          <td colspan="7">
            <button
              type="button"
              class="group-toggle-btn"
              data-action="toggle-group"
              data-group-name="${escapeHtml(groupName)}"
              aria-expanded="${collapsed ? "false" : "true"}"
              aria-label="${collapsed ? "Expand class group" : "Collapse class group"}"
              title="${collapsed ? "Expand" : "Collapse"}"
            ><span class="group-caret" aria-hidden="true">${collapsed ? "▸" : "▾"}</span></button>
            <span class="group-label">${escapeHtml(groupName)} (${groupCount})</span>
          </td>
        </tr>
        ${collapsed ? "" : groupRows}
      `;
    })
    .join("");

  reportsTableBody.innerHTML = html;
  updateSortHeaderUi();
}

function renderMetadataOptions() {
  const unique = (values) =>
    [...new Set(values.map((value) => String(value || "").trim()).filter(Boolean))].sort((a, b) =>
      a.localeCompare(b)
    );

  const classNames = unique(reportsCache.map((report) => report.className));
  const professorNames = unique(reportsCache.map((report) => report.professorName));
  const universities = unique(reportsCache.map((report) => report.university));

  classNameOptions.innerHTML = classNames.map((value) => `<option value="${escapeHtml(value)}"></option>`).join("");
  professorNameOptions.innerHTML = professorNames
    .map((value) => `<option value="${escapeHtml(value)}"></option>`)
    .join("");
  universityNameOptions.innerHTML = universities
    .map((value) => `<option value="${escapeHtml(value)}"></option>`)
    .join("");
}

function closeRowActionMenu(fromElement) {
  const menu = fromElement?.closest("details.row-actions-menu");
  if (menu) {
    menu.removeAttribute("open");
  }
}

function updateRowActionMenuPlacement(menu) {
  if (!menu) {
    return;
  }
  const trigger = menu.querySelector(".row-actions-trigger");
  const panel = menu.querySelector(".row-actions-panel");
  if (!(panel instanceof HTMLElement) || !(trigger instanceof HTMLElement)) {
    return;
  }

  const margin = 8;
  panel.style.visibility = "hidden";
  panel.style.left = "-9999px";
  panel.style.top = "-9999px";
  const panelWidth = panel.offsetWidth || 180;
  const panelHeight = panel.offsetHeight || 120;
  const triggerRect = trigger.getBoundingClientRect();

  let left = triggerRect.right - panelWidth;
  if (left < margin) {
    left = margin;
  }
  const maxLeft = window.innerWidth - panelWidth - margin;
  if (left > maxLeft) {
    left = Math.max(margin, maxLeft);
  }

  let top = triggerRect.bottom + 4;
  const maxTop = window.innerHeight - panelHeight - margin;
  if (top > maxTop) {
    top = triggerRect.top - panelHeight - 4;
  }
  if (top < margin) {
    top = margin;
  }

  panel.style.left = `${Math.round(left)}px`;
  panel.style.top = `${Math.round(top)}px`;
  panel.style.visibility = "";
}

async function loadReports() {
  try {
    const response = await apiFetch("/api/reports");
    const data = await readApiJson(response);
    reportsCache = Array.isArray(data.reports) ? data.reports : [];
    renderMetadataOptions();
    renderReportsTable();
  } catch (error) {
    setStatus(error.message || "Failed to load reports.", true);
  }
}

async function loadReportForInlineAction(reportId) {
  const response = await apiFetch(`/api/reports/${encodeURIComponent(reportId)}`);
  const data = await readApiJson(response);
  const report = data.report;
  if (!report) {
    throw new Error("Report payload missing.");
  }
  activeReport = report;
  actualScoreInput.value = report.actualOutcome?.trueScore || "";
  actualFeedbackInput.value = report.actualOutcome?.trueFeedback || "";
  return report;
}

async function openSavedReport(reportId) {
  try {
    const response = await apiFetch(`/api/reports/${encodeURIComponent(reportId)}`);
    const data = await readApiJson(response);

    const report = data.report;
    if (!report) {
      throw new Error("Report payload missing.");
    }
    activeReport = report;

    viewerTitle.textContent = report.name || "Saved Report";
    viewerMeta.textContent = `${report.className || ""} | ${report.professorName || ""} | ${report.university || ""} | ${formatDate(
      report.runDate
    )}`;
    viewerScore.textContent = `Predicted Score (AI): ${report.result?.pointsEarned ?? "N/A"}/${report.result?.pointsPossible ?? "N/A"} (${report.result?.letterGrade ?? "N/A"}) | Word Count: ${report.result?.submissionWordCount ?? "N/A"}`;
    renderDeductionsInto(viewerDeductions, report.result?.deductions || []);
    renderRubricTableInto(viewerRubric, report.result?.rubricBreakdown || []);

    const hasActualOutcome = Boolean(
      String(report.actualOutcome?.trueScore || "").trim() || String(report.actualOutcome?.trueFeedback || "").trim()
    );
    outcomeComparison.classList.toggle("hidden", !hasActualOutcome);
    if (hasActualOutcome) {
      trueScoreDisplay.textContent = `True Score: ${report.actualOutcome?.trueScore || "N/A"}`;
      trueFeedbackDisplay.textContent = report.actualOutcome?.trueFeedback || "No true feedback recorded.";
      predictedScoreDisplay.textContent = `Predicted Score: ${report.result?.pointsEarned ?? "N/A"}/${report.result?.pointsPossible ?? "N/A"}`;
      renderPredictedFeedbackSummary(report.result || {});
    } else {
      trueScoreDisplay.textContent = "";
      trueFeedbackDisplay.textContent = "";
      predictedScoreDisplay.textContent = "";
      predictedFeedbackDisplay.innerHTML = "";
    }

    const isGrammarOnly = String(report.gradingContext?.gradingOption || "standard") === "grammar_only";
    const canRegrade = isGrammarOnly
      ? Boolean(report.gradingContext?.promptInstructions)
      : Boolean(report.gradingContext?.rubric && report.gradingContext?.instructions && report.gradingContext?.promptInstructions);
    closeRegradeDialog();
    closeActualOutcomeDialog();
    actualScoreInput.value = report.actualOutcome?.trueScore || "";
    actualFeedbackInput.value = report.actualOutcome?.trueFeedback || "";

    viewerSection.classList.remove("hidden");
    setStatus("Saved report loaded.");
  } catch (error) {
    setStatus(error.message || "Failed to open saved report.", true);
  }
}

async function deleteSavedReport(reportId, reportName) {
  const name = String(reportName || "this report");
  const confirmed = window.confirm(`Delete saved report "${name}"? This cannot be undone.`);
  if (!confirmed) {
    return;
  }

  try {
    const response = await apiFetch(`/api/reports/${encodeURIComponent(reportId)}`, {
      method: "DELETE"
    });
    await readApiJson(response);

    if (activeReport && String(activeReport.id) === String(reportId)) {
      activeReport = null;
      closeReportViewer();
    }

    await loadReports();
    setStatus("Saved report deleted.");
  } catch (error) {
    setStatus(error.message || "Failed to delete report.", true);
  }
}

async function regradeActiveReport() {
  if (!activeReport?.id) {
    setStatus("Open a saved report first.", true);
    return;
  }
  const file = regradeSubmissionFile.files?.[0];
  if (!file) {
    setStatus("Please upload the updated submission file.", true);
    return;
  }

  regradeBtn.disabled = true;
  try {
    const formData = new FormData();
    formData.set("submission", file);
    const response = await apiFetch(`/api/reports/${encodeURIComponent(activeReport.id)}/regrade`, {
      method: "POST",
      body: formData
    });
    const data = await readApiJson(response);

    latestResult = ensureResultModelUsed(data.result || {});
    latestRunDate = data.report?.runDate || new Date().toISOString();
    latestGradingContext = activeReport.gradingContext || null;
    setSaveReportButtonSaved(false);
    renderCurrentResult(latestResult);
    resultsSection.classList.remove("hidden");

    if (data.report) {
      reportNameInput.value = data.report.name || reportNameInput.value;
      classNameInput.value = data.report.className || classNameInput.value;
      professorNameInput.value = data.report.professorName || professorNameInput.value;
      universityNameInput.value = data.report.university || universityNameInput.value;
    }

    await loadReports();
    await openSavedReport(activeReport.id);
    closeRegradeDialog();
    setStatus("Regrade complete. The saved report was replaced with the new result.");
  } catch (error) {
    setStatus(error.message || "Failed to regrade report.", true);
  } finally {
    regradeBtn.disabled = false;
  }
}

async function saveActualOutcome() {
  if (!activeReport?.id) {
    setStatus("Open a saved report first.", true);
    return;
  }

  const trueScore = String(actualScoreInput.value || "").trim();
  const trueFeedback = String(actualFeedbackInput.value || "").trim();
  const clearing = !trueScore && !trueFeedback;
  if (!clearing && (!trueScore || !trueFeedback)) {
    setStatus("Provide both true score and professor feedback, or clear both fields to remove true outcome.", true);
    return;
  }

  saveActualBtn.disabled = true;
  try {
    const response = await apiFetch(`/api/reports/${encodeURIComponent(activeReport.id)}/actual-outcome`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ trueScore, trueFeedback })
    });
    const data = await readApiJson(response);

    activeReport = data.report || activeReport;
    actualScoreInput.value = activeReport.actualOutcome?.trueScore || trueScore;
    actualFeedbackInput.value = activeReport.actualOutcome?.trueFeedback || trueFeedback;

    await loadReports();
    setSaveActualButtonSaved(true);
    setStatus(data.cleared ? "True outcome cleared." : "True outcome saved.");
  } catch (error) {
    setStatus(error.message || "Failed to save true outcome.", true);
  } finally {
    saveActualBtn.disabled = false;
  }
}

async function saveCurrentReport() {
  if (!latestResult) {
    setStatus("Grade a submission first before saving a report.", true);
    return;
  }

  const name = String(reportNameInput.value || "").trim();
  const className = String(classNameInput.value || "").trim();
  const professorName = String(professorNameInput.value || "").trim();
  const university = String(universityNameInput.value || "").trim();

  if (!name || !className || !professorName || !university) {
    setStatus("Report name, class name, professor name, and university are required to save.", true);
    return;
  }
  const isGrammarOnly = String(latestGradingContext?.gradingOption || "standard") === "grammar_only";
  if (!latestGradingContext?.promptInstructions) {
    setStatus("This result is missing prompt context. Please grade again before saving.", true);
    return;
  }
  if (!isGrammarOnly && (!latestGradingContext?.rubric || !latestGradingContext?.instructions)) {
    setStatus("This result is missing rubric/instructions context. Please grade again before saving.", true);
    return;
  }

  saveReportBtn.disabled = true;
  try {
    const resultToSave = ensureResultModelUsed(latestResult || {});
    const response = await apiFetch("/api/reports", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        name,
        className,
        professorName,
        university,
        runDate: latestRunDate || new Date().toISOString(),
        result: resultToSave,
        gradingContext: latestGradingContext
      })
    });
    await readApiJson(response);

    await loadReports();
    setSaveReportButtonSaved(true);
    setStatus("Feedback report saved.");
  } catch (error) {
    setStatus(error.message || "Failed to save report.", true);
  } finally {
    saveReportBtn.disabled = false;
  }
}

async function saveCustomPrompt() {
  const name = String(savedPromptName.value || "").trim();
  const text = String(promptInstructions.value || "").trim();

  if (!text) {
    setStatus("Enter prompt instructions before saving.", true);
    return;
  }

  if (!name) {
    setStatus("Enter a name for your saved prompt.", true);
    return;
  }

  savePromptBtn.disabled = true;
  try {
    const response = await apiFetch("/api/custom-prompts", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name, text })
    });
    const data = await readApiJson(response);
    const newPromptId = String(data?.prompt?.id || "").trim();

    savedPromptName.value = "";
    await loadSavedPrompts();
    if (newPromptId) {
      promptProfile.value = `saved:${newPromptId}`;
      togglePromptControls();
    }
    setStatus("Prompt saved. You can now select it from Evaluation Prompt.");
  } catch (error) {
    setStatus(error.message || "Failed to save prompt.", true);
  } finally {
    savePromptBtn.disabled = false;
  }
}

async function saveSelectedPromptEdits() {
  if (!isSavedPromptSelection()) {
    setStatus("Select a saved prompt to save edits.", true);
    return;
  }
  const savedId = promptProfile.value.slice("saved:".length);
  const text = String(promptInstructions.value || "").trim();
  if (!text) {
    setStatus("Prompt Instructions cannot be empty.", true);
    return;
  }

  const name = String(savedPromptNameMap[savedId] || "").trim();
  const promptType = String(savedPromptTypeMap[savedId] || "custom");
  if (promptType === "professor_calibrated") {
    const confirmed = window.confirm(
      "Save changes to this professor prompt? This will overwrite the existing professor prompt text."
    );
    if (!confirmed) {
      return;
    }
  }

  savePromptEditsBtn.disabled = true;
  try {
    const response = await apiFetch(`/api/custom-prompts/${encodeURIComponent(savedId)}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name, text })
    });
    await readApiJson(response);
    await loadSavedPrompts();
    promptProfile.value = `saved:${savedId}`;
    syncPromptEditorFromSelection();
    togglePromptControls();
    updateSavePromptEditsButtonState();
    setStatus("Saved prompt updated.");
  } catch (error) {
    setStatus(error.message || "Failed to save prompt changes.", true);
  } finally {
    if (isSavedPromptSelection()) {
      updateSavePromptEditsButtonState();
    } else {
      savePromptEditsBtn.disabled = false;
    }
  }
}

async function deleteSelectedPrompt() {
  if (!isSavedPromptSelection()) {
    return;
  }

  const savedId = promptProfile.value.slice("saved:".length);
  const savedName = savedPromptNameMap[savedId] || "this saved prompt";
  const confirmed = window.confirm(`Delete saved prompt \"${savedName}\"? This cannot be undone.`);
  if (!confirmed) {
    return;
  }

  deletePromptBtn.disabled = true;
  try {
    const response = await apiFetch(`/api/custom-prompts/${encodeURIComponent(savedId)}`, {
      method: "DELETE"
    });
    await readApiJson(response);

    await loadSavedPrompts();
    setStatus("Saved prompt deleted.");
  } catch (error) {
    setStatus(error.message || "Failed to delete saved prompt.", true);
  } finally {
    deletePromptBtn.disabled = false;
  }
}

promptProfile.addEventListener("change", handlePromptProfileChange);
gradingOptionGrammarOnly.addEventListener("change", handleGrammarOnlyToggle);
promptInstructions.addEventListener("input", () => {
  if (isSavedPromptSelection()) {
    updateSavePromptEditsButtonState();
  }
});
otherFile.addEventListener("change", toggleOtherRelevance);
submissionFileInput.addEventListener("change", () => {
  setSaveReportButtonSaved(false);
  if (!String(reportNameInput.value || "").trim()) {
    reportNameInput.value = suggestReportName();
  }
});
reportNameInput.addEventListener("input", () => setSaveReportButtonSaved(false));
classNameInput.addEventListener("input", () => setSaveReportButtonSaved(false));
professorNameInput.addEventListener("input", () => setSaveReportButtonSaved(false));
universityNameInput.addEventListener("input", () => setSaveReportButtonSaved(false));
savePromptBtn.addEventListener("click", saveCustomPrompt);
savePromptEditsBtn.addEventListener("click", saveSelectedPromptEdits);
deletePromptBtn.addEventListener("click", deleteSelectedPrompt);
saveReportBtn.addEventListener("click", saveCurrentReport);
actualScoreInput.addEventListener("input", () => setSaveActualButtonSaved(false));
actualFeedbackInput.addEventListener("input", () => setSaveActualButtonSaved(false));
closePromptBuilderDialogBtn.addEventListener("click", closePromptBuilderDialog);
attachSafeBackdropClose(promptBuilderDialog, closePromptBuilderDialog);
promptBuilderProfessorSelect.addEventListener("change", () => {
  void loadPromptBuilderFeedbackForProfessor().catch((error) => {
    setPromptBuilderStatus(error.message || "Failed to load professor feedback.", true);
  });
});
promptBuilderSelectAllFeedback.addEventListener("change", () => {
  const checked = Boolean(promptBuilderSelectAllFeedback.checked);
  const checks = promptBuilderDialog.querySelectorAll(".prompt-builder-feedback-check");
  checks.forEach((item) => {
    item.checked = checked;
  });
  updatePromptBuilderSelectAllState();
});
promptBuilderFeedbackBody.addEventListener("change", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLInputElement) || !target.classList.contains("prompt-builder-feedback-check")) {
    return;
  }
  updatePromptBuilderSelectAllState();
});
promptBuilderOriginalPromptSelect.addEventListener("change", updatePromptBuilderOriginalPromptText);
promptBuilderBuildBtn.addEventListener("click", buildPromptFromFeedback);
promptBuilderSaveBtn.addEventListener("click", saveGeneratedPromptFromBuilder);
reportSortHeaders.forEach((button) => {
  button.addEventListener("click", () => {
    handleSortHeaderClick(button.getAttribute("data-sort-key"));
  });
});
reportsTableBody.addEventListener("click", async (event) => {
  const rawTarget = event.target;
  if (!(rawTarget instanceof HTMLElement)) {
    return;
  }
  const target = rawTarget.closest("[data-action]");
  if (!(target instanceof HTMLElement)) {
    return;
  }

  const action = target.getAttribute("data-action");
  if (action === "toggle-group") {
    const groupName = String(target.getAttribute("data-group-name") || "").trim();
    if (!groupName) {
      return;
    }
    if (collapsedClassGroups[groupName]) {
      delete collapsedClassGroups[groupName];
    } else {
      collapsedClassGroups[groupName] = true;
    }
    saveCollapsedGroupsState();
    renderReportsTable();
    return;
  }

  const reportId = target.getAttribute("data-report-id");
  if (!reportId) {
    return;
  }
  if (action === "open") {
    closeRowActionMenu(target);
    await openSavedReport(reportId);
    return;
  }
  if (action === "regrade") {
    try {
      closeRowActionMenu(target);
      const report = await loadReportForInlineAction(reportId);
      const isGrammarOnly = String(report.gradingContext?.gradingOption || "standard") === "grammar_only";
      const canRegrade = isGrammarOnly
        ? Boolean(report.gradingContext?.promptInstructions)
        : Boolean(report.gradingContext?.rubric && report.gradingContext?.instructions && report.gradingContext?.promptInstructions);
      if (!canRegrade) {
        setStatus("This report cannot be re-graded because rubric/instructions context is missing.", true);
        return;
      }
      openRegradeDialog();
    } catch (error) {
      setStatus(error.message || "Failed to prepare re-grade.", true);
    }
    return;
  }
  if (action === "actual") {
    try {
      closeRowActionMenu(target);
      await loadReportForInlineAction(reportId);
      openActualOutcomeDialog();
    } catch (error) {
      setStatus(error.message || "Failed to open true grade dialog.", true);
    }
    return;
  }
  if (action === "delete") {
    closeRowActionMenu(target);
    const reportName = target.getAttribute("data-report-name") || "this report";
    await deleteSavedReport(reportId, reportName);
  }
});
reportsTableBody.addEventListener(
  "toggle",
  (event) => {
    const target = event.target;
    if (!(target instanceof HTMLDetailsElement) || !target.classList.contains("row-actions-menu")) {
      return;
    }
    if (!target.open) {
      return;
    }
    document.querySelectorAll("details.row-actions-menu[open]").forEach((menu) => {
      if (menu !== target) {
        menu.removeAttribute("open");
      }
    });
    updateRowActionMenuPlacement(target);
  },
  true
);
window.addEventListener("resize", () => {
  document.querySelectorAll("details.row-actions-menu[open]").forEach((menu) => {
    menu.removeAttribute("open");
  });
});
window.addEventListener(
  "scroll",
  () => {
    document.querySelectorAll("details.row-actions-menu[open]").forEach((menu) => {
      menu.removeAttribute("open");
    });
  },
  true
);
document.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof Element)) {
    return;
  }
  if (target.closest(".row-actions-menu")) {
    return;
  }
  document.querySelectorAll("details.row-actions-menu[open]").forEach((menu) => {
    menu.removeAttribute("open");
  });
});
regradeBtn.addEventListener("click", regradeActiveReport);
closeRegradeDialogBtn.addEventListener("click", closeRegradeDialog);
closeActualOutcomeDialogBtn.addEventListener("click", closeActualOutcomeDialog);
saveActualBtn.addEventListener("click", saveActualOutcome);
tabGraderBtn.addEventListener("click", () => switchTab("grader"));
tabReportsBtn.addEventListener("click", () => switchTab("reports"));
tabSettingsBtn.addEventListener("click", () => switchTab("settings"));
  llmModelSelect.addEventListener("change", () => {
  applyLlmKeyHint();
  const changed = activeModelId && llmModelSelect.value !== activeModelId;
  setSettingsStatus(changed ? "Unsaved change: click Save LLM Settings to make this the active model." : "", false);
  setSettingsTestStatus("", false);
});
testLlmModelBtn.addEventListener("click", testSelectedLlmModel);
saveLlmSettingsBtn.addEventListener("click", saveLlmSettings);
saveGraderTimeoutBtn.addEventListener("click", saveGraderTimeoutSetting);
llmApiKeyInput.addEventListener("input", () => {
  setSettingsTestStatus("", false);
});
graderTimeoutSecondsInput.addEventListener("input", () => {
  setGraderTimeoutStatus("", false);
});
studentProfileSelect.addEventListener("change", handleProfileSelectionChange);
closeViewerBtn.addEventListener("click", closeReportViewer);
attachSafeBackdropClose(viewerSection, closeReportViewer);
attachSafeBackdropClose(regradeDialog, closeRegradeDialog);
attachSafeBackdropClose(actualOutcomeDialog, closeActualOutcomeDialog);

togglePromptControls();
toggleOtherRelevance();
switchTab("grader");
initializeFileUploadControls();
loadGraderTimeoutSetting();

async function initializeApp() {
  try {
    await loadProfiles();
    await reloadProfileData();
    await loadLlmSettings();
    await loadPromptBuilderProfessors();
    renderPromptBuilderOriginalPromptOptions();
    promptBuilderInstructions.value = DEFAULT_PROMPT_BUILD_INSTRUCTIONS;
    promptBuilderOutput.value = "";
    promptBuilderSaveName.value = "";
    promptBuilderSaveBtn.disabled = true;
    loadGraderTimeoutSetting();
  } catch (error) {
    setStatus(error.message || "Failed to initialize profiles.", true);
  }
}

initializeApp();

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const formData = new FormData(form);
  const promptText = String(formData.get("promptInstructions") || "").trim();
  const submissionType = getSubmissionType();
  const gradingOption = getGradingOption();

  if (submissionType === "invalid") {
    setStatus("Select only one submission type: Discussion or Discussion Reply.", true);
    return;
  }
  formData.set("submissionType", submissionType);
  formData.set("gradingOption", gradingOption);

  if (!promptText && gradingOption !== "grammar_only") {
    setStatus("Prompt Instructions is required.", true);
    return;
  }

  if (gradingOption !== "grammar_only") {
    const hasRubric = Boolean(rubricFileInput.files && rubricFileInput.files.length > 0);
    const hasInstructions = Boolean(instructionsFileInput.files && instructionsFileInput.files.length > 0);
    if (!hasRubric || !hasInstructions) {
      setStatus("Please upload rubric and assignment instructions, or enable Spelling and Grammar Only mode.", true);
      return;
    }
  }

  if (otherFile.files && otherFile.files.length > 0 && !String(formData.get("otherRelevance") || "").trim()) {
    setStatus("Please explain why the Other supporting file is relevant.", true);
    return;
  }

  submitBtn.disabled = true;
  const originalSubmitLabel = submitBtn.textContent;
  submitBtn.textContent = "Grading...";
  resultsSection.classList.add("hidden");
  latestResult = null;
  latestRunDate = null;
  latestGradingContext = null;
  const startedAt = Date.now();
  const timeoutSeconds = Math.floor(gradeRequestTimeoutMs / 1000);
  setStatus(`Grading in progress... 0s elapsed (timeout ${timeoutSeconds}s).`, false, true);
  const progressTimer = setInterval(() => {
    const elapsedSeconds = Math.floor((Date.now() - startedAt) / 1000);
    setStatus(`Grading in progress... ${elapsedSeconds}s elapsed (timeout ${timeoutSeconds}s).`, false, true);
  }, 1000);

  try {
    const response = await apiFetchWithTimeout(
      "/api/grade",
      {
      method: "POST",
      body: formData
      },
      gradeRequestTimeoutMs
    );
    const data = await readApiJson(response);

    latestResult = ensureResultModelUsed(data.result || {});
    latestRunDate = new Date().toISOString();
    latestGradingContext = data.gradingContext || null;
    setSaveReportButtonSaved(false);
    renderCurrentResult(latestResult);

    if (!String(reportNameInput.value || "").trim()) {
      reportNameInput.value = suggestReportName();
    }

    resultsSection.classList.remove("hidden");
    setStatus("Grading complete.", false, true);
  } catch (error) {
    if (error?.name === "AbortError") {
      setStatus(
        `Grading timed out after ${Math.floor(gradeRequestTimeoutMs / 1000)}s. The model may be rate-limited or overloaded. Try again or switch models in Settings.`,
        true
      );
      return;
    }
    setStatus(error.message || "Unexpected error occurred.", true);
  } finally {
    clearInterval(progressTimer);
    submitBtn.disabled = false;
    submitBtn.textContent = originalSubmitLabel;
  }
});
