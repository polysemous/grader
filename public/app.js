const form = document.getElementById("grade-form");
const promptProfile = document.getElementById("prompt-profile");
const promptInstructions = document.getElementById("prompt-instructions");
const savePromptWrap = document.getElementById("save-prompt-wrap");
const savedPromptName = document.getElementById("saved-prompt-name");
const savePromptBtn = document.getElementById("save-prompt-btn");
const deletePromptWrap = document.getElementById("delete-prompt-wrap");
const deletePromptBtn = document.getElementById("delete-prompt-btn");
const otherFile = document.getElementById("other-file");
const otherRelevanceWrap = document.getElementById("other-relevance-wrap");
const otherRelevance = document.getElementById("other-relevance");
const submissionTypeDiscussion = document.getElementById("submission-type-discussion");
const submissionTypeDiscussionReply = document.getElementById("submission-type-discussion-reply");

const submitBtn = document.getElementById("submit-btn");
const statusEl = document.getElementById("status");

const resultsSection = document.getElementById("results");
const scorePill = document.getElementById("score-pill");
const wordCountDisplay = document.getElementById("word-count-display");
const deductionsWrap = document.getElementById("deductions-wrap");
const rubricWrap = document.getElementById("rubric-table-wrap");

const PROMPT_PREVIEWS = {
  graduate_professor:
    "You are a rigorous graduate-level professor. Evaluate with high standards for depth, originality, precision, and scholarly quality.",
  high_school_teacher:
    "You are a high school teacher. Evaluate for clarity, structure, factual correctness, and age-appropriate expectations.",
  supportive_tutor:
    "You are a supportive tutor. Grade fairly and provide highly actionable feedback that helps the student improve quickly.",
  strict_examiner:
    "You are a strict examiner. Prioritize adherence to rubric criteria and assignment requirements over stylistic generosity."
};

let savedPromptMap = {};
let savedPromptNameMap = {};

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.className = isError ? "status-error" : "";
}

function renderSavedPromptOptions(prompts) {
  const oldGroup = document.getElementById("saved-prompt-group");
  if (oldGroup) {
    oldGroup.remove();
  }

  if (!Array.isArray(prompts) || prompts.length === 0) {
    return;
  }

  const group = document.createElement("optgroup");
  group.id = "saved-prompt-group";
  group.label = "Saved Custom Prompts";

  prompts.forEach((prompt) => {
    const option = document.createElement("option");
    option.value = `saved:${prompt.id}`;
    option.textContent = `Saved: ${prompt.name}`;
    group.appendChild(option);
  });

  promptProfile.insertBefore(group, promptProfile.querySelector('option[value="custom"]'));
}

function isSavedPromptSelection() {
  return promptProfile.value.startsWith("saved:");
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

async function loadSavedPrompts() {
  try {
    const currentValue = promptProfile.value;
    const response = await fetch("/api/custom-prompts");
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Failed to load saved prompts.");
    }

    const prompts = Array.isArray(data.prompts) ? data.prompts : [];
    savedPromptMap = {};
    savedPromptNameMap = {};
    prompts.forEach((item) => {
      savedPromptMap[item.id] = item.text;
      savedPromptNameMap[item.id] = item.name;
    });

    renderSavedPromptOptions(prompts);
    const canKeepCurrent =
      currentValue === "custom" ||
      currentValue in PROMPT_PREVIEWS ||
      (currentValue.startsWith("saved:") && savedPromptMap[currentValue.slice("saved:".length)]);
    promptProfile.value = canKeepCurrent ? currentValue : "graduate_professor";
    syncPromptEditorFromSelection();
    togglePromptControls();
  } catch (error) {
    setStatus(error.message || "Failed to load saved prompts.", true);
  }
}

function togglePromptControls() {
  const showCustomTools = promptProfile.value === "custom";
  savePromptWrap.classList.toggle("hidden", !showCustomTools);
  deletePromptWrap.classList.toggle("hidden", !isSavedPromptSelection());
  syncPromptEditorFromSelection();
}

function toggleOtherRelevance() {
  const hasOtherFile = Boolean(otherFile.files && otherFile.files.length > 0);
  otherRelevanceWrap.classList.toggle("hidden", !hasOtherFile);
  otherRelevance.required = hasOtherFile;
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
    const response = await fetch("/api/custom-prompts", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name, text })
    });

    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Failed to save prompt.");
    }

    savedPromptName.value = "";
    await loadSavedPrompts();
    setStatus("Prompt saved. You can now select it from Evaluation Prompt.");
  } catch (error) {
    setStatus(error.message || "Failed to save prompt.", true);
  } finally {
    savePromptBtn.disabled = false;
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
    const response = await fetch(`/api/custom-prompts/${encodeURIComponent(savedId)}`, {
      method: "DELETE"
    });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Failed to delete saved prompt.");
    }

    await loadSavedPrompts();
    setStatus("Saved prompt deleted.");
  } catch (error) {
    setStatus(error.message || "Failed to delete saved prompt.", true);
  } finally {
    deletePromptBtn.disabled = false;
  }
}

function renderRubricTable(breakdown) {
  if (!Array.isArray(breakdown) || breakdown.length === 0) {
    rubricWrap.innerHTML = "<p>No rubric breakdown was returned.</p>";
    return;
  }

  const rows = breakdown
    .map(
      (item) => `
      <tr>
        <td>${item.criterion}</td>
        <td>${item.score}</td>
        <td>${item.maxScore}</td>
        <td>${item.rationale}</td>
      </tr>
    `
    )
    .join("");

  rubricWrap.innerHTML = `
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

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function renderDeductions(deductions) {
  if (!Array.isArray(deductions) || deductions.length === 0) {
    deductionsWrap.innerHTML = "<p>No specific deductions were reported.</p>";
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

  deductionsWrap.innerHTML = `<div class="feedback-stack">${cards}</div>`;
}

promptProfile.addEventListener("change", togglePromptControls);
otherFile.addEventListener("change", toggleOtherRelevance);
savePromptBtn.addEventListener("click", saveCustomPrompt);
deletePromptBtn.addEventListener("click", deleteSelectedPrompt);

togglePromptControls();
toggleOtherRelevance();
loadSavedPrompts();

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const formData = new FormData(form);
  const promptText = String(formData.get("promptInstructions") || "").trim();
  const submissionType = getSubmissionType();

  if (submissionType === "invalid") {
    setStatus("Select only one submission type: Discussion or Discussion Reply.", true);
    return;
  }
  formData.set("submissionType", submissionType);

  if (!promptText) {
    setStatus("Prompt Instructions is required.", true);
    return;
  }

  if (otherFile.files && otherFile.files.length > 0 && !String(formData.get("otherRelevance") || "").trim()) {
    setStatus("Please explain why the Other supporting file is relevant.", true);
    return;
  }

  submitBtn.disabled = true;
  resultsSection.classList.add("hidden");
  setStatus("Grading in progress...");

  try {
    const response = await fetch("/api/grade", {
      method: "POST",
      body: formData
    });

    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "Failed to grade assignment.");
    }

    const result = data.result || {};

    scorePill.textContent = `${result.pointsEarned ?? "N/A"}/${result.pointsPossible ?? "N/A"} (${result.letterGrade ?? "N/A"})`;
    wordCountDisplay.textContent = `Authoritative submission word count: ${result.submissionWordCount ?? "N/A"}`;
    renderDeductions(result.deductions);
    renderRubricTable(result.rubricBreakdown);

    resultsSection.classList.remove("hidden");
    setStatus("Grading complete.");
  } catch (error) {
    setStatus(error.message || "Unexpected error occurred.", true);
  } finally {
    submitBtn.disabled = false;
  }
});
