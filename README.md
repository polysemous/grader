# AI Assignment Grader

A simple web app that grades a student assignment using either OpenAI or Gemini based on:

1. Grading Rubric
2. Assignment Instructions
3. Student Submission
4. Optional Other Supporting File (with relevance explanation)

Supported upload formats: `.pdf`, `.doc`, `.docx`

## Features

- Required upload of all three grading inputs.
- Optional fourth upload (`Other`) with required relevance explanation when used.
- Prompt profile dropdown with useful defaults:
  - Graduate Professor
  - High School Teacher
  - Supportive Tutor
  - Strict Examiner
  - Add your own prompt
- Save custom prompts for future grading reuse.
- Select saved custom prompts directly from the same evaluation dropdown.
- Delete saved custom prompts with confirmation.
- Single editable `Prompt Instructions` field that auto-loads selected prompt text and can be edited per grading run (not saved unless explicitly saved as a custom prompt).
- AI-powered grading output:
  - Points earned / points possible (example: `24/25`)
  - Letter grade
  - Summary
  - Strengths
  - Areas to improve
  - Rubric breakdown table

## Setup

1. Install dependencies:

```bash
npm install
```

2. Create an environment file:

```bash
cp .env.example .env
```

3. Configure provider and key in `.env`:

### Gemini (recommended if OpenAI quota is exceeded)

```env
LLM_PROVIDER=gemini
GEMINI_API_KEY=your_gemini_api_key_here
GEMINI_MODEL=gemini-2.5-flash
PORT=3000
```

### OpenAI

```env
LLM_PROVIDER=openai
OPENAI_API_KEY=your_openai_api_key_here
OPENAI_MODEL=gpt-4.1-mini
PORT=3000
```

4. Start the app:

```bash
npm start
```

5. Open in browser:

`http://localhost:3000`

## Notes

- Prompt Instructions must be filled before grading.
- Saved prompts are stored locally in `data/custom-prompts.json`.
- Some legacy `.doc` files may parse imperfectly depending on source encoding; `.docx` and `.pdf` are typically more reliable.
