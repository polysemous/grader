# AI Assignment Grader

A local web app for assignment feedback and grading workflows using uploaded submission files, rubric/instructions context, reusable prompts, and saved reports.

Supported upload formats: `.pdf`, `.doc`, `.docx`

## Setup

1. Install dependencies.

```bash
npm install
```

2. Create your local environment file from the template.

```bash
cp .env.example .env
```

3. Add your own API key values to `.env`.

4. Start the app.

```bash
npm start
```

5. Open `http://localhost:3000`

## Security and Local Data

- `.env` is local-only and must never be committed.
- Runtime app data is local-only and stored under `data/`.
- The repository is configured to ignore `data/*.json`, which may contain:
  - API-key settings metadata
  - saved prompts
  - saved reports
  - profile data
- If you want example configuration for collaborators, use `.env.example` rather than committing a real `.env`.

## Notes

- Some grading modes require rubric and instructions; grammar-only mode does not.
- Saved prompts and reports are created locally at runtime.
- Some legacy `.doc` files may parse imperfectly depending on source encoding; `.docx` and `.pdf` are typically more reliable.
