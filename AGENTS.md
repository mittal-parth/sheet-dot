# AGENTS.md

Notes for anyone (human or coding agent) setting up or changing Sheetdot. Read [README.md](README.md) for what the product does, [PLAN.md](PLAN.md) for the decisions behind it, and [docs/sheetdot-data-flow.svg](docs/sheetdot-data-flow.svg) for how data moves.

## How it fits together

| Part | Where | What it does |
|---|---|---|
| Excel add-in | `addin/public/` | Plain JavaScript, no build step. Custom functions (`functions.js`, `functions.json`), the task pane (`taskpane.html`), column suggestions, About me and Automations sheets, bill tracking (`bills.js`). |
| Backend | `backend/worker.ts` | One Cloudflare Worker. Serves the add-in files from `addin/public` and handles every request from Excel. |
| Workbook store | `Workbook` Durable Object in `worker.ts` | One per workbook. Saved answers, watches, About me preferences. |
| Bill watch | `BillWatch` Durable Object in `bills.ts` | One per workbook, holding a watch for each tracked table. Sets its own alarm, searches Gmail, reconciles bill rows. |
| OpenAI | `ask.ts`, `watch.ts`, `suggestions.ts`, `gmail.ts` | Responses API for ASK, LIST, suggestions and reading emails. Decisions API for CHECK. OpenAI-hosted agents (Agents API automations) for WATCH. Called over REST; there is no SDK dependency. |
| Composio | `gmail.ts` | Gmail OAuth and read-only Gmail calls through Composio's proxy. |
| Manifests | `addin/manifest.xml`, `addin/manifest.hosted.xml` | Local (`https://localhost:3000`) and deployed (`https://sheet-dot.songguesser.workers.dev`) versions of the same add-in. |

Model names live at the top of `backend/ask.ts`.

## What you need

- Node 24 and npm.
- A personal Microsoft account that can open Excel for the web.
- An OpenAI API key with Agents API access (beta header `OpenAI-Beta: agents=v1`).
- A Composio API key, plus a Gmail auth config that requests only `gmail.readonly`.
- A Cloudflare account (free tier works) if you are deploying.

## 1. Keys

Create `.env` in the repo root. It is gitignored; never commit it.

```
OPENAI_API_KEY=...
COMPOSIO_API_KEY=...
COMPOSIO_GMAIL_AUTH_CONFIG_ID=...
```

Formulas work with only `OPENAI_API_KEY`. Bill tracking also needs both Composio values.

## 2. Run locally

Office add-ins must load over HTTPS, so install the Office dev certificate once:

```bash
npx office-addin-dev-certs install
```

Then start the backend. It serves both the API and the add-in files at `https://localhost:3000`:

```bash
cd backend && npm install && npm run dev
```

Check it is up at `https://localhost:3000/health`.

## 3. Load the add-in in Excel for the web

1. Open a workbook at excel.cloud.microsoft.
2. Go to **Home → Add-ins → More add-ins → My Add-ins → Upload My Add-in**.
3. Upload `addin/manifest.xml` for local development, or `addin/manifest.hosted.xml` for the deployed backend.
4. Open the **Sheetdot** task pane from the ribbon.

To check it works end to end:

- Type `=AI.ASK("What is the capital of Karnataka?")` in a cell. You should see ⏳, then an answer with a purple source comment.
- In the task pane, choose **Add demo sheets**. It adds **Weekend Plans**, **Find a Flat**, **Family Bills** and a **Demo Guide**. **Refresh demo data** pulls current rows from public sources into the first two.
- On **Family Bills**, select the table, click **Track bills**, connect a Gmail account that has a real bill and its receipt, and pick **every 1 min**. Within a couple of minutes the table should gain the bill row, then mark it **Paid**.

## 4. Deploy

```bash
cd backend
npx wrangler login
npx wrangler deploy
npx wrangler secret put OPENAI_API_KEY
npx wrangler secret put COMPOSIO_API_KEY
npx wrangler secret put COMPOSIO_GMAIL_AUTH_CONFIG_ID
```

If your Worker URL differs from `sheet-dot.songguesser.workers.dev`, replace that host everywhere in `addin/manifest.hosted.xml` before uploading it. Watch live logs with `npx wrangler tail`.

## Checks

- `cd backend && npm run check` builds the Worker without deploying (`wrangler deploy --dry-run`). Run it after every backend change.
- There is no unit test suite yet. Verify add-in changes in real Excel for the web, using the steps in section 3.

## Rules for changes

- **Excel is the only interface.** No chat panels. Prefer native Excel things: formulas, tables, comments, conditional formatting, checkboxes.
- **Gmail is read-only.** The agent never sends, deletes or pays. Any future write action must be proposed in the sheet and approved by a tick from the user.
- **Keep the last good result.** If a source fails or Gmail disconnects, leave the existing cell values alone and show the problem in Automations (for example, `Reconnect Gmail`).
- **A WATCH only edits its own table.**
- **Never log or commit keys**, and never put user data in URLs.
- Update [PLAN.md](PLAN.md) when a product decision changes, and update the diagram when the data flow changes.

## Updating the diagram

`docs/sheetdot-data-flow.excalidraw` is the source. Open it at [excalidraw.com](https://excalidraw.com), edit, then export **SVG** (with "Embed scene" off) over `docs/sheetdot-data-flow.svg` so the README stays in sync.
