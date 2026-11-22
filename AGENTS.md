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

## Just want to play with it?

You don't need keys, a clone or any setup. A copy is already running at `https://sheet-dot.songguesser.workers.dev`.

1. Follow [Load the add-in](#load-the-add-in-in-excel-for-the-web) with `addin/manifest.hosted.xml`. The README's [Try it now](README.md#try-it-now) has the same steps, written for non-developers.
2. In the Sheetdot task pane, choose **Add demo sheets**. You get **Weekend Plans**, **Find a Flat** and **Family Bills**, plus a **Demo Guide** sheet that walks through them.
3. Try the checks in [Check it works](#check-it-works), then write your own formulas on any table.

The hosted copy runs on the maintainers' keys with no sign-in or limits. It's meant for trying things out, not for real data you care about. Set up your own copy (below) only if you want to change the code or use your own keys.

## Set up your own copy

Any person or coding agent can follow these steps in order. Steps marked **🧑 human** need a browser login or a dashboard, so an agent should stop there and ask the user to do them.

### What you need

| Need | Used for | Without it |
|---|---|---|
| Node 24 and npm | Running and deploying the backend | Nothing runs |
| OpenAI API key with access to the models in `backend/ask.ts`, the Decisions API, and the Agents API beta (`OpenAI-Beta: agents=v1`) | Every formula | ASK and LIST need only the Responses API. CHECK fails without Decisions. WATCH fails without the Agents API. |
| Composio API key and a Gmail auth config | Bill tracking | Formulas still work; **Track bills** fails |
| Cloudflare account (free tier works) | Hosting | You can still run locally |
| Personal Microsoft account | Excel for the web | Nothing to load the add-in into |

Check model and API access first. It's the step most likely to block you.

### Step 1. Get the keys (🧑 human)

Follow [docs/env-setup.md](docs/env-setup.md) for click-by-click steps with links to each provider's official docs. In short:

1. Create an OpenAI API key at platform.openai.com/api-keys.
2. In the Composio dashboard, copy the project API key. Then create a **Gmail** auth config that requests only the `https://www.googleapis.com/auth/gmail.readonly` scope, and copy its ID (it usually starts with `ac_`).
3. Copy the example file in the repo root and fill in the three values. `.env` is gitignored; never commit it.

```bash
cp .env.example .env
```

### Step 2. Run locally

Office add-ins must load over HTTPS. Install the Office dev certificate once (it asks for your password to trust it):

```bash
npx office-addin-dev-certs install
```

Start the backend. It serves the API and the add-in files at `https://localhost:3000`:

```bash
cd backend && npm install && npm run dev
```

`https://localhost:3000/health` should return OK.

### Step 3. Deploy (optional)

Log in once (🧑 human, opens a browser):

```bash
cd backend && npx wrangler login
```

Then deploy and add the keys as secrets:

```bash
cd backend && npx wrangler deploy
```

```bash
cd backend && npx wrangler secret put OPENAI_API_KEY
```

```bash
cd backend && npx wrangler secret put COMPOSIO_API_KEY
```

```bash
cd backend && npx wrangler secret put COMPOSIO_GMAIL_AUTH_CONFIG_ID
```

`wrangler deploy` prints your URL, for example `https://sheet-dot.<your-subdomain>.workers.dev`. Make a manifest for it from the local one, with a new add-in ID so it doesn't clash with ours:

```bash
sed -e "s#https://localhost:3000#https://sheet-dot.YOUR-SUBDOMAIN.workers.dev#g" -e "s#<Id>[^<]*</Id>#<Id>$(uuidgen | tr A-Z a-z)</Id>#" addin/manifest.xml > addin/manifest.mine.xml
```

Watch live logs with `npx wrangler tail`.

### Load the add-in in Excel for the web

All of this is 🧑 human: it needs a browser and a Microsoft sign-in. An agent should give the user these steps and the path of the manifest to upload.

1. Go to [excel.cloud.microsoft](https://excel.cloud.microsoft) in Chrome, Edge or Safari.
2. Sign in with a personal Microsoft account (Outlook.com, Hotmail, or one created with any email address via **Create one!**). Excel for the web is free.
3. Choose **Blank workbook**.
4. On the **Home** tab, click **Add-ins**, then **More Add-ins**.
5. Go to the **My Add-ins** tab and click **Upload My Add-in**.
6. Click **Browse** and pick the manifest, then **Upload**:
   - `addin/manifest.hosted.xml` for the hosted copy,
   - `addin/manifest.xml` while `npm run dev` is running locally,
   - `addin/manifest.mine.xml` for your own deployment.
7. Click **Open Sheetdot** in the **Sheetdot** group at the right end of the **Home** tab.

An uploaded add-in may need uploading again in another workbook or after clearing browser data. If the button doesn't appear, refresh the page and upload again. For the local manifest, the browser must trust `https://localhost:3000`. If the pane is blank, open that URL in a tab once to check.

### Check it works

- Type `=AI.ASK("What is the capital of Karnataka?")` in a cell. You should see ⏳, then an answer with a purple source comment.
- Type `=AI.CHECK("Bengaluru is in Karnataka")`. You should get `TRUE`. This confirms Decisions API access.
- Type `=AI.WATCH("What time is it in Tokyo?", A1, "every 1 min")`. A row should appear on the **⏰ Automations** sheet. This confirms Agents API access.
- In the task pane, choose **Add demo sheets**. On **Family Bills**, select the table, click **Track bills**, connect a Gmail account that has a real bill and its receipt (🧑 human), and pick **every 1 min**. Within a couple of minutes the bill row should appear, then turn **Paid**.

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
