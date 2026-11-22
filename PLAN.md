# Plan: Sheetdot, Excel as the control plane for personal agents

> [Clickable product mockup](docs/mockup.html)

## Context

A hackathon entry for **Equal AI x OpenAI, "Build Creative: personal context, into everyday AI products"**.
- The brief: a real user, a working path from raw input to a useful outcome, an original idea.

- **One-liner:** OpenAI Dots, but the interface is a spreadsheet. A control plane for personal agents.
  - Rows are what agents look after, columns are what they check, and schedules are when they check.
- Excel is the only interface. There is **no chat** anywhere. Everything is as Excel-native as possible.
- Agents keep working while Excel is closed. Schedules are the core feature.
- We don't build an agent harness. We sit on the **OpenAI Agents API** and use **Composio** for app connections.

### Why this, and not what exists

| Product | Shape | Our angle |
|---|---|---|
| ChatGPT / Claude / Copilot for Excel | A chat side panel; Excel is the LLM's notebook | No chat. The sheet *is* the instruction. |
| Autosheet (`=GPT`, `GPT_WEB`, `GPT_LIST`…) | Cell functions; the LLM is Excel's calculator; each call is a one-off | **Standing orders** that run on schedules, remember the last run, use personal sources, and take real actions |
| `=COPILOT()` | Retired 14 Sep 2026 | — |

**The test:** close Excel for a week. With every other product, nothing happens. With Sheetdot, the bill row turned red, the payment got matched, a new subscription row appeared, and the email to the LIC agent is waiting for your tick.

## Decisions already made

**Product**
- **Demo scenarios:**
  - **Family accounts:** EMIs, premiums, subscriptions, checked against bank and bill alert emails.
  - **Flat hunt:** the second act.
  - Positioned as a general agent control plane; these are just the demos. No user interviews.
- **Formulas:** `AI.ASK`, `AI.CHECK`, `AI.LIST`, and `AI.WATCH`.
- **Column suggestions:** Opt in once per workbook from the Sheetdot task pane. Selecting a populated Excel Table or a contiguous range with headers shows up to two grey suggestions in empty cells to its right. Tick a suggestion's checkbox to append a working column.
- **The workbook is the agent.** No named bots. One agent per workbook, configured by the 🙋 About me sheet.
- **The agent picks its own sources,** within what's connected.
- **What the agent sees in v0:** the whole sheet the formula is on, plus About me.
- **AI answers are saved.** Unchanged inputs reuse the saved answer. WATCH can update its result on a schedule.
- **Schedules in plain words** ("daily 8am", "every 1 min") in both `AI.WATCH` and ⏰ Automations. Editing the Schedule cell reschedules the job.
- **A WATCH can add rows,** but only to its own table.
- **Real actions only, approved by a human.** The agent proposes an action in the row (amber cell + native checkbox). Ticking the box runs it for real.
  - v0 actions: send a Gmail email, create a Calendar event, a real UPI payment link or QR. We never move money.
- **Unticked actions are re-decided on every WATCH run** (kept, replaced or cleared). A stale action never runs.
- **No inbox or "needs you" sheet. Colour is the signal** (see the colour language below).
- **Two special sheets:**
  - 🙋 **About me:** facts and preferences, plus a small Connections table.
  - ⏰ **Automations:** one row per WATCH.
- **Onboarding:**
  - Formulas work immediately with web search.
  - When a question needs Gmail, the cell says `🔌 Needs Gmail`. Ribbon **Connect** opens Composio's OAuth in an Office pop-up.
- **Distribution:**
  - Hackathon: sideload, plus a "try it yourself" page.
  - Later: Microsoft Marketplace.
- **Name: Sheetdot** (working name).

**Tech**
- **Engine:** OpenAI Agents API. No fallback engine.
- **One light, fast model for everything.** No model picker.
- **App access:** Composio. The agent gets **read-only** tools through MCP. Writes only happen through approved actions that our backend runs.
- **Hackathon scope:**
  - One user, our own keys, no sign-in, no limits or quotas, no local model.
  - Later, the user ID is the Microsoft account ID, through Office single sign-on.
- **Host:** Excel for the web first (free, personal Microsoft account). The same add-in works on desktop Excel later.
- **Sources for v0:** Gmail + Google Calendar. **Stretch:** bank SMS via a small Mac helper.
- **Demo:** real automations on an "every 1 min" schedule.

## The product, in one screen

| Excel thing | What it does in Sheetdot |
|---|---|
| `=AI.ASK(question, row)` | One value per row |
| `=AI.CHECK(condition, row)` | TRUE/FALSE, which drives Excel's own conditional formatting |
| `=AI.LIST(request, n)` | Fills in a table out of nothing |
| `=AI.WATCH(question, row, "daily 8am")` | A standing order. The formula is the schedule. |
| Grey suggestions beside a table | Tick one to add a column with Excel or AI formulas |
| Sheetdot task pane | Enable column suggestions and manage schedules |
| Excel **note** on a cell | Where the answer came from, and when |
| ⏰ Automations sheet | Every WATCH. Edit Schedule to change it, delete the row to stop it. |
| 🙋 About me sheet | The agent's memory and instructions, plus the Connections table |

**Colour language**

| Look | Means |
|---|---|
| ⏳ | Working |
| Excel note marker | Filled by an agent (hover to see the source) |
| Grey header, description and ☐ beside the table | Suggested column. Tick to add it and calculate the rows. |
| **Yellow fill** | Changed since you last looked. Clears when you select it. |
| **Amber cell + ☐** | The agent wants to do something. Tick to approve. |
| **Green** | Done ("✓ email sent 9:02") |
| Grey italic | `🔌 Needs Gmail` / `⚠ failed`; hover for why |
| Red / green rows | *Your* rules: Excel's own conditional formatting |

## Demo script (3 minutes)

1. Open "Family accounts" and enable column suggestions. Select the table and accept a suggested column. Connect Gmail for the payment-check formulas once that integration is available.
2. Use a TRUE/FALSE check to drive native conditional formatting for rows that need attention.
3. Enter an `AI.WATCH` formula with `"every 1 min"`. Its schedule appears in ⏰ Automations.
4. **Close the tab.** Wait for a real run. Reopen and show:
   - yellow cells: BESCOM is now paid, so its row is no longer red;
   - a **new Spotify row**;
   - amber **"📧 Email Ramesh (LIC) for the premium link ☐"**.
5. Tick it. The email really arrives on a phone on stage. The cell turns green.
6. **Second act, Flat hunt:** `=AI.LIST(...)` spills listings. Select the results to see suggested columns, then accept one from the grid.

### Try column suggestions locally

Open the workbook through `excel.cloud.microsoft` in Arc with the local Sheetdot add-in loaded. Paste this sample into **A1** on a blank sheet.

```tsv
Listing	Monthly rent	Deposit	Bedrooms
Maple Court	1500	2200	2
Oak Terrace	1850	900	3
Birch Place	1300	1500	1
```

1. Open the Sheetdot task pane and enable **Suggest columns in this workbook**.
2. Select **B2**. Grey suggestions should appear to the right, starting at **E1**. Exact names vary. In the live test, Sheetdot proposed **Rent per bedroom** and **Upfront cost**.
3. Before accepting, select **A10** and then **B2** again. The grey previews should clear and then return.
4. Tick a suggestion's checkbox. The chosen column becomes **E**, even if it was the second preview. If you picked **Upfront cost**, Maple Court should show `3,700`. If you picked **Rent per bedroom**, it should show `750`.

## Architecture

```
Excel for the web
  └─ Sheetdot add-in (Office.js, shared runtime, loads on document open)
       • custom functions: AI.ASK / AI.CHECK / AI.LIST / AI.WATCH (streaming)
       • column suggestions appear in the grid after workbook opt-in
       • on open: pulls patches (new rows, actions, changed marks) and applies them with Excel.run
       • watches checkbox changes → POST /actions/:id/approve
       │  HTTPS (polls every few seconds while open; no websockets)
       ▼
Sheetdot backend: Cloudflare Worker + one Durable Object per workbook (SQLite inside)
       • /ask      cache-first; on miss → model call with web search + Composio MCP (read-only)
       • /suggest-columns   propose columns from headers, sample rows and About me preferences
       • /watch    create / update / delete Agents API automations
       • /webhook  Agents API events → answer function calls from the mirror
       • /sync     add-in pushes the watched ranges + About me; backend keeps the copy
       • /patches  results waiting to be applied, plus "last seen" values for yellow
       • /actions  proposed actions; on approve → run through Composio (Gmail send, Calendar create)
       • /connect  makes a Composio OAuth link for the Office pop-up
       • /pay/:id  a UPI QR + link page for payment actions
       • /debug/:workbook  JSON dump of cache, automations, runs, patches, actions (for verification)
       │                                   │
       ▼                                   ▼
OpenAI                                    Composio
  per-cell answers: one model call          OAuth for Gmail + Google Calendar (user_id = "me" for now)
    with web search + MCP tools             read tools exposed as an MCP server
  WATCH: Agents API agent per workbook,     write tools called only by our backend after a tick
    automation "every 1 min" → new session,
    function tools: read_rows, write_cells, add_rows, propose_action
```

### How each feature works

- **Column suggestions**
  - Workbook opt-in allows Sheetdot to send the headers, up to five sample rows, and About me preferences to OpenAI. Selecting the table generates suggestions automatically.
  - A suggestion contains a short header, a description, and an accept checkbox. Previews stay outside the table and use only empty space. Existing data stays in place.
  - Acceptance appends the next real column and fills every data row. Exact calculations use native Excel formulas. Interpretation and public-web lookups use `AI.CHECK` or `AI.ASK`. Suggestions use available inputs and sources.
  - Unused previews clear on leaving the table and reappear from the session cache when you return. When you accept the second suggestion first, the remaining preview moves to the right of the new column.
- **`AI.ASK`**
  - The function sends (workbook, question, row values) to `/ask`.
  - A cache hit returns at once. A miss shows ⏳ (a streaming function), then one model call returns a typed value plus a source string.
  - The source string becomes an Excel note.
- **`AI.CHECK`:** uses the Decisions API to evaluate a condition against the referenced range and About me preferences. Returns native Excel TRUE when the supplied evidence supports the condition, and FALSE when the evidence contradicts it or is insufficient. Checks that depend on web or Gmail results need those results as input.
- **`AI.LIST`:** uses the Responses API with web search and returns a header row followed by up to the requested number of data rows (1–50). The result spills like any native array formula.
- **`AI.WATCH`:**
  - First evaluation: it creates one Agents API automation (key = workbook + cell range + question), runs it once now, and adds a row to ⏰ Automations.
  - Each run starts a new session with the workbook's agent. The agent calls:
    - `read_rows`: our backend answers from its copy, including the last-run values;
    - `write_cells` / `add_rows` / `propose_action`: our backend stores them as patches.
  - Deleting the formula, or the Automations row, deletes the automation.
- **Excel closed:**
  - Patches wait in the Durable Object.
  - On open, the custom functions re-evaluate and get the latest cached values. The shared runtime then applies new rows, action cells and yellow marks.
- **Actions:**
  - `propose_action` creates `{id, row, kind, preview, args}`.
  - The add-in writes the amber text cell and a checkbox cell.
  - A tick sends `POST /actions/:id/approve`. The backend runs it through Composio, and the cell turns green with the time.
  - The next WATCH run sees open actions and may keep, replace or clear them.
- **Yellow "changed":** the backend remembers what each cell showed when you last looked. A different value gets a yellow fill. Selecting the cell clears it, and the add-in reports "seen".
- **Why both a direct model call and the Agents API:** a WATCH runs on a schedule and needs function calls and sessions. A fill-down of 50 rows needs 50 *parallel*, fast, stateless answers. A single Agents API session handles one turn at a time, so per-cell answers use direct model calls with the same model and tools. Both are OpenAI, and neither is our own harness.

### Repo layout (proposed)

```
sheetdot/
  PLAN.md
  docs/mockup.html
  addin/      Office add-in (Yeoman "Excel Custom Functions, shared runtime", TypeScript)
    src/functions/   AI.* custom functions
    public/suggestions.js   in-grid column previews and acceptance
    src/core/        pure logic (patch apply, colour rules): no Office.js imports
    src/excel/       thin Excel.run adapter
  backend/    Cloudflare Worker + Durable Object (TypeScript, Hono)
    src/schedule.ts  plain words → cron
    src/agent.ts     Agents API: agent, automations, webhook handling, function results
    src/answer.ts    per-cell model call
    src/composio.ts  connect links, MCP URL, action execution
  scripts/    verification scripts (below)
  templates/  Family accounts + Flat hunt workbooks (built by a script)
```

## Work, in order of risk

### 0. The free-Excel test (do first, about 15 minutes)

- Scaffold the add-in and sideload it into a blank workbook on excel.com (personal account).
- Confirm:
  1. a custom function returns a value, a streaming value, and a spilled array;
  2. the shared runtime loads on open and the ribbon button runs;
  3. Excel.run can add a **note**, set a fill, add a **conditional format**, and insert a **checkbox**, with a TRUE/FALSE dropdown as the fallback;
  4. a checkbox change fires an event in the add-in;
  5. an Office pop-up dialog opens an outside URL.
- **If any of these fail, we change the product before building.** For example: notes fail → use threaded comments; checkbox fails → use the dropdown.

### 1. OpenAI + Composio test (in parallel with 0)

- Script: create an agent, run one turn with web search, and create an "every 1 min" automation. Receive the webhook at the deployed worker and answer one function call.
- Composio: connect Gmail for `user_id=me`. Read the last 5 bank alert emails through MCP from one model call. Send one test email through the backend.

### 2. Backend core

- Durable Object per workbook: cache, mirror, automations, patches, actions, last-seen values.
- `/ask` with typed outputs (string / number / boolean / 2-D array) plus a source string.
- `/suggest-columns` with validated column recipes for a small worksheet sample.
- Plain words → cron. Automation create, update and delete. Webhook handler with signature check.
- Composio connect link, MCP config, action runner. `/pay/:id` page.
- `/debug/:workbook`.

### 3. Add-in core

- The four functions, in-grid column suggestions, and Connect.
- Notes, colour language, checkbox → approve, yellow clears on select.
- The Automations sheet kept in sync both ways: edit Schedule → reschedule; delete row → delete automation.
- About me sheet sync, including the Connections table.

### 4. Templates and demo data

- A script that builds "Family accounts" and "Flat hunt" (with its conditional formatting rules) in any workbook.
- Demo Gmail with real alert and bill emails (see "What you need to supply").

### 5. Stretch

- Mac SMS helper: reads new bank SMS from Messages' local database (`chat.db`) and posts them to the backend as a source.
- "Try it yourself" page with the manifest and a 4-step guide.

## Verification

The aim: **I can check almost everything myself, without you clicking in Excel.** You only do the final pass in real Excel.

| # | What | How | Who runs it |
|---|---|---|---|
| V1 | Pure logic | Vitest unit tests: schedule parsing, cache keys, patch merge, the action state machine (proposed → approved → done / cleared), suggestion recipes, colour rules | Me, on every change |
| V2 | Add-in against a fake Excel | Microsoft's `office-addin-mock` for the `Excel.run` adapter: assert that notes, fills, conditional formats and checkboxes are requested correctly | Me |
| V3 | Real OpenAI + Composio | `scripts/verify-agents.ts`: one ASK with web search, one Gmail read, one automation fired by webhook → patch stored. Uses real keys. | Me (needs your keys) |
| V4 | Full flow with no Excel | `scripts/fake-workbook.ts` plays the add-in: register a workbook → sync rows → ASK → create a one-minute WATCH → wait for a real run → pull patches → approve an action → check the email was sent. It prints a pass/fail checklist. | Me |
| V5 | Backend state | `GET /debug/:workbook` shows the cache, automations (with next run), runs, patches, actions and last-seen values. Every script asserts against it. | Me |
| V6 | Real Excel | Follow the demo script in Excel for the web. Check workbook opt-in, preview cleanup and return, and acceptance of the second suggestion first on both a plain range and a native Excel Table. | Either of us in a signed-in Excel session |
| V7 | Demo rehearsal | Fresh browser, sideload, template, demo script with the tab closed during the 1-minute run. Twice in a row without a failure. | Both |

Definition of done for the hackathon: V1–V5 green, V6 passes in real Excel, V7 passes twice.

## What you need to supply

| # | What | Why | Notes |
|---|---|---|---|
| 1 | **Microsoft account** (personal) signed in at excel.com, plus a blank workbook in OneDrive | Sideloading and the real-Excel tests | Any outlook.com / hotmail / Gmail-backed Microsoft account |
| 2 | **OpenAI API key** with Agents API access (beta, `OpenAI-Beta: agents=v1`) and some credits | Engine | Check whether hackathon credits come with it. Put it in `sheetdot/backend/.dev.vars` (I'll add it to `.gitignore`); never paste it in chat. |
| 3 | **OpenAI webhook signing secret** for our worker URL | Automations → our backend | Created in the OpenAI dashboard once the worker URL exists (after step 2). I'll tell you the URL. |
| 4 | **Composio account + API key**, with Gmail and Google Calendar enabled (Composio's managed OAuth app is fine for now) | Connections and actions | Same `.dev.vars` file |
| 5 | **A Gmail account for the demo** with real bank-alert and bill emails | The personal source | Recommended: a fresh demo Gmail that you forward real HDFC/LIC/BESCOM/Netflix alerts into, so your real inbox isn't on stage. Your call. |
| 6 | **Cloudflare account** (free) | Hosting the backend | You'll run `npx wrangler login` once; it opens the browser. |
| 7 | **Household rows for the template** (expense, amount, due day) | Family accounts sheet | Real or realistic numbers; you choose what's shown on stage |
| 8 | **A UPI ID (VPA) to receive the demo payment** | The UPI action | Your own is fine; the demo amount can be ₹1 |
| 9 | **A phone on stage** showing the demo Gmail | The "email really arrived" moment | — |
| 10 | (Stretch) **Mac with iPhone Text Message Forwarding on**, plus Full Disk Access for the helper | Bank SMS source | Only if we reach the stretch |

Already available on this machine: Node 24, npm 12. Installed on first use, no action needed: Yeoman Office generator, wrangler, office-addin-dev-certs.

## Risks

| Risk | If it happens |
|---|---|
| Free Excel web blocks sideloading or custom functions on personal accounts | Test 0 finds it on day one. Fallback: a Microsoft 365 trial account. |
| No Agents API access on our key | Ask the organisers. WATCH can't be faked; this blocks the core feature. |
| Notes, checkbox or conditional-format APIs missing on Excel web | Swap in comments, a TRUE/FALSE dropdown, or a fill colour set by us |
| Indian bank alert formats vary | Demo with the banks in your demo Gmail; the model reads free text, so there are no parsers to write |
| A slow model on stage | Open the demo table once to generate suggestions, then use the cached suggestions during the demo. Run WATCH live. |
| An email sent to the wrong person | Actions only send to addresses in About me or the original thread. The amber preview shows the recipient. |

## Later (not for the hackathon)

- Office single sign-on → Microsoft account ID as the user ID; multi-user Composio.
- Limiting what the agent sees to exactly the referenced ranges.
- Marketplace listing; desktop Excel; more sources (SMS on Android, Drive, WhatsApp); more action kinds.
