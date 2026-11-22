# Sheetdot

**Excel as the control plane for always-on personal agents.**

## Brief

In India, and especially among Indian dads, people love their good old Excel. From quick calculations to monthly expenses, flights, job hunts and life savings, there is an Excel sheet for everything. It has long been the richest source of personal context.

Sheetdot brings always-running personal agents right inside Excel, connected to your favourite tools. There is no chat. Everything happens through Excel formulas, and they keep running even when Excel is closed. Imagine OpenAI Dots, but used natively through Excel.

## Data flow diagram

![How data moves through Sheetdot](docs/sheetdot-data-flow.svg)

Follow the numbers: the red path is an AI formula and the green path is a bill tracked from Gmail. The editable source is [docs/sheetdot-data-flow.excalidraw](docs/sheetdot-data-flow.excalidraw).

## FAQs

**1. Why not connect Claude (or ChatGPT, or Copilot) to Excel?**

Those are chat panels beside the sheet. You ask, they edit the sheet, and then they stop. The instruction lives in a chat history, and nothing happens once you close Excel. In Sheetdot the instruction is a formula saved in the cell. It re-runs on its own schedule and reads your own sources, like Gmail.

**2. Why not use plugins like [Autosheet](https://autosheet.com/)?**

Functions like `=GPT()` work like a calculator: each call answers once, from the model or the web. They don't run on a schedule, don't remember the last run, can't read your personal sources, and do nothing while Excel is closed.

**3. So how is Sheetdot different?**

- **Standing orders, not one-off answers.** `=AI.WATCH(..., "daily 8am")` keeps checking and updates the cell when the answer changes.
- **Runs with Excel closed.** Agents run in the cloud. Open the workbook and the changes are waiting, highlighted in yellow.
- **Personal sources.** Connect Gmail, and Sheetdot can add a row for a new bill and mark it Paid when the receipt arrives.
- **Excel-native.** No chat and no new app to learn. Results are cells, rows, comments that cite their source, and your own conditional formatting.

## Try it now

Sheetdot is already deployed at **https://sheet-dot.songguesser.workers.dev**. You need a browser and a free Microsoft account. No keys, no install, no code.

**Get the add-in file**

1. Open [addin/manifest.hosted.xml](addin/manifest.hosted.xml) on GitHub.
2. Click the **Download raw file** button (the down-arrow icon above the file). Save `manifest.hosted.xml` somewhere you can find it.

**Open Excel**

3. Go to [excel.cloud.microsoft](https://excel.cloud.microsoft) in Chrome, Edge or Safari.
4. Sign in with a Microsoft account. Any personal account works (Outlook.com, Hotmail, or one you create with your Gmail address). If you don't have one, choose **Create one!** on the sign-in page. Excel for the web is free.
5. Choose **Blank workbook**.

**Add Sheetdot to Excel**

6. On the **Home** tab, click **Add-ins**, then **More Add-ins** at the bottom of the menu.
7. In the window that opens, go to the **My Add-ins** tab and click **Upload My Add-in** (top right).
8. Click **Browse**, pick `manifest.hosted.xml`, then click **Upload**.
9. A **Sheetdot** group with an **Open Sheetdot** button appears on the right of the **Home** tab. Click it to open the Sheetdot pane.

**Play**

10. In the Sheetdot pane, click **Add demo sheets**. You get **Weekend Plans**, **Find a Flat** and **Family Bills**, plus a **Demo Guide** sheet that walks you through them.
11. Or type a formula into any cell, for example `=AI.ASK("What is the capital of Karnataka?")`. You'll see ⏳ for a few seconds, then the answer.

If the **Open Sheetdot** button is missing in another workbook, or after you clear your browser data, repeat steps 6–8. If it doesn't appear at all, refresh the page and try again.

The hosted copy has no sign-in or limits and runs on the maintainers' keys. Use it to try things, not for data you care about.

## Three examples

### 1. Track household bills from Gmail

You have a **Family Bills** table with *Merchant*, *Amount due* and *Due date*. Select it, click **Track bills** in the Sheetdot task pane, connect Gmail (read-only) and choose **daily 8am**.

- A December BESCOM bill lands in your inbox. At 8am, with Excel closed, Sheetdot reads it and adds a row: `BESCOM · ₹1,842 · due 12 Dec`.
- Two days later the payment receipt arrives. The same row turns **Paid**.
- When you next open Excel, both changes are highlighted in yellow, and each cell's comment links to the email it came from.

### 2. Ask a question of every row

On a **Find a Flat** table, type this in a new column and fill it down:

```
=AI.ASK("Does this listing allow pets?", A2:D2)
```

Each row gets its own answer, with the source in a purple comment. Use `AI.CHECK` when you want TRUE/FALSE that Excel's conditional formatting can colour:

```
=AI.CHECK("Rent is under ₹35,000 and it has 2+ bedrooms", A2:D2)
```

Select the table and Sheetdot also suggests columns in grey to the right, such as **Rent per bedroom**. Tick one to add it as a real column.

### 3. A standing order that re-runs on its own

```
=AI.WATCH("Will it rain on Saturday at this venue?", A2:C2, "daily 8am")
```

The formula creates an OpenAI-hosted agent that re-checks every morning. The **⏰ Automations** sheet lists it; edit its *Schedule* cell to change when it runs, or delete the row to stop it.

## The formulas

| Formula | What you get |
|---|---|
| `=AI.ASK(question, row)` | One value per row |
| `=AI.CHECK(condition, row)` | TRUE or FALSE |
| `=AI.LIST(request, n)` | A new table of up to *n* rows, spilled into the sheet |
| `=AI.WATCH(question, row, schedule)` | `AI.ASK`, re-run on a schedule like `"daily 8am"` or `"every 1 min"` |

Preferences you write on the **🙋 About me** sheet are given to every answer.

## Run it yourself

To run your own copy instead of the hosted one, see [AGENTS.md](AGENTS.md) for setup from scratch: keys, local development, sideloading into Excel for the web, and deploying. The product plan and decisions are in [PLAN.md](PLAN.md).
