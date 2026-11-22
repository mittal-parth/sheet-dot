# Getting the environment variables

Sheetdot needs three values in `.env` (copied from [.env.example](../.env.example)). This page shows where each one comes from. Dashboards change often, so if a label here doesn't match, follow the linked official page.

| Variable | Provider | Needed for |
|---|---|---|
| `OPENAI_API_KEY` | OpenAI | Every formula |
| `COMPOSIO_API_KEY` | Composio | Track bills (Gmail) |
| `COMPOSIO_GMAIL_AUTH_CONFIG_ID` | Composio | Track bills (Gmail) |

If you only want to try formulas, you can skip both Composio steps.

## 1. `OPENAI_API_KEY`

Official guide: [OpenAI developer quickstart](https://developers.openai.com/api/docs/quickstart).

1. Sign in at [platform.openai.com](https://platform.openai.com), or create an account.
2. Add credits at [Billing](https://platform.openai.com/account/billing/overview). Sheetdot runs many model calls (each row, each scheduled run), and free-tier limits are too low.
3. Open [API keys](https://platform.openai.com/api-keys) and create a new secret key. Pick the project it belongs to and keep permissions at **All**.
4. Copy the key straight away. It starts with `sk-`, and OpenAI shows it only once.
5. Paste it into `.env`:

   ```
   OPENAI_API_KEY=sk-...
   ```

**Check your access.** Sheetdot uses the models named at the top of [backend/ask.ts](../backend/ask.ts), the Decisions API (for `AI.CHECK`) and the Agents API beta (for `AI.WATCH`). Not every account has all of these. After setup, run the three formula checks in [AGENTS.md → Check it works](../AGENTS.md#check-it-works). Each one confirms one piece of access.

## 2. `COMPOSIO_API_KEY`

Official guide: [Authenticating to Composio](https://docs.composio.dev/reference/authenticating-to-composio).

1. Sign up or log in at [dashboard.composio.dev](https://dashboard.composio.dev).
2. Select a project, or create one (for example, "Sheetdot").
3. Open **Project Settings** and go to the **API Keys** tab.
4. Copy the **project** API key. Don't use an organization access token: Sheetdot sends the key as `x-api-key`, which needs a project key.
5. Paste it into `.env`:

   ```
   COMPOSIO_API_KEY=...
   ```

## 3. `COMPOSIO_GMAIL_AUTH_CONFIG_ID`

Official guide: [Authenticating tools](https://docs.composio.dev/docs/authenticating-tools).

An auth config tells Composio how Sheetdot asks users for Gmail access. Create it in the same project as the API key.

1. In the dashboard, open **Auth Configs** ([direct link](https://dashboard.composio.dev/~/project/auth-configs)).
2. Click **Create Auth Config** and choose **Gmail**.
3. Choose **OAuth** as the authentication method.
4. Keep **Composio managed** authentication. It needs no Google Cloud setup and is fine for trying Sheetdot. Use your own Google OAuth app only if you need your own branding or higher limits.
5. Under scopes, keep only **`https://www.googleapis.com/auth/gmail.readonly`** and remove the rest. Sheetdot only reads mail. It never sends, deletes or labels anything.
6. Click **Create Auth Configuration**.
7. Copy the auth config's ID from the Auth Configs list. It usually starts with `ac_`.
8. Paste it into `.env`:

   ```
   COMPOSIO_GMAIL_AUTH_CONFIG_ID=ac_...
   ```

You don't connect a Gmail account here. Each user connects their own Gmail later from Excel by clicking **Connect Gmail** in the Sheetdot pane.

## 4. Use the values

- **Running locally:** `npm run dev` in `backend/` reads the root `.env` automatically.
- **Deployed:** store each value as a Worker secret. Run each command from `backend/`, and paste the value when asked:

  ```bash
  npx wrangler secret put OPENAI_API_KEY
  ```

  ```bash
  npx wrangler secret put COMPOSIO_API_KEY
  ```

  ```bash
  npx wrangler secret put COMPOSIO_GMAIL_AUTH_CONFIG_ID
  ```

Never commit `.env` or paste keys into issues, chats or screenshots. If a key leaks, revoke it in the same dashboard and create a new one.
