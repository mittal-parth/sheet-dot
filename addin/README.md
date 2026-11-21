# Sheetdot Excel compatibility check

This add-in exercises the Excel features needed for Sheetdot. It creates a worksheet with test formulas and controls and can run without API keys.

## Build and serve

Use Node.js 24 or later. Build from the `addin` directory.

```sh
npm ci
npm run build
npm test
```

The build writes `dist/manifest.xml` for `https://localhost:3000`. Serve it with a certificate trusted by the browser running Excel. Microsoft's Office Add-in developer certificates are one way to set up localhost HTTPS.

```sh
npx office-addin-dev-certs install
SHEETDOT_TLS_CERT="$HOME/.office-addin-dev-certs/localhost.crt" \
SHEETDOT_TLS_KEY="$HOME/.office-addin-dev-certs/localhost.key" npm run serve
```

If using an existing HTTPS origin, run `SHEETDOT_PUBLIC_URL=https://your-host.example npm run build` and serve `dist` at that origin. `npm run serve` listens on port 3000 by default and serves HTTP when TLS files are omitted, which works behind an HTTPS reverse proxy. The same origin must serve the runtime, functions metadata, icons, and dialog.

## Check in Excel for the web

1. Open a blank workbook in Excel for the web with a personal Microsoft account. Under **Home → Add-ins → More Settings → Upload My Add-in**, choose `dist/manifest.xml`. [Microsoft sideload instructions](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#manually-sideload-an-add-in-to-office-on-the-web)
2. Select **Home → Sheetdot → Run checks**. Verify the **Sheetdot Compatibility** worksheet opens.
3. In B2, verify `Value returned`. In B3, verify `Stream finished` after `Working…`. In B4:C5, verify a two-by-two result. Re-enter the formula in B3 to see the streaming state again.
4. Open the note in B6 and verify it contains `Source: Sheetdot compatibility test`. If Excel reports that notes are unavailable, record that result. In B7, verify yellow fill, then enter `2` to see the conditional format turn it green.
5. Toggle the checkbox in B8 and check that B9 reports the new value. If the worksheet reports that the checkbox API is unavailable, use the TRUE/FALSE dropdown and check the same result.
6. Once B10 says startup is enabled, clear the formulas in B2:B4 so formula recalculation cannot start the runtime. Note the time in B11, then save, close, and reopen the workbook. B11 should update before you select a ribbon command or edit the worksheet.
7. Select **Home → Sheetdot → Test dialog**, then **Open test dialog**. Verify the pop-up opens and choose **Open Microsoft Learn** inside it. Verify the same pop-up navigates to Microsoft Learn.

## OpenAI automation check

From the repository root, with `OPENAI_API_KEY` available in the shell, check the project's Agents API access. Pass `--schedule` to create a temporary one-minute automation. It waits for an actual scheduled delivery and removes the recurring schedule as soon as that delivery is recorded. It then checks for a completed web search and answer. If the process is forcibly stopped before cleanup, remove the named test schedule from the Agents API.

```sh
node --use-env-proxy scripts/verify-agents.mjs
node --use-env-proxy scripts/verify-agents.mjs --schedule
```
