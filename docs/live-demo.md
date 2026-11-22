# Sheetdot live demo

Open **Book 10** in Excel for the web. Start on **Gmail Bill Test**, then switch to **Live Demo India** while Gmail checks for the receipt.

## 1. Set up Gmail and track one bill

Before recording, pick a real bill and its matching payment receipt. In Gmail, create the label `sheetdot-demo` and a filter that applies it to a distinctive subject phrase such as `Sheetdot demo`. Forward the bill with that phrase. Keep the receipt out of the label until the recording.

1. On **Gmail Bill Test**, enter the real merchant, amount, and due date in the existing bill table. Use the merchant name from the bill and leave its payment status empty.
2. Open Sheetdot and reuse the existing Gmail connection in Book 10. If the table is already being tracked on the right label, use that schedule. To start tracking, select a cell in the bill table, set the Gmail label to `sheetdot-demo` and the schedule to `every 1 min`, and click **Track bills in selected table**.
3. Show **No payment evidence found**, then forward the real receipt so the filter labels it. Switch to **Live Demo India**.

If the **Open Sheetdot** button is missing and formulas show `#NAME?`, use the browser profile where Sheetdot was uploaded. To register it again, use **Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in** and select [the hosted manifest](../addin/manifest.hosted.xml). Confirm that typing `=AI.` shows the available formulas before recording.

## 2. Start the trip research

In **K4**, type this on the **Live Demo India** sheet:

```excel
=AI.WATCH("Search the web for direct trains from Hyderabad to Pune Junction. Compare the supplied departure stations using published timetables. Append up to three verified trains as rows with Departure station, Code, Train, Departs / arrives and Source. Include train number and published local times. Skip options you cannot verify.",G4:G6,"in 1 min")
```

Show the queued run in **Automations**. The schedule starts one run in about a minute. The job starts in the background while you continue.

If Excel needs the trip table rebuilt, enter the headers `Departure station`, `Code`, `Train`, `Departs / arrives`, `Source`, `Research` in **F3:K3**. Enter **Hyderabad Deccan / HYB**, **Secunderabad / SC**, and **Kacheguda / KCG** in **F4:G6**. Select **F3:K6**, choose **Insert → Table**, and leave **My table has headers** selected.

## 3. Look up a closing price and fill down

In **D4**, type:

```excel
=AI.ASK("Last Nasdaq closing price in USD for this ticker. Include the trading date.",C4)
```

Show Apple’s result, then drag the fill handle down through **D8**. Each row uses its own ticker. The date identifies the trading session on a weekend.

To flag stocks that closed at $300 or above, type this in **E4** after ASK returns a price. You can fill it down through **E8**:

```excel
=AI.CHECK("The closing price in this cell is at least 300 USD. Ignore the trading date.",D4)
```

## 4. List live trades, then watch the same market

In **B12**, type:

```excel
=AI.LIST("Search Coinbase Exchange for the five latest trades for BTC-USD. List trade time in IST with seconds, and price in USD, newest first.",5)
```

Show the price and time spilling into rows. Edit **B12** in place:

```excel
=AI.WATCH("Search the web for the latest trade for BTC-USD on Coinbase Exchange. Return the price in USD and the trade time in IST with seconds on one line.",C11:D11,"every 1 min")
```

Open **Automations**, point to **Live Demo India!B12** and **Last run**, and return to **B12**. The newest trade’s timestamp makes the refresh visible even if two trades have the same price. The cell now shows the latest trade in one line.

## 5. Show the train rows and the payment

Return to the trip table. The WATCH researched more than one station and added rows with train numbers, times, and sources. This is the point to show that the job kept working while you moved on, then wrote its results into the existing Excel table.

Switch back to **Gmail Bill Test**. Show the receipt in Gmail, then **Paid** on the tracked row and the matching receipt in the cell’s comment. Allow another check if the receipt arrived just after a run. Stop the demo schedules in **Automations** once you finish recording.

Sources: [Nasdaq](https://www.nasdaq.com/market-activity/stocks), [Coinbase Exchange trades](https://docs.cdp.coinbase.com/api-reference/exchange-api/rest-api/products/get-product-trades), and [Indian Railways](https://www.indianrail.gov.in/mail_express_trn_list.html).
