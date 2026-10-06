# Grab Receipts Matcher

A static site with no build step and no backend. Every file is processed in the browser.

1. **Grab for Business exports**: add every category file (Transport, Food...). They stack in a list, and you can remove one with ×. The category comes from the file name.
2. **Brex accounting export (CSV)**: columns are detected automatically and can be adjusted under "Brex column mapping". That panel also has the card filter, **only card ending 1521** by default.
3. **Match & generate**, then download:
   - `Matched_Receipts_YYYY-MM-DD.xlsx`: one row per Brex charge, with Status, Match Method, Date, Grab Category, Booking ID, Receipt PDF, Integration ERP (QB), Brex Dashboard, amounts, Location, USD, **Memo**, Current Brex Memo, Name Check, Note
   - `Grab_Receipts_YYYY-MM-DD.zip`: one PDF per trip or order (ride and food layouts)

## Results table
- **Copy** buttons in the Memo, Location, Class (`travel`) and GL Account cells (`TR03 - Transportation` for rides, `TR04 - Meals` for food), for pasting into Brex.
- **PDF** downloads that row's receipt. **open** links go to QuickBooks (Integration ERP) and Brex.
- The status chips at the top filter the table.
- **Uploaded** checkbox (first column, stays visible while scrolling): tick each Brex charge once its receipt and memo are in Brex. Ticked rows turn light green, the bar shows "N of M uploaded", and **Hide uploaded** leaves only what's left. Ticks are remembered by Brex expense ID, even across months.
- **Save session** stores the files, settings and ticks in this browser. Reopen it from **Saved sessions** at the top, where you can also **Export** it to a file (to back up or share) and **Import** it on another computer.

## Memo format
- `Name, Grab Transpo` or `Name, Grab Food`. The name always comes from the Grab report (Januario → Ino, Thareq → Abel).
- When a trip is split across several Brex charges, each charge gets its own row and memo: `… 1/2`, `… 2/2`.
- A refund line gets `… refund`.

## Matching (`match.js`)
Brex lines that share the booking ID in their **Statement Descriptor** (`Grab* A-…`) count as one charge. The matcher then tries these in order:
1. **Booking ID**: the charge's ID equals the trip ID. If the sum is short, ID-less top-ups charged within 3 hours are added.
2. **Amount + time**: exact amount, charged within 30 minutes of the trip.
3. **Anchored split**: one charge near the trip time plus 1–3 top-ups within 3 hours add up to the fare.
4. **Amount + date** (same person), then **combined** charges, then **near amount** (within 3%), which is marked for review.
5. Refunds are attached to the order they belong to. Anything left over is labelled: Outside Brex range, Other card, Grab Food/Grab-ec, Not in Grab report, or Unmatched.

## Files
- `index.html`, `style.css`: the page layout and styling
- `match.js`: parsing and matching logic, with no DOM code
- `receipt.js`: draws the receipt PDFs with jsPDF
- `app.js`: connects file uploads, the column mapping, the results table and the downloads
- `samples/`: fictional example exports in the real formats (Grab Transport, Grab Food, Brex). Drop all three in to see every case: booking-ID match, split charge, ride with no booking ID, food order, refund, other card, and outside the Brex range

## Run locally
```
cd ~/Downloads/grab-receipts-app && python3 -m http.server 5173
```
Then open http://127.0.0.1:5173. You need to be online, because the libraries load from cdnjs.

## Deploy
Static site, no build step: on Vercel, import the GitHub repo with the framework preset **Other** and no build command. Each push redeploys it.

## Where data is stored
Nothing is uploaded anywhere. Saved sessions and upload ticks live in the browser (IndexedDB, with localStorage as a fallback), per browser and per site address. Use **Export / Import** in "Saved sessions" to move them between browsers, or from the local site to the deployed one.
