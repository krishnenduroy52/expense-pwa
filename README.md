# Expense PWA

A client-side personal expense dashboard for importing bank statements, identifying debit transactions, assigning merchant categories, and reviewing spending summaries. The app is built with React and Vite and can be installed as a Progressive Web App (PWA).

## Contents

- [Highlights](#highlights)
- [Technology](#technology)
- [Requirements](#requirements)
- [Run locally](#run-locally)
- [Production build and preview](#production-build-and-preview)
- [Using the dashboard](#using-the-dashboard)
- [Importing a statement](#importing-a-statement)
- [Supported columns](#supported-columns)
- [How transaction classification works](#how-transaction-classification-works)
- [Local data and privacy](#local-data-and-privacy)
- [Progressive Web App behavior](#progressive-web-app-behavior)
- [Project structure](#project-structure)
- [Available scripts](#available-scripts)
- [Known limitations](#known-limitations)

## Highlights

- Import comma-separated CSV files and Excel `.xls` or `.xlsx` workbooks.
- Read statement headings even when introductory rows appear above the table, as long as the app can recognize the header row.
- Show total spend, the largest spending category, unresolved entries, a category bar chart, and a spending-share donut chart.
- Categorize recognized merchants using built-in keyword rules.
- Resolve unknown merchants, add custom keyword rules, and change the category of an individual transaction.
- Display the transaction date, optional value date, reference, debit amount, and available balance.
- Store transactions and custom rules in this browser using `localStorage`.
- Use responsive layouts for desktop and mobile screens.

## Technology

- **React 19** for the user interface.
- **Vite 8** for local development and production builds.
- **SheetJS (`xlsx`)** for reading Excel workbooks in the browser.
- **Oxlint** for linting.
- Browser APIs including `FileReader`, `localStorage`, Cache Storage, and Service Workers.

There is no server or database component in this project. Statement parsing, categorization, and storage happen in the browser.

## Requirements

- Node.js and npm versions compatible with the Vite version declared in `package.json`.
- A modern browser with JavaScript enabled. PWA installation and offline caching require Service Worker and Cache Storage support and work in secure contexts (normally HTTPS, or localhost during development).

## Run locally

From the project directory, install dependencies and start the Vite development server:

```sh
npm install
npm run dev
```

Vite prints a local URL in the terminal (usually `http://localhost:5173`). Open that URL in a browser. Changes to source files are served using Vite's development-time hot module replacement.

## Production build and preview

Create an optimized production build:

```sh
npm run build
```

The generated static site is written to `dist/`. Preview that build locally with:

```sh
npm run preview
```

Deploy the contents of `dist/` to a static web host. Configure the host to serve the app's `index.html` for the root route. Service Workers require HTTPS in production; localhost is treated as a secure context by browsers for development.

## Using the dashboard

1. Start the app. On a new browser profile, sample transactions are shown so that the dashboard is not empty.
2. Choose **Import CSV / XLS** and select a `.csv`, `.xls`, or `.xlsx` statement file.
3. The app reads the first worksheet in an Excel workbook, locates a recognizable header row, and prepares the debit transactions it can identify.
4. Review the summary cards and category charts. Transactions classified as `unknown` appear in **Unknown merchant review**.
5. Choose **Resolve** to assign an unknown merchant to one of the available categories. This also creates a keyword rule from that transaction's narration and reference, so matching transactions can be classified later.
6. Use the category selector in **Monthly expense list** to change one transaction's category, or add a merchant keyword and category under **Category rules**.

Importing a new statement replaces the currently displayed transaction list; it does not append the new rows to the previous import. The current custom rules are used to classify the new import.

### Categories

The built-in categories are:

- `chai & smoke`
- `fuel`
- `food`
- `shopping`
- `groceries`
- `travel`
- `bills`
- `health`
- `entertainment`
- `other`
- `unknown` (used for transactions that do not match a rule)

The category selector includes the predefined categories and `unknown`. The custom-rule form provides the predefined categories other than `unknown`.

## Importing a statement

### CSV example

The parser expects a header row and comma-separated data rows. For example:

```csv
Date,Narration,Withdrawal Amt,Deposit Amt,Chq./Ref.No.,Balance
2026-09-01,UPI/CHAI WALA,120.00,,UPI00001234,24600.00
2026-09-02,UPI/FUEL STATION,1850.00,,,22750.00
2026-09-03,Salary credit,,50000.00,NEFT000099,72750.00
```

The example uses common statement headings; the app recognizes multiple spelling variants (see [Supported columns](#supported-columns)). CSV fields containing commas should be quoted, for example `"Cafe, Main Street"`. The CSV reader handles quoted cells and escaped quotes.

### Excel workbooks

Excel `.xls` and `.xlsx` files are read in the browser with SheetJS. The app reads **the first worksheet only** and searches its rows for a header row containing a recognized date column, narration/description column, and amount/debit/credit column. Introductory text above the header row is allowed. The column names do not have to be in a particular order.

### Supported columns

Column names are matched without regard to case, spaces, punctuation, underscores, or parentheses. For example, `Withdrawal Amt`, `withdrawal_amt`, and `WITHDRAWAL-AMT` normalize to the same key.

| Data used by the app | Recognized heading examples                                                                                                                                              |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Transaction date     | `Date`, `Txn Date`, `Transaction Date`, `Value Dt`, `Value Date`, `Posting Date`                                                                                         |
| Value date           | `Value Dt`, `Value Date`, `Valuedate`                                                                                                                                    |
| Narration / merchant | `Narration`, `Description`, `Particulars`, `Remarks`, `Transaction Details`, `Details`, `UPI ID`, `Bank ID`, `Merchant`, `Reference`, `Name`                             |
| Debit / expense      | `Debit Amount`, `Debit (Dr)`, `Withdrawal Amt`, `Withdrawal Amount`, `Withdrawal (Dr)`, `Withdrawal`, `Debit`                                                            |
| Credit / deposit     | `Credit Amount`, `Credit (Cr)`, `Deposit Amt`, `Deposit Amount`, `Deposit (Cr)`, `Deposit`, `Credit`                                                                     |
| Transaction amount   | `Amount`, `Txn Amount`, `Value`                                                                                                                                          |
| Transaction type     | `Type`, `Transaction Type`, `Txn Type`, `Dr/Cr`, `Debit/Credit`                                                                                                          |
| Reference            | `Ref No`, `Chq Ref No`, `Chq/Ref No`, `Chq./Ref.No.`, `Cheque No`, `Cheque Number`, `Cheque Reference Number`, `Reference`, `UPI ID`, `Bank ID`, `Transaction ID`, `UTR` |
| Balance              | `Closing Balance`, `Balance`, `Available Balance`                                                                                                                        |

For finding the header row, the parser requires all three of the following: a date heading (`Date`, `Txn Date`, `Transaction Date`, `Value Dt`, `Value Date`, or `Posting Date`); a narration heading (`Narration`, `Description`, `Particulars`, `Transaction Details`, `Remarks`, or `Details`); and at least one debit, credit, or amount heading. The remaining aliases in the table map values after a qualifying header row has been found; by themselves, they do not necessarily identify the header row. If no qualifying row is found, the import yields no parsed rows.

### Debit, credit, and amount handling

- A non-empty debit/withdrawal amount is treated as an expense and displayed as a positive amount.
- A non-empty credit/deposit amount is treated as income and excluded from the expense dashboard.
- If a statement uses a single amount column, transaction type headings such as `Dr/Cr` can indicate whether an amount is a debit (`Dr`/`Debit`) or a credit (`Cr`/`Credit`/`Deposit`). Without a debit/credit column or type, a positive amount is treated as an expense.
- Currency symbols and thousands separators are stripped before parsing numeric amounts. The dashboard formats amounts as Indian rupees (`INR`) with no fractional digits.
- Rows without a usable positive expense amount are not shown as expenses.

### Narration cleanup

The app removes common payment-channel suffixes when they occur after a dash at the end of narration. For example:

```text
UPI/SHOP NAME - Paid via CRED  ->  UPI/SHOP NAME
UPI/SHOP NAME - via PhonePe    ->  UPI/SHOP NAME
```

The cleanup recognizes phrases such as `Paid via`, `Payment via`, `via`, and provider names such as CRED, PhonePe, Google Pay/GPay, Paytm, Amazon Pay, and UPI. It trims from the matching dash to the end of the narration.

## How transaction classification works

The app checks the narration and reference text for category keywords. Matching is case-insensitive and uses substring containment. When a merchant does not match, it is assigned to `unknown` and appears in the review section.

Examples of built-in matching keywords include:

| Category        | Example matching text                                                  |
| --------------- | ---------------------------------------------------------------------- |
| `chai & smoke`  | chai, tea, pan, cafe, coffee, cigarette, smoke                         |
| `fuel`          | fuel, petrol, diesel, BPCL, Shell, filling station                     |
| `food`          | Zomato, Swiggy, restaurant, mess, dinner, food                         |
| `shopping`      | Amazon, Flipkart, Myntra, shop, market, store, bazaar, Paytm           |
| `groceries`     | grocery, supermarket, BigBasket, Zepto, fresh, Reliance Retail         |
| `travel`        | Uber, Ola, rail, train, bus, taxi                                      |
| `bills`         | electricity, bill, internet, water, rent, Razorpay, NoBroker, JioFiber |
| `health`        | medicine, pharmacy, hospital, diagnostic, MediBuddy                    |
| `entertainment` | Netflix, Spotify, cinema, CRED                                         |

The app checks saved user-created rules before the built-in rules. A custom rule is a case-insensitive `contains` match. Resolved unknown transactions create a rule using the selected transaction's narration and reference, if available. Changing a single transaction's category in the expense list changes that transaction; it does not create a reusable rule.

At a high level, the import and classification pipeline is:

```js
const rows = parseCsv(csvText); // Excel rows are mapped through the same row builder
const expenses = rows.filter((row) => row.isExpense);
const categorized = expenses.map((row) => ({
  ...row,
  category: getCategoryFromRule(`${row.description} ${row.reference}`, rules),
}));
```

An imported transaction is stored in a shape similar to this (empty or unavailable fields can be blank or `null`):

```js
{
	id: "0-2026-09-01-UPI/CHAI WALA",
	date: "2026-09-01",
	valueDate: "",
	description: "UPI/CHAI WALA",
	merchant: "UPI/CHAI WALA",
	amount: 120,
	debit: 120,
	credit: null,
	balance: 24600,
	reference: "UPI00001234",
	isExpense: true,
	category: "chai & smoke",
}
```

The browser stores the rule list and current transaction list together as JSON:

```js
{
	"rules": [
		{ "id": "rule-user-example", "pattern": "market", "category": "shopping", "matchType": "contains" }
	],
	"transactions": [
		{ "id": "...", "description": "...", "amount": 120, "category": "..." }
	]
}
```

The full transaction object contains additional statement values where available; the example shows only enough fields to illustrate the persisted structure.

## Local data and privacy

- Imported statement files are read locally by the browser using `FileReader` and SheetJS; this project contains no API endpoint or upload service.
- Transactions and category rules are saved in the current browser's `localStorage` under the key `expense-pwa-entries-v1`.
- Data is not synchronized between browsers or devices. Clearing the site's browser storage removes the saved transactions and custom rules.
- On first launch (or if no saved transaction list exists), the app initializes sample transactions. Importing a statement replaces that displayed list and then saves it locally.
- The application does not currently provide an in-app export, delete-all, or reset-data control. To remove local saved data, clear this site's local storage in the browser's developer tools or site settings.

Do not treat the app as a bank-connected service: it does not sign in to a bank, fetch transactions, or verify financial data. Check imported results against the original statement.

## Progressive Web App behavior

The page links to `public/manifest.webmanifest` and registers `public/sw.js` on window load. The manifest sets the app name, standalone display mode, theme/background colors, start URL, and an SVG icon. The Service Worker precaches `/`, removes older versions of its named cache during activation, and uses a cache-first lookup for GET requests while saving successful network responses for later use.

Offline behavior depends on what the browser has cached. The Service Worker does not synchronize data, implement background uploads, or provide push notifications. Serve the production app over HTTPS to enable PWA installation outside localhost.

## Project structure

```text
expense-pwa/
├── index.html                  # HTML entry point, manifest link, service-worker registration
├── package.json                # Dependencies and npm scripts
├── vite.config.js               # Vite React plugin configuration
├── public/
│   ├── favicon.svg              # Browser favicon
│   ├── icons.svg                # SVG icon asset
│   ├── manifest.webmanifest     # Installable PWA metadata
│   └── sw.js                    # Service-worker cache behavior
└── src/
	├── App.jsx                  # Dashboard, import/parser, rules, summaries, local persistence
	├── App.css                  # Dashboard and responsive component styles
	├── index.css                # Global styles and page defaults
	├── main.jsx                 # React root and app bootstrap
	└── assets/                  # Source assets
```

## Available scripts

| Command           | Description                                           |
| ----------------- | ----------------------------------------------------- |
| `npm install`     | Install dependencies declared in `package.json`.      |
| `npm run dev`     | Start the Vite development server with hot reload.    |
| `npm run build`   | Generate the production site in `dist/`.              |
| `npm run preview` | Serve the production build locally for a final check. |
| `npm run lint`    | Run Oxlint on the project.                            |

## Known limitations

- Only CSV, legacy Excel `.xls`, and `.xlsx` files are accepted by the import control; the first worksheet is used for Excel files.
- The CSV parser is a small in-app parser, not a full CSV standard implementation. Unusual delimiters, encodings, malformed quoted rows, and locale-specific number formats may need preprocessing.
- The import replaces rather than merges transaction lists. There is no duplicate detection, account selection, date-range filtering, or date-based monthly grouping yet; “monthly” is the dashboard label for the currently loaded transaction set.
- Built-in and user-defined keyword matches may need review because short or broad words can match unrelated merchants.
- The app stores its data only in browser storage, has no multi-user authentication, and does not back up or synchronize records.
- Imported data may contain sensitive financial information. Use a trusted device and browser profile, and clear browser storage when appropriate.
