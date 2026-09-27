import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import "./App.css";

const STORAGE_KEY = "expense-pwa-entries-v1";
const DEFAULT_CATEGORY_ORDER = [
  "chai & smoke",
  "fuel",
  "food",
  "shopping",
  "groceries",
  "travel",
  "bills",
  "health",
  "entertainment",
  "other",
];

const DEFAULT_RULES = [
  {
    id: "rule-chai",
    pattern: "chai",
    category: "chai & smoke",
    matchType: "contains",
  },
  {
    id: "rule-chaurasiya",
    pattern: "chaurasiya",
    category: "chai & smoke",
    matchType: "contains",
  },
  {
    id: "rule-tea",
    pattern: "tea",
    category: "chai & smoke",
    matchType: "contains",
  },
  {
    id: "rule-tea-stall",
    pattern: "tea stall",
    category: "chai & smoke",
    matchType: "contains",
  },
  {
    id: "rule-hangout",
    pattern: "hangout tea",
    category: "chai & smoke",
    matchType: "contains",
  },
  {
    id: "rule-pan",
    pattern: "pan",
    category: "chai & smoke",
    matchType: "contains",
  },
  {
    id: "rule-cafe",
    pattern: "cafe",
    category: "chai & smoke",
    matchType: "contains",
  },
  {
    id: "rule-coffee",
    pattern: "coffee",
    category: "chai & smoke",
    matchType: "contains",
  },
  {
    id: "rule-sweets",
    pattern: "sweets",
    category: "chai & smoke",
    matchType: "contains",
  },
  {
    id: "rule-cigarette",
    pattern: "cigarette",
    category: "chai & smoke",
    matchType: "contains",
  },
  {
    id: "rule-smoke",
    pattern: "smoke",
    category: "chai & smoke",
    matchType: "contains",
  },
  { id: "rule-fuel", pattern: "fuel", category: "fuel", matchType: "contains" },
  {
    id: "rule-petrol",
    pattern: "petrol",
    category: "fuel",
    matchType: "contains",
  },
  {
    id: "rule-diesel",
    pattern: "diesel",
    category: "fuel",
    matchType: "contains",
  },
  { id: "rule-bpcl", pattern: "bpcl", category: "fuel", matchType: "contains" },
  {
    id: "rule-shell",
    pattern: "shell",
    category: "fuel",
    matchType: "contains",
  },
  {
    id: "rule-energy-station",
    pattern: "energy station",
    category: "fuel",
    matchType: "contains",
  },
  {
    id: "rule-filling-station",
    pattern: "filling station",
    category: "fuel",
    matchType: "contains",
  },
  {
    id: "rule-zomato",
    pattern: "zomato",
    category: "food",
    matchType: "contains",
  },
  {
    id: "rule-swiggy",
    pattern: "swiggy",
    category: "food",
    matchType: "contains",
  },
  {
    id: "rule-dominos",
    pattern: "dominos",
    category: "food",
    matchType: "contains",
  },
  {
    id: "rule-restaurant",
    pattern: "restaurant",
    category: "food",
    matchType: "contains",
  },
  { id: "rule-mess", pattern: "mess", category: "food", matchType: "contains" },
  { id: "rule-food", pattern: "food", category: "food", matchType: "contains" },
  {
    id: "rule-dinner",
    pattern: "dinner",
    category: "food",
    matchType: "contains",
  },
  {
    id: "rule-barbeque",
    pattern: "barbeque",
    category: "food",
    matchType: "contains",
  },
  {
    id: "rule-hotel",
    pattern: "hotel",
    category: "food",
    matchType: "contains",
  },
  {
    id: "rule-khana",
    pattern: "khana",
    category: "food",
    matchType: "contains",
  },
  {
    id: "rule-amazon",
    pattern: "amazon",
    category: "shopping",
    matchType: "contains",
  },
  {
    id: "rule-flipkart",
    pattern: "flipkart",
    category: "shopping",
    matchType: "contains",
  },
  {
    id: "rule-myntra",
    pattern: "myntra",
    category: "shopping",
    matchType: "contains",
  },
  {
    id: "rule-shop",
    pattern: "shop",
    category: "shopping",
    matchType: "contains",
  },
  {
    id: "rule-market",
    pattern: "market",
    category: "shopping",
    matchType: "contains",
  },
  {
    id: "rule-store",
    pattern: "store",
    category: "shopping",
    matchType: "contains",
  },
  {
    id: "rule-bazaar",
    pattern: "bazaar",
    category: "shopping",
    matchType: "contains",
  },
  {
    id: "rule-grocery",
    pattern: "grocery",
    category: "groceries",
    matchType: "contains",
  },
  {
    id: "rule-supermarket",
    pattern: "supermarket",
    category: "groceries",
    matchType: "contains",
  },
  {
    id: "rule-bigbasket",
    pattern: "bigbasket",
    category: "groceries",
    matchType: "contains",
  },
  {
    id: "rule-fresh",
    pattern: "fresh",
    category: "groceries",
    matchType: "contains",
  },
  {
    id: "rule-zepto",
    pattern: "zepto",
    category: "groceries",
    matchType: "contains",
  },
  {
    id: "rule-reliance",
    pattern: "reliance retail",
    category: "groceries",
    matchType: "contains",
  },
  {
    id: "rule-uber",
    pattern: "uber",
    category: "travel",
    matchType: "contains",
  },
  { id: "rule-ola", pattern: "ola", category: "travel", matchType: "contains" },
  {
    id: "rule-rail",
    pattern: "rail",
    category: "travel",
    matchType: "contains",
  },
  {
    id: "rule-train",
    pattern: "train",
    category: "travel",
    matchType: "contains",
  },
  { id: "rule-bus", pattern: "bus", category: "travel", matchType: "contains" },
  {
    id: "rule-taxi",
    pattern: "taxi",
    category: "travel",
    matchType: "contains",
  },
  {
    id: "rule-electricity",
    pattern: "electricity",
    category: "bills",
    matchType: "contains",
  },
  {
    id: "rule-bill",
    pattern: "bill",
    category: "bills",
    matchType: "contains",
  },
  {
    id: "rule-internet",
    pattern: "internet",
    category: "bills",
    matchType: "contains",
  },
  {
    id: "rule-water",
    pattern: "water",
    category: "bills",
    matchType: "contains",
  },
  {
    id: "rule-rent",
    pattern: "rent",
    category: "bills",
    matchType: "contains",
  },
  {
    id: "rule-flat-rent",
    pattern: "flat rent",
    category: "bills",
    matchType: "contains",
  },
  {
    id: "rule-razorpay",
    pattern: "razorpay",
    category: "bills",
    matchType: "contains",
  },
  {
    id: "rule-medicine",
    pattern: "medicine",
    category: "health",
    matchType: "contains",
  },
  {
    id: "rule-pharmacy",
    pattern: "pharmacy",
    category: "health",
    matchType: "contains",
  },
  {
    id: "rule-hospital",
    pattern: "hospital",
    category: "health",
    matchType: "contains",
  },
  {
    id: "rule-diagnostic",
    pattern: "diagnostic",
    category: "health",
    matchType: "contains",
  },
  {
    id: "rule-medibuddy",
    pattern: "medibuddy",
    category: "health",
    matchType: "contains",
  },
  {
    id: "rule-netflix",
    pattern: "netflix",
    category: "entertainment",
    matchType: "contains",
  },
  {
    id: "rule-spotify",
    pattern: "spotify",
    category: "entertainment",
    matchType: "contains",
  },
  {
    id: "rule-cinema",
    pattern: "cinema",
    category: "entertainment",
    matchType: "contains",
  },
  {
    id: "rule-cred",
    pattern: "cred",
    category: "entertainment",
    matchType: "contains",
  },
  {
    id: "rule-nobroker",
    pattern: "nobroker",
    category: "bills",
    matchType: "contains",
  },
  {
    id: "rule-paytm",
    pattern: "paytm",
    category: "shopping",
    matchType: "contains",
  },
  {
    id: "rule-jiofiber",
    pattern: "jiofiber",
    category: "bills",
    matchType: "contains",
  },
];

const SAMPLE_CSV = `Date,Description,Debit,Credit,Balance
2024-09-01,UPI/CHAI WALA,120.00,,24600.00
2024-09-02,UPI/FUEL STATION,1850.00,,22750.00
2024-09-03,UPI/ZOMATO FOOD,420.00,,22330.00
2024-09-08,UPI/AMAZON MKT,1480.00,,20850.00
2024-09-10,UPI/UBER RIDE,390.00,,20460.00
2024-09-13,UPI/NETFLIX SUB,499.00,,19961.00
2024-09-18,UPI/PETROL BUNK,2102.00,,17859.00
2024-09-22,UPI/SHOPPING MALL,960.00,,16899.00
2024-09-27,UPI/NEW MERCHANT,550.00,,16349.00`;

function normalizeText(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase();
}

function cleanNarration(value) {
  return String(value ?? "")
    .replace(
      /\s*[-–—]\s*(?:(?:paid|payment)\s+via|via)\s+.+$|\s*[-–—]\s*(?:cred|phonepe|google\s*pay|gpay|paytm|amazon\s*pay|upi)\b.*$/i,
      ""
    )
    .trim();
}

function parseCurrency(value) {
  if (value === undefined || value === null || value === "") return null;
  const cleaned = String(value)
    .replace(/[^0-9.,-]/g, "")
    .replace(/,/g, "");
  if (!cleaned) return null;
  const parsed = Number.parseFloat(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function splitCsvRow(line) {
  const cells = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];

    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === "," && !inQuotes) {
      cells.push(current);
      current = "";
    } else {
      current += char;
    }
  }

  cells.push(current);
  return cells.map((cell) => cell.trim());
}

function normalizeKey(value) {
  return normalizeText(value).replace(/[^a-z0-9]/g, "");
}

function getRecordValue(record, variants) {
  const compactMap = Object.fromEntries(
    Object.entries(record).map(([key, value]) => [normalizeKey(key), value])
  );

  for (const variant of variants) {
    const match = compactMap[normalizeKey(variant)];
    if (match !== undefined && match !== null && match !== "") {
      return match;
    }
  }

  return "";
}

function findStatementHeaderIndex(rows) {
  return rows.findIndex((row) => {
    const values = Array.isArray(row) ? row : Object.keys(row);
    const headers = values.map(normalizeKey);
    const hasDate = headers.some((header) =>
      [
        "date",
        "txndate",
        "transactiondate",
        "valuedt",
        "valuedate",
        "postingdate",
      ].includes(header)
    );
    const hasDescription = headers.some((header) =>
      [
        "narration",
        "description",
        "particulars",
        "transactiondetails",
        "remarks",
        "details",
      ].includes(header)
    );
    const hasAmount = headers.some((header) =>
      [
        "withdrawal",
        "withdrawaldr",
        "withdrawalamt",
        "withdrawalamount",
        "debit",
        "debitdr",
        "debitamount",
        "deposit",
        "depositcr",
        "depositamt",
        "depositamount",
        "credit",
        "creditcr",
        "creditamount",
        "amount",
        "txnamount",
      ].includes(header)
    );

    return hasDate && hasDescription && hasAmount;
  });
}

function buildStatementRow(item, rowIndex) {
  const date = getRecordValue(item, [
    "date",
    "txn date",
    "transaction date",
    "value dt",
    "value date",
    "posting date",
  ]);
  const valueDate = getRecordValue(item, [
    "value dt",
    "value date",
    "valuedate",
  ]);
  const description = cleanNarration(
    getRecordValue(item, [
      "description",
      "narration",
      "particulars",
      "remarks",
      "transaction details",
      "details",
      "upi id",
      "bank id",
      "merchant",
      "reference",
      "name",
    ])
  );
  const reference = getRecordValue(item, [
    "ref no",
    "chq ref no",
    "chq/ref no",
    "chq./ref.no.",
    "cheque no",
    "cheque number",
    "cheque reference number",
    "reference",
    "upi id",
    "bank id",
    "transaction id",
    "utr",
  ]);
  const debit = parseCurrency(
    getRecordValue(item, [
      "debit amount",
      "debit (dr)",
      "withdrawal amt",
      "withdrawal amount",
      "withdrawal (dr)",
      "withdrawal",
      "dr",
      "debit",
    ])
  );
  const credit = parseCurrency(
    getRecordValue(item, [
      "credit amount",
      "credit (cr)",
      "deposit amt",
      "deposit amount",
      "deposit (cr)",
      "deposit",
      "cr",
      "credit",
    ])
  );
  const rawAmount = parseCurrency(
    getRecordValue(item, ["amount", "txn amount", "value"])
  );
  const balance = parseCurrency(
    getRecordValue(item, ["closing balance", "balance", "available balance"])
  );
  const transactionType = normalizeKey(
    getRecordValue(item, [
      "type",
      "transaction type",
      "txn type",
      "dr/cr",
      "debit/credit",
    ])
  );

  let amountValue = rawAmount;
  if (debit !== null) amountValue = Math.abs(debit);
  if (credit !== null) amountValue = -Math.abs(credit);

  if (["debit", "dr"].includes(transactionType) && amountValue !== null) {
    amountValue = Math.abs(amountValue);
  }

  if (
    ["credit", "cr", "deposit"].includes(transactionType) &&
    amountValue !== null
  ) {
    amountValue = -Math.abs(amountValue);
  }

  if (amountValue === null) {
    amountValue = rawAmount;
  }

  return {
    id: `${rowIndex}-${date}-${description}`,
    date,
    valueDate,
    description,
    merchant: description,
    amount: amountValue,
    debit,
    credit,
    balance,
    reference,
    isExpense:
      debit !== null ||
      (credit === null && amountValue !== null && amountValue > 0),
  };
}

function parseCsv(csvText) {
  const lines = csvText
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  if (!lines.length) {
    return [];
  }

  const rows = lines.map((line) => splitCsvRow(line));
  const headerIndex = findStatementHeaderIndex(rows);
  if (headerIndex === -1) {
    return [];
  }

  const headers = rows[headerIndex].map((header) =>
    normalizeText(String(header ?? ""))
  );
  const dataRows = rows.slice(headerIndex + 1);

  return dataRows.map((values, rowIndex) => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = values[index] ?? "";
    });

    return buildStatementRow(item, rowIndex);
  });
}

function getCategoryFromRule(text, customRules) {
  const normalizedText = normalizeText(text);
  if (!normalizedText) return "unknown";

  const customMatch = customRules.find((rule) => {
    const pattern = normalizeText(rule.pattern);
    if (!pattern) return false;
    return rule.matchType === "exact"
      ? normalizedText === pattern
      : normalizedText.includes(pattern);
  });

  if (customMatch) {
    return customMatch.category;
  }

  const defaultMatch = DEFAULT_RULES.find((rule) => {
    const pattern = normalizeText(rule.pattern);
    return normalizedText.includes(pattern);
  });

  return defaultMatch?.category ?? "unknown";
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 0,
  }).format(value || 0);
}

function makeId(prefix) {
  if (typeof crypto !== "undefined" && crypto.randomUUID) {
    return `${prefix}-${crypto.randomUUID()}`;
  }

  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function createSampleTransactions() {
  return parseCsv(SAMPLE_CSV)
    .filter((row) => row.isExpense)
    .map((row) => ({
      ...row,
      category: getCategoryFromRule(row.description, DEFAULT_RULES),
    }));
}

function App() {
  const [rules, setRules] = useState(DEFAULT_RULES);
  const [transactions, setTransactions] = useState([]);
  const [pendingUnknowns, setPendingUnknowns] = useState([]);
  const [selectedUnknown, setSelectedUnknown] = useState(null);
  const [categoryMap, setCategoryMap] = useState({});

  useEffect(() => {
    try {
      const stored = JSON.parse(localStorage.getItem(STORAGE_KEY) || "null");

      if (stored?.transactions?.length) {
        const nextRules = stored.rules?.length ? stored.rules : DEFAULT_RULES;
        setRules(nextRules);
        setTransactions(stored.transactions);
        setPendingUnknowns(
          stored.transactions.filter((entry) => entry.category === "unknown")
        );
      } else {
        const initialTransactions = createSampleTransactions();
        setTransactions(initialTransactions);
        setRules(DEFAULT_RULES);
        setPendingUnknowns(
          initialTransactions.filter((entry) => entry.category === "unknown")
        );
        localStorage.setItem(
          STORAGE_KEY,
          JSON.stringify({
            rules: DEFAULT_RULES,
            transactions: initialTransactions,
          })
        );
      }
    } catch (error) {
      const fallback = createSampleTransactions();
      setTransactions(fallback);
      setRules(DEFAULT_RULES);
      setPendingUnknowns(
        fallback.filter((entry) => entry.category === "unknown")
      );
    }
  }, []);

  useEffect(() => {
    if (!transactions.length) {
      return;
    }

    const persisted = {
      rules,
      transactions,
    };

    localStorage.setItem(STORAGE_KEY, JSON.stringify(persisted));
  }, [rules, transactions]);

  const totalsByCategory = useMemo(() => {
    return transactions.reduce((accumulator, item) => {
      const category = item.category || "unknown";
      accumulator[category] =
        (accumulator[category] || 0) + Number(item.amount || 0);
      return accumulator;
    }, {});
  }, [transactions]);

  const chartData = useMemo(() => {
    const entries = Object.entries(totalsByCategory).filter(
      ([, total]) => total > 0
    );
    const maxValue = Math.max(...entries.map(([, total]) => total), 1);

    return entries.map(([category, total]) => ({
      category,
      total,
      width: Math.max((total / maxValue) * 100, 6),
    }));
  }, [totalsByCategory]);

  const totalExpense = useMemo(
    () => transactions.reduce((sum, item) => sum + Number(item.amount || 0), 0),
    [transactions]
  );

  const categoryColorMap = {
    "chai & smoke": "#f97316",
    fuel: "#facc15",
    food: "#f43f5e",
    shopping: "#8b5cf6",
    groceries: "#10b981",
    travel: "#38bdf8",
    bills: "#14b8a6",
    health: "#22c55e",
    entertainment: "#ec4899",
    other: "#94a3b8",
    unknown: "#ef4444",
  };

  const categoryList = Object.keys(totalsByCategory)
    .sort((a, b) => totalsByCategory[b] - totalsByCategory[a])
    .slice(0, 6);

  const handleCsvImport = (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = () => {
      let parsedRows = [];

      if (
        file.name.toLowerCase().endsWith(".xlsx") ||
        file.name.toLowerCase().endsWith(".xls")
      ) {
        const workbook = XLSX.read(reader.result, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: "",
          raw: false,
          blankrows: false,
        });

        const headerIndex = findStatementHeaderIndex(rows);
        if (headerIndex !== -1) {
          const headers = rows[headerIndex].map((header) =>
            normalizeText(String(header ?? ""))
          );
          parsedRows = rows.slice(headerIndex + 1).map((values, rowIndex) => {
            const row = {};
            headers.forEach((header, index) => {
              row[header] = values[index] ?? "";
            });
            return buildStatementRow(row, rowIndex);
          });
        }
      } else {
        parsedRows = parseCsv(String(reader.result || ""));
      }

      const classified = parsedRows
        .filter((row) => row.isExpense)
        .map((row) => ({
          ...row,
          category: getCategoryFromRule(
            `${row.description} ${row.reference}`,
            rules
          ),
        }));

      const nextTransactions = classified.filter((row) => row.amount > 0);
      setTransactions(nextTransactions);
      const unresolved = nextTransactions.filter(
        (entry) => entry.category === "unknown"
      );
      setPendingUnknowns(unresolved);
      if (!unresolved.length) {
        setSelectedUnknown(null);
      }
    };

    if (
      file.name.toLowerCase().endsWith(".xlsx") ||
      file.name.toLowerCase().endsWith(".xls")
    ) {
      reader.readAsArrayBuffer(file);
    } else {
      reader.readAsText(file);
    }

    event.target.value = "";
  };

  const addUserRule = (patternValue, categoryValue) => {
    const pattern = patternValue.trim();
    const category = categoryValue.trim();
    if (!pattern || !category) return;

    const rule = {
      id: makeId("rule-user"),
      pattern,
      category,
      matchType: "contains",
    };

    setRules((currentRules) => {
      const nextRules = [rule, ...currentRules];
      setTransactions((currentTransactions) =>
        currentTransactions.map((entry) => ({
          ...entry,
          category:
            entry.category === "unknown" &&
            `${entry.description} ${entry.reference}`
              .toLowerCase()
              .includes(pattern.toLowerCase())
              ? category
              : entry.category,
        }))
      );
      return nextRules;
    });
  };

  const resolveUnknown = (unknownId, category) => {
    const transaction = transactions.find((entry) => entry.id === unknownId);
    if (!transaction) return;

    const description =
      `${transaction.description} ${transaction.reference}`.trim();
    const pattern =
      description || transaction.reference || transaction.description;

    if (pattern && category !== "unknown") {
      setRules((currentRules) => [
        {
          id: makeId("rule-user"),
          pattern,
          category,
          matchType: "contains",
        },
        ...currentRules,
      ]);
    }

    setTransactions((currentTransactions) =>
      currentTransactions.map((entry) =>
        entry.id === unknownId
          ? {
              ...entry,
              category,
            }
          : entry
      )
    );

    setPendingUnknowns((current) =>
      current.filter((entry) => entry.id !== unknownId)
    );
    setSelectedUnknown(null);
  };

  const donutSegments = categoryList.map((category) => {
    const value = totalsByCategory[category] || 0;
    const percent = totalExpense ? (value / totalExpense) * 100 : 0;
    return {
      category,
      value,
      percent,
      color: categoryColorMap[category] || "#94a3b8",
    };
  });

  const largestCategory = chartData[0];

  return (
    <div className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">Monthly expense dashboard</p>
          <h1>Bank statement classifier</h1>
        </div>
        <label className="upload-button">
          <input
            type="file"
            accept=".csv,text/csv,.xls,.xlsx"
            onChange={handleCsvImport}
          />
          Import CSV / XLS
        </label>
      </header>

      <main className="dashboard">
        <section className="summary-grid">
          <article className="summary-card highlight">
            <span>Total spend</span>
            <strong>{formatCurrency(totalExpense)}</strong>
            <small>{transactions.length} classified entries</small>
          </article>
          <article className="summary-card">
            <span>Largest category</span>
            <strong>
              {largestCategory ? largestCategory.category : "N/A"}
            </strong>
            <small>
              {largestCategory ? formatCurrency(largestCategory.total) : "0"}
            </small>
          </article>
          <article className="summary-card">
            <span>Unknown entries</span>
            <strong>{pendingUnknowns.length}</strong>
            <small>
              {pendingUnknowns.length ? "Needs review" : "All mapped"}{" "}
            </small>
          </article>
        </section>

        <section className="content-grid">
          <article className="panel chart-panel">
            <div className="panel-header">
              <h2>Expense by category</h2>
            </div>

            <div className="bar-chart">
              {chartData.map(({ category, total, width }) => (
                <div key={category} className="bar-row">
                  <span className="bar-label">{category}</span>
                  <div className="bar-track">
                    <div
                      className="bar-fill"
                      style={{
                        width: `${width}%`,
                        background: categoryColorMap[category] || "#94a3b8",
                      }}
                    />
                  </div>
                  <strong>{formatCurrency(total)}</strong>
                </div>
              ))}
            </div>
          </article>

          <article className="panel donut-panel">
            <div className="panel-header">
              <h2>Share of spend</h2>
            </div>
            <div className="donut-wrap">
              <svg
                viewBox="0 0 120 120"
                className="donut-chart"
                role="img"
                aria-label="Expense category chart"
              >
                <circle cx="60" cy="60" r="42" className="donut-bg" />
                {donutSegments.map((segment, index) => {
                  const circumference = 2 * Math.PI * 42;
                  const dash = (segment.percent / 100) * circumference;
                  const gap = circumference - dash;
                  const offset = donutSegments
                    .slice(0, index)
                    .reduce(
                      (total, item) =>
                        total + (item.percent / 100) * circumference,
                      0
                    );
                  return (
                    <circle
                      key={segment.category}
                      cx="60"
                      cy="60"
                      r="42"
                      className="donut-segment"
                      stroke={segment.color}
                      strokeDasharray={`${dash} ${gap}`}
                      strokeDashoffset={-offset}
                    />
                  );
                })}
                <text
                  x="60"
                  y="55"
                  textAnchor="middle"
                  className="donut-total-label"
                >
                  {formatCurrency(totalExpense)}
                </text>
                <text
                  x="60"
                  y="70"
                  textAnchor="middle"
                  className="donut-total-sub"
                >
                  Spend
                </text>
              </svg>
            </div>
            <div className="legend-list">
              {donutSegments.map((segment) => (
                <div key={segment.category} className="legend-item">
                  <span
                    className="legend-swatch"
                    style={{ background: segment.color }}
                  />
                  <span>{segment.category}</span>
                  <strong>{segment.percent.toFixed(0)}%</strong>
                </div>
              ))}
            </div>
          </article>
        </section>

        <section className="panel">
          <div className="panel-header split-header">
            <h2>Unknown merchant review</h2>
            <button
              type="button"
              className="ghost-button"
              onClick={() => {
                setTransactions((current) =>
                  current.map((entry) => ({
                    ...entry,
                    category: entry.category || "unknown",
                  }))
                );
              }}
            >
              Keep existing mapping
            </button>
          </div>

          {pendingUnknowns.length === 0 ? (
            <div className="empty-state">
              No unresolved expenses. Everything is classified.
            </div>
          ) : (
            <div className="unknown-list">
              {pendingUnknowns.map((entry) => (
                <div key={entry.id} className="unknown-row">
                  <div>
                    <strong>{entry.description}</strong>
                    <small>{entry.reference || "No reference ID"}</small>
                  </div>
                  <div className="money-wrap">
                    <span>{formatCurrency(entry.amount)}</span>
                    <button
                      type="button"
                      onClick={() => setSelectedUnknown(entry)}
                    >
                      Resolve
                    </button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </section>

        <section className="panel">
          <div className="panel-header">
            <h2>Category rules</h2>
          </div>
          <form
            className="rule-form"
            onSubmit={(event) => {
              event.preventDefault();
              const form = new FormData(event.currentTarget);
              addUserRule(form.get("pattern"), form.get("category"));
              event.currentTarget.reset();
            }}
          >
            <input name="pattern" placeholder="merchant or UPI keyword" />
            <select name="category" defaultValue="food">
              {DEFAULT_CATEGORY_ORDER.map((category) => (
                <option key={category} value={category}>
                  {category}
                </option>
              ))}
            </select>
            <button type="submit">Save rule</button>
          </form>
        </section>

        <section className="panel table-panel">
          <div className="panel-header">
            <h2>Monthly expense list</h2>
          </div>
          <div className="expense-table">
            {transactions.map((entry) => (
              <div key={entry.id} className="table-row">
                <div>
                  <strong>{entry.description}</strong>
                  <small>
                    {entry.date || "No date"}
                    {entry.valueDate && entry.valueDate !== entry.date
                      ? ` · Value date: ${entry.valueDate}`
                      : ""}
                    {entry.reference ? ` · Ref: ${entry.reference}` : ""}
                  </small>
                </div>
                <select
                  className="category-select"
                  value={entry.category || "unknown"}
                  onChange={(event) => {
                    const nextCategory = event.target.value;
                    setTransactions((currentTransactions) =>
                      currentTransactions.map((transaction) =>
                        transaction.id === entry.id
                          ? { ...transaction, category: nextCategory }
                          : transaction
                      )
                    );
                  }}
                  style={{
                    borderColor: categoryColorMap[entry.category] || "#94a3b8",
                    color: categoryColorMap[entry.category] || "#94a3b8",
                  }}
                >
                  {DEFAULT_CATEGORY_ORDER.map((category) => (
                    <option key={category} value={category}>
                      {category}
                    </option>
                  ))}
                  <option value="unknown">unknown</option>
                </select>
                <strong>{formatCurrency(entry.amount)}</strong>
                <small className="balance-value">
                  {entry.balance !== null && entry.balance !== undefined
                    ? `Balance ${formatCurrency(entry.balance)}`
                    : ""}
                </small>
              </div>
            ))}
          </div>
        </section>
      </main>

      {selectedUnknown && (
        <div
          className="modal-backdrop"
          onClick={() => setSelectedUnknown(null)}
        >
          <div
            className="modal-card"
            onClick={(event) => event.stopPropagation()}
          >
            <h3>Classify this merchant</h3>
            <p>
              <strong>{selectedUnknown.description}</strong>
              <br />
              {selectedUnknown.reference || "No reference"}
            </p>

            <div className="modal-actions">
              {DEFAULT_CATEGORY_ORDER.map((category) => (
                <button
                  key={category}
                  type="button"
                  className="tag-button"
                  onClick={() => resolveUnknown(selectedUnknown.id, category)}
                >
                  {category}
                </button>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
