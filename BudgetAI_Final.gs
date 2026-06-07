const GROQ_API_KEY = "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";
const GROQ_URL = "https://api.groq.com/openai/v1/chat/completions";
const SPIKE_THRESHOLD_PERCENT = 50;
const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

// ─── MENU ─────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💰 Budget Tools")
    .addItem("🚀 Run Everything", "runAll")
    .addSeparator()
    .addItem("1. Categorize Transactions", "categorizeTransactions")
    .addItem("2. Update Fund Flow", "updateFundFlow")
    .addItem("3. Rebuild Quality Control", "buildQualityControl")
    .addSeparator()
    .addItem("ℹ️ View Sheet Status", "showStatus")
    .addToUi();
}

function runAll() {
  categorizeTransactions();
  updateFundFlow();
  buildQualityControl();
  SpreadsheetApp.getUi().alert("✅ All done! Open your web dashboard URL to see results.");
}

// ─── HELPERS ──────────────────────────────────────────────────────
function columnToLetter(col) {
  let letter = "";
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

function parseHDFCDate(dateVal) {
  if (!dateVal) return null;
  if (dateVal instanceof Date && !isNaN(dateVal)) {
    return { day: dateVal.getDate(), month: dateVal.getMonth(), year: dateVal.getFullYear() };
  }
  const str = dateVal.toString().trim();
  const parts = str.split(/[-\/]/);
  if (parts.length === 3) {
    const p0 = parseInt(parts[0]), p1 = parseInt(parts[1]), p2 = parseInt(parts[2]);
    if (p2 > 1900) return { day: p0, month: p1 - 1, year: p2 };
    if (p0 > 1900) return { day: p2, month: p1 - 1, year: p0 };
  }
  return null;
}

function toMonthYear(val) {
  if (!val) return "";
  if (val instanceof Date && !isNaN(val)) {
    return MONTH_NAMES[val.getMonth()] + "-" + val.getFullYear();
  }
  const cleaned = val.toString().trim().replace(/^'+/, "");
  if (cleaned.includes("-") && cleaned.length <= 10) return cleaned;
  return "";
}

function saveMetadata(lastTxnDate, txnCount, closingBalance) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("lastUpdated",    new Date().toLocaleString("en-IN"));
  props.setProperty("lastTxnDate",    lastTxnDate    || "");
  props.setProperty("txnCount",       String(txnCount || 0));
  props.setProperty("closingBalance", String(closingBalance || 0));
}

function getMetadata() {
  const props = PropertiesService.getScriptProperties();
  return {
    lastUpdated:    props.getProperty("lastUpdated")    || "Never",
    lastTxnDate:    props.getProperty("lastTxnDate")    || "Unknown",
    txnCount:       props.getProperty("txnCount")       || "0",
    closingBalance: props.getProperty("closingBalance") || "0",
    uncategorized:  JSON.parse(props.getProperty("uncategorized") || "[]"),
  };
}

function showStatus() {
  const meta = getMetadata();
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Categorized").getDataRange().getValues();
  const totalTxns = data.length - 1;
  const monthCounts = {};
  for (let i = 1; i < data.length; i++) {
    const my = toMonthYear(data[i][7]);
    if (my) monthCounts[my] = (monthCounts[my] || 0) + 1;
  }
  const monthSummary = Object.entries(monthCounts)
    .map(function(entry) { return entry[0] + ": " + entry[1] + " txns"; }).join("\n");
  SpreadsheetApp.getUi().alert(
    "📊 SHEET STATUS\n\n" +
    "🕐 Last Updated: " + meta.lastUpdated + "\n" +
    "📅 Last Transaction: " + meta.lastTxnDate + "\n" +
    "🏦 Closing Balance: ₹" + parseFloat(meta.closingBalance).toLocaleString("en-IN") + "\n" +
    "📋 Total Transactions: " + totalTxns + "\n" +
    "⚠️ Uncategorized: " + meta.uncategorized.length + "\n\n" +
    "📆 Transactions per Month:\n" + monthSummary
  );
}

// ─── STEP 1: CATEGORIZE ───────────────────────────────────────────
function categorizeTransactions() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Raw Imports");
  const catSheet = ss.getSheetByName("Categorized");

  if (!rawSheet || !catSheet) {
    SpreadsheetApp.getUi().alert("❌ Missing tabs: Raw Imports or Categorized.");
    return;
  }

  catSheet.clearContents();
  catSheet.clearFormats();
  catSheet.getRange(1, 1, 1, 8).setValues([
    ["Date","Narration","Type","Amount","Category","Subcategory","Confidence","Month-Year"]
  ]);
  catSheet.getRange(1, 1, 1, 8)
    .setBackground("#0f4c81").setFontColor("#ffffff").setFontWeight("bold");

  const rawData = rawSheet.getDataRange().getValues();
  let colDate = 0, colNarr = 1, colWd = 4, colDep = 5, colClosing = 6, headerRow = -1;

  for (let i = 0; i < rawData.length; i++) {
    const rowStr = rawData[i].map(function(c) { return (c||"").toString(); }).join(" ").toLowerCase();
    if (rowStr.includes("narration") && rowStr.includes("withdrawal")) {
      headerRow = i;
      rawData[i].forEach(function(cell, c) {
        const v = (cell||"").toString().toLowerCase().trim();
        if (v === "date") colDate = c;
        if (v === "narration" || v === "description" || v === "particulars") colNarr = c;
        if (v.includes("withdrawal") || v === "debit") colWd = c;
        if (v.includes("deposit")    || v === "credit") colDep = c;
        if (v.includes("closing")    || v.includes("balance")) colClosing = c;
      });
      break;
    }
  }

  const startRow = headerRow >= 0 ? headerRow + 1 : 1;
  const results  = [];
  const needsAI  = []; // {idx, narration, amount, type}
  let skipped = 0;
  let lastTxnDate = "";
  let lastClosingBalance = 0;
  let latestDateObj = null;

  // Pass 1 — rule-based (instant)
  for (let i = startRow; i < rawData.length; i++) {
    const row       = rawData[i];
    const narration = (row[colNarr] || "").toString().trim();
    const wd        = parseFloat(row[colWd])      || 0;
    const dep       = parseFloat(row[colDep])     || 0;
    const closing   = parseFloat(row[colClosing]) || 0;

    if (!row[colDate] || (!wd && !dep) || !narration) { skipped++; continue; }

    const d = parseHDFCDate(row[colDate]);
    if (!d || !d.year) { skipped++; continue; }

    const dateStr   = String(d.day).padStart(2,"0") + "/" + String(d.month+1).padStart(2,"0") + "/" + d.year;
    const monthYear = MONTH_NAMES[d.month] + "-" + d.year;
    const type      = wd > 0 ? "Debit" : "Credit";
    const amount    = wd > 0 ? wd : dep;

    const txnDate = new Date(d.year, d.month, d.day);
    if (!latestDateObj || txnDate > latestDateObj) {
      latestDateObj = txnDate;
      lastTxnDate   = dateStr;
    }
    if (closing > 0) lastClosingBalance = closing;

    const cat = ruleBasedCategorize(narration, amount, type);

    if (cat.main === "NEEDS_AI") {
      needsAI.push({ idx: results.length, narration: narration, amount: amount, type: type });
      results.push([dateStr, narration, type, amount, "Pending", "Pending", "Low", monthYear]);
    } else {
      results.push([dateStr, narration, type, amount, cat.main, cat.sub, cat.confidence, monthYear]);
    }
  }

  // Pass 2 — batch AI for unknowns only
  if (needsAI.length > 0) {
    const aiResults = callGroqBatch(needsAI);
    for (let i = 0; i < needsAI.length; i++) {
      const cat = aiResults[i] || { main:"Other", sub:"Unknown", confidence:"Low" };
      results[needsAI[i].idx][4] = cat.main;
      results[needsAI[i].idx][5] = cat.sub;
      results[needsAI[i].idx][6] = cat.confidence;
    }
  }

  // Collect uncategorized
  const uncategorized = [];
  for (let i = 0; i < results.length; i++) {
    if (results[i][4] === "Other" || results[i][5] === "Unknown") {
      uncategorized.push({
        date: results[i][0], narration: results[i][1],
        type: results[i][2], amount:    results[i][3],
        main: results[i][4], sub:       results[i][5]
      });
    }
  }

  // Write all at once
  if (results.length > 0) {
    catSheet.getRange(2, 1, results.length, 8).setValues(results);
    for (let r = 0; r < results.length; r++) {
      catSheet.getRange(2 + r, 1, 1, 8)
        .setBackground(r % 2 === 0 ? "#f8f9fa" : "#ffffff");
    }
  }

  saveMetadata(lastTxnDate, results.length, lastClosingBalance);
  PropertiesService.getScriptProperties()
    .setProperty("uncategorized", JSON.stringify(uncategorized));
  updateStatusTab();

  rawSheet.clearContents();
  rawSheet.getRange(1, 1, 1, 7).setValues([
    ["Date","Narration","Ref No","Value Dt","Withdrawal Amt.","Deposit Amt.","Closing Balance"]
  ]);

  SpreadsheetApp.getUi().alert(
    "✅ Done!\n" +
    "• " + results.length + " transactions processed\n" +
    "• " + (results.length - needsAI.length) + " matched by rules instantly\n" +
    "• " + needsAI.length + " sent to AI\n" +
    "• " + uncategorized.length + " flagged for review\n" +
    "• Last txn: " + lastTxnDate + "\n" +
    "• Balance: ₹" + lastClosingBalance.toLocaleString("en-IN")
  );
}

// ─── RULE-BASED CATEGORIZER (no API, instant) ─────────────────────
function ruleBasedCategorize(narration, amount, type) {
  const n = narration.toUpperCase();

  if (type === "Credit") {
    if (/SALARY|PAYROLL|PAYSLIP|SALARY SEP|SALARY OCT|SALARY NOV|SALARY DEC|SALARY JAN|SALARY FEB|SALARY MAR|SALARY APR|SALARY MAY|SALARY JUN|SALARY JUL|SALARY AUG/.test(n)) return {main:"Income", sub:"Salary", confidence:"High"};
    if (/ITDTAX|TAX REFUND|INCOME TAX/.test(n))     return { main:"Income", sub:"Tax Refund",         confidence:"High" };
    if (/DIVIDEND|DIV |FNLDIV|INTDIV/.test(n))      return { main:"Income", sub:"Dividend",           confidence:"High" };
    if (/ZERODHA PAYOUT|ZERODHA BROKING/.test(n))   return { main:"Income", sub:"Investment Returns", confidence:"High" };
    if (/REFUND/.test(n))                            return { main:"Income", sub:"Reimbursement",      confidence:"High" };
    if (/ACH C-/.test(n))                           return { main:"Income", sub:"Dividend",           confidence:"Medium" };
  }

  if (type === "Debit") {
    if (/TRULIV/.test(n))                                                              return { main:"Fixed Exp", sub:"Rent",                 confidence:"High" };
    if (/NET BANKING SI|STANDING INSTRUCTION/.test(n))                                return { main:"Fixed Exp", sub:"Standing Instruction", confidence:"High" };
    if (/IB BILLPAY|BILLPAY DR/.test(n))                                             return { main:"Fixed Exp", sub:"CC Bill Payment",      confidence:"High" };
    if (/ACH D-.*IDFC|ACH D-.*HDFC|ACH D-.*ICICI|ACH D-.*SBI|ACH D-.*AXIS/.test(n)) return { main:"Fixed Exp", sub:"Loan EMI",             confidence:"High" };
    if (/ACH D- TP ACH INDIANESIGN/.test(n))                                          return { main:"Fixed Exp", sub:"Insurance",            confidence:"High" };
    if (/INDMONEY.*INSURANCE|INSURANCE.*INDMONEY/.test(n))                            return { main:"Fixed Exp", sub:"Insurance",            confidence:"High" };
    if (/NETFLIX/.test(n))                                                             return { main:"Fixed Exp", sub:"Subscription",         confidence:"High" };
    if (/APPLE.*MEDIA|APPLE.*SERVICE|APPLESERVICES/.test(n))                          return { main:"Fixed Exp", sub:"Subscription",         confidence:"High" };
    if (/SPOTIFY|HOTSTAR|PRIME.*VIDEO|DISNEY/.test(n))                                return { main:"Fixed Exp", sub:"Subscription",         confidence:"High" };
    if (/DISTRICT M/.test(n))                                                          return { main:"Fixed Exp", sub:"Subscription",         confidence:"High" };
    if (/ZERODHA|GROWW|COIN BY ZERODHA/.test(n))                                      return { main:"Investing",  sub:"Stocks",              confidence:"High" };
    if (/INDMONEY.*RZP|INDMONEY1X/.test(n))                                           return { main:"Investing",  sub:"Stocks",              confidence:"High" };
    if (/SIP |MUTUAL FUND|MIRAE|AXIS.*FUND|HDFC.*FUND|SBI.*FUND/.test(n))            return { main:"Investing",  sub:"SIP",                 confidence:"High" };
    if (/EPFO|PF |PROVIDENT FUND/.test(n))                                            return { main:"Investing",  sub:"EPFO",                confidence:"High" };
    if (/ZOMATO|ETERNAL LIMITED|SWIGGY|JUICY POP|TEJASFOODS|VENKATESHWARA|METRO JUICE|SRI KRISHNA BEKARY/.test(n)) return { main:"Personal Exp", sub:"Food",      confidence:"High" };
    if (/BLINKIT|BIGBASKET|DMART|SM CHENNAI|ZEPTO|INSTAMART/.test(n))                return { main:"Personal Exp", sub:"Groceries",         confidence:"High" };
    if (/KSRTC|INDIAN RAILWAY|IRCTC|AIRPORT METRO|METRO STATION/.test(n))            return { main:"Personal Exp", sub:"Travel",            confidence:"High" };
    if (/UBER|OLA CABS|HIGHWAY TOURIST|RAPIDO/.test(n))                              return { main:"Personal Exp", sub:"Travel",            confidence:"High" };
    if (/MEDICAL|PHARMACY|HOSPITAL|MEDICALS|APOLLO|MEDPLUS/.test(n))                 return { main:"Personal Exp", sub:"Medical",           confidence:"High" };
    if (/AMAZON|FLIPKART|MYNTRA|NYKAA|MEESHO/.test(n))                               return { main:"Personal Exp", sub:"Shopping",          confidence:"High" };
    if (/KWABEY|ALMA SOURCING/.test(n))                                               return { main:"Personal Exp", sub:"Shopping",          confidence:"High" };
    if (/TASMAC|AGS CINEMA|PVR|INOX|BOOKMYSHOW/.test(n))                             return { main:"Personal Exp", sub:"Entertainment",     confidence:"High" };
    if (/LAKSHARA|ACADEMY|UDEMY|COURSERA/.test(n))                                   return { main:"Personal Exp", sub:"Education",         confidence:"High" };
    if (/URBAN COMPANY|URBANCOMPANY/.test(n))                                         return { main:"Personal Exp", sub:"Home Services",     confidence:"High" };
    if (/BHARAT PETROLEUM|INDIAN OIL|BPCL|HPCL|IOCL|PETROL/.test(n))               return { main:"Personal Exp", sub:"Transport",         confidence:"High" };
    if (/^RTGS DR/.test(n) && amount > 50000)                                        return { main:"Transfer",     sub:"Self Transfer",      confidence:"Medium" };
    if (/^IMPS-|^NEFT-/.test(n))                                                      return { main:"Transfer",     sub:"UPI Transfer",       confidence:"Medium" };

    if (/^UPI-/.test(n)) {
      const merchantMatch = n.match(/UPI-([A-Z0-9& ]+)-/);
      const merchant = merchantMatch ? merchantMatch[1].trim() : "";
      const isP2B = /LTD|LIMITED|PRIVATE|PVT|SHOP|STORE|MART|CAFE|RESTAURANT|FOODS|KITCHEN|MEDICAL|PHARMACY|SERVICES|ACADEMY|CINEMA|HOTEL|TRAVELS|AGENCY|ENTERPRISE|COMPANY/.test(merchant);
      const isP2P = !isP2B && merchant.split(" ").length <= 4 && /^[A-Z ]+$/.test(merchant);
      if (isP2P) return { main:"Transfer", sub:"UPI Transfer", confidence:"Medium" };
    }
  }

  return { main:"NEEDS_AI", sub:"NEEDS_AI", confidence:"Low" };
}

// ─── BATCH GROQ CALL (20 transactions per request) ────────────────
function callGroqBatch(transactions) {
  if (!transactions || transactions.length === 0) return [];

  const chunkSize = 20;
  const allResults = [];

  for (let c = 0; c < transactions.length; c += chunkSize) {
    const chunk = transactions.slice(c, c + chunkSize);

    const txnList = chunk.map(function(t, i) {
      return (i+1) + ". Narration: \"" + t.narration + "\" | Amount: ₹" + t.amount + " | Type: " + t.type;
    }).join("\n");

    const prompt =
      "You are an expert Indian personal finance categorizer.\n\n" +
      "Categorize each transaction. Reply ONLY with a JSON array, no other text.\n\n" +
      "Categories:\n" +
      "Income > Salary|Bonus|Tax Refund|Investment Returns|Reimbursement|Dividend|Other Income\n" +
      "Fixed Exp > Rent|Loan EMI|Insurance|CC Bill Payment|Subscription|Standing Instruction\n" +
      "Personal Exp > Food|Groceries|Travel|Transport|Medical|Shopping|Entertainment|Education|Home Services|Miscellaneous\n" +
      "Investing > SIP|Stocks|Mutual Funds|EPFO|Gold|Fixed Deposit\n" +
      "Transfer > Self Transfer|Family Transfer|UPI Transfer\n\n" +
      "Transactions:\n" + txnList + "\n\n" +
      "Reply ONLY with JSON array:\n" +
      "[{\"main\":\"Personal Exp\",\"sub\":\"Food\",\"confidence\":\"High\"},...]";

    try {
      const response = UrlFetchApp.fetch(GROQ_URL, {
        method: "POST",
        headers: {
          "Authorization": "Bearer " + GROQ_API_KEY,
          "Content-Type": "application/json"
        },
        payload: JSON.stringify({
          model: "llama-3.3-70b-versatile",
          messages: [{ role: "user", content: prompt }],
          temperature: 0.1,
          max_tokens: 1500
        }),
        muteHttpExceptions: true
      });

      if (response.getResponseCode() !== 200) {
        for (let i = 0; i < chunk.length; i++) {
          allResults.push({ main:"Other", sub:"Unknown", confidence:"Low" });
        }
        continue;
      }

      const responseText = JSON.parse(response.getContentText()).choices[0].message.content;
      const jsonMatch = responseText.match(/\[[\s\S]*\]/);

      if (!jsonMatch) {
        for (let i = 0; i < chunk.length; i++) {
          allResults.push({ main:"Other", sub:"Unknown", confidence:"Low" });
        }
        continue;
      }

      const parsed = JSON.parse(jsonMatch[0]);
      for (let i = 0; i < parsed.length; i++) {
        allResults.push({
          main:       parsed[i].main       || "Other",
          sub:        parsed[i].sub        || "Unknown",
          confidence: parsed[i].confidence || "Medium"
        });
      }

      // Brief pause between chunks only
      if (c + chunkSize < transactions.length) Utilities.sleep(300);

    } catch(e) {
      for (let i = 0; i < chunk.length; i++) {
        allResults.push({ main:"Other", sub:"Unknown", confidence:"Low" });
      }
    }
  }

  return allResults;
}

// ─── STATUS TAB ───────────────────────────────────────────────────
function updateStatusTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let statusSheet = ss.getSheetByName("Status");
  if (!statusSheet) statusSheet = ss.insertSheet("Status");
  statusSheet.clearContents();
  statusSheet.clearFormats();

  const meta = getMetadata();
  const data = ss.getSheetByName("Categorized").getDataRange().getValues();
  const totalTxns = data.length - 1;

  const monthCounts = {};
  for (let i = 1; i < data.length; i++) {
    const my = toMonthYear(data[i][7]);
    if (my) monthCounts[my] = (monthCounts[my] || 0) + 1;
  }

  statusSheet.getRange(1,1).setValue("📊 BUDGET.AI — SHEET STATUS")
    .setFontSize(16).setFontWeight("bold").setFontColor("#0f4c81");
  statusSheet.getRange(2,1).setValue("Auto-updated every time you run categorization")
    .setFontColor("#888").setFontStyle("italic");

  statusSheet.getRange(4,1,1,2).setValues([["Metric","Value"]])
    .setBackground("#0f4c81").setFontColor("#ffffff").setFontWeight("bold");

  const metrics = [
    ["🕐 Last Sheet Updated",    meta.lastUpdated],
    ["📅 Last Transaction Date", meta.lastTxnDate],
    ["🏦 Closing Balance",       "₹" + parseFloat(meta.closingBalance).toLocaleString("en-IN")],
    ["📋 Total Transactions",    totalTxns + " transactions"],
    ["⚠️ Uncategorized",         meta.uncategorized.length + " need review"],
  ];

  for (let i = 0; i < metrics.length; i++) {
    statusSheet.getRange(5+i, 1).setValue(metrics[i][0]).setFontWeight("bold");
    statusSheet.getRange(5+i, 2).setValue(metrics[i][1]);
    statusSheet.getRange(5+i, 1, 1, 2).setBackground(i % 2 === 0 ? "#f8f9fa" : "#ffffff");
  }

  statusSheet.getRange(12,1).setValue("📆 Transactions Per Month")
    .setFontWeight("bold").setFontColor("#0f4c81").setFontSize(12);
  statusSheet.getRange(13,1,1,2).setValues([["Month","Count"]])
    .setBackground("#34495e").setFontColor("#ffffff").setFontWeight("bold");

  const sortedMonths = Object.keys(monthCounts).sort(function(a, b) {
    const partsA = a.split("-"); const partsB = b.split("-");
    return new Date("01 " + partsA[0] + " " + partsA[1]) - new Date("01 " + partsB[0] + " " + partsB[1]);
  });

  for (let i = 0; i < sortedMonths.length; i++) {
    const count = monthCounts[sortedMonths[i]];
    statusSheet.getRange(14+i, 1).setValue(sortedMonths[i]);
    statusSheet.getRange(14+i, 2).setValue(count);
    statusSheet.getRange(14+i, 1, 1, 2)
      .setBackground(count < 10 ? "#fff3cd" : "#d4edda")
      .setFontColor(count  < 10 ? "#856404" : "#155724");
  }

  statusSheet.setColumnWidth(1, 220);
  statusSheet.setColumnWidth(2, 200);
}

// ─── STEP 2: UPDATE FUND FLOW ─────────────────────────────────────
function updateFundFlow() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const rawData = ss.getSheetByName("Categorized").getDataRange().getValues();

  // Normalize Month-Year to string
  const data = rawData.map(function(row, idx) {
    if (idx === 0) return row;
    const newRow = row.slice();
    newRow[7] = toMonthYear(row[7]);
    return newRow;
  });

  // Detect FY from data
  let fyStart = new Date().getFullYear();
  for (let i = 1; i < data.length; i++) {
    const my = data[i][7];
    if (!my || !my.includes("-")) continue;
    const parts  = my.split("-");
    const monIdx = MONTH_NAMES.indexOf(parts[0]);
    if (monIdx >= 3) { fyStart = parseInt(parts[1]); break; }
  }

  const fyLabel  = "FY" + String(fyStart).slice(2) + "-" + String(fyStart+1).slice(2);
  const fyMonths = ["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"];

  const monthColMap = {};
  for (let i = 0; i < fyMonths.length; i++) {
    const yr = i <= 8 ? fyStart : fyStart + 1;
    monthColMap[fyMonths[i] + "-" + yr] = i + 2;
  }

  let ffSheet = ss.getSheetByName(fyLabel);
  if (!ffSheet) ffSheet = ss.insertSheet(fyLabel);
  ffSheet.clearContents();
  ffSheet.clearFormats();

  ffSheet.getRange(1, 1, 1, 13).setValues([["Category / Month"].concat(fyMonths)])
    .setBackground("#0f4c81").setFontColor("#ffffff").setFontWeight("bold");

  const rowLabels = [
    ["💰 Income"],["🔒 Fixed Expenses"],["🛒 Personal Expenses"],
    ["📈 Investments"],["💵 Savings"],[""],
    ["── Fixed Breakdown ──"],["  → Rent"],["  → Loan EMI"],
    ["  → CC Bill Payment"],["  → Insurance"],["  → Subscription"],
    ["  → Standing Instruction"],[""],
    ["── Personal Breakdown ──"],["  → Food"],["  → Groceries"],
    ["  → Travel"],["  → Medical"],["  → Shopping"],
    ["  → Entertainment"],["  → Education"],["  → Home Services"],
    ["  → Miscellaneous"]
  ];

  ffSheet.getRange(2, 1, rowLabels.length, 1).setValues(rowLabels);

  const rowMap = {
    "Income":2, "Fixed Exp":3, "Personal Exp":4, "Investing":5,
    "Rent":9, "Loan EMI":10, "CC Bill Payment":11, "Insurance":12,
    "Subscription":13, "Standing Instruction":14,
    "Food":17, "Groceries":18, "Travel":19, "Medical":20,
    "Shopping":21, "Entertainment":22, "Education":23,
    "Home Services":24, "Miscellaneous":25
  };

  const trackedSubs = ["Rent","Loan EMI","CC Bill Payment","Insurance","Subscription",
    "Standing Instruction","Food","Groceries","Travel","Medical","Shopping",
    "Entertainment","Education","Home Services","Miscellaneous"];

  const totals = {};
  function tKey(label, col) { return label + "||" + col; }

  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const my     = (row[7] || "").toString().trim();
    const col    = monthColMap[my];
    if (!col) continue;

    const amount = parseFloat(row[3]) || 0;
    const main   = (row[4] || "").toString();
    const sub    = (row[5] || "").toString();
    const type   = (row[2] || "").toString();

    if (type === "Credit" && main === "Income")
      totals[tKey("Income",col)] = (totals[tKey("Income",col)] || 0) + amount;
    if (type === "Debit" && main === "Fixed Exp")
      totals[tKey("Fixed Exp",col)] = (totals[tKey("Fixed Exp",col)] || 0) + amount;
    if (type === "Debit" && main === "Personal Exp")
      totals[tKey("Personal Exp",col)] = (totals[tKey("Personal Exp",col)] || 0) + amount;
    if (type === "Debit" && main === "Investing")
      totals[tKey("Investing",col)] = (totals[tKey("Investing",col)] || 0) + amount;

    if (trackedSubs.indexOf(sub) >= 0)
      totals[tKey(sub,col)] = (totals[tKey(sub,col)] || 0) + amount;
  }

  const totalKeys = Object.keys(totals);
  for (let i = 0; i < totalKeys.length; i++) {
    const parts = totalKeys[i].split("||");
    const label = parts[0];
    const col   = parseInt(parts[1]);
    const rowNum = rowMap[label];
    if (rowNum && col) ffSheet.getRange(rowNum, col).setValue(Math.round(totals[totalKeys[i]]));
  }

  for (let i = 0; i < fyMonths.length; i++) {
    const col = i + 2;
    const cl  = columnToLetter(col);
    ffSheet.getRange(6, col).setFormula(
      "=IF(" + cl + "2>0," + cl + "2-" + cl + "3-" + cl + "4-" + cl + "5,\"\")"
    );
  }

  ffSheet.getRange(2,1).setFontColor("#27ae60").setFontWeight("bold");
  ffSheet.getRange(3,1).setFontColor("#e74c3c").setFontWeight("bold");
  ffSheet.getRange(4,1).setFontColor("#e74c3c").setFontWeight("bold");
  ffSheet.getRange(5,1).setFontColor("#2980b9").setFontWeight("bold");
  ffSheet.getRange(6,1).setFontColor("#27ae60").setFontWeight("bold");
  ffSheet.setColumnWidth(1, 220);

  const meta = getMetadata();
  ffSheet.getRange(rowLabels.length + 3, 1)
    .setValue("Last updated: " + meta.lastUpdated + " | Last txn: " + meta.lastTxnDate)
    .setFontColor("#aaa").setFontStyle("italic");

  SpreadsheetApp.getUi().alert("✅ Fund Flow updated: " + fyLabel);
}

// ─── FIND DUPLICATES ──────────────────────────────────────────────
function findDuplicates() {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Categorized").getDataRange().getValues();
  const duplicates = [];

  for (let i = 1; i < data.length; i++) {
    for (let j = i + 1; j < data.length; j++) {
      const amt1  = parseFloat(data[i][3]);
      const amt2  = parseFloat(data[j][3]);
      const narr1 = (data[i][1] || "").toString().substring(0,30).toLowerCase();
      const narr2 = (data[j][1] || "").toString().substring(0,30).toLowerCase();
      const type1 = data[i][2];
      const type2 = data[j][2];
      const dp1   = (data[i][0] || "").toString().split("/");
      const dp2   = (data[j][0] || "").toString().split("/");
      if (!dp1 || dp1.length < 3) continue;
      const date1    = new Date(dp1[2] + "-" + dp1[1] + "-" + dp1[0]);
      const date2    = new Date(dp2[2] + "-" + dp2[1] + "-" + dp2[0]);
      const daysDiff = Math.abs((date1 - date2) / 86400000);
      if (amt1 === amt2 && type1 === type2 && daysDiff <= 3 && narr1 === narr2) {
        duplicates.push({
          date1:     data[i][0], date2:   data[j][0],
          narration: data[i][1], amount:  amt1,
          type:      type1,      days:    daysDiff
        });
      }
    }
  }
  return duplicates;
}

// ─── DETECT SPIKES ────────────────────────────────────────────────
function detectSpikes() {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Categorized").getDataRange().getValues();
  const monthly = {};

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] !== "Debit") continue;
    const sub = (data[i][5] || "").toString();
    const my  = toMonthYear(data[i][7]);
    const amt = parseFloat(data[i][3]) || 0;
    if (!my || !sub) continue;
    if (!monthly[sub]) monthly[sub] = {};
    monthly[sub][my] = (monthly[sub][my] || 0) + amt;
  }

  const spikes = [];
  const subKeys = Object.keys(monthly);
  for (let i = 0; i < subKeys.length; i++) {
    const sub    = subKeys[i];
    const months = monthly[sub];
    const vals   = Object.values(months);
    if (vals.length < 2) continue;
    let total = 0;
    for (let v = 0; v < vals.length; v++) total += vals[v];
    const avg     = total / vals.length;
    const mKeys   = Object.keys(months);
    for (let m = 0; m < mKeys.length; m++) {
      const monthTotal = months[mKeys[m]];
      const pct = ((monthTotal - avg) / avg) * 100;
      if (pct >= SPIKE_THRESHOLD_PERCENT) {
        spikes.push({
          month: mKeys[m], sub:   sub,
          total: Math.round(monthTotal),
          avg:   Math.round(avg),
          pct:   Math.round(pct)
        });
      }
    }
  }
  return spikes;
}

// ─── STEP 3: QUALITY CONTROL ──────────────────────────────────────
function buildQualityControl() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const qcSheet = ss.getSheetByName("Quality Control");
  qcSheet.clearContents();
  qcSheet.clearFormats();

  const duplicates    = findDuplicates();
  const spikes        = detectSpikes();
  const meta          = getMetadata();
  const uncategorized = meta.uncategorized;
  let row = 1;

  // ── SUMMARY ──
  qcSheet.getRange(row, 1, 1, 5).setBackground("#0f4c81");
  qcSheet.getRange(row, 1).setValue("🔍 QUALITY CONTROL SUMMARY")
    .setFontSize(14).setFontWeight("bold").setFontColor("#ffffff");
  row++;
  qcSheet.getRange(row, 1)
    .setValue("Last updated: " + meta.lastUpdated + "  |  Last txn: " + meta.lastTxnDate + "  |  Balance: ₹" + parseFloat(meta.closingBalance).toLocaleString("en-IN"))
    .setFontColor("#7f8c8d");
  row += 2;

  qcSheet.getRange(row, 1, 1, 3).setValues([["Metric","Count","Status"]])
    .setBackground("#34495e").setFontColor("#ffffff").setFontWeight("bold");
  row++;

  const summaryRows = [
    ["❓ Uncategorized",   uncategorized.length, uncategorized.length > 0 ? "⚠️ Review below" : "✅ All good",    uncategorized.length > 0 ? "#fff3cd" : "#d4edda"],
    ["🔴 Spending Spikes", spikes.length,        spikes.length > 0        ? "⚠️ Review below" : "✅ Consistent",  spikes.length > 0        ? "#fff3cd" : "#d4edda"],
    ["🔁 Duplicates",      duplicates.length,    duplicates.length > 0    ? "⚠️ Review below" : "✅ None found",  duplicates.length > 0    ? "#fff3cd" : "#d4edda"]
  ];
  for (let i = 0; i < summaryRows.length; i++) {
    qcSheet.getRange(row, 1, 1, 3)
      .setValues([[summaryRows[i][0], summaryRows[i][1], summaryRows[i][2]]])
      .setBackground(summaryRows[i][3]);
    row++;
  }
  row += 2;

  // ── UNCATEGORIZED ──
  qcSheet.getRange(row, 1, 1, 6).setBackground("#6c3483");
  qcSheet.getRange(row, 1).setValue("❓ UNCATEGORIZED TRANSACTIONS")
    .setFontSize(12).setFontWeight("bold").setFontColor("#ffffff");
  row++;

  if (uncategorized.length > 0) {
    const patterns = {};
    for (let i = 0; i < uncategorized.length; i++) {
      const key = uncategorized[i].narration.substring(0, 25);
      patterns[key] = (patterns[key] || 0) + 1;
    }
    const patternEntries = Object.keys(patterns).map(function(k) { return [k, patterns[k]]; });
    patternEntries.sort(function(a, b) { return b[1] - a[1]; });
    const topPatterns = patternEntries.slice(0, 5);

    qcSheet.getRange(row, 1).setValue("📊 Top unidentified patterns — add these as rules to improve accuracy:")
      .setFontWeight("bold").setFontColor("#6c3483");
    row++;
    qcSheet.getRange(row, 1, 1, 2).setValues([["Pattern","Count"]])
      .setBackground("#a569bd").setFontColor("#ffffff").setFontWeight("bold");
    row++;
    for (let i = 0; i < topPatterns.length; i++) {
      qcSheet.getRange(row, 1, 1, 2).setValues([[topPatterns[i][0], topPatterns[i][1]]])
        .setBackground("#f4ecf7");
      row++;
    }
    row++;

    qcSheet.getRange(row, 1, 1, 5).setValues([["Date","Narration","Type","Amount","AI Guess"]])
      .setBackground("#a569bd").setFontColor("#ffffff").setFontWeight("bold");
    row++;
    for (let i = 0; i < uncategorized.length; i++) {
      const t = uncategorized[i];
      qcSheet.getRange(row, 1, 1, 5).setValues([[
        t.date, t.narration.substring(0, 70), t.type,
        "₹" + t.amount, t.main + " > " + t.sub
      ]]).setBackground(i % 2 === 0 ? "#f9f0ff" : "#f4ecf7");
      row++;
    }
  } else {
    qcSheet.getRange(row, 1).setValue("✅ All transactions categorized!")
      .setBackground("#d4edda").setFontColor("#155724");
    row++;
  }
  row += 2;

  // ── SPIKES ──
  qcSheet.getRange(row, 1, 1, 5).setBackground("#c0392b");
  qcSheet.getRange(row, 1).setValue("🔴 SPENDING SPIKES")
    .setFontSize(12).setFontWeight("bold").setFontColor("#ffffff");
  row++;
  if (spikes.length > 0) {
    qcSheet.getRange(row, 1, 1, 5).setValues([["Month","Category","This Month","Avg","% Above"]])
      .setBackground("#e74c3c").setFontColor("#ffffff").setFontWeight("bold");
    row++;
    for (let i = 0; i < spikes.length; i++) {
      const sp = spikes[i];
      qcSheet.getRange(row, 1, 1, 5).setValues([[
        sp.month, sp.sub,
        "₹" + sp.total.toLocaleString("en-IN"),
        "₹" + sp.avg.toLocaleString("en-IN"),
        "+" + sp.pct + "% 🔴"
      ]]).setBackground(i % 2 === 0 ? "#fff3cd" : "#ffeeba");
      row++;
    }
  } else {
    qcSheet.getRange(row, 1).setValue("✅ No spending spikes.")
      .setBackground("#d4edda").setFontColor("#155724");
    row++;
  }
  row += 2;

  // ── DUPLICATES ──
  qcSheet.getRange(row, 1, 1, 6).setBackground("#856404");
  qcSheet.getRange(row, 1).setValue("🔁 DUPLICATE TRANSACTIONS")
    .setFontSize(12).setFontWeight("bold").setFontColor("#ffffff");
  row++;
  if (duplicates.length > 0) {
    qcSheet.getRange(row, 1, 1, 6).setValues([["Date 1","Date 2","Narration","Amount","Type","Days Apart"]])
      .setBackground("#f39c12").setFontColor("#ffffff").setFontWeight("bold");
    row++;
    for (let i = 0; i < duplicates.length; i++) {
      const d = duplicates[i];
      qcSheet.getRange(row, 1, 1, 6).setValues([[
        d.date1, d.date2, d.narration.toString().substring(0, 60),
        "₹" + d.amount, d.type, d.days + " day(s)"
      ]]).setBackground(i % 2 === 0 ? "#fffdf0" : "#fff9e6");
      row++;
    }
  } else {
    qcSheet.getRange(row, 1).setValue("✅ No duplicates.")
      .setBackground("#d4edda").setFontColor("#155724");
  }

  var colWidths = [180, 180, 320, 110, 90, 100];
  for (var i = 0; i < colWidths.length; i++) {
    qcSheet.setColumnWidth(i + 1, colWidths[i]);
  }

  SpreadsheetApp.getUi().alert(
    "✅ Quality Control done!\n" +
    "• " + uncategorized.length + " uncategorized\n" +
    "• " + duplicates.length + " duplicate(s)\n" +
    "• " + spikes.length + " spike(s)"
  );
}

// ─── WEB APP ──────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile("dashboard")
    .setTitle("Budget.AI Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName("Categorized").getDataRange().getValues();

  const props = PropertiesService.getScriptProperties();
  const meta  = {
    lastUpdated:    props.getProperty("lastUpdated")    || "Never",
    lastTxnDate:    props.getProperty("lastTxnDate")    || "Unknown",
    closingBalance: props.getProperty("closingBalance") || "0",
    uncategorized:  JSON.parse(props.getProperty("uncategorized") || "[]").length,
  };

  const monthlyTotals = {};
  for (let i = 1; i < data.length; i++) {
    const my   = toMonthYear(data[i][7]);
    const amt  = parseFloat(data[i][3]) || 0;
    const main = (data[i][4] || "").toString();
    const type = (data[i][2] || "").toString();
    if (!my) continue;
    if (!monthlyTotals[my]) monthlyTotals[my] = { Income:0, Fixed:0, Personal:0, Investing:0 };
    if (type === "Credit" && main === "Income")       monthlyTotals[my].Income    += amt;
    if (type === "Debit"  && main === "Fixed Exp")    monthlyTotals[my].Fixed     += amt;
    if (type === "Debit"  && main === "Personal Exp") monthlyTotals[my].Personal  += amt;
    if (type === "Debit"  && main === "Investing")    monthlyTotals[my].Investing += amt;
  }

  const sortedMonths = Object.keys(monthlyTotals).sort(function(a, b) {
    const pa = a.split("-"); const pb = b.split("-");
    return new Date("01 " + pa[0] + " " + pa[1]) - new Date("01 " + pb[0] + " " + pb[1]);
  });

  const monthly = sortedMonths.map(function(m) {
    return {
      month:    m,
      income:   Math.round(monthlyTotals[m].Income),
      fixed:    Math.round(monthlyTotals[m].Fixed),
      personal: Math.round(monthlyTotals[m].Personal),
      investing:Math.round(monthlyTotals[m].Investing)
    };
  });

  let tIncome = 0, tFixed = 0, tPersonal = 0, tInvest = 0;
  for (let i = 0; i < monthly.length; i++) {
    tIncome   += monthly[i].income;
    tFixed    += monthly[i].fixed;
    tPersonal += monthly[i].personal;
    tInvest   += monthly[i].investing;
  }
  const tSavings = tIncome - tFixed - tPersonal - tInvest;

  // 3-year data from FY sheets
  const fyNames   = ["FY23-24","FY24-25","FY25-26"];
  const threeYear = [];
  for (let f = 0; f < fyNames.length; f++) {
    const sheet = ss.getSheetByName(fyNames[f]);
    if (!sheet) { threeYear.push({ fy: fyNames[f], totalSavings:0, totalInvesting:0 }); continue; }
    const rows = sheet.getDataRange().getValues();
    let totalSavings = 0, totalInvesting = 0;
    for (let r = 0; r < rows.length; r++) {
      const label = (rows[r][0] || "").toString();
      const rowVals = rows[r].slice(1);
      let rowTotal = 0;
      for (let v = 0; v < rowVals.length; v++) rowTotal += parseFloat(rowVals[v]) || 0;
      if (label.includes("Savings"))     totalSavings   = rowTotal;
      if (label.includes("Investments")) totalInvesting = rowTotal;
    }
    threeYear.push({ fy: fyNames[f], totalSavings: totalSavings, totalInvesting: totalInvesting });
  }

  const today   = new Date();
  const fyStart = today.getMonth() >= 3 ? today.getFullYear() : today.getFullYear() - 1;
  const fyLabel = "FY" + String(fyStart).slice(2) + "-" + String(fyStart+1).slice(2);

  return { monthly: monthly, tIncome: tIncome, tFixed: tFixed, tPersonal: tPersonal,
           tInvest: tInvest, tSavings: tSavings, fyLabel: fyLabel,
           threeYear: threeYear, meta: meta };
}
