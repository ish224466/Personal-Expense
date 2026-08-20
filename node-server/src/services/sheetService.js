import { getSheetsClient } from "../config/google.js";
import dotenv from "dotenv";

dotenv.config();
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

export const EXPENSE_CATEGORIES = [
  "Food",
  "Commute",
  "Utilities",
  "Shopping",
  "Rent",
  "Entertainment",
  "Trip",
  "Subscription",
  "Home",
  "Other",
];

const HEADER_ROW = ["Date", ...EXPENSE_CATEGORIES, "Total"];
const SUMMARY_ROW_LABEL = "Total";

function getColumnLetter(index) {
  let result = "";
  let value = index;

  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }

  return result;
}

function parseLocalDate(dateStr) {
  let date;
  const ymdMatch = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateStr);

  if (ymdMatch) {
    const year = Number(ymdMatch[1]);
    const monthIndex = Number(ymdMatch[2]) - 1;
    const day = Number(ymdMatch[3]);
    date = new Date(year, monthIndex, day);

    if (
      date.getFullYear() !== year ||
      date.getMonth() !== monthIndex ||
      date.getDate() !== day
    ) {
      throw new Error("Invalid date");
    }
  } else {
    date = new Date(dateStr);

    if (Number.isNaN(date.getTime())) {
      throw new Error("Invalid date");
    }
  }

  return date;
}

function buildSummaryRow() {
  return [
    SUMMARY_ROW_LABEL,
    ...EXPENSE_CATEGORIES.map((_, index) => `=SUM(${getColumnLetter(index + 2)}3:${getColumnLetter(index + 2)})`),
    `=SUM(${getColumnLetter(EXPENSE_CATEGORIES.length + 2)}3:${getColumnLetter(EXPENSE_CATEGORIES.length + 2)})`,
  ];
}

function isNewStructure(rows) {
  if (!rows || rows.length < 2) {
    return false;
  }

  const headerRow = rows[1] || [];
  return HEADER_ROW.every((value, index) => headerRow[index] === value);
}

async function writeNewStructure(sheets, sheetName, dataRows) {
  const lastColumn = getColumnLetter(HEADER_ROW.length);

  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A:${lastColumn}`,
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A1`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [buildSummaryRow(), HEADER_ROW, ...dataRows],
    },
  });
}

async function ensureSheetStructure(sheets, sheetName) {
  const lastColumn = getColumnLetter(HEADER_ROW.length);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A1:${lastColumn}`,
  });

  const rows = res.data.values || [];

  if (!rows.length) {
    await writeNewStructure(sheets, sheetName, []);
    return [];
  }

  if (isNewStructure(rows)) {
    return rows.slice(2);
  }

  const migratedRows = [];

  for (let i = 2; i < rows.length; i++) {
    const row = rows[i];

    if (!row || !row[0]) {
      continue;
    }

    const legacyAmount = Number(row[1] || 0) || 0;
    migratedRows.push([
      row[0],
      ...EXPENSE_CATEGORIES.slice(0, -1).map(() => 0),
      legacyAmount,
      legacyAmount,
    ]);
  }

  await writeNewStructure(sheets, sheetName, migratedRows);
  return migratedRows;
}

export async function addExpense(amount, dateStr, category) {
  const sheets = await getSheetsClient();

  if (!EXPENSE_CATEGORIES.includes(category)) {
    throw new Error("Invalid category");
  }

  const date = parseLocalDate(dateStr);

  const month = date.toLocaleString("default", { month: "long" });
  const year = date.getFullYear();

  const sheetName = `${month}/${year}`;
  const formattedDate = `${date.getDate()}/${date.getMonth() + 1}/${year}`;
  const categoryIndex = EXPENSE_CATEGORIES.indexOf(category);
  const categoryColumnIndex = categoryIndex + 2;
  const totalColumnIndex = EXPENSE_CATEGORIES.length + 2;

  const dataRows = await ensureSheetStructure(sheets, sheetName);

  let rowIndex = -1;
  let currentRows = dataRows;

  for (let i = 0; i < currentRows.length; i++) {
    if (currentRows[i][0] === formattedDate) {
      rowIndex = i + 3;
      break;
    }
  }

  // CASE 1: Row exists -> update the selected category and total.
  if (rowIndex !== -1) {
    const existingRow = currentRows[rowIndex - 3] || [];
    const row = HEADER_ROW.map((_, index) => existingRow[index] ?? 0);
    const existingCategoryAmount = Number(row[categoryColumnIndex - 1] || 0) || 0;
    const existingTotal = Number(row[totalColumnIndex - 1] || 0) || 0;
    const updatedCategoryAmount = existingCategoryAmount + amount;
    const updatedTotal = existingTotal + amount;

    row[categoryColumnIndex - 1] = updatedCategoryAmount;
    row[totalColumnIndex - 1] = updatedTotal;

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A${rowIndex}:${getColumnLetter(HEADER_ROW.length)}${rowIndex}`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [row],
      },
    });

    return "Updated existing date";
  }

  // CASE 2: Row does NOT exist -> create a new row.
  const newRow = [
    formattedDate,
    ...EXPENSE_CATEGORIES.map(() => 0),
    0,
  ];

  newRow[categoryColumnIndex - 1] = amount;
  newRow[totalColumnIndex - 1] = amount;

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A:${getColumnLetter(HEADER_ROW.length)}`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [newRow],
    },
  });

  return "Created new date entry";
}

export async function getMonthlyExpenses() {
  const sheets = await getSheetsClient();

  const today = new Date();
  const month = today.toLocaleString("default", { month: "long" });
  const year = today.getFullYear();

  const sheetName = `${month}/${year}`;

  const rows = await ensureSheetStructure(sheets, sheetName);

  let expenses = [];
  let total = 0;
  let categoryTotals = Object.fromEntries(EXPENSE_CATEGORIES.map((name) => [name, 0]));

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    if (!row || !row[0]) continue;

    const date = row[0];
    const rowCategories = Object.fromEntries(EXPENSE_CATEGORIES.map((name, index) => [name, Number(row[index + 1] || 0) || 0]));
    const expense = Number(row[EXPENSE_CATEGORIES.length + 1] || 0) || 0;

    total += expense;
    for (const category of EXPENSE_CATEGORIES) {
      categoryTotals[category] += rowCategories[category];
    }

    expenses.push({
      date,
      categories: rowCategories,
      expense,
    });
  }

  return {
    total,
    categoryTotals,
    expenses,
  };
}