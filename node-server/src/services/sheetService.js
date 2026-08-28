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
  "Medical",
  "Other",
];

const NOTES_HEADER = "Notes";
const HEADER_ROW = ["Date", ...EXPENSE_CATEGORIES, "Total", NOTES_HEADER];
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

function getDaysInMonth(year, monthIndex) {
  return new Date(year, monthIndex + 1, 0).getDate();
}

function formatExpenseDate(day, monthIndex, year) {
  return `${day}/${monthIndex + 1}/${year}`;
}

function buildEmptyMonthRows(year, monthIndex) {
  const daysInMonth = getDaysInMonth(year, monthIndex);
  const rows = [];

  for (let day = 1; day <= daysInMonth; day++) {
    rows.push([
      formatExpenseDate(day, monthIndex, year),
      ...EXPENSE_CATEGORIES.map(() => 0),
      0,
      "",
    ]);
  }

  return rows;
}

function buildSummaryRow(lastDataRowIndex) {
  return [
    SUMMARY_ROW_LABEL,
    ...EXPENSE_CATEGORIES.map((_, index) => `=SUM(${getColumnLetter(index + 2)}3:${getColumnLetter(index + 2)}${lastDataRowIndex})`),
    `=SUM(${getColumnLetter(EXPENSE_CATEGORIES.length + 2)}3:${getColumnLetter(EXPENSE_CATEGORIES.length + 2)}${lastDataRowIndex})`,
  ];
}

async function ensureMonthlySheetExists(sheets, sheetName) {
  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
  });

  const existingSheet = spreadsheet.data.sheets?.find(
    (sheet) => sheet.properties?.title === sheetName,
  );

  if (existingSheet) {
    const sheetId = existingSheet.properties?.sheetId;
    const columnCount = existingSheet.properties?.gridProperties?.columnCount || 0;

    if (sheetId !== undefined && columnCount < HEADER_ROW.length) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          requests: [
            {
              updateSheetProperties: {
                properties: {
                  sheetId,
                  gridProperties: {
                    columnCount: HEADER_ROW.length,
                  },
                },
                fields: "gridProperties.columnCount",
              },
            },
          ],
        },
      });
    }

    return;
  }

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          addSheet: {
            properties: {
              title: sheetName,
              gridProperties: {
                rowCount: 200,
                columnCount: HEADER_ROW.length,
              },
            },
          },
        },
      ],
    },
  });
}

function mergeExistingRowsIntoMonthGrid(existingRows, year, monthIndex) {
  const monthRows = buildEmptyMonthRows(year, monthIndex);
  const existingHeader = Array.isArray(existingRows[1]) ? existingRows[1] : [];
  const existingColumnIndexes = Object.fromEntries(
    existingHeader.map((header, index) => [header, index]),
  );
  const hasNamedHeader = existingHeader.length > 0;

  for (let i = 2; i < existingRows.length; i++) {
    const row = existingRows[i];

    if (!row || !row[0]) {
      continue;
    }

    const day = Number(String(row[0]).split("/")[0]);

    if (!day || day < 1 || day > monthRows.length) {
      continue;
    }

    const targetRow = monthRows[day - 1];

    if (hasNamedHeader) {
      for (let columnIndex = 1; columnIndex < HEADER_ROW.length; columnIndex++) {
        const header = HEADER_ROW[columnIndex];
        const sourceIndex = existingColumnIndexes[header];

        if (sourceIndex === undefined) {
          continue;
        }

        targetRow[columnIndex] = header === NOTES_HEADER
          ? String(row[sourceIndex] || "")
          : Number(row[sourceIndex] || 0) || 0;
      }
    } else if (row.length >= HEADER_ROW.length - 1) {
      for (let columnIndex = 1; columnIndex < HEADER_ROW.length; columnIndex++) {
        targetRow[columnIndex] = columnIndex === HEADER_ROW.length - 1
          ? String(row[columnIndex] || "")
          : Number(row[columnIndex] || 0) || 0;
      }
    } else {
      const legacyAmount = Number(row[1] || 0) || 0;
      targetRow[EXPENSE_CATEGORIES.length] = legacyAmount;
      targetRow[EXPENSE_CATEGORIES.length + 1] = legacyAmount;
    }
  }

  return monthRows;
}

async function writeNewStructure(sheets, sheetName, dataRows) {
  const lastColumn = getColumnLetter(HEADER_ROW.length);
  const lastDataRowIndex = dataRows.length + 2;

  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A:${lastColumn}`,
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A1`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [buildSummaryRow(lastDataRowIndex), HEADER_ROW, ...dataRows],
    },
  });
}

async function ensureSheetStructure(sheets, sheetName, year, monthIndex) {
  const lastColumn = getColumnLetter(HEADER_ROW.length);
  await ensureMonthlySheetExists(sheets, sheetName);

  let res;
  try {
    res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A1:${lastColumn}`,
    });
  } catch (error) {
    if (String(error?.message || "").includes("Unable to parse range")) {
      await ensureMonthlySheetExists(sheets, sheetName);
      res = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!A1:${lastColumn}`,
      });
    } else {
      throw error;
    }
  }

  const rows = res.data.values || [];

  const normalizedRows = mergeExistingRowsIntoMonthGrid(rows, year, monthIndex);

  await writeNewStructure(sheets, sheetName, normalizedRows);
  return normalizedRows;
}

export async function addExpense(amount, dateStr, category, notes = "") {
  const sheets = await getSheetsClient();

  if (!EXPENSE_CATEGORIES.includes(category)) {
    throw new Error("Invalid category");
  }

  const date = parseLocalDate(dateStr);

  const month = date.toLocaleString("default", { month: "long" });
  const year = date.getFullYear();

  const sheetName = `${month}/${year}`;
  const formattedDate = formatExpenseDate(date.getDate(), date.getMonth(), year);
  const categoryIndex = EXPENSE_CATEGORIES.indexOf(category);
  const categoryColumnIndex = categoryIndex + 2;
  const totalColumnIndex = EXPENSE_CATEGORIES.length + 2;

  const dataRows = await ensureSheetStructure(sheets, sheetName, year, date.getMonth());
  const rowIndex = date.getDate() + 2;
  const currentRows = dataRows;

  // CASE 1: Row exists -> update the selected category and total.
  const existingRow = currentRows[date.getDate() - 1] || [];
  const row = HEADER_ROW.map((_, index) => existingRow[index] ?? 0);
  const existingCategoryAmount = Number(row[categoryColumnIndex - 1] || 0) || 0;
  const existingTotal = Number(row[totalColumnIndex - 1] || 0) || 0;
  const updatedCategoryAmount = existingCategoryAmount + amount;
  const updatedTotal = existingTotal + amount;

  row[0] = formattedDate;
  row[categoryColumnIndex - 1] = updatedCategoryAmount;
  row[totalColumnIndex - 1] = updatedTotal;
  const notesColumnIndex = HEADER_ROW.length - 1;
  const existingNotes = String(row[notesColumnIndex] || "").trim();
  const cleanNotes = String(notes || "").trim().slice(0, 500);
  row[notesColumnIndex] = [existingNotes, cleanNotes].filter(Boolean).join(" | ");

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

export async function getMonthlyExpenses(monthParam) {
  const sheets = await getSheetsClient();

  const today = new Date();
  const requestedMonth = /^\d{4}-\d{2}$/.test(monthParam || "")
    ? parseLocalDate(`${monthParam}-01`)
    : today;
  const month = requestedMonth.toLocaleString("default", { month: "long" });
  const year = requestedMonth.getFullYear();

  const sheetName = `${month}/${year}`;

  const rows = await ensureSheetStructure(sheets, sheetName, year, requestedMonth.getMonth());

  let expenses = [];
  let total = 0;
  let categoryTotals = Object.fromEntries(EXPENSE_CATEGORIES.map((name) => [name, 0]));

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    if (!row || !row[0]) continue;

    const date = row[0];
    const rowCategories = Object.fromEntries(EXPENSE_CATEGORIES.map((name, index) => [name, Number(row[index + 1] || 0) || 0]));
    const expense = Number(row[EXPENSE_CATEGORIES.length + 1] || 0) || 0;

    if (expense === 0 && Object.values(rowCategories).every((value) => value === 0)) {
      continue;
    }

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
    month: `${year}-${String(requestedMonth.getMonth() + 1).padStart(2, "0")}`,
    total,
    categoryTotals,
    expenses,
  };
}