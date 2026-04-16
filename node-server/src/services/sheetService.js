import { getSheetsClient } from "../config/google.js";
import dotenv from "dotenv";

dotenv.config();
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

export async function addExpense(amount, dateStr) {
  const sheets = await getSheetsClient();

  // Parse YYYY-MM-DD as a local calendar date to avoid timezone shifts.
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

  const month = date.toLocaleString("default", { month: "long" });
  const year = date.getFullYear();

  const sheetName = `${month}/${year}`;
  const formattedDate = `${date.getDate()}/${date.getMonth() + 1}/${year}`;

  // 1. Get sheet data
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A:C`,
  });

  const rows = res.data.values || [];

  let rowIndex = -1;

  // 2. Find row
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === formattedDate) {
      rowIndex = i + 1;
      break;
    }
  }

  // ✅ CASE 1: Row exists → update
  if (rowIndex !== -1) {
    const row = rows[rowIndex - 1];

    let expense = parseInt(row[1] || 0);
    expense += amount;

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!B${rowIndex}`,
      valueInputOption: "RAW",
      requestBody: {
        values: [[expense]],
      },
    });

    return "Updated existing date";
  }

  // ✅ CASE 2: Row does NOT exist → create new row
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A:C`,
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [[formattedDate, amount]],
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

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A:C`,
  });

  const rows = res.data.values;

  if (!rows) throw new Error("No data");

  let expenses = [];
  let total = 0;

  // skip header rows (assumes row 1-2 headers)
  for (let i = 2; i < rows.length; i++) {
    const row = rows[i];

    if (!row || !row[0]) continue;

    const date = row[0];
    const expense = parseInt(row[1] || 0);

    total += expense;

    expenses.push({ date, expense });
  }

  return {
    total,
    expenses
  };
}