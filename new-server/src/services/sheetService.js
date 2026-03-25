import { getSheetsClient } from "../config/google.js";
import dotenv from "dotenv";

dotenv.config();
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

export async function addExpense(amount, dateStr) {
  const sheets = await getSheetsClient();

  const date = new Date(dateStr);

  const month = date.toLocaleString("default", { month: "long" });
  const year = date.getFullYear();

  const sheetName = `${month}/${year}`;
  const formattedDate = `${date.getDate()}/${date.getMonth() + 1}/${year}`;

  // 1. Get sheet data
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A:C`,
  });

  const rows = res.data.values;

  if (!rows) throw new Error("Sheet not found");

  // 2. Find row
  let rowIndex = -1;

  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === formattedDate) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) throw new Error("Date row not found");

  const row = rows[rowIndex - 1];

  let expense = parseInt(row[1] || 0);
  expense += amount;

  // 3. Update ONLY column B (important)
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!B${rowIndex}`,
    valueInputOption: "RAW",
    requestBody: {
      values: [[expense]],
    },
  });

  return "Updated successfully";
}