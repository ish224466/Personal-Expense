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
const TRIPS_SHEET_NAME = "Trips";
export const TRIP_CATEGORIES = [
  "Food",
  "Stay",
  "Travel",
  "Booking/Tickets",
  "Shopping",
  "Utilities",
  "Rent",
  "Entertainment",
  "Subscription",
  "Home",
  "Medical",
  "Other",
];
const TRIPS_HEADER_ROW = ["Date", "Destination", "Trip Start", "Trip End", ...TRIP_CATEGORIES, "Total", NOTES_HEADER, "Photo"];
const LEGACY_CURRENT_TRIPS_HEADER_ROW = ["Date", "Destination", ...TRIP_CATEGORIES, "Total", NOTES_HEADER, "Photo"];
const PREVIOUS_TRIPS_HEADER_ROW = [
  "Date", "Destination", "Food", "Stay", "Travel", "Shopping", "Utilities", "Rent",
  "Entertainment", "Subscription", "Home", "Medical", "Other", "Total", NOTES_HEADER,
];
const TRIPS_START_INDEX = 2;
const TRIPS_END_INDEX = 3;
const TRIPS_TOTAL_INDEX = 4 + TRIP_CATEGORIES.length;
const TRIPS_NOTES_INDEX = TRIPS_TOTAL_INDEX + 1;
const TRIPS_PHOTO_INDEX = TRIPS_TOTAL_INDEX + 2;
let tripsPreparation = null;

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
  const normalizedDate = String(dateStr ?? "").trim();
  let date;
  const ymdMatch = /^(\d{4})-(\d{2})-(\d{2})$/.exec(normalizedDate);

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
    date = new Date(normalizedDate);

    if (Number.isNaN(date.getTime())) {
      throw new Error(`Invalid date: ${normalizedDate || "missing"}`);
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

function formatTripDate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function parseSheetDate(value, monthIndex, year) {
  const match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(value || ""));
  if (!match) return null;

  const date = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
  return date.getFullYear() === year && date.getMonth() === monthIndex && date.getDate() === Number(match[1])
    ? formatTripDate(date)
    : null;
}

function normalizeTripSheetDate(value) {
  const text = String(value || "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  const match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);
  if (!match) return "";

  const date = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
  if (date.getFullYear() !== Number(match[3]) || date.getMonth() !== Number(match[2]) - 1 || date.getDate() !== Number(match[1])) {
    return "";
  }

  return formatTripDate(date);
}

async function ensureTripsSheet(sheets) {
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const tripSheet = spreadsheet.data.sheets?.find(
    (sheet) => sheet.properties?.title === TRIPS_SHEET_NAME,
  );

  if (!tripSheet) throw new Error(`Missing ${TRIPS_SHEET_NAME} sheet`);
  return tripSheet;
}

async function getTripRows(sheets) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${TRIPS_SHEET_NAME}!A2:${getColumnLetter(TRIPS_HEADER_ROW.length)}`,
  });
  return response.data.values || [];
}

async function getExistingTripRows(sheets) {
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const tripSheetExists = spreadsheet.data.sheets?.some(
    (sheet) => sheet.properties?.title === TRIPS_SHEET_NAME,
  );

  if (!tripSheetExists) return [];
  return getTripRows(sheets);
}

async function migrateLegacyTrips(sheets) {
  throw new Error("Legacy trip migration is disabled");
  /*
  await ensureTripsSheet(sheets);
  const existingResponse = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${TRIPS_SHEET_NAME}!A1:${getColumnLetter(TRIPS_HEADER_ROW.length)}`,
  });
  const existingValues = existingResponse.data.values || [];
  const existingHeader = existingValues[0] || [];
  const isCurrentFormat = JSON.stringify(existingHeader) === JSON.stringify(TRIPS_HEADER_ROW);
  const isLegacyCurrentFormat = JSON.stringify(existingHeader) === JSON.stringify(LEGACY_CURRENT_TRIPS_HEADER_ROW);
  const isPreviousFormat = JSON.stringify(existingHeader) === JSON.stringify(PREVIOUS_TRIPS_HEADER_ROW);
  const groupedRows = new Map();
  let legacyImported = false;

  const addToGroup = (destination, date, category, amount, notes = "", tripStart = "", tripEnd = "", photo = "") => {
    const cleanDate = normalizeTripSheetDate(date) || String(date || "").trim() || "undated";
    const cleanDestination = String(destination || "Unassigned").trim() || "Unassigned";
    const numericAmount = Number(amount || 0) || 0;
    const key = `${cleanDestination}\u0000${cleanDate}`;
    if (!TRIP_CATEGORIES.includes(category) || (numericAmount <= 0 && cleanDate !== "undated")) return;
    if (!groupedRows.has(key)) {
      groupedRows.set(key, {
        destination: cleanDestination,
        date: cleanDate,
        categories: Object.fromEntries(TRIP_CATEGORIES.map((name) => [name, 0])),
        notes: [],
        tripStart,
        tripEnd,
        photo,
      });
    }
    const group = groupedRows.get(key);
    group.categories[category] += numericAmount;
    if (notes && !group.notes.includes(notes)) group.notes.push(notes);
    if (tripStart) group.tripStart = tripStart;
    if (tripEnd) group.tripEnd = tripEnd;
    if (photo) group.photo = photo;
  };

  for (const row of existingValues.slice(1)) {
    if (isCurrentFormat || isLegacyCurrentFormat || isPreviousFormat) {
      const notes = String(row[existingHeader.indexOf(NOTES_HEADER)] || "");
      const tripStart = isCurrentFormat ? String(row[existingHeader.indexOf("Trip Start")] || "") : "";
      const tripEnd = isCurrentFormat ? String(row[existingHeader.indexOf("Trip End")] || "") : "";
      const photo = isCurrentFormat ? String(row[existingHeader.indexOf("Photo")] || "") : "";
      TRIP_CATEGORIES.forEach((category) => {
        const sourceIndex = existingHeader.indexOf(category);
        if (sourceIndex >= 0) addToGroup(row[1], row[0], category, row[sourceIndex], notes, tripStart, tripEnd, photo);
      });
    } else {
      const isSimpleLegacyFormat = existingHeader[0] === "Date" && existingHeader[1] === "Destination";
      const date = isSimpleLegacyFormat ? row[0] : row[2];
      const destination = isSimpleLegacyFormat ? row[1] : row[1];
      const rawCategory = isSimpleLegacyFormat ? row[2] : row[3];
      const category = rawCategory === "Commute" || rawCategory === "Travel"
        ? "Travel"
        : TRIP_CATEGORIES.includes(rawCategory) ? rawCategory : "Other";
      const amount = isSimpleLegacyFormat ? row[3] : row[4];
      const notes = isSimpleLegacyFormat ? row[4] : row[5];
      const source = !isSimpleLegacyFormat && row[6] ? `Legacy source: ${row[6]}` : "";
      addToGroup(destination || "Unassigned", date, category, amount, [notes, source].filter(Boolean).join(" | "));
    }
  }

  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
  ];
  const legacySheets = (spreadsheet.data.sheets || [])
    .map((sheet) => sheet.properties?.title)
    .filter((title) => {
      const [month, year] = String(title || "").split("/");
      return monthNames.includes(month) && /^\d{4}$/.test(year);
    });

  for (const sheetName of legacySheets) {
    const [monthName, yearText] = sheetName.split("/");
    const monthIndex = monthNames.indexOf(monthName);
    const year = Number(yearText);
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A1:N`,
    });

    for (const row of (response.data.values || []).slice(2)) {
      const tripAmount = Number(row[EXPENSE_CATEGORIES.indexOf("Trip") + 1] || 0) || 0;
      const date = parseSheetDate(row[0], monthIndex, year);
      if (!date || tripAmount <= 0) continue;

      const source = `Legacy source: ${sheetName}:${row[0]}`;
      const previousSource = `Legacy source: legacy:${sheetName}:${row[0]}`;
      const alreadyImported = [...groupedRows.values()].some((group) => (
        group.notes.includes(source) || group.notes.includes(previousSource)
      ));
      if (!alreadyImported) {
        addToGroup("Unassigned", date, "Other", tripAmount, source);
        legacyImported = true;
      }
    }
  }

  const outputRows = [...groupedRows.values()].map((group) => [
    group.date,
    group.destination,
    group.tripStart || "",
    group.tripEnd || "",
    ...TRIP_CATEGORIES.map((category) => group.categories[category] || 0),
    TRIP_CATEGORIES.reduce((sum, category) => sum + (group.categories[category] || 0), 0),
    group.notes.join(" | "),
    group.photo || "",
  ]);
  const needsRewrite = !isCurrentFormat || legacyImported || outputRows.length !== existingValues.slice(1).length;

  if (needsRewrite) {
    await sheets.spreadsheets.values.clear({
      spreadsheetId: SPREADSHEET_ID,
      range: `${TRIPS_SHEET_NAME}!A:Z`,
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${TRIPS_SHEET_NAME}!A1:${getColumnLetter(TRIPS_HEADER_ROW.length)}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [TRIPS_HEADER_ROW, ...outputRows] },
    });
  }
  */
}

async function prepareTrips(sheets) {
  if (!tripsPreparation) {
    tripsPreparation = migrateLegacyTrips(sheets).finally(() => {
      tripsPreparation = null;
    });
  }

  await tripsPreparation;
}

function getTripSummary(rows, month) {
  const monthRows = rows.filter((row) => {
    const date = normalizeTripSheetDate(row[0]);
    const hasDestination = String(row[1] || "").trim();
    return hasDestination && (date.startsWith(`${month}-`) || !date);
  });
  const byDate = {};

  for (const row of monthRows) {
    const date = normalizeTripSheetDate(row[0]) || "undated";
    const destination = row[1] || "Unassigned";
    const tripStart = row[TRIPS_START_INDEX] || "";
    const tripEnd = row[TRIPS_END_INDEX] || "";
    const notes = row[TRIPS_NOTES_INDEX] || "";
    const photo = row[TRIPS_PHOTO_INDEX] || "";
    const categories = Object.fromEntries(TRIP_CATEGORIES.map((category, index) => [
      category,
      Number(row[index + 4] || 0) || 0,
    ]));
    const numericAmount = Object.values(categories).reduce((sum, amount) => sum + amount, 0);
    if (!String(row[1] || "").trim()) continue;
    byDate[date] ||= { total: 0, destinations: {} };
    byDate[date].total += numericAmount;
    byDate[date].destinations[destination] ||= { total: 0, categories: {}, entries: [], photo: "", tripStart, tripEnd };
    const destinationData = byDate[date].destinations[destination];
    if (tripStart) destinationData.tripStart = tripStart;
    if (tripEnd) destinationData.tripEnd = tripEnd;
    if (photo) destinationData.photo = photo;
    destinationData.total += numericAmount;
    Object.entries(categories).forEach(([category, amount]) => {
      if (amount > 0) destinationData.categories[category] = (destinationData.categories[category] || 0) + amount;
    });
    destinationData.entries.push({ date, categories, amount: numericAmount, notes, photo });
  }

  return byDate;
}

export async function addTripExpense({ amount, dateStr, date: requestDate, destination, category, notes = "", photoPath = "", tripStart = "", tripEnd = "" }) {
  const sheets = await getSheetsClient();
  const tripDateValue = dateStr || requestDate;
  const date = parseLocalDate(tripDateValue);
  const cleanDestination = String(destination || "").trim().slice(0, 120);
  const cleanNotes = String(notes || "").trim().slice(0, 500);
  const cleanTripStart = String(tripStart || "").trim();
  const cleanTripEnd = String(tripEnd || "").trim();

  if (!cleanDestination) throw new Error("Destination required");
  if (!TRIP_CATEGORIES.includes(category)) throw new Error("Invalid trip category");
  if (cleanTripStart) parseLocalDate(cleanTripStart);
  if (cleanTripEnd) parseLocalDate(cleanTripEnd);
  if (cleanTripStart && cleanTripEnd && parseLocalDate(cleanTripEnd) < parseLocalDate(cleanTripStart)) {
    throw new Error("Trip end date must be on or after the start date");
  }

  await ensureTripsSheet(sheets);
  const rows = await getTripRows(sheets);
  const tripDate = formatTripDate(date);
  const rowIndex = rows.findIndex((row) => normalizeTripSheetDate(row[0]) === tripDate && row[1] === cleanDestination);
  const outputRow = rowIndex >= 0
    ? [...rows[rowIndex], ...Array(Math.max(0, TRIPS_HEADER_ROW.length - rows[rowIndex].length)).fill(0)]
    : [tripDate, cleanDestination, cleanTripStart, cleanTripEnd, ...TRIP_CATEGORIES.map(() => 0), 0, "", ""];
  outputRow[0] = tripDate;
  outputRow[1] = cleanDestination;
  if (cleanTripStart) outputRow[TRIPS_START_INDEX] = cleanTripStart;
  if (cleanTripEnd) outputRow[TRIPS_END_INDEX] = cleanTripEnd;
  const categoryIndex = TRIP_CATEGORIES.indexOf(category) + 4;
  outputRow[categoryIndex] = (Number(outputRow[categoryIndex] || 0) || 0) + Number(amount);
  outputRow[TRIPS_TOTAL_INDEX] = TRIP_CATEGORIES.reduce((sum, name, index) => sum + (Number(outputRow[index + 4] || 0) || 0), 0);
  outputRow[TRIPS_NOTES_INDEX] = [outputRow[TRIPS_NOTES_INDEX], cleanNotes].filter(Boolean).join(" | ");
  if (photoPath) outputRow[TRIPS_PHOTO_INDEX] = photoPath;

  if (rowIndex >= 0) {
    const sheetRow = rowIndex + 2;
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${TRIPS_SHEET_NAME}!A${sheetRow}:${getColumnLetter(TRIPS_HEADER_ROW.length)}${sheetRow}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [outputRow] },
    });
  } else {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${TRIPS_SHEET_NAME}!A:${getColumnLetter(TRIPS_HEADER_ROW.length)}`,
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: { values: [outputRow] },
    });
  }

  return "Trip expense added";
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

    const tripRows = await getExistingTripRows(sheets);
  const tripSummary = getTripSummary(tripRows, `${year}-${String(requestedMonth.getMonth() + 1).padStart(2, "0")}`);
  const rows = await ensureSheetStructure(sheets, sheetName, year, requestedMonth.getMonth());

  let expenses = [];
  let total = 0;
  let categoryTotals = Object.fromEntries(EXPENSE_CATEGORIES.map((name) => [name, 0]));

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    if (!row || !row[0]) continue;

    const date = row[0];
    const dateParts = String(date).split("/").map(Number);
    const isoDate = dateParts.length === 3
      ? `${dateParts[2]}-${String(dateParts[1]).padStart(2, "0")}-${String(dateParts[0]).padStart(2, "0")}`
      : "";
    const rowCategories = Object.fromEntries(EXPENSE_CATEGORIES.map((name, index) => [
      name,
      name === "Trip" ? Number(tripSummary[isoDate]?.total || 0) : Number(row[index + 1] || 0) || 0,
    ]));
    const expense = Object.values(rowCategories).reduce((sum, value) => sum + value, 0);

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
      trip: tripSummary[isoDate] || null,
    });
  }

  const trips = Object.entries(tripSummary).map(([date, summary]) => ({ date, ...summary }));

  return {
    month: `${year}-${String(requestedMonth.getMonth() + 1).padStart(2, "0")}`,
    total,
    categoryTotals,
    expenses,
    trips,
  };
}
