package com.tanmay.expense_tracker.service;

import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.List;
import java.util.Locale;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.ValueRange;

@Service
public class ExpenseService {

    @Autowired
    private Sheets sheets;

    private final String SPREADSHEET_ID = "1SxhtzmKLZ-vcYSddr-BvW5IjirU_e1SIhKEZgizne_M";

    public String processExpense(int amount, String dateStr) {
        try {
            LocalDate date = LocalDate.parse(dateStr);

            String sheetName = date.getMonth()
                    .getDisplayName(TextStyle.FULL, Locale.ENGLISH)
                    + "/" + date.getYear();

            String formattedDate = date.getDayOfMonth() + "/"
                    + date.getMonthValue() + "/" + date.getYear();

            // 1. Fetch sheet data
            ValueRange response = sheets.spreadsheets().values()
                    .get(SPREADSHEET_ID, sheetName + "!A:C")
                    .execute();

            List<List<Object>> rows = response.getValues();

            if (rows == null) {
                return "Sheet not found";
            }

            // 2. Find row
            int rowIndex = -1;
            for (int i = 0; i < rows.size(); i++) {
                if (rows.get(i).get(0).toString().equals(formattedDate)) {
                    rowIndex = i + 1;
                    break;
                }
            }

            if (rowIndex == -1) {
                return "Date row not found";
            }

            List<Object> row = rows.get(rowIndex - 1);

            int expense = Integer.parseInt(row.get(1).toString());

            expense += amount;
            // 3. Update
            ValueRange body = new ValueRange()
                    .setValues(List.of(List.of(expense)));

            sheets.spreadsheets().values()
                    .update(SPREADSHEET_ID,
                            sheetName + "!B" + rowIndex,
                            body)
                    .setValueInputOption("RAW")
                    .execute();

            return "Updated successfully";

        } catch (Exception e) {
            return "Error: " + e.getMessage();
        }
    }
}
