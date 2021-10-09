package com.tahwtice.apps;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelReporter {

    public void exportOrigin(Workbook workbook) {
        try {
            Sheet sheet = workbook.getSheetAt(0);
            int count = sheet.getLastRowNum();
            Row row;
            while (count > 0) {
                count--;
                row = sheet.getRow(count);
                if (row == null) {
                    sheet.shiftRows(count + 1, sheet.getLastRowNum(), -1);
                }
            }

            FileOutputStream out = new FileOutputStream(Constants.EXCEL_PATH_BILLING);
            workbook.write(out);
            out.close();
            System.out.println("File written successfully");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
