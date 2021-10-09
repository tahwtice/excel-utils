package com.tahwtice.apps;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.tahwtice.apps.model.Billing;
import com.tahwtice.apps.model.Excel;

public class ExcelReporter {

    public void exportOrigin(Excel<Billing> billingExcel) {
        try {
            Sheet sheet = billingExcel.getWorkbook().getSheetAt(0);
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
            billingExcel.getWorkbook().write(out);
            out.close();
            System.out.println("File written successfully");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
