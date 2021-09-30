package com.tahwtice.apps;

import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelParser {

    private static final String EXCEL_PATH = "src/main/resources/src.xlsx";

    public void parse() {
        try {
            FileInputStream file = new FileInputStream(EXCEL_PATH);

            // Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            // Iterate through each rows one by one
            for (Row row : sheet) {
                // For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    System.out.println(cell.getStringCellValue() + "   ");
                }
            }

            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
