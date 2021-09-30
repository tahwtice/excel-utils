package com.tahwtice.apps;

import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReporter {

    private static final String EXCEL_PATH = "src/main/resources/dist.xlsx";

    private static final String[] columns = {"ID", "FirstName", "LastName"};

    public void export() {
        // Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        // Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Cricketer Data");

        // This data needs to be written (Object[])
        Map<String, Object[]> data = new TreeMap<>();

        data.put("2", new Object[] {1, "Sachin", "Tendulkar"});
        data.put("3", new Object[] {2, "Saurav", "Ganguly"});
        data.put("4", new Object[] {3, "Rahul", "Dravid"});
        data.put("5", new Object[] {4, "Virat", "Kohli"});

        // Iterate over data and write to sheet
        Set<String> keySet = data.keySet();
        int rowNum = 1;

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Header Row
        Row headerRow = sheet.createRow(0);

        // Create cells
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        for (String key : keySet) {
            Row row = sheet.createRow(rowNum++);
            Object[] objArr = data.get(key);
            int cellNum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellNum++);
                if (obj instanceof String)
                    cell.setCellValue((String) obj);
                else if (obj instanceof Integer)
                    cell.setCellValue((Integer) obj);
            }
        }

        try {
            // Write the workbook in file system
            FileOutputStream out = new FileOutputStream(EXCEL_PATH);
            workbook.write(out);
            out.close();
            System.out.println("****File written successfully*****");
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
