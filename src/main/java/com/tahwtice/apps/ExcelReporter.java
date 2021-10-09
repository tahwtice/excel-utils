package com.tahwtice.apps;

import java.io.FileOutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.tahwtice.apps.model.Billing;
import com.tahwtice.apps.model.Excel;

public class ExcelReporter {

    private static final String EXCEL_PATH = "src/main/resources/Billing copy.xlsx";

    private static final String[] columns = {"SAP Billing", "Sales Order", "Customer PO", "Payer", "Payer name",
            "Material code", "Material description", "Unit price", "Quantity", "Total Value", "*结算单号", "数量", "含税金额",
            "PO(New)", "Item"};

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

            FileOutputStream out = new FileOutputStream(EXCEL_PATH);
            billingExcel.getWorkbook().write(out);
            out.close();
            System.out.println("File written successfully");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void export(List<Billing> billingList) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        // headerFont.setFontHeightInPoints((short) 14);
        // headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Header Row
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Create contents
        int rowNum = 1;
        for (Billing billing : billingList) {
            Row row = sheet.createRow(rowNum++);
            int cellNum = 0;
            Cell cell = row.createCell(cellNum++);
            cell.setCellValue(billing.getBilling());
            cell = row.createCell(cellNum++);
            cell.setCellValue(billing.getSalesOrder());
            cell = row.createCell(cellNum++);
            cell.setCellValue(billing.getCustomerPO());
            cell = row.createCell(cellNum++);
            cell.setCellValue(billing.getPayer());
            cell = row.createCell(cellNum++);
            cell.setCellValue(billing.getPayerName());
            cell = row.createCell(cellNum++);
            cell.setCellValue(billing.getMaterialCode());
            cell = row.createCell(cellNum++);
            cell.setCellValue(billing.getMaterialDescription());
            cell = row.createCell(cellNum++);
            cell.setCellValue(billing.getUnitPrice());
            cell = row.createCell(cellNum++);
            cell.setCellValue(billing.getQuantity());
            cell = row.createCell(cellNum++);
            cell.setCellValue(billing.getTotalValue());
            cell = row.createCell(cellNum++);
            cell.setCellValue(billing.getFinalBilling());
            cell = row.createCell(cellNum++);
            cell.setCellValue(billing.getFinalQuantity());
            cell = row.createCell(cellNum++);
            cell.setCellValue(billing.getFinalTotalValue());
            cell = row.createCell(cellNum++);
            cell.setCellValue(billing.getFinalPO());
            cell = row.createCell(cellNum);
            cell.setCellValue(billing.getItem());
        }

        try {
            // Write the workbook in file system
            FileOutputStream out = new FileOutputStream(EXCEL_PATH);
            workbook.write(out);
            out.close();
            System.out.println("File written successfully");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
