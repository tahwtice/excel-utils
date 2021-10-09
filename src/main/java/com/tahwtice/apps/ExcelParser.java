package com.tahwtice.apps;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.tahwtice.apps.model.Billing;
import com.tahwtice.apps.model.Excel;
import com.tahwtice.apps.model.Order;

public class ExcelParser {

    public void copyExcel(String srcPath, String destPath) {
        try {
            FileInputStream fis = new FileInputStream(srcPath);
            FileOutputStream fos = new FileOutputStream(destPath);

            Workbook workbook;
            if (srcPath.toLowerCase().endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else {
                workbook = new HSSFWorkbook(fis);
            }

            workbook.write(fos);
            fos.close();
            fis.close();
            System.out.println("File copied");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public Excel<Order> parseOrder() {
        Excel<Order> excel = new Excel<>();
        List<Order> list = new ArrayList<>();

        try {
            FileInputStream fis = new FileInputStream(Constants.EXCEL_PATH_ORDER);
            excel.setWorkbook(new XSSFWorkbook(fis));
            Sheet sheet = excel.getWorkbook().getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                Order order = new Order();
                order.setSalesOrder(row.getCell(0).getStringCellValue());
                order.setMaterialCode(row.getCell(5).getStringCellValue());
                order.setUnitPrice(Double.parseDouble(row.getCell(7).getStringCellValue()));
                order.setQuantity(Double.parseDouble(row.getCell(8).getStringCellValue()));
                order.setTotalValue(Double.parseDouble(row.getCell(9).getStringCellValue()));

                order.setGuid(order.getSalesOrder() + order.getMaterialCode());
                order.setRowIndex(row.getRowNum());

                list.add(order);
            }

            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        excel.setItems(list);
        return excel;
    }

    public Excel<Billing> parseBilling() {
        Excel<Billing> excel = new Excel<>();
        List<Billing> list = new ArrayList<>();

        try {
            FileInputStream fis = new FileInputStream(Constants.EXCEL_PATH_BILLING);
            excel.setWorkbook(new XSSFWorkbook(fis));
            Sheet sheet = excel.getWorkbook().getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                Billing billing = new Billing();
                billing.setSalesOrder(row.getCell(2).getStringCellValue());
                billing.setMaterialCode(row.getCell(5).getStringCellValue());
                billing.setUnitPrice(Double.parseDouble(row.getCell(7).getStringCellValue()));
                billing.setQuantity(Double.parseDouble(row.getCell(8).getStringCellValue()));
                billing.setTotalValue(Double.parseDouble(row.getCell(9).getStringCellValue()));

                billing.setGuid(billing.getSalesOrder() + billing.getMaterialCode().substring(2));
                billing.setRowIndex(row.getRowNum());

                list.add(billing);
            }

            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        excel.setItems(list);
        return excel;
    }

    public void exportOrigin(Workbook workbook, String path) {
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

            FileOutputStream fos = new FileOutputStream(path);
            workbook.write(fos);
            fos.close();
            System.out.println("File written successfully");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
