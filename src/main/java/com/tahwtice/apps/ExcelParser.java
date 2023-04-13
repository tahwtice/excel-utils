package com.tahwtice.apps;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.tahwtice.apps.model.Billing;
import com.tahwtice.apps.model.Excel;
import com.tahwtice.apps.model.Order;

public class ExcelParser {

    private static final Map<String, Integer> INDEX_MAP = Constants.PROJECT_CELL_INDEX_MAP.get(Constants.PROJECT_ID);

    public void copyExcel(String srcPath, String destPath) {
        try {
            FileInputStream fis = new FileInputStream(srcPath);
            FileOutputStream fos = new FileOutputStream(destPath);

            Workbook workbook = WorkbookFactory.create(fis);
            workbook.write(fos);

            fos.close();
            fis.close();
            System.out.println("File copied!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public Excel<Order> parseOrder() {
        Excel<Order> excel = new Excel<>();
        List<Order> list = new ArrayList<>();

        try {
            FileInputStream fis = new FileInputStream(Constants.EXCEL_PATH_ORDER);
            excel.setWorkbook(WorkbookFactory.create(fis));
            Sheet sheet = excel.getWorkbook().getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                Order order = new Order();
                order.setSalesOrder(
                        row.getCell(INDEX_MAP.get(Constants.CELL_INDEX_ORDER_SALES_ORDER)).getStringCellValue());
                order.setMaterialCode(
                        row.getCell(INDEX_MAP.get(Constants.CELL_INDEX_ORDER_MATERIAL_CODE)).getStringCellValue());
                order.setUnitPrice(Double.parseDouble(
                        row.getCell(INDEX_MAP.get(Constants.CELL_INDEX_ORDER_UNIT_PRICE)).getStringCellValue()));
                order.setQuantity(Double.parseDouble(
                        row.getCell(INDEX_MAP.get(Constants.CELL_INDEX_ORDER_QUANTITY)).getStringCellValue()));
                order.setTotalValue(Double.parseDouble(
                        row.getCell(INDEX_MAP.get(Constants.CELL_INDEX_ORDER_TOTAL_VALUE)).getStringCellValue()));

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
            excel.setWorkbook(WorkbookFactory.create(fis));
            Sheet sheet = excel.getWorkbook().getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                Billing billing = new Billing();
                billing.setSalesOrder(
                        row.getCell(INDEX_MAP.get(Constants.CELL_INDEX_BILLING_SALES_ORDER)).getStringCellValue());
                billing.setMaterialCode(
                        row.getCell(INDEX_MAP.get(Constants.CELL_INDEX_BILLING_MATERIAL_CODE)).getStringCellValue());
                billing.setUnitPrice(Double.parseDouble(
                        row.getCell(INDEX_MAP.get(Constants.CELL_INDEX_BILLING_UNIT_PRICE)).getStringCellValue()));
                billing.setQuantity(Double.parseDouble(
                        row.getCell(INDEX_MAP.get(Constants.CELL_INDEX_BILLING_QUANTITY)).getStringCellValue()));
                billing.setTotalValue(Double.parseDouble(
                        row.getCell(INDEX_MAP.get(Constants.CELL_INDEX_BILLING_TOTAL_VALUE)).getStringCellValue()));

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
            System.out.println("File written successfully!");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
