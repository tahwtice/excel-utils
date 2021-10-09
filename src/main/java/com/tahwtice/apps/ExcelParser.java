package com.tahwtice.apps;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.tahwtice.apps.model.Billing;
import com.tahwtice.apps.model.Excel;
import com.tahwtice.apps.model.Order;

public class ExcelParser {

    public Excel<Order> parseOrder() {
        Excel<Order> excel = new Excel<>();
        List<Order> list = new ArrayList<>();

        try {
            FileInputStream file = new FileInputStream(Constants.EXCEL_PATH_ORDER);
            excel.setWorkbook(new HSSFWorkbook(file));
            Sheet sheet = excel.getWorkbook().getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                Order order = new Order();
                order.setMaterial(row.getCell(8).getStringCellValue());
                order.setUnitPrice(Double.parseDouble(row.getCell(12).getStringCellValue()));
                order.setQuantity(Double.parseDouble(row.getCell(13).getStringCellValue()));
                order.setTotalPrice(Double.parseDouble(row.getCell(14).getStringCellValue()));
                order.setSalesOrder(row.getCell(15).getStringCellValue());

                order.setGuid(order.getSalesOrder() + order.getMaterial());
                order.setRowIndex(row.getRowNum());

                list.add(order);
            }

            file.close();
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
            FileInputStream file = new FileInputStream(Constants.EXCEL_PATH_BILLING);
            excel.setWorkbook(new XSSFWorkbook(file));
            Sheet sheet = excel.getWorkbook().getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                Billing billing = new Billing();
                billing.setSalesOrder(row.getCell(1).getStringCellValue());
                billing.setMaterialCode(row.getCell(5).getStringCellValue());
                billing.setUnitPrice(Double.parseDouble(row.getCell(7).getStringCellValue()));
                billing.setQuantity(Double.parseDouble(row.getCell(8).getStringCellValue()));
                billing.setTotalValue(Double.parseDouble(row.getCell(9).getStringCellValue()));

                billing.setGuid(billing.getSalesOrder() + billing.getMaterialCode().substring(2));
                billing.setRowIndex(row.getRowNum());

                list.add(billing);
            }

            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        excel.setItems(list);
        return excel;
    }
}
