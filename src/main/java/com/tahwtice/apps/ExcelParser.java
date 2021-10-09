package com.tahwtice.apps;

import java.io.FileInputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
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

    private static final String EXCEL_PATH_ORDER = "src/main/resources/Order.xls";
    private static final String EXCEL_PATH_BILLING = "src/main/resources/Billing.xlsx";

    public Excel<Order> parseOrder() {
        Excel<Order> excel = new Excel<>();
        List<Order> list = new ArrayList<>();

        try {
            FileInputStream file = new FileInputStream(EXCEL_PATH_ORDER);
            excel.setWorkbook(new HSSFWorkbook(file));
            Sheet sheet = excel.getWorkbook().getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                int count = 0;
                Order order = new Order();
                order.setGuid(row.getCell(count++).getStringCellValue());
                order.setSystemId(row.getCell(count++).getNumericCellValue());
                order.setOrderDate(LocalDate.parse(row.getCell(count++).getStringCellValue(),
                        DateTimeFormatter.ofPattern("yyyy-MM-dd")));
                order.setGroup(row.getCell(count++).getStringCellValue());
                order.setPerson(row.getCell(count++).getStringCellValue());
                order.setPhone(row.getCell(count++).getStringCellValue());
                order.setCellphone(row.getCell(count++).getStringCellValue());
                order.setMail(row.getCell(count++).getStringCellValue());
                order.setBrand(row.getCell(count++).getStringCellValue());
                order.setMaterial(row.getCell(count++).getStringCellValue());
                order.setPackageSize(row.getCell(count++).getStringCellValue());
                order.setName(row.getCell(count++).getStringCellValue());
                order.setLevel(row.getCell(count++).getStringCellValue());
                order.setUnitPrice(Double.parseDouble(row.getCell(count++).getStringCellValue()));
                order.setQuantity(Double.parseDouble(row.getCell(count++).getStringCellValue()));
                order.setTotalPrice(Double.parseDouble(row.getCell(count++).getStringCellValue()));
                order.setSalesOrder(row.getCell(count).getStringCellValue());

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
            FileInputStream file = new FileInputStream(EXCEL_PATH_BILLING);
            excel.setWorkbook(new XSSFWorkbook(file));
            Sheet sheet = excel.getWorkbook().getSheetAt(0);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                int count = 0;
                Billing billing = new Billing();
                billing.setBilling(row.getCell(count++).getStringCellValue());
                billing.setSalesOrder(row.getCell(count++).getStringCellValue());
                billing.setCustomerPO(row.getCell(count++).getStringCellValue());
                billing.setPayer(row.getCell(count++).getStringCellValue());
                billing.setPayerName(row.getCell(count++).getStringCellValue());
                billing.setMaterialCode(row.getCell(count++).getStringCellValue());
                billing.setMaterialDescription(row.getCell(count++).getStringCellValue());
                billing.setUnitPrice(Double.parseDouble(row.getCell(count++).getStringCellValue()));
                billing.setQuantity(Double.parseDouble(row.getCell(count++).getStringCellValue()));
                billing.setTotalValue(Double.parseDouble(row.getCell(count++).getStringCellValue()));
                count += 4;
                billing.setItem(Double.parseDouble(row.getCell(count).getStringCellValue()));

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
