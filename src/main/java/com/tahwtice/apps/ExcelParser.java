package com.tahwtice.apps;

import java.io.FileInputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.tahwtice.apps.model.Order;

public class ExcelParser {

    private static final String EXCEL_PATH = "src/main/resources/Order.xls";

    // private static final

    public void parse() {
        try {
            FileInputStream file = new FileInputStream(EXCEL_PATH);

            // Create Workbook instance holding reference to .xlsx file
            Workbook workbook = new HSSFWorkbook(file);

            // Get first/desired sheet from the workbook
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate through each rows one by one
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

                System.out.println(order);
            }

            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
