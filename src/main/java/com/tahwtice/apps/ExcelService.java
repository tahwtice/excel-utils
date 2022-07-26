package com.tahwtice.apps;

import java.util.Optional;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.tahwtice.apps.model.Billing;
import com.tahwtice.apps.model.Excel;
import com.tahwtice.apps.model.Order;

public class ExcelService {
    private final ExcelParser parser;
    private final BillingService billingService;

    private Excel<Order> orderExcel;
    private Excel<Billing> billingExcel;

    public ExcelService() {
        this.parser = new ExcelParser();
        this.billingService = new BillingService();
    }

    private void copy() {
        this.parser.copyExcel(Constants.EXCEL_PATH_TEMPLATE_ORDER, Constants.EXCEL_PATH_ORDER);
        this.parser.copyExcel(Constants.EXCEL_PATH_TEMPLATE_BILLING, Constants.EXCEL_PATH_BILLING);
    }

    public final void parse() {
        this.copy();

        this.orderExcel = this.parser.parseOrder();
        this.billingExcel = this.parser.parseBilling();

        this.billingService.setBillingList(this.billingExcel.getItems());
    }

    public final void vLookUp() {
        this.orderExcel.getItems().forEach(order -> {
            Optional<Billing> billingOptional = this.billingService.findBillingByGuid(order.getGuid());
            if (billingOptional.isEmpty()) {
                order.setDeleted(false);
                return;
            }

            Billing billing = billingOptional.get();
            // billing's delete flag being already set to false means duplicate matched
            if (!billing.isDeleted()) {
                System.out.println(billing.getGuid());
            }
            billing.setDeleted(false);

            Sheet sheet = this.billingExcel.getWorkbook().getSheetAt(0);
            Row row = sheet.getRow(billing.getRowIndex());
            row.getCell(11).setCellValue(billing.getQuantity());
            row.getCell(12).setCellValue(order.getTotalValue());

            if (Math.abs(billing.getTotalValue() - order.getTotalValue()) <= Constants.DISCREPANCY) {
                row.getCell(10).setCellValue(Constants.FINAL_BILLING);
            }
        });

        System.out.println("order sum:" + this.orderExcel.getItems().toArray().length);
        System.out.println("order deleted:" + this.orderExcel.getItems().stream().filter(Order::isDeleted).toArray().length);
        System.out.println("billing sum:" + this.billingExcel.getItems().toArray().length);
        System.out.println("billing deleted:" + this.billingExcel.getItems().stream().filter(Billing::isDeleted).toArray().length);

        this.orderExcel.getItems().stream().filter(Order::isDeleted).forEach(item -> {
            Sheet sheet = this.orderExcel.getWorkbook().getSheetAt(0);
            Row row = sheet.getRow(item.getRowIndex());
            sheet.removeRow(row);
        });

        this.billingExcel.getItems().stream().filter(Billing::isDeleted).forEach(item -> {
            Sheet sheet = this.billingExcel.getWorkbook().getSheetAt(0);
            Row row = sheet.getRow(item.getRowIndex());
            sheet.removeRow(row);
        });
    }

    public final void export() {
        this.billingExcel.getItems().stream().filter(item -> !item.isDeleted()).forEach(System.out::println);
        this.parser.exportOrigin(this.billingExcel.getWorkbook(), Constants.EXCEL_PATH_BILLING);
        this.parser.exportOrigin(this.orderExcel.getWorkbook(), Constants.EXCEL_PATH_ORDER);
    }
}
