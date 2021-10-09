package com.tahwtice.apps;

import java.util.Optional;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.tahwtice.apps.model.Billing;
import com.tahwtice.apps.model.Excel;
import com.tahwtice.apps.model.Order;

public class ExcelService {
    private static final double DISCREPANCY = 0.50;
    private static final String FINAL_BILLING = "AG1004";

    private final ExcelParser parser;
    private final ExcelReporter reporter;
    private final BillingService billingService;

    private Excel<Order> orderExcel;
    private Excel<Billing> billingExcel;

    public ExcelService() {
        this.parser = new ExcelParser();
        this.reporter = new ExcelReporter();
        this.billingService = new BillingService();
    }

    public final void parse() {
        this.orderExcel = this.parser.parseOrder();
        this.billingExcel = this.parser.parseBilling();

        this.billingService.setBillingList(this.billingExcel.getItems());
    }

    public final void vLookUp() {
        this.orderExcel.getItems().forEach(order -> {
            Optional<Billing> billingOptional = this.billingService.findBillingByGuid(order.getGuid());
            if (!billingOptional.isPresent()) {
                order.setDeleted(false);
                return;
            }

            Billing billing = billingOptional.get();
            billing.setDeleted(false);

            Sheet sheet = this.billingExcel.getWorkbook().getSheetAt(0);
            Row row = sheet.getRow(billing.getRowIndex());
            row.getCell(11).setCellValue(billing.getQuantity());
            row.getCell(12).setCellValue(order.getTotalValue());

            if (Math.abs(billing.getTotalValue() - order.getTotalValue()) <= DISCREPANCY) {
                row.getCell(10).setCellValue(FINAL_BILLING);
            }
        });

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
        this.reporter.exportOrigin(this.billingExcel.getWorkbook(), Constants.EXCEL_PATH_BILLING);
        this.reporter.exportOrigin(this.orderExcel.getWorkbook(), Constants.EXCEL_PATH_ORDER);
    }
}
