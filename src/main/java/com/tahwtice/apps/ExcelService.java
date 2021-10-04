package com.tahwtice.apps;

import java.util.ArrayList;
import java.util.List;

import com.tahwtice.apps.model.Billing;
import com.tahwtice.apps.model.Order;

public class ExcelService {
    private static final double DISCREPANCY = 0.50;
    private static final String FINAL_BILLING = "AG1004";

    private final ExcelParser parser;
    private final ExcelReporter reporter;
    private final BillingService billingService;

    private List<Order> orderList;
    private final List<Billing> finalBillingList;

    public ExcelService() {
        this.parser = new ExcelParser();
        this.reporter = new ExcelReporter();
        this.billingService = new BillingService();

        this.finalBillingList = new ArrayList<>();
    }

    public final void parse() {
        this.orderList = this.parser.parseOrder();
        this.billingService.setBillingList(this.parser.parseBilling());
    }

    public final void vLookUp() {
        this.orderList.forEach(order -> {
            Billing billing = this.billingService.findBillingByGuid(order.getGuid());
            double finalTotalValue = billing.getTotalValue();
            if (Math.abs(billing.getTotalValue() - order.getTotalPrice()) <= DISCREPANCY) {
                finalTotalValue = order.getTotalPrice();
            }

            billing.setFinalBilling(FINAL_BILLING);
            billing.setFinalTotalValue(finalTotalValue);
            billing.setFinalQuantity(billing.getQuantity());
            this.finalBillingList.add(billing);
        });
    }

    public final void export() {
        this.finalBillingList.forEach(System.out::println);
        this.reporter.export(this.finalBillingList);
    }
}
