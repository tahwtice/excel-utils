package com.tahwtice.apps;

import java.util.List;

import com.tahwtice.apps.model.Billing;
import com.tahwtice.apps.model.Order;

public class ExcelService {
    private final ExcelParser parser;
    private final ExcelReporter reporter;
    private List<Order> orderList;
    private List<Billing> billingList;
    private List<Billing> finalBillingList;

    public ExcelService() {
        this.parser = new ExcelParser();
        this.reporter = new ExcelReporter();
    }

    public final void parse() {
        this.orderList = this.parser.parseOrder();
        this.billingList = this.parser.parseBilling();

        this.orderList.forEach(System.out::println);
        this.billingList.forEach(System.out::println);
    }

    public final void vLookUp() {
        this.finalBillingList = this.billingList;
    }

    public final void export() {
        this.finalBillingList.forEach(System.out::println);
        this.reporter.export(this.finalBillingList);
    }
}
