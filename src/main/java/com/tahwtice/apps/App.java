package com.tahwtice.apps;

// https://techblogstation.com/java/read-and-write-excel-file-in-java/

import java.util.List;

import com.tahwtice.apps.model.Billing;
import com.tahwtice.apps.model.Order;

public class App {
    public static void main(String[] args) {
        System.out.println("Hello World!");

        ExcelParser parser = new ExcelParser();
        List<Order> orderList = parser.parseOrder();
        orderList.forEach(System.out::println);
        List<Billing> billingList = parser.parseBilling();
        billingList.forEach(System.out::println);

        ExcelReporter reporter = new ExcelReporter();
        reporter.export(billingList);
    }
}
