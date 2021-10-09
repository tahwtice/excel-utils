package com.tahwtice.apps.model;

import lombok.Data;

@Data
public class Billing {
    private String guid;

    private String billing;
    private String salesOrder;
    private String customerPO;
    private String payer;
    private String payerName;
    private String materialCode;
    private String materialDescription;
    private double unitPrice;
    private double quantity;
    private double totalValue;
    private double item;

    private String finalBilling;
    private double finalQuantity;
    private double finalTotalValue;
    private String finalPO;

    private int rowIndex;
    private boolean deleted = true;
}
