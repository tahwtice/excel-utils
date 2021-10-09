package com.tahwtice.apps.model;

import lombok.Data;

@Data
public class Billing {
    private String guid;

    private String salesOrder;
    private String materialCode;
    private double unitPrice;
    private double quantity;
    private double totalValue;

    private int rowIndex;
    private boolean deleted = true;
}
