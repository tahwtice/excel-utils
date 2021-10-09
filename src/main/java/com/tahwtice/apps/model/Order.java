package com.tahwtice.apps.model;

import lombok.Data;

@Data
public class Order {
    private String guid;

    private String material;
    private double unitPrice;
    private double quantity;
    private double totalPrice;
    private String salesOrder;

    private int rowIndex;
}
