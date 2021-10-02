package com.tahwtice.apps.model;

import java.time.LocalDate;

import lombok.Data;

@Data
public class Order {
    private String guid;
    private double systemId;
    private LocalDate orderDate;
    private String group;
    private String person;
    private String phone;
    private String cellphone;
    private String mail;
    private String brand;
    private String material;
    private String packageSize;
    private String name;
    private String level;
    private double unitPrice;
    private double quantity;
    private double totalPrice;
    private String salesOrder;
}
