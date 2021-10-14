package com.tahwtice.apps;

import java.util.AbstractMap;
import java.util.Map;

public class Constants {
    public static final String PROJECT_ID = "XMU";

    public static final String EXCEL_PATH_ROOT = "src/main/resources/" + "XMU";

    public static final String EXCEL_PATH_TEMPLATE_ORDER = EXCEL_PATH_ROOT + "/template/Order.xlsx";
    public static final String EXCEL_PATH_TEMPLATE_BILLING = EXCEL_PATH_ROOT + "/template/Billing.xlsx";
    public static final String EXCEL_PATH_ORDER = EXCEL_PATH_ROOT + "/Order.xlsx";
    public static final String EXCEL_PATH_BILLING = EXCEL_PATH_ROOT + "/Billing.xlsx";

    public static final String FINAL_BILLING = "AG1004";
    public static final double DISCREPANCY = 0.50;

    public static final String CELL_INDEX_ORDER_SALES_ORDER = "oso";
    public static final String CELL_INDEX_ORDER_MATERIAL_CODE = "omc";
    public static final String CELL_INDEX_ORDER_UNIT_PRICE = "oup";
    public static final String CELL_INDEX_ORDER_QUANTITY = "oqt";
    public static final String CELL_INDEX_ORDER_TOTAL_VALUE = "otv";

    public static final String CELL_INDEX_BILLING_SALES_ORDER = "bso";
    public static final String CELL_INDEX_BILLING_MATERIAL_CODE = "bmc";
    public static final String CELL_INDEX_BILLING_UNIT_PRICE = "bup";
    public static final String CELL_INDEX_BILLING_QUANTITY = "bqt";
    public static final String CELL_INDEX_BILLING_TOTAL_VALUE = "btv";

    public static final Map<String, Integer> CELL_INDEX_MAP_XMU = Map.ofEntries(

            new AbstractMap.SimpleEntry<>(CELL_INDEX_ORDER_SALES_ORDER, 0),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_ORDER_MATERIAL_CODE, 5),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_ORDER_UNIT_PRICE, 7),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_ORDER_QUANTITY, 8),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_ORDER_TOTAL_VALUE, 9),

            new AbstractMap.SimpleEntry<>(CELL_INDEX_BILLING_SALES_ORDER, 2),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_BILLING_MATERIAL_CODE, 5),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_BILLING_UNIT_PRICE, 7),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_BILLING_QUANTITY, 8),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_BILLING_TOTAL_VALUE, 9)

    );

    public static final Map<String, Map<String, Integer>> PROJECT_CELL_INDEX_MAP = Map.of("XMU", CELL_INDEX_MAP_XMU);
}
