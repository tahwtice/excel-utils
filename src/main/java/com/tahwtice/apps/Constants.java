package com.tahwtice.apps;

import java.util.AbstractMap;
import java.util.Map;

public class Constants {
    // Step 1: change the project id
    public static final String PROJECT_ID = "XMU";

    public static final String EXCEL_PATH_ROOT = "src/main/resources/" + PROJECT_ID;

    // Step 2: check the suffix of file (.xls or .xlsx)
    public static final String EXCEL_PATH_TEMPLATE_ORDER = EXCEL_PATH_ROOT + "/template/Order.xlsx";
    public static final String EXCEL_PATH_TEMPLATE_BILLING = EXCEL_PATH_ROOT + "/template/Billing.xlsx";
    public static final String EXCEL_PATH_ORDER = EXCEL_PATH_ROOT + "/Order.xlsx";
    public static final String EXCEL_PATH_BILLING = EXCEL_PATH_ROOT + "/Billing.xlsx";

    // Step 3: change the final billing code
    public static final String FINAL_BILLING = "GZ220722-0526";
    // Step 4: check the discrepancy value
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

    // Step 5: check the related cell index for excel columns
    public static final Map<String, Integer> CELL_INDEX_MAP_XMU = Map.ofEntries(
            // cell index of Order for project XMU
            new AbstractMap.SimpleEntry<>(CELL_INDEX_ORDER_SALES_ORDER, 0),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_ORDER_MATERIAL_CODE, 5),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_ORDER_UNIT_PRICE, 7),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_ORDER_QUANTITY, 8),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_ORDER_TOTAL_VALUE, 9),
            // cell index of Billing for project XMU
            new AbstractMap.SimpleEntry<>(CELL_INDEX_BILLING_SALES_ORDER, 2),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_BILLING_MATERIAL_CODE, 5),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_BILLING_UNIT_PRICE, 7),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_BILLING_QUANTITY, 8),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_BILLING_TOTAL_VALUE, 9)

    );
    public static final Map<String, Integer> CELL_INDEX_MAP_BJMU = Map.ofEntries(

            // cell index of Order for project BJMU
            new AbstractMap.SimpleEntry<>(CELL_INDEX_ORDER_SALES_ORDER, 15),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_ORDER_MATERIAL_CODE, 8),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_ORDER_UNIT_PRICE, 12),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_ORDER_QUANTITY, 13),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_ORDER_TOTAL_VALUE, 14),
            // cell index of Billing for project BJMU
            new AbstractMap.SimpleEntry<>(CELL_INDEX_BILLING_SALES_ORDER, 1),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_BILLING_MATERIAL_CODE, 5),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_BILLING_UNIT_PRICE, 7),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_BILLING_QUANTITY, 8),
            new AbstractMap.SimpleEntry<>(CELL_INDEX_BILLING_TOTAL_VALUE, 9)

    );

    public static final Map<String, Map<String, Integer>> PROJECT_CELL_INDEX_MAP =
            Map.of("XMU", CELL_INDEX_MAP_XMU, "BJMU", CELL_INDEX_MAP_BJMU);
}
