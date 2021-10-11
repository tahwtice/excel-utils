package com.tahwtice.apps;

import java.util.Map;

public class Constants {
    public static final String PROJECT_ID = "XMU";

    public static final String EXCEL_PATH_ROOT = "src/main/resources/" + PROJECT_ID;

    public static final String EXCEL_PATH_TEMPLATE_ORDER = EXCEL_PATH_ROOT + "/template/Order.xlsx";
    public static final String EXCEL_PATH_TEMPLATE_BILLING = EXCEL_PATH_ROOT + "/template/Billing.xlsx";
    public static final String EXCEL_PATH_ORDER = EXCEL_PATH_ROOT + "/Order.xlsx";
    public static final String EXCEL_PATH_BILLING = EXCEL_PATH_ROOT + "/Billing.xlsx";

    public static final String FINAL_BILLING = "AG1004";
    public static final double DISCREPANCY = 0.50;

    private static final Object[][] array = {{"a", 0}, {"b", 0}, {"c", 0}, {"d", 0}, {"f", 0}};
    public static final Map<String, Integer> CELL_INDEX_MAP_XMU = Map.of((String) array[0][0], (Integer) array[0][1],
            (String) array[1][0], (Integer) array[1][1], (String) array[2][0], (Integer) array[2][1],
            (String) array[3][0], (Integer) array[3][1], (String) array[4][0], (Integer) array[4][1]);



    public static final Map<String, Map<String, Integer>> PROJECT_CELL_INDEX_MAP =
            Map.of(PROJECT_ID, CELL_INDEX_MAP_XMU);

    public static final int CELL_INDEX_ORDER_SALES_ORDER = 0;
    public static final int CELL_INDEX_ORDER_MATERIAL_CODE = 0;
    public static final int CELL_INDEX_ORDER_UNIT_PRICE = 0;
    public static final int CELL_INDEX_ORDER_QUANTITY = 0;
    public static final int CELL_INDEX_ORDER_TOTAL_VALUE = 0;

    public static final int CELL_INDEX_BILLING_SALES_ORDER = 0;
    public static final int CELL_INDEX_BILLING_MATERIAL_CODE = 0;
    public static final int CELL_INDEX_BILLING_UNIT_PRICE = 0;
    public static final int CELL_INDEX_BILLING_QUANTITY = 0;
    public static final int CELL_INDEX_BILLING_TOTAL_VALUE = 0;
}
