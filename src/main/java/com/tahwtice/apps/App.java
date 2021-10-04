package com.tahwtice.apps;

// https://techblogstation.com/java/read-and-write-excel-file-in-java/


public class App {
    public static void main(String[] args) {
        System.out.println("Hello World!");

        ExcelService service = new ExcelService();
        service.parse();
        service.vLookUp();
        service.export();
    }
}
