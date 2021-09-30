package com.tahwtice.apps;

// https://techblogstation.com/java/read-and-write-excel-file-in-java/

public class App {
    public static void main(String[] args) {
        System.out.println("Hello World!");

        ExcelParser parser = new ExcelParser();
        parser.parse();

        System.out.println("Begin to generate Excel...");

        ExcelReporter reporter = new ExcelReporter();
        reporter.export();
    }
}
