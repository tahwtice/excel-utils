package com.tahwtice.apps;

// https://techblogstation.com/java/read-and-write-excel-file-in-java/

import java.time.Duration;
import java.time.Instant;

public class App {
    public static void main(String[] args) {
        System.out.println("Hello World!");

        Instant begin = Instant.now();
        ExcelService service = new ExcelService();
        service.parse();
        service.vLookUp();
        service.export();
        Instant end = Instant.now();

        System.out.printf("Total time cost: %d seconds", Duration.between(begin, end).toSeconds());
    }
}
