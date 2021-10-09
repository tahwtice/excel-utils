package com.tahwtice.apps.model;

import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

import lombok.Data;

@Data
public class Excel<T> {
    private Workbook workbook;
    private List<T> items;
}
