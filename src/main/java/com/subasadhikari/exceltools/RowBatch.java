package com.subasadhikari.exceltools;

import java.util.List;

public record RowBatch(List<List<CellValue>> rows, boolean poison) {
    public static RowBatch of(List<List<CellValue>> rows) {
        return new RowBatch(rows, false);
    }
    public  static RowBatch withPoison() {
        return new RowBatch(List.of(), true);
    }
}