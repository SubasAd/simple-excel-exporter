package com.subasadhikari.exceltools;

import java.util.List;

@FunctionalInterface

public interface MergeStrategy {
    List<MergeRegion> compute(List<List<CellValue>> rows, int rowOffset);
}
