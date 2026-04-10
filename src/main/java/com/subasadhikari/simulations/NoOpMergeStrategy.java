package com.subasadhikari.simulations;

import com.subasadhikari.exceltools.CellValue;
import com.subasadhikari.exceltools.MergeRegion;
import com.subasadhikari.exceltools.MergeStrategy;

import java.util.List;

public class NoOpMergeStrategy implements MergeStrategy {
    @Override
    public List<MergeRegion> compute(List<List<CellValue>> rows, int rowOffset) {
        return List.of();
    }
}
