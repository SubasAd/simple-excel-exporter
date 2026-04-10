package com.subasadhikari.exceltools;
public record MergeRegion(int firstRow, int lastRow, int firstCol, int lastCol) {

    public String toCellRange() {
        return colLetter(firstCol) + (firstRow + 1) + ":" +
                colLetter(lastCol)  + (lastRow  + 1);
    }
    static String colLetter(int col) {
        StringBuilder sb = new StringBuilder();
        int c = col;
        while (c >= 0) { sb.insert(0, (char)('A' + c % 26)); c = c / 26 - 1; }
        return sb.toString();
    }
}