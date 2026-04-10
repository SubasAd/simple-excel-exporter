package com.subasadhikari.exceltools;

import java.time.LocalDate;

public class CellValue {

        final CellType type;
        final Object value;
        final String format;
        public CellValue(CellType type, Object value, String format) {
            this.type = type;
            this.value = value;
            this.format = format;
        }
    public CellValue(CellType type, Object value) {
        this.type = type;
        this.value = value;
        this.format = null;
    }

    public static CellValue string(String s) { return new CellValue(CellType.STRING, s); }
    public static CellValue number(Number n) { return new CellValue(CellType.NUMBER, n ); }
    public static CellValue date(LocalDate d) { return new CellValue(CellType.DATE, d ); }
    public static CellValue bool(boolean b) { return new CellValue(CellType.BOOLEAN, b ); }

}

