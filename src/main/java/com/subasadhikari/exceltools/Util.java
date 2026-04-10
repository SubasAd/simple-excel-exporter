package com.subasadhikari.exceltools;

public class Util {
    static String colIndexToLetter(int idx) {
        StringBuilder sb = new StringBuilder(3);
        while (idx >= 0) {
            sb.insert(0, (char) ('A' + (idx % 26)));
            idx = idx / 26 - 1;
        }
        return sb.toString();
    }
}
