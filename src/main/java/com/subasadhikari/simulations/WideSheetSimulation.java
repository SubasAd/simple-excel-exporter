package com.subasadhikari.simulations;

import com.subasadhikari.exceltools.CellValue;
import com.subasadhikari.exceltools.StreamingExcelWriter;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ThreadLocalRandom;

public class WideSheetSimulation {
    private static final int ROWS = 10000;
    private static final int COLS = 50;
    private static final int BATCH_SIZE = 500;

    private static final String[] NEPALI_NAMES = {
            "राम बहादुर", "सीता देवी", "कृष्ण प्रसाद", "लक्ष्मी कुमारी",
            "बिष्णु हरि", "गणेश बहादुर", "पार्वती देवी", "सूर्य प्रसाद"
    };

    private static final String[] DISTRICTS = {
            "काठमाडौं", "ललितपुर", "भक्तपुर", "पोखरा", "विराटनगर"
    };

    public static void main(String[] args) throws Exception {
        String filePath = "wide_devanagari_test.xlsx";
        int queueCapacity = 100;

        NoOpMergeStrategy strategy = new NoOpMergeStrategy();

        System.out.printf("Starting export: %,d rows × %,d columns = %,d cells%n",
                ROWS, COLS, (long) ROWS * COLS);

        long start = System.currentTimeMillis();

        try (StreamingExcelWriter writer = new StreamingExcelWriter(filePath, queueCapacity, strategy)) {
            for (int rowStart = 0; rowStart < ROWS; rowStart += BATCH_SIZE) {
                int batchRows = Math.min(BATCH_SIZE, ROWS - rowStart);
                List<List<CellValue>> batch = generateBatch(rowStart, batchRows, COLS);
                writer.writeTyped(batch);

                if ((rowStart + batchRows) % 100_000 == 0) {
                    long elapsed = System.currentTimeMillis() - start;
                    System.out.printf("Written %,d rows in %.2f sec (%,.0f rows/sec)%n",
                            rowStart + batchRows, elapsed / 1000.0,
                            (rowStart + batchRows) / (elapsed / 1000.0));
                }
            }
        }

        long duration = System.currentTimeMillis() - start;
        double fileSizeMB = new java.io.File(filePath).length() / (1024.0 * 1024.0);
        System.out.printf("Completed: %s (%.2f MB) in %.2f seconds%n",
                filePath, fileSizeMB, duration / 1000.0);
        System.out.printf("Throughput: %,.0f rows/sec, %,.0f cells/sec%n",
                ROWS / (duration / 1000.0),
                ((long) ROWS * COLS) / (duration / 1000.0));
    }

    private static List<List<CellValue>> generateBatch(int startRow, int rowCount, int cols) {
        List<List<CellValue>> batch = new ArrayList<>(rowCount);
        ThreadLocalRandom rnd = ThreadLocalRandom.current();

        for (int r = 0; r < rowCount; r++) {
            int absoluteRow = startRow + r;
            List<CellValue> row = new ArrayList<>(cols);

            for (int c = 0; c < cols; c++) {
                if (c == 0) {
                    row.add(CellValue.number(absoluteRow + 1));
                } else if (c == 1) {
                    row.add(CellValue.string(NEPALI_NAMES[absoluteRow % NEPALI_NAMES.length]));
                } else if (c == 2) {
                    row.add(CellValue.date(LocalDate.of(2020, 1, 1).plusDays(rnd.nextInt(1000))));
                } else if (c == 3) {
                    row.add(CellValue.string(DISTRICTS[absoluteRow % DISTRICTS.length]));
                } else if (c == 4) {
                    row.add(CellValue.string("वडा-" + (absoluteRow % 32 + 1)));
                } else if (c < 10) {
                    row.add(CellValue.number(rnd.nextDouble(100, 10000)));
                } else if (c % 10 == 0) {
                    row.add(CellValue.bool(rnd.nextBoolean()));
                } else {
                    row.add(CellValue.string("Data_" + absoluteRow + "_" + c));
                }
            }
            batch.add(row);
        }
        return batch;
    }
}