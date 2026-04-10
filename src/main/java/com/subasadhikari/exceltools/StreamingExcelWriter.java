package com.subasadhikari.exceltools;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.List;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static com.subasadhikari.exceltools.XmlHeaders.*;
import static java.util.zip.Deflater.BEST_SPEED;

public class StreamingExcelWriter implements Closeable {
    private static final int OVERLAP = 3;
    private final BlockingQueue<RowBatch> queue;
    private final MergeStrategy strategy;
    private final Thread consumerThread;
    private volatile Exception consumerError;

    private final List<MergeRegion> allMerges = new ArrayList<>();
    private final Deque<List<CellValue>> overlap = new ArrayDeque<>();
    private int absoluteRow = 0;
    private final ByteArrayOutputStream rowBuffer = new ByteArrayOutputStream(32*1024);
    private static final byte[][] COL_LETTERS = new byte[16_384][];
    private static final LocalDate EXCEL_EPOCH = LocalDate.of(1899, 12, 30);
    private final byte[] charBuf = new byte[4];

    static {for (int i = 0; i < COL_LETTERS.length; i++) {COL_LETTERS[i] = Util.colIndexToLetter(i).getBytes(StandardCharsets.UTF_8);}}

    public StreamingExcelWriter(String filePath, int queueCapacity, MergeStrategy strategy) throws IOException {
        this.queue    = new LinkedBlockingQueue<>(queueCapacity);
        this.strategy = strategy;
        ZipOutputStream zos = new ZipOutputStream(new BufferedOutputStream(new FileOutputStream(filePath)));
        zos.setLevel(BEST_SPEED);
        consumerThread = new Thread(() -> consume(zos), "excel-consumer");
        consumerThread.start();
    }

    @Override
    public void close() throws IOException {
        try {
            queue.put(RowBatch.withPoison());
            consumerThread.join();
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
        }
        if (consumerError != null) throw new IOException(consumerError);
    }

    private void consume(ZipOutputStream zos) {
            try {
                writeBoilerplate(zos);
                zos.putNextEntry(new ZipEntry("xl/worksheets/sheet1.xml"));

                zos.write(XML_DECL);
                zos.write(WORKSHEET_OPEN);
                zos.write(SHEET_DATA_OPEN);

                while (true) {
                    RowBatch batch = queue.take();
                    boolean isFinal = batch.poison();
                    List<List<CellValue>> rows = batch.rows();

                    List<List<CellValue>> window = new ArrayList<>();
                    int overlapSize = overlap.size();
                    window.addAll(overlap);
                    window.addAll(rows);
                    int windowOffset = absoluteRow - overlapSize;

                    List<MergeRegion> regions = strategy.compute(window, windowOffset);
                    for (MergeRegion r : regions) {
                        if (isFinal || r.lastRow() < absoluteRow) {
                            allMerges.add(r);
                        }
                    }

                    for (List<CellValue> row : rows) {
                        writeTypedRowToStream(zos, row, absoluteRow++);
                    }

                    overlap.clear();
                    int start = Math.max(0, window.size() - OVERLAP);
                    for (int i = start; i < window.size(); i++) {
                        overlap.add(window.get(i));
                    }

                    if (isFinal) break;
                }
                zos.write(SHEET_DATA_CLOSE);
                if (!allMerges.isEmpty()) {
                    zos.write(MERGE_CELLS_OPEN);
                    zos.write(Integer.toString(allMerges.size()).getBytes(StandardCharsets.UTF_8));
                    zos.write(MERGE_CELLS_MID);
                    for (MergeRegion r : allMerges) {
                        zos.write(MERGE_CELL_OPEN);
                        zos.write(r.toCellRange().getBytes(StandardCharsets.UTF_8));
                        zos.write(MERGE_CELL_CLOSE);
                    }
                    zos.write(MERGE_CELLS_CLOSE);
                }
                zos.write(WORKSHEET_CLOSE);
                zos.closeEntry();
                zos.close();

            } catch (Exception e) {
                consumerError = e;
            }
        }


    public void writeTyped(List<List<CellValue>> rows) throws InterruptedException {
        if (consumerError != null) throw new RuntimeException(consumerError);
        List<List<CellValue>> safeCopy = new ArrayList<>(rows);
        queue.put(RowBatch.of(safeCopy));
    }


    private void escapeXmlToBytes(String s, ByteArrayOutputStream out) throws IOException {
        if (s == null) return;
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            switch (c) {
                case '&':  out.write(AMP);  break;
                case '<':  out.write(LT);   break;
                case '>':  out.write(GT);   break;
                case '"':  out.write(QUOT); break;
                case '\'': out.write(APOS); break;
                default:
                    if (c < 0x80) {
                        out.write(c);
                    } else {
                        int codePoint;
                        if (Character.isHighSurrogate(c) && i + 1 < s.length()
                                && Character.isLowSurrogate(s.charAt(i + 1))) {
                            codePoint = Character.toCodePoint(c, s.charAt(++i));
                        } else {
                            codePoint = c;
                        }
                        int len = encodeUtf8(codePoint, charBuf);
                        out.write(charBuf, 0, len);
                    }
            }
        }
    }

    private static int encodeUtf8(int cp, byte[] buf) {
        if (cp < 0x80) {
            buf[0] = (byte) cp;
            return 1;
        } else if (cp < 0x800) {
            buf[0] = (byte) (0xC0 | (cp >> 6));
            buf[1] = (byte) (0x80 | (cp & 0x3F));
            return 2;
        } else if (cp < 0x10000) {
            buf[0] = (byte) (0xE0 | (cp >> 12));
            buf[1] = (byte) (0x80 | ((cp >> 6) & 0x3F));
            buf[2] = (byte) (0x80 | (cp & 0x3F));
            return 3;
        } else {
            buf[0] = (byte) (0xF0 | (cp >> 18));
            buf[1] = (byte) (0x80 | ((cp >> 12) & 0x3F));
            buf[2] = (byte) (0x80 | ((cp >> 6) & 0x3F));
            buf[3] = (byte) (0x80 | (cp & 0x3F));
            return 4;
        }
    }


    public static long toExcelDate(LocalDate date) {
        return date.toEpochDay() - EXCEL_EPOCH.toEpochDay();
    }

    private void writeTypedRowToStream(OutputStream out, List<CellValue> row, int rowIdx) throws IOException {
        rowBuffer.reset();
        byte[] rowNumBytes = Integer.toString(rowIdx + 1).getBytes(StandardCharsets.UTF_8);
        rowBuffer.write(ROW_OPEN);
        rowBuffer.write(rowNumBytes);
        rowBuffer.write(ROW_MID);

        for (int c = 0; c < row.size(); c++) {
            CellValue cv = row.get(c);
            if (cv == null || cv.value == null) continue;

            rowBuffer.write(CELL_OPEN_NUMBER);
            rowBuffer.write(COL_LETTERS[c]);
            rowBuffer.write(rowNumBytes);

            switch (cv.type) {
                case NUMBER:
                    rowBuffer.write(CELL_TYPE_NUMBER);
                    rowBuffer.write(cv.value.toString().getBytes(StandardCharsets.UTF_8));
                    rowBuffer.write(CELL_CLOSE_V);
                    break;
                case DATE:
                    rowBuffer.write(DATE_STYLE_ATTR);
                    long excelDate = toExcelDate((LocalDate) cv.value);
                    rowBuffer.write(Long.toString(excelDate).getBytes(StandardCharsets.UTF_8));
                    rowBuffer.write(CELL_CLOSE_V);
                    break;
                case BOOLEAN:
                    rowBuffer.write(CELL_TYPE_BOOLEAN);
                    rowBuffer.write(((Boolean) cv.value) ? "1".getBytes() : "0".getBytes());
                    rowBuffer.write(CELL_CLOSE_V);
                    break;
                case STRING:
                default:
                    rowBuffer.write(CELL_MID);
                    escapeXmlToBytes(cv.value.toString(), rowBuffer);
                    rowBuffer.write(CELL_CLOSE);
                    break;
            }
        }
        rowBuffer.write(ROW_CLOSE);
        rowBuffer.writeTo(out);
    }

    public void writeBoilerplate(ZipOutputStream zos) throws IOException {
            addFile(zos, "[Content_Types].xml", CONTENT_TYPES_XML);
            addFile(zos, "_rels/.rels", ROOT_RELS_XML);
            addFile(zos, "xl/workbook.xml", WORKBOOK_XML);
            addFile(zos, "xl/_rels/workbook.xml.rels", WORKBOOK_RELS_XML);
            addFile(zos,"xl/styles.xml",STYLES_XML);
    }


    private void addFile(ZipOutputStream zos, String fileName, byte [] content) throws IOException {
        ZipEntry entry = new ZipEntry(fileName);
        zos.putNextEntry(entry);
        zos.write(content);
        zos.closeEntry();
    }
}