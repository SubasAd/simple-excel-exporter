package com.subasadhikari.exceltools;

import java.nio.charset.StandardCharsets;

public class XmlHeaders {
    public static final byte[] XML_DECL = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n".getBytes(StandardCharsets.UTF_8);
    public static final byte[] WORKSHEET_OPEN = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n".getBytes(StandardCharsets.UTF_8);
    public static final byte[] SHEET_DATA_OPEN = "  <sheetData>\n".getBytes(StandardCharsets.UTF_8);
    public static final byte[] SHEET_DATA_CLOSE = "  </sheetData>\n".getBytes(StandardCharsets.UTF_8);
    public static final byte[] WORKSHEET_CLOSE = "</worksheet>".getBytes(StandardCharsets.UTF_8);
    public static final byte[] MERGE_CELLS_OPEN  = "  <mergeCells count=\"".getBytes(StandardCharsets.UTF_8);
    public static final byte[] MERGE_CELLS_MID   = "\">\n".getBytes(StandardCharsets.UTF_8);
    public static final byte[] MERGE_CELLS_CLOSE = "  </mergeCells>\n".getBytes(StandardCharsets.UTF_8);
    public static final byte[] MERGE_CELL_OPEN   = "    <mergeCell ref=\"".getBytes(StandardCharsets.UTF_8);
    public static final byte[] MERGE_CELL_CLOSE  = "\"/>\n".getBytes(StandardCharsets.UTF_8);
    public static final byte[] ROW_OPEN  = "    <row r=\"".getBytes(StandardCharsets.UTF_8);
    public static final byte[] ROW_MID   = "\">\n".getBytes(StandardCharsets.UTF_8);
    public static final byte[] ROW_CLOSE = "    </row>\n".getBytes(StandardCharsets.UTF_8);

    public static final byte[] CELL_OPEN  = "      <c r=\"".getBytes(StandardCharsets.UTF_8);
    public static final byte[] CELL_MID   = "\" t=\"inlineStr\"><is><t>".getBytes(StandardCharsets.UTF_8);
    public static final byte[] CELL_CLOSE = "</t></is></c>\n".getBytes(StandardCharsets.UTF_8);


    public static final byte[] CELL_OPEN_NUMBER = "      <c r=\"".getBytes(StandardCharsets.UTF_8);
    public static final byte[] CELL_TYPE_NUMBER = "\" t=\"n\"><v>".getBytes(StandardCharsets.UTF_8);
    public static final byte[] CELL_TYPE_BOOLEAN = "\" t=\"b\"><v>".getBytes(StandardCharsets.UTF_8);
    public static final byte[] CELL_CLOSE_V = "</v></c>\n".getBytes(StandardCharsets.UTF_8);
    public static final byte [] DATE_STYLE_ATTR = "\" s=\"1\"><v>".getBytes(StandardCharsets.UTF_8);
    public static final byte[] CELL_OPEN_NUMBER_STYLED = "      <c r=\"".getBytes(StandardCharsets.UTF_8);
    public static final byte[] CELL_TYPE_NUMBER_STYLED = "\" s=\"%d\" t=\"n\"><v>".getBytes(StandardCharsets.UTF_8);


    // XML escape sequences
    public static final byte[] AMP  = "&amp;".getBytes(StandardCharsets.UTF_8);
    public static final byte[] LT   = "&lt;".getBytes(StandardCharsets.UTF_8);
    public static final byte[] GT   = "&gt;".getBytes(StandardCharsets.UTF_8);
    public static final byte[] QUOT = "&quot;".getBytes(StandardCharsets.UTF_8);
    public static final byte[] APOS = "&apos;".getBytes(StandardCharsets.UTF_8);

    public static final byte[] CONTENT_TYPES_XML = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n" +
                    "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n" +
                    "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n" +
                    "  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n" +
                    "  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n" +
                    "  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n" +
                    "</Types>"
    ).getBytes(StandardCharsets.UTF_8);

    public static final byte[] ROOT_RELS_XML = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                    "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\n" +
                    "</Relationships>"
    ).getBytes(StandardCharsets.UTF_8);

    public static final byte[] WORKBOOK_XML = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n" +
                    "  <bookViews><workbookView/></bookViews>\n" +  // good practice
                    "  <sheets>\n" +
                    "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n" +
                    "  </sheets>\n" +
                    "</workbook>"
    ).getBytes(StandardCharsets.UTF_8);

    public static final byte[] WORKBOOK_RELS_XML = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                    "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n" +
                    "  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n" +
                    "</Relationships>"
    ).getBytes(StandardCharsets.UTF_8);

    public static final byte[] STYLES_XML = (
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n" +
                    "  <numFmts count=\"2\">\n" +
                    "    <numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd\"/>\n" +
                    "    <numFmt numFmtId=\"165\" formatCode=\"#,##0.00\"/>\n" +
                    "  </numFmts>\n" +
                    "  <fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>\n" +
                    "  <fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>\n" +
                    "  <borders count=\"1\"><border/></borders>\n" +
                    "  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\n" +
                    "  <cellXfs count=\"3\">\n" +
                    "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>\n" +
                    "    <xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>\n" +
                    "    <xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>\n" +
                    "  </cellXfs>\n" +
                    "</styleSheet>"
    ).getBytes(StandardCharsets.UTF_8);
}
