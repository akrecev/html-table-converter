package com.kretsev.htmltableconverter.converter.impl;

import com.kretsev.htmltableconverter.converter.HtmlToXlsConverter;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

@Service
public class HtmlToXlsConverterImpl implements HtmlToXlsConverter {
    @Override
    public byte[] convertHtmlTable(String htmlData, HSSFWorkbook workbook) throws IOException {
        HSSFSheet sheet = workbook.getSheetAt(0);
        Document doc = Jsoup.parse(htmlData);
        CellStyle headerStyle = getHeaderStyle(workbook);
        CellStyle dataStyle = getDataCellStyle(workbook);
        int rowCount = 1;

        String className;
        String elementId;
        if (doc.select("div").first() != null) {
            className = Objects.requireNonNull(doc.select("div").first()).className();
            elementId = Objects.requireNonNull(doc.select("div").first()).id();
            if (!className.contains("table") && !elementId.contains("table")) {
                rowCount = getRowTitleCreating(workbook, sheet, doc, rowCount);
            }
        }

        for (Element table : doc.select("table")) {
            Element headerTable = table.getElementsByTag("thead").first();
            Element dataTable = table.getElementsByTag("tbody").first();
            if (headerTable != null && dataTable != null) {
                rowCount = getRowTableCreating(sheet, headerStyle, rowCount, headerTable);
                rowCount = getRowTableCreating(sheet, dataStyle, rowCount - 1, dataTable);
            } else {
                rowCount = getRowTableCreating(sheet, dataStyle, rowCount, table);
            }
            setBordersToMergedCells(sheet, dataStyle);

        }
        ByteArrayOutputStream outByteStream = new ByteArrayOutputStream();
        workbook.write(outByteStream);

        return outByteStream.toByteArray();
    }

    private static int getRowTitleCreating(HSSFWorkbook workbook, HSSFSheet sheet, Document doc, int rowCount) {
        Elements title = Objects.requireNonNull(doc.select("div").first()).children();
        makeCell(
                sheet.createRow(++rowCount),
                0,
                Objects.requireNonNull(doc.select("div").first()).ownText(),
                getTitleStyle(workbook)
        );
        for (Element element : title) {
            makeCell(sheet.createRow(++rowCount), 0, element.text(), getTitleStyle(workbook));
        }
        return rowCount;
    }

    private static int getRowTableCreating(HSSFSheet sheet, CellStyle style, int rowCount, Element table) {
        Row tableRow;
        int colspanValue;
        int rowLength;
        int rowspanValue;
        int rowStartTable = rowCount - 1;
        rowCount++;
        List<List<Integer>> rowspanCatalog = new ArrayList<>();
        List<Integer> rowspanList = new ArrayList<>(Collections.nCopies(1000, 0));
        for (Element row : table.select("tr")) {
            tableRow = sheet.createRow(rowCount);

            int thCount = 0;
            int position = 0;
            int indent = 0;
            int value;

            Elements ths = row.select("th, td");
            rowLength = ths.size();

            rowspanList = rowspanList.stream().map(integer -> --integer).collect(Collectors.toList());
            rowspanCatalog.add(rowspanList);

            for (int i = 0; i < rowLength; i++) {
                value = 0;
                if (rowCount - rowStartTable > 2) {
                    for (int k = position; k < i + indent + value + 1; k++) {
                        if (rowspanCatalog.get(rowCount - rowStartTable - 3).get(k) > 0) {
                            value++;
                        }
                        position = k + 1;
                    }
                }
                indent += value;

                Element element = ths.get(thCount);
                String rowspan = element.attr("rowspan");
                String colspan = element.attr("colspan");

                if (!colspan.isEmpty()) {
                    makeCell(tableRow, i + indent, element.text(), style);
                }
                colspanValue = 0;
                if (!colspan.isEmpty() && colspan.matches("-?\\d+(.\\d+)?")) {
                    colspanValue = Integer.parseInt(colspan);
                } else {
                    makeCell(tableRow, i + indent, element.text(), style);
                }
                rowspanValue = 0;
                if (!rowspan.isEmpty() && rowspan.matches("-?\\d+(.\\d+)?")) {
                    rowspanValue = Integer.parseInt(rowspan);
                    rowspanList.set(i + indent, rowspanValue - 1);
                    if (colspanValue > 1) {
                        for (int j = 1; j < colspanValue; j++) {
                            rowspanList.set(i + indent + j, rowspanValue - 1);
                        }
                    }
                }
                if (colspanValue > 1 && rowspanValue <= 1) {
                    for (int j = 0; j < colspanValue - 1; j++) {
                        i++;
                        rowLength++;
                        tableRow.createCell(i + indent);
                    }
                    sheet.addMergedRegion(new CellRangeAddress(
                            rowCount, rowCount,
                            i - colspanValue + 1 + indent, i + indent)
                    );
                }
                if (rowspanValue > 1 && colspanValue <= 1) {
                    sheet.addMergedRegion(new CellRangeAddress(
                            rowCount, rowCount + rowspanValue - 1,
                            i + indent, i + indent)
                    );
                }
                if (colspanValue > 1 && rowspanValue > 1) {
                    for (int j = 0; j < colspanValue - 1; j++) {
                        i++;
                        rowLength++;
                        tableRow.createCell(i + indent);
                    }
                    sheet.addMergedRegion(new CellRangeAddress(
                            rowCount, rowCount + rowspanValue - 1,
                            i - colspanValue + 1 + indent, i + indent)
                    );
                }
                thCount++;
            }
            rowCount++;
        }

        return rowCount;
    }

    private static void makeCell(Row row, int columnNumber, String cellValue, CellStyle style) {
        Cell cell = row.createCell(columnNumber);
        cell.setCellValue(cellValue);
        cell.setCellStyle(style);
    }

    private static Font getHeaderFont(HSSFWorkbook workbook) {
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setFontName("Calibri");
        font.setBold(true);

        return font;
    }

    private static Font getDataSellFont(HSSFWorkbook workbook) {
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setFontName("Calibri");

        return font;
    }

    private static void setBordersToMergedCells(HSSFSheet sheet, CellStyle style) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress rangeAddress : mergedRegions) {
            RegionUtil.setBorderTop(style.getBorderBottomEnum(), rangeAddress, sheet);
            RegionUtil.setBorderLeft(style.getBorderLeftEnum(), rangeAddress, sheet);
            RegionUtil.setBorderRight(style.getBorderRightEnum(), rangeAddress, sheet);
            RegionUtil.setBorderBottom(style.getBorderBottomEnum(), rangeAddress, sheet);
        }
    }

    private static CellStyle getTitleStyle(HSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFont(getDataSellFont(workbook));
        style.setVerticalAlignment(VerticalAlignment.TOP);

        return style;
    }

    private static CellStyle getHeaderStyle(HSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFont(getHeaderFont(workbook));
        style.setWrapText(true);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        setBorder(style);

        return style;
    }

    private static CellStyle getDataCellStyle(HSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFont(getDataSellFont(workbook));
        style.setWrapText(true);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        setBorder(style);

        return style;
    }

    private static void setBorder(CellStyle style) {
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
    }
}
