package com.kretsev.htmltableconverter.converter.impl;

import com.kretsev.htmltableconverter.converter.HtmlXlsConverter;
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
public class HtmlXlsConverterImpl implements HtmlXlsConverter {
    @Override
    public byte[] convertTable(String htmlData) throws IOException {
        return getTable(htmlData);
    }

    private byte[] getTable(String htmlData) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        Document doc = Jsoup.parse(htmlData);
        CellStyle headerStyle = getHeaderStyle(workbook);
        CellStyle dataStyle = getDataCellStyle(workbook);
        Row tableRow;
        int rowCount = 1;
        int rowspanValue;
        int colspanValue;
        int rowLength;
        int headerHeight;
        int rowStartTable;

        String className;
        String elementId;
        if (doc.select("div").first() != null) {
            className = Objects.requireNonNull(doc.select("div").first()).className();
            elementId = Objects.requireNonNull(doc.select("div").first()).id();
            if (!className.contains("table") && !elementId.contains("table")) {
                Elements title = Objects.requireNonNull(doc.select("div").first()).children();
                makeCell(
                        sheet.createRow(++rowCount),
                        0,
                        Objects.requireNonNull(doc.select("div").first()).ownText(),
                        getTitleStyle(workbook)
                );
                sheet.addMergedRegion(new CellRangeAddress(
                        rowCount, rowCount,
                        0, 17)
                );
                for (Element element : title) {
                    makeCell(sheet.createRow(++rowCount), 0, element.text(), getTitleStyle(workbook));
                    sheet.addMergedRegion(new CellRangeAddress(
                            rowCount, rowCount,
                            0, 17)
                    );
                }
            }
        }

        int rowEndTitle = rowCount - 1;
        for (Element table : doc.select("table")) {
            headerHeight = 0;
            rowStartTable = rowCount - 1;
            rowCount++;
            List<List<Integer>> rowspanCatalog = new ArrayList<>();
            List<Integer> rowspanList = new ArrayList<>(Collections.nCopies(1000, 0));
            for (Element row : table.select("tr")) {
                tableRow = sheet.createRow(rowCount);

                int thCount = 0;
                int position = 0;
                int indent = 0;
                int value;

                Elements ths = row.select("th");
                rowLength = ths.size();

                rowspanList = rowspanList.stream().map(integer -> --integer).collect(Collectors.toList());
                rowspanCatalog.add(rowspanList);
                headerHeight++;

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
                        makeCell(tableRow, i + indent, element.ownText(), headerStyle);
                        if ((headerHeight == 1 || headerHeight == 2) && rowEndTitle == rowStartTable) {
                            sheet.autoSizeColumn(i);
                        }
                    }
                    colspanValue = 0;
                    if (!colspan.isEmpty() && colspan.matches("-?\\d+(.\\d+)?")) {
                        colspanValue = Integer.parseInt(colspan);
                    } else {
                        makeCell(tableRow, i + indent, element.ownText(), headerStyle);
                        if ((headerHeight == 1 || headerHeight == 2) && rowEndTitle == rowStartTable) {
                            sheet.autoSizeColumn(i + indent, true);
                        }
                    }
                    if (colspanValue > 1) {
                        for (int j = 0; j < colspanValue - 1; j++) {
                            i++;
                            rowLength++;
                            tableRow.createCell(i);
                        }
                        sheet.addMergedRegion(new CellRangeAddress(
                                rowCount, rowCount,
                                i - colspanValue + 1 + indent, i + indent)
                        );
                    }
                    rowspanValue = 0;
                    if (!rowspan.isEmpty() && rowspan.matches("-?\\d+(.\\d+)?")) {
                        rowspanValue = Integer.parseInt(rowspan);
                        rowspanList.set(i + indent, rowspanValue);
                    }
                    if (rowspanValue > 1) {
                        sheet.addMergedRegion(new CellRangeAddress(
                                rowCount, rowCount + rowspanValue - 1,
                                i + indent, i + indent)
                        );
                    }
                    thCount++;
                }

                Elements tds = row.select("td");
                indent = 0;
                for (int i = 0; i < tds.size(); i++) {
                    if (i == 0) {
                        rowspanCatalog.clear();
                        rowspanCatalog.add(rowspanList);
                    }
                    Element element = tds.get(i);
                    value = 0;
                    for (int k = position; k < i + indent + value + 1; k++) {
                        if (rowspanCatalog.get(rowCount - rowStartTable - headerHeight - 1).get(k) > 0) {
                            value++;
                        }
                        position = k + 1;
                    }
                    indent += value;
                    makeCell(tableRow, i + indent, element.ownText(), dataStyle);
                    String rowspan = element.attr("rowspan");
                    rowspanValue = 0;
                    if (!rowspan.isEmpty() && rowspan.matches("-?\\d+(.\\d+)?")) {
                        rowspanValue = Integer.parseInt(rowspan);
                        rowspanList.set(i + indent, rowspanValue);
                    }
                    if (rowspanValue > 1) {
                        sheet.addMergedRegion(new CellRangeAddress(
                                rowCount, rowCount + rowspanValue - 1,
                                i + indent, i + indent)
                        );
                    }
                }
                rowCount++;
            }
            setBordersToMergedCells(sheet);
        }
        ByteArrayOutputStream outByteStream = new ByteArrayOutputStream();
        workbook.write(outByteStream);

        return outByteStream.toByteArray();
    }

    private void makeCell(Row row, int columnNumber, String cellValue, CellStyle style) {
        Cell cell = row.createCell(columnNumber);
        cell.setCellValue(cellValue);
        cell.setCellStyle(style);
    }

    private CellStyle getHeaderStyle(HSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFont(getHeaderFont(workbook));
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);

        return style;
    }

    private CellStyle getDataCellStyle(HSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFont(getDataSellFont(workbook));
        style.setWrapText(true);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);

        return style;
    }

    private CellStyle getTitleStyle(HSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFont(getDataSellFont(workbook));
        style.setWrapText(true);
        style.setVerticalAlignment(VerticalAlignment.TOP);

        return style;
    }

    private Font getHeaderFont(HSSFWorkbook workbook) {
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setFontName("Calibri");
        font.setBold(true);

        return font;
    }

    private Font getDataSellFont(HSSFWorkbook workbook) {
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setFontName("Calibri");

        return font;
    }

    private static void setBordersToMergedCells(HSSFSheet sheet) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress rangeAddress : mergedRegions) {
            RegionUtil.setBorderTop(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderBottom(BorderStyle.THIN, rangeAddress, sheet);
        }
    }
}
