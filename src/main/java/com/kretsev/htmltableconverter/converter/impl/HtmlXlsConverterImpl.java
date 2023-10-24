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
        int rowspanValue;
        int colspanValue;
        int rowLength;
        int rowCount = 1;
        int headerHeight = 0;

        Row titleRow;
        Row tableRow;
        Cell titleCell;

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        Document doc = Jsoup.parse(htmlData);

        CellStyle headerStyle = getHeaderStyle(workbook);
        CellStyle dataStyle = getDataCellStyle(workbook);

        if (!Objects.requireNonNull(doc.select("div").first()).hasClass("table-responsive")) {
            Elements title = Objects.requireNonNull(doc.select("div").first()).children();
            titleRow = sheet.createRow(++rowCount);
            titleCell = titleRow.createCell(0);
            titleCell.setCellValue(Objects.requireNonNull(doc.select("div").first()).ownText());
            sheet.addMergedRegion(new CellRangeAddress(
                    rowCount, rowCount,
                    0, 10)
            );
            for (Element element : title) {
                titleRow = sheet.createRow(++rowCount);
                titleCell = titleRow.createCell(0);
                titleCell.setCellValue(element.text());
                sheet.addMergedRegion(new CellRangeAddress(
                        rowCount, rowCount,
                        0, 10)
                );
            }
        }

        int rowStartTable = rowCount - 1;
        for (Element table : doc.select("table")) {
            rowCount++;
            Cell cell;
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
                        cell = tableRow.createCell(i + indent);
                        cell.setCellValue(element.ownText());
                        cell.setCellStyle(headerStyle);
                    }
                    colspanValue = 0;
                    if (!colspan.isEmpty() && colspan.matches("-?\\d+(.\\d+)?")) {
                        colspanValue = Integer.parseInt(colspan);
                    } else {
                        cell = tableRow.createCell(i + indent);
                        cell.setCellValue(element.ownText());
                        cell.setCellStyle(headerStyle);
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
                    sheet.autoSizeColumn(i, true);

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

                    cell = tableRow.createCell(i + indent);
                    cell.setCellValue(element.ownText());
                    cell.setCellStyle(dataStyle);

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
