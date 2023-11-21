package com.kretsev.htmltableconverter.converter;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.IOException;

public interface HtmlToXlsConverter {
    byte[] convertHtmlTable(String htmlData, HSSFWorkbook workbook) throws IOException;
}
