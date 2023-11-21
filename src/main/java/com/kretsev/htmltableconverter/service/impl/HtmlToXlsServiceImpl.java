package com.kretsev.htmltableconverter.service.impl;

import com.kretsev.htmltableconverter.converter.HtmlToXlsConverter;
import com.kretsev.htmltableconverter.service.HtmlToXlsService;
import lombok.RequiredArgsConstructor;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;
import java.util.Properties;

@Service
@RequiredArgsConstructor
public class HtmlToXlsServiceImpl implements HtmlToXlsService {
    private final HtmlToXlsConverter converter;

    @Override
    public byte[] getXlsForm(String htmlData, String index, String variant) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        Properties tableProperties = new Properties();
        InputStream inputStream = getClass().getClassLoader().getResourceAsStream("static/table-columns-size.yaml");
        tableProperties.load(inputStream);
        String tableColumnsSize = tableProperties.getProperty("form" + index + "v" + variant);
        if (tableColumnsSize == null) {
            tableColumnsSize = tableProperties.getProperty("default");
        }
        if (!tableColumnsSize.isEmpty()) {
            List<Float> columnsSizeList = Arrays.stream(tableColumnsSize.split(";"))
                    .map(String::trim)
                    .map(s -> s.replace(',', '.'))
                    .map(Float::parseFloat)
                    .toList();
            for (int i = 0; i < columnsSizeList.size(); i++) {
                sheet.setColumnWidth(i, (int) (columnsSizeList.get(i) * 256));
            }
        }

        return converter.convertHtmlTable(htmlData, workbook);
    }
}
