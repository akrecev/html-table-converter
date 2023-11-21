package com.kretsev.htmltableconverter.service;

import java.io.IOException;

public interface HtmlToXlsService {
    byte[] getXlsForm(String htmlData, String index, String variant) throws IOException;
}
