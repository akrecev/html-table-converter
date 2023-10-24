package com.kretsev.htmltableconverter.converter;

import java.io.IOException;

public interface HtmlXlsConverter {
    byte[] convertTable(String htmlData) throws IOException;
}
