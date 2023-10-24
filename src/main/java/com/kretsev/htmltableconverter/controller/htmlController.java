package com.kretsev.htmltableconverter.controller;

import com.kretsev.htmltableconverter.converter.HtmlXlsConverter;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

@RestController
@RequestMapping("/html")
@RequiredArgsConstructor
public class htmlController {
    private final HtmlXlsConverter converter;

    @PostMapping("/to-xls")
    public ResponseEntity<?> ListHydrologicalPosts(
            @RequestBody String htmlData) throws IOException {
        LocalDate date = LocalDate.now();
        String fileName = "report_" + date.format(DateTimeFormatter.ofPattern("dd-MM-yyyy")) + ".xls";
        String headerValue = "attachment; filename=\"" + fileName + "\"";
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, headerValue)
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(converter.convertTable(htmlData));
    }
}
