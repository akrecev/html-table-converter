package com.kretsev.htmltableconverter.controller;

import com.kretsev.htmltableconverter.service.HtmlToXlsService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

@RestController
@RequestMapping("/html_table")
@RequiredArgsConstructor
public class htmlController {
    private final HtmlToXlsService htmlToXlsService;

    @PostMapping("/{index}")
    public ResponseEntity<?> getTableForm(
            @PathVariable(required = false) String index,
            @RequestParam(value = "variant", defaultValue = "") String variant,
            @RequestBody String htmlData) throws IOException {
        LocalDate date = LocalDate.now();
        String fileName = "report_" + date.format(DateTimeFormatter.ofPattern("yyyy-MM-dd")) + ".xls";
        String headerValue = "attachment; filename=\"" + fileName + "\"";
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, headerValue)
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(htmlToXlsService.getXlsForm(htmlData, index, variant));
    }
}
