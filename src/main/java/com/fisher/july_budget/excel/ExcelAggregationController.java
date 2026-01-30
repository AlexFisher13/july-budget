package com.fisher.july_budget.excel;

import lombok.RequiredArgsConstructor;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import static org.springframework.http.MediaType.parseMediaType;

@RestController
@RequestMapping("/api/excel")
@RequiredArgsConstructor
public class ExcelAggregationController {

    private final ExcelAggregationService excelAggregationService;

    @PostMapping(value = "/aggregate", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<byte[]> aggregate(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty()) {
            return ResponseEntity.badRequest().build();
        }

        byte[] result;
        try {
            result = excelAggregationService.aggregateByCategory(file.getInputStream());
        } catch (Exception ex) {
            return ResponseEntity.badRequest().build();
        }

        return ResponseEntity.ok()
                .headers(buildExcelHeaders())
                .body(result);
    }

    private HttpHeaders buildExcelHeaders() {
        return new HttpHeaders() {{
            setContentType(parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            setContentDisposition(ContentDisposition.attachment().filename("summary.xlsx").build());
        }};
    }
}
