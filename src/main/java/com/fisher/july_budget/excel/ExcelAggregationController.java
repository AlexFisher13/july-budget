package com.fisher.july_budget.excel;

import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/api/excel")
public class ExcelAggregationController {

	private final ExcelAggregationService excelAggregationService;

	public ExcelAggregationController(ExcelAggregationService excelAggregationService) {
		this.excelAggregationService = excelAggregationService;
	}

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

		HttpHeaders headers = new HttpHeaders();
		headers.setContentType(
				MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
		headers.setContentDisposition(ContentDisposition.attachment().filename("summary.xlsx").build());
		return ResponseEntity.ok().headers(headers).body(result);
	}
}
