package com.fisher.july_budget.excel;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class ExcelAggregationService {

	private static final int DEFAULT_NAME_COLUMN = 0;
	private static final int DEFAULT_PRICE_COLUMN = 1;
	private static final int DEFAULT_CATEGORY_COLUMN = 2;

	public byte[] aggregateByCategory(InputStream inputStream) {
		Map<String, BigDecimal> totals = new LinkedHashMap<>();
		try (Workbook workbook = WorkbookFactory.create(inputStream)) {
			Sheet sheet = workbook.getSheetAt(0);
			if (sheet == null) {
				throw new IllegalArgumentException("Excel файл пустой.");
			}

			DataFormatter formatter = new DataFormatter();
			int rowStart = detectHeaderRow(sheet, formatter).map(index -> index + 1).orElse(0);
			for (int i = rowStart; i <= sheet.getLastRowNum(); i++) {
				Row row = sheet.getRow(i);
				if (row == null) {
					continue;
				}
				String category = formatter.formatCellValue(row.getCell(DEFAULT_CATEGORY_COLUMN)).trim();
				if (category.isBlank()) {
					continue;
				}
				BigDecimal price = parsePrice(row.getCell(DEFAULT_PRICE_COLUMN), formatter);
				if (price == null) {
					continue;
				}
				totals.merge(category, price, BigDecimal::add);
			}
		} catch (IOException ex) {
			throw new IllegalStateException("Не удалось прочитать Excel файл.", ex);
		}

		return buildSummaryWorkbook(totals);
	}

	private Optional<Integer> detectHeaderRow(Sheet sheet, DataFormatter formatter) {
		Row firstRow = sheet.getRow(0);
		if (firstRow == null) {
			return Optional.empty();
		}
		String nameHeader = formatter.formatCellValue(firstRow.getCell(DEFAULT_NAME_COLUMN)).trim().toLowerCase();
		String priceHeader = formatter.formatCellValue(firstRow.getCell(DEFAULT_PRICE_COLUMN)).trim().toLowerCase();
		String categoryHeader = formatter.formatCellValue(firstRow.getCell(DEFAULT_CATEGORY_COLUMN)).trim().toLowerCase();
		if (nameHeader.contains("name") || nameHeader.contains("покупк")
				|| priceHeader.contains("price") || priceHeader.contains("цен")
				|| categoryHeader.contains("category") || categoryHeader.contains("категор")) {
			return Optional.of(0);
		}
		return Optional.empty();
	}

	private BigDecimal parsePrice(Cell cell, DataFormatter formatter) {
		if (cell == null) {
			return null;
		}
		if (cell.getCellType() == CellType.NUMERIC) {
			return BigDecimal.valueOf(cell.getNumericCellValue());
		}
		String value = formatter.formatCellValue(cell).trim();
		if (value.isBlank()) {
			return null;
		}
		value = value.replace(",", ".").replaceAll("\\s+", "");
		try {
			return new BigDecimal(value);
		} catch (NumberFormatException ex) {
			return null;
		}
	}

	private byte[] buildSummaryWorkbook(Map<String, BigDecimal> totals) {
		try (Workbook outWorkbook = new XSSFWorkbook();
				ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
			Sheet sheet = outWorkbook.createSheet("Сводка");
			Row header = sheet.createRow(0);
			header.createCell(0).setCellValue("Категория");
			header.createCell(1).setCellValue("Сумма");

			int rowIndex = 1;
			for (Map.Entry<String, BigDecimal> entry : totals.entrySet()) {
				Row row = sheet.createRow(rowIndex++);
				row.createCell(0).setCellValue(entry.getKey());
				row.createCell(1).setCellValue(entry.getValue().setScale(2, RoundingMode.HALF_UP).doubleValue());
			}
			sheet.autoSizeColumn(0);
			sheet.autoSizeColumn(1);

			outWorkbook.write(outputStream);
			return outputStream.toByteArray();
		} catch (IOException ex) {
			throw new IllegalStateException("Не удалось сформировать Excel файл.", ex);
		}
	}
}
