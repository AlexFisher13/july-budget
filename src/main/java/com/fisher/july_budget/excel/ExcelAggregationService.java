package com.fisher.july_budget.excel;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class ExcelAggregationService {

    private static final int ROW_START = 1;
    private static final int PRICE_CELL = 4;

    public byte[] aggregateByCategory(InputStream inputStream) {
        Map<String, BigDecimal> totals = new LinkedHashMap<>();

        try (Workbook workbook = WorkbookFactory.create(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);

            for (int i = ROW_START; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);

                BigDecimal price = parsePriceToPositive(row.getCell(PRICE_CELL));
                if (price == null) {
                    continue;
                }
                totals.merge("category", price, BigDecimal::add);
            }
        } catch (IOException ex) {
            throw new IllegalStateException("Не удалось прочитать Excel файл.", ex);
        }

        return buildSummaryWorkbook(totals);
    }


    private BigDecimal parsePriceToPositive(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC) {
            return BigDecimal.valueOf(cell.getNumericCellValue()).abs();
        } else {
            throw new RuntimeException(cell.getColumnIndex() + "invalid value");
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
