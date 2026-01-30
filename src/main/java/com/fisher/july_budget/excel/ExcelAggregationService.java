package com.fisher.july_budget.excel;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

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

    private static final int PRICE_CELL = 4;
    private static final int DESC_CELL = 11;

    private final static Map<String, String> CATEGORIES = new HashMap<>() {{
        put("перекрёсток доставка", "продукты");
        put("dostavka perekrestka_sdk", "продукты");
        put("перекрёсток", "продукты");
        put("продуктовый магазин", "продукты");
        put("дикси", "продукты");
        put("микс фрукт", "продукты");
        put("самокат", "продукты");
        put("куулклевер", "продукты");
        put("пятёрочка доставка", "продукты");
        put("пятёрочка", "продукты");
        put("вкусвилл", "продукты");
        put("лукойл", "бензин");
    }};

    public byte[] aggregateByCategory(InputStream inputStream) {
        Map<String, BigDecimal> totals = new LinkedHashMap<>();
        List<Row> noCategoryRows = new LinkedList<>();

        try (Workbook workbook = WorkbookFactory.create(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                var row = sheet.getRow(i);
                var price = parsePriceToPositive(row.getCell(PRICE_CELL));
                var description = row.getCell(DESC_CELL);

                var category = CATEGORIES.get(description.getStringCellValue().toLowerCase());

                if (category == null) {
                    noCategoryRows.add(row);
                    continue;
                }

                totals.merge(category, price, BigDecimal::add);
            }
        } catch (IOException ex) {
            throw new IllegalStateException("Не удалось прочитать Excel файл.", ex);
        }

        return buildSummaryWorkbook(totals, noCategoryRows);
    }


    private BigDecimal parsePriceToPositive(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC) {
            return BigDecimal.valueOf(cell.getNumericCellValue()).abs();
        } else {
            throw new RuntimeException(cell.getColumnIndex() + "invalid value");
        }
    }

    private byte[] buildSummaryWorkbook(Map<String, BigDecimal> totals, List<Row> noCategoryRows) {
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

            // ===== Блок "без категории" =====
            if (!noCategoryRows.isEmpty()) {

                rowIndex++;

                Row noCatHeader = sheet.createRow(rowIndex++);
                noCatHeader.createCell(0).setCellValue("Операции без категории");

                for (Row sourceRow : noCategoryRows) {
                    Row targetRow = sheet.createRow(rowIndex++);
                    copyRowValues(sourceRow, targetRow);
                }
            }

            outWorkbook.write(outputStream);
            return outputStream.toByteArray();
        } catch (IOException ex) {
            throw new IllegalStateException("Не удалось сформировать Excel файл.", ex);
        }
    }

    private void copyRowValues(Row sourceRow, Row targetRow) {
        if (sourceRow == null) {
            return;
        }

        DataFormatter formatter = new DataFormatter();

        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            if (sourceCell == null) {
                continue;
            }

            Cell targetCell = targetRow.createCell(i);
            String value = formatter.formatCellValue(sourceCell);
            targetCell.setCellValue(value);
        }
    }
}
