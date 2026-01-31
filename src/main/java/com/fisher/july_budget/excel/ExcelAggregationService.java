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
import org.apache.poi.ss.usermodel.CellStyle;
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
        put("яндекс такси", "такси");
        put("парковк", "парковки");
        put("мтс", "связь");
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

                var category = resolveCategory(description.getStringCellValue().toLowerCase());

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

            // ===== Сводка =====
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Категория");
            header.createCell(1).setCellValue("Сумма");

            int rowIndex = 1;
            for (Map.Entry<String, BigDecimal> entry : totals.entrySet()) {
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(entry.getKey());
                row.createCell(1).setCellValue(entry.getValue().setScale(2, RoundingMode.HALF_UP).doubleValue());
            }

            // Автосайз только для колонок сводки
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);

            // ===== Блок "без категории" =====
            if (!noCategoryRows.isEmpty()) {
                rowIndex++;

                Row noCatHeader = sheet.createRow(rowIndex++);
                noCatHeader.createCell(0).setCellValue("Операции без категории");

                int noCatStartRowIndex = rowIndex; // первая строка с данными "без категории"

                for (Row sourceRow : noCategoryRows) {
                    Row targetRow = sheet.createRow(rowIndex++);
                    copyRow(sourceRow, targetRow, outWorkbook);
                }

                // Автосайз ТОЛЬКО по контенту блока "без категории"
                // (чтобы сводка не влияла на ширину этих колонок)
                int maxCols = 0;
                for (int r = noCatStartRowIndex; r < rowIndex; r++) {
                    Row r0 = sheet.getRow(r);
                    if (r0 != null && r0.getLastCellNum() > maxCols) {
                        maxCols = r0.getLastCellNum();
                    }
                }

                for (int c = 0; c < maxCols; c++) {
                    int maxWidth = sheet.getColumnWidth(c); // текущее (после сводки)
                    sheet.autoSizeColumn(c);

                    int widthByNoCat = sheet.getColumnWidth(c);

                    // откатим "влияние сводки": берём максимум из
                    // - ширины, которая была до автосайза (сводка)
                    // - ширины, полученной по "без категории"
                    sheet.setColumnWidth(c, Math.max(maxWidth, widthByNoCat));
                }
            }

            outWorkbook.write(outputStream);
            return outputStream.toByteArray();
        } catch (IOException ex) {
            throw new IllegalStateException("Не удалось сформировать Excel файл.", ex);
        }
    }

    private String resolveCategory(String description) {
        if (description == null) {
            return null;
        }

        String normalized = description.trim().toLowerCase();
        if (normalized.isEmpty()) {
            return null;
        }

        // 1) точное совпадение
        String exact = CATEGORIES.get(normalized);
        if (exact != null) {
            return exact;
        }

        // 2) по префиксу (startsWith)
        for (Map.Entry<String, String> entry : CATEGORIES.entrySet()) {
            if (normalized.startsWith(entry.getKey())) {
                return entry.getValue();
            }
        }

        return null;
    }

    private void copyRow(Row srcRow, Row destRow, Workbook destWb) {
        if (srcRow == null) return;

        destRow.setHeight(srcRow.getHeight());

        for (int i = 0; i < srcRow.getLastCellNum(); i++) {
            Cell srcCell = srcRow.getCell(i);
            if (srcCell == null) continue;

            Cell destCell = destRow.createCell(i);

            // 1) копируем значение по типу
            copyCellValue(srcCell, destCell);

            // 2) копируем стиль
            copyCellStyle(srcCell, destCell, destWb);
        }
    }

    private void copyCellValue(Cell srcCell, Cell destCell) {
        switch (srcCell.getCellType()) {
            case STRING -> destCell.setCellValue(srcCell.getRichStringCellValue());
            case NUMERIC -> destCell.setCellValue(srcCell.getNumericCellValue());
            case BOOLEAN -> destCell.setCellValue(srcCell.getBooleanCellValue());
            case FORMULA -> destCell.setCellFormula(srcCell.getCellFormula());
            case BLANK -> destCell.setBlank();
            default -> {
                // ERROR / _NONE — на всякий случай
                destCell.setCellValue(srcCell.toString());
            }
        }
    }

    private void copyCellStyle(Cell srcCell, Cell destCell, Workbook destWb) {
        // ВАЖНО: CellStyle нельзя просто взять из другого Workbook и присвоить.
        // Нужно создать новый стиль в destWb и клонировать в него.
        CellStyle newStyle = destWb.createCellStyle();
        newStyle.cloneStyleFrom(srcCell.getCellStyle());
        destCell.setCellStyle(newStyle);
    }

    private void autoSizeAllColumns(Sheet sheet) {
        if (sheet.getPhysicalNumberOfRows() == 0) {
            return;
        }

        Row firstRow = sheet.getRow(0);
        if (firstRow == null) {
            return;
        }

        int lastCellNum = firstRow.getLastCellNum();
        for (int i = 0; i < lastCellNum; i++) {
            sheet.autoSizeColumn(i);
        }
    }

}
