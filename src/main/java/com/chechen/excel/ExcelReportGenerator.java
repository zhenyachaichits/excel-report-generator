package com.chechen.excel;

import com.chechen.excel.annotation.ExcelColumn;
import com.chechen.excel.annotation.ExcelSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.Arrays;
import java.util.Collection;
import java.util.Map;

public class ExcelReportGenerator {

    private String resultPath;
    private String fileName;

    public ExcelReportGenerator(String resultPath, String fileName) {
        this.resultPath = resultPath;
        this.fileName = fileName;
    }

    private Workbook createWorkBook() {
        Workbook workbook = null;

        if (fileName.contains(".xls")) {
            workbook = new HSSFWorkbook();
        } else if (fileName.contains(".xlsx")) {
            workbook = new XSSFWorkbook();
        }

        return workbook;
    }

    private Workbook saveWorkbook(Workbook workbook) {
        try (FileOutputStream fileOutputStream = new FileOutputStream(resultPath + fileName)) {
            workbook.write(fileOutputStream);
        } catch (IOException e) {
            e.printStackTrace();
            // TODO: 11/29/2016 log
        }

        return workbook;
    }

    public Workbook generateReport(Object... sheets) {
        Workbook workbook = createWorkBook();

        Arrays.stream(sheets).forEach(sheetObject -> {
            Class<?> clazz = sheetObject.getClass();
            ExcelSheet excelSheet = clazz.getDeclaredAnnotation(ExcelSheet.class);

            if (excelSheet == null) {
                throw new IllegalArgumentException(clazz.getSimpleName() + " is not annotated as ExcelSheet");
            }

            Sheet sheet = workbook.createSheet(excelSheet.name());
            final int[] cellCounter = {0};

            Arrays.stream(clazz.getDeclaredFields())
                    .forEach(field -> fieldToColumn(sheetObject, sheet, field, cellCounter));

        });

        return saveWorkbook(workbook);
    }

    private void fieldToColumn(Object sheetObject, Sheet sheet, Field field, int[] cellCounter) {
        try {
            field.setAccessible(true);
            ExcelColumn excelColumn = field.getDeclaredAnnotation(ExcelColumn.class);
            if (excelColumn != null) {
                Type fieldType = field.getGenericType();
                if (fieldType instanceof ParameterizedType) {
                    Class<?> typeClass = field.getType();
                    if (Map.class.equals(typeClass)) {
                        Map<Object, Object> fieldMap = (Map<Object, Object>) field.get(sheetObject);
                        processMapColumns(fieldMap, sheet, cellCounter, excelColumn);
                    } else {
                        Collection<Object> fieldList = (Collection<Object>) field.get(sheetObject);
                        processCollectionColumn(fieldList, sheet, cellCounter, excelColumn);
                    }
                } else {
                    Object fieldObject = field.get(sheetObject);
                    processBasicColumn(fieldObject, sheet, cellCounter, excelColumn);
                }
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
            // TODO: 11/29/2016 log
        }
    }

    private void processCollectionColumn(Collection<Object> fieldList, Sheet sheet,
                                         int[] cellCounter, ExcelColumn excelColumn) {
        final int[] rowCounter = {setColumnTitle(sheet, cellCounter[0], 0, excelColumn.name())};

        for (Object value : fieldList) {
            Row currentRow = getRow(sheet, rowCounter[0]++);

            Class<?> clazz = value.getClass();
            ExcelSheet excelSheet = clazz.getDeclaredAnnotation(ExcelSheet.class);

            if (excelSheet != null) {
                currentRow.createCell(cellCounter[0]).setCellValue(excelSheet.name());
            }

            if ("".equals(value.toString())) {
                currentRow.createCell(cellCounter[0] + 1).setCellValue(excelColumn.emptyCellMessage());
            } else {
                currentRow.createCell(cellCounter[0]).setCellValue(value.toString());
            }
        }

        if (fieldList.isEmpty()) {
            Row currentRow = getRow(sheet, rowCounter[0]);
            currentRow.createCell(cellCounter[0]).setCellValue(excelColumn.messageSuccess());
        }

        cellCounter[0] += 2;
    }

    private void processBasicColumn(Object fieldObject, Sheet sheet,
                                    int[] cellCounter, ExcelColumn excelColumn) {

        int rowCounter = setColumnTitle(sheet, cellCounter[0], 0, excelColumn.name());

        if (fieldObject instanceof Boolean || boolean.class.equals(fieldObject.getClass())) {
            Row currentRow = getRow(sheet, rowCounter);
            if ((Boolean) fieldObject) {
                currentRow.createCell(cellCounter[0]).setCellValue(excelColumn.messageSuccess());
            } else {
                currentRow.createCell(cellCounter[0]).setCellValue(excelColumn.messageError());
            }
        } else {
            Row currentRow = getRow(sheet, rowCounter);
            currentRow.createCell(cellCounter[0]).setCellValue(fieldObject.toString());
        }

        cellCounter[0]++;
    }

    private int setColumnTitle(Sheet sheet, int cellCounter, int rowCounter, String value) {
        CellStyle style = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style.setFont(font);
        style.setAlignment(CellStyle.ALIGN_CENTER);

        int width = ((int) (value.length() * 1.14388)) * 256;
        sheet.setColumnWidth(cellCounter, width);

        Row currentRow = getRow(sheet, rowCounter++);
        Cell cell = currentRow.createCell(cellCounter);
        cell.setCellValue(value);
        cell.setCellStyle(style);

        return rowCounter;
    }

    private Row getRow(Sheet sheet, int index) {
        Row row = sheet.getRow(index);
        if (row == null) {
            row = sheet.createRow(index);
        }

        return row;
    }

    private void processMapColumns(Map fieldMap, Sheet sheet,
                                   int[] cellCounter, ExcelColumn excelColumn) {

        int startRowCounter = setColumnTitle(sheet, cellCounter[0], 0, excelColumn.name());
        int startCellCounter = cellCounter[0];


        String[] subColumnNames = excelColumn.subColumn();
        Arrays.stream(subColumnNames).forEach(name -> setColumnTitle(sheet, cellCounter[0]++, startRowCounter, name));

        cellCounter[0] = startCellCounter;


        sheet.addMergedRegion(new CellRangeAddress(startRowCounter - 1, startRowCounter - 1,
                startCellCounter, subColumnNames.length + startCellCounter - 1));


        int[] rowCounter = {startRowCounter};
        if (subColumnNames.length > 0) {
            rowCounter[0]++;
        }

        fieldMap.forEach((key, value) -> {
            Class<?> clazz = key.getClass();
            ExcelSheet excelSheet = clazz.getDeclaredAnnotation(ExcelSheet.class);

            String keyValue;
            if (excelSheet != null) {
                keyValue = excelSheet.name();
            } else {
                keyValue = key.toString();
            }

            if (value instanceof Collection) {
                Collection valueList = (Collection) value;
                valueList.forEach(currentValue -> {
                    Row listRow = getRow(sheet, rowCounter[0]++);
                    listRow.createCell(cellCounter[0]).setCellValue(keyValue);

                    listRow.createCell(cellCounter[0] + 1).setCellValue(currentValue.toString());

                    if (!"".equals(excelColumn.optionalCellMessage())) {
                        if ("".equals(currentValue.toString())) {
                            listRow.createCell(cellCounter[0] + 2).setCellValue(excelColumn.emptyCellMessage());
                        } else {
                            listRow.createCell(cellCounter[0] + 2).setCellValue(excelColumn.optionalCellMessage());
                        }
                    }
                });

            } else {
                Row currentRow = getRow(sheet, rowCounter[0]++);
                currentRow.createCell(cellCounter[0]).setCellValue(keyValue);

                currentRow.createCell(cellCounter[0] + 1).setCellValue(value.toString());

                if (!"".equals(excelColumn.optionalCellMessage())) {
                    if ("".equals(value.toString())) {
                        currentRow.createCell(cellCounter[0] + 2).setCellValue(excelColumn.emptyCellMessage());
                    } else {
                        currentRow.createCell(cellCounter[0] + 2).setCellValue(excelColumn.optionalCellMessage());
                    }
                }
            }
        });

        cellCounter[0] += subColumnNames.length + 1;
    }
}
