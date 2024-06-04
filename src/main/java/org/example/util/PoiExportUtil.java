package org.example.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
/**
 * @author: Alex Hu
 * @createTime: 2024/06/03 19:02
 * @description:
 */
public class PoiExportUtil {

    /**
     * Custom annotation for exporting Excel file
     */
    @Target(ElementType.FIELD)
    @Retention(RetentionPolicy.RUNTIME)
    public @interface PoiExportField {
        // Label of the column
        String label();

        // Order of the column,default 0,means the first column
        int order() default 0;

        // If true, this field will be used to create subgroup rows
        boolean subGroup() default false;

        // Width of the column
        int width() default 20;

        // Alignment of the column
        HorizontalAlignment align() default HorizontalAlignment.LEFT;
    }

    /**
     * Export data to excel file
     *
     * @param list     List of data
     * @param fileName File name
     * @param <T>      Type of data
     */
    public <T> void exportToExcel(List<T> list, String fileName) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Data");
        writeSheet(sheet, list, null, null, null);

        // Write to file
        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
            fileOut.flush();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Write data to a sheet
     *
     * @param sheet         Sheet
     * @param dataList      List of data
     * @param headerStyle   Header style
     * @param subGroupStyle Subgroup style
     * @param dataCellStyle Data cell style
     * @param <T>           Type of data
     */
    public <T> void writeSheet(Sheet sheet, List<T> dataList, CellStyle headerStyle, CellStyle subGroupStyle, CellStyle dataCellStyle) {
        // If data list is empty, return
        if (dataList == null || dataList.isEmpty()) {
            return;
        }

        // If styles are not provided, use default styles
        if (headerStyle == null) {
            headerStyle = createDefaultHeaderStyle(sheet.getWorkbook());
        }
        if (subGroupStyle == null) {
            subGroupStyle = createDefaultSubGroupStyle(sheet.getWorkbook());
        }
        if (dataCellStyle == null) {
            dataCellStyle = createDefaultDataCellStyle(sheet.getWorkbook());
        }

        Field[] fields = dataList.get(0).getClass().getDeclaredFields();
        // Filter fields with PoiExportField annotation
        List<Field> annotatedFields = new ArrayList<>();
        // Filter fields with PoiExportField annotation and subGroup is true, for creating subgroup rows
        List<Field> subGroupFields = new ArrayList<>();

        // Filter fields with PoiExportField annotation and sort them by order attribute
        for (Field field : fields) {
            PoiExportField annotation = field.getAnnotation(PoiExportField.class);
            if (annotation != null) {
                if (annotation.subGroup()) {
                    subGroupFields.add(field);
                } else {
                    annotatedFields.add(field);
                }
            }
        }
        // Sort fields by order attribute
        annotatedFields.sort(Comparator.comparingInt(field -> {
            PoiExportField annotation = field.getAnnotation(PoiExportField.class);
            return annotation.order();
        }));

        //annotated fields is empty, return
        if (annotatedFields.isEmpty()) {
            return;
        }
        // Create header row
        createHeaderRow(sheet, annotatedFields, headerStyle);

        // Create data rows
        createSheetWithData(sheet, dataList, annotatedFields, subGroupFields, subGroupStyle, dataCellStyle);
    }

    /**
     * Create header row
     *
     * @param sheet           Sheet
     * @param annotatedFields List of annotated fields
     * @param headerStyle     Header style
     */
    private void createHeaderRow(Sheet sheet, List<Field> annotatedFields, CellStyle headerStyle) {
        int lastRowNum = sheet.getLastRowNum();
        Row headerRow = sheet.createRow(lastRowNum + 1);
        for (int i = 0; i < annotatedFields.size(); i++) {
            Field field = annotatedFields.get(i);
            PoiExportField annotation = field.getAnnotation(PoiExportField.class);
            Cell headerCell = headerRow.createCell(i);
            headerCell.setCellValue(annotation.label());
            headerCell.setCellStyle(headerStyle);
            // Set column width
            sheet.setColumnWidth(i, annotation.width() * 256);
        }
    }

    /**
     * Create data rows
     *
     * @param sheet           Sheet
     * @param dataList        List of data
     * @param annotatedFields List of annotated fields
     * @param subGroupFields  List of subgroup fields
     * @param subGroupStyle   Subgroup style
     * @param dataCellStyle   Data cell style
     * @param <T>             Type of data
     */

    private <T> void createSheetWithData(Sheet sheet, List<T> dataList, List<Field> annotatedFields, List<Field> subGroupFields, CellStyle subGroupStyle, CellStyle dataCellStyle) {
        String lastSubGroupValue = null;
        int rowIndex = sheet.getLastRowNum() + 1;
        for (T data : dataList) {
            // Create subgroup row
            if (subGroupFields != null && !subGroupFields.isEmpty()) {
                String currentSubGroupValue = getSubGroupValue(data, subGroupFields);
                if (!currentSubGroupValue.equals(lastSubGroupValue)) {
                    Row subGroupRow = sheet.createRow(rowIndex++);
                    Cell subGroupCell = subGroupRow.createCell(0);
                    subGroupCell.setCellValue(currentSubGroupValue);
                    CellRangeAddress mergeRegion = new CellRangeAddress(rowIndex - 1, rowIndex - 1, 0, annotatedFields.size() - 1);
                    sheet.addMergedRegion(mergeRegion);
                    //CellRangeAddress  mergeRegion = sheet.getMergedRegion(sheet.getNumMergedRegions() - 1);
                    RegionUtil.setBorderBottom(BorderStyle.THIN, mergeRegion, sheet);
                    RegionUtil.setBorderTop(BorderStyle.THIN, mergeRegion, sheet);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, mergeRegion, sheet);
                    RegionUtil.setBorderRight(BorderStyle.THIN, mergeRegion, sheet);
                    subGroupCell.setCellStyle(subGroupStyle);
                    lastSubGroupValue = currentSubGroupValue;
                }
            }
            // Create data row
            Row row = sheet.createRow(rowIndex++);
            for (int j = 0; j < annotatedFields.size(); j++) {
                Field field = annotatedFields.get(j);
                PoiExportField annotation = field.getAnnotation(PoiExportField.class);
                Cell cell = row.createCell(j);
                try {
                    // Get field value from getter method
                    field.setAccessible(true);
                    String getterName = "get" + Character.toUpperCase(field.getName().charAt(0)) + field.getName().substring(1);
                    Method getterMethod = data.getClass().getMethod(getterName);
                    Object value = getterMethod.invoke(data);
                    if (value != null) {
                        cell.setCellValue(value.toString());
                    }
                    // Set cell style alignment
                    CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
                    cellStyle.cloneStyleFrom(dataCellStyle);
                    cellStyle.setAlignment(annotation.align());
                    cell.setCellStyle(cellStyle);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * Get subgroup value
     *
     * @param data           Data
     * @param subGroupFields List of subgroup fields
     * @param <T>            Type of data
     * @return Subgroup value
     */
    private <T> String getSubGroupValue(T data, List<Field> subGroupFields) {
        StringBuilder subGroupValue = new StringBuilder();
        for (Field field : subGroupFields) {
            try {
                field.setAccessible(true);
                String getterName = "get" + Character.toUpperCase(field.getName().charAt(0)) + field.getName().substring(1);
                Method getterMethod = data.getClass().getMethod(getterName);
                Object value = getterMethod.invoke(data);
                if (value != null) {
                    subGroupValue.append(value.toString()).append(" ");
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return subGroupValue.toString().trim();
    }

    /**
     * Create and return default header style
     *
     * @param workbook Workbook
     * @return CellStyle
     */
    private CellStyle createDefaultHeaderStyle(Workbook workbook) {
        Font fontBold = workbook.createFont();
        fontBold.setBold(true);
        fontBold.setFontHeightInPoints((short) 12);
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setWrapText(true);
        headerStyle.setFont(fontBold);
        return headerStyle;
    }

    /**
     * Create and return default subgroup style
     *
     * @param workbook Workbook
     * @return CellStyle
     */
    private CellStyle createDefaultSubGroupStyle(Workbook workbook) {
        // Create and return default subgroup style
        Font fontBold = workbook.createFont();
        fontBold.setBold(true);
        fontBold.setFontHeightInPoints((short) 11);
        CellStyle subGroupStyle = workbook.createCellStyle();
        subGroupStyle.setBorderTop(BorderStyle.THIN);
        subGroupStyle.setBorderBottom(BorderStyle.THIN);
        subGroupStyle.setBorderLeft(BorderStyle.THIN);
        subGroupStyle.setBorderRight(BorderStyle.THIN);
        subGroupStyle.setAlignment(HorizontalAlignment.CENTER);
        subGroupStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        subGroupStyle.setFont(fontBold);
        return subGroupStyle;
    }

    /**
     * Create and return default data cell style
     *
     * @param workbook Workbook
     * @return CellStyle
     */
    private CellStyle createDefaultDataCellStyle(Workbook workbook) {
        // Create and return default data cell style
        CellStyle dataCellStyle = workbook.createCellStyle();
        Font fontBold = workbook.createFont();
        fontBold.setBold(true);
        fontBold.setFontHeightInPoints((short) 11);
        dataCellStyle.setBorderTop(BorderStyle.THIN);
        dataCellStyle.setBorderBottom(BorderStyle.THIN);
        dataCellStyle.setBorderLeft(BorderStyle.THIN);
        dataCellStyle.setBorderRight(BorderStyle.THIN);
        dataCellStyle.setFont(fontBold);
        return dataCellStyle;
    }
}
