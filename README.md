自定义注解实现Excel 导出
======================
### 1. 注解定义,定义导出Excel的字段
```java
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
```
### 2. 实体类，使用注解定义导出字段，不导出的字段不用加注解
```java
@Data
public class OrderVO {
    @PoiExportUtil.PoiExportField(label = "订单编号", order = 1, align = HorizontalAlignment.CENTER)
    private String orderNo;
    @PoiExportUtil.PoiExportField(label = "订单用户", order = 2, align = HorizontalAlignment.CENTER)
    private String orderUser;
    @PoiExportUtil.PoiExportField(label = "订单时间", order = 3, align = HorizontalAlignment.CENTER)
    private String orderTime;
    @PoiExportUtil.PoiExportField(label = "订单金额", order = 4, width = 15, align = HorizontalAlignment.RIGHT)
    private String orderAmount;
    private String orderDesc;
    private String orderRemark;
    private String orderPhone;
    private String orderZipCode;
    @PoiExportUtil.PoiExportField(label = "订单国家", subGroup = true)
    private String orderCountry;
    @PoiExportUtil.PoiExportField(label = "订单省份", subGroup = true)
    private String orderProvince;
    @PoiExportUtil.PoiExportField(label = "订单城市", order = 6)
    private String orderCity;
    @PoiExportUtil.PoiExportField(label = "详细地址", order = 7)
    private String orderAddressDetail;

}
```
### 3. Excel导出工具类，注解定义放到工具类中，方便使用
```java
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
```

### 4. 测试
```java
public class PoiExportUtilTest {
    @Test
    public void exportToExcel() {
        PoiExportUtil poiExportUtil = new PoiExportUtil();
        List<OrderVO> orderVOList = generateOrders();
        poiExportUtil.exportToExcel(orderVOList, "order.xlsx");
    }

    public List<OrderVO> generateOrders() {
        String[] COUNTRIES = {"China", "Japan", "Canada"};
        Random RANDOM = new Random();
        List<OrderVO> orders = new ArrayList<>();
        for (int i = 0; i < 30; i++) {
            String orderNo = "OrderNo" + (i + 1);
            String orderUser = "User" + (i + 1);
            String orderTime = "Time" + (i + 1);
            String orderAmount = RANDOM.nextInt(10000) + ".00";
            String orderDesc = "Desc" + (i + 1);
            String orderRemark = "Remark" + (i + 1);
            String orderPhone = "Phone" + (i + 1);
            String orderZipCode = "ZipCode" + (i + 1);
            String orderCountry = COUNTRIES[RANDOM.nextInt(COUNTRIES.length)];
            String orderProvince = "Province" + (i + 1) % 3;
            String orderCity = "City" + (i + 1);
            String orderAddressDetail = "AddressDetail" + (i + 1);
            OrderVO order = createOrder(orderNo, orderUser, orderTime, orderAmount, orderDesc, orderRemark, orderPhone,
                    orderZipCode, orderCountry, orderProvince, orderCity, orderAddressDetail);
            orders.add(order);
        }
        // Sort by orderCountry and orderTime
        return orders.stream()
                .sorted(Comparator.comparing(OrderVO::getOrderCountry)
                        .thenComparing(OrderVO::getOrderProvince)
                        .thenComparing(OrderVO::getOrderTime))
                .collect(Collectors.toList());
    }


    private OrderVO createOrder(String orderNo, String orderUser, String orderTime, String orderAmount,
                                String orderDesc, String orderRemark, String orderPhone, String orderZipCode,
                                String orderCountry, String orderProvince, String orderCity, String orderAddressDetail) {
        OrderVO order = new OrderVO();
        order.setOrderNo(orderNo);
        order.setOrderUser(orderUser);
        order.setOrderTime(orderTime);
        order.setOrderAmount(orderAmount);
        order.setOrderDesc(orderDesc);
        order.setOrderRemark(orderRemark);
        order.setOrderPhone(orderPhone);
        order.setOrderZipCode(orderZipCode);
        order.setOrderCountry(orderCountry);
        order.setOrderProvince(orderProvince);
        order.setOrderCity(orderCity);
        order.setOrderAddressDetail(orderAddressDetail);
        return order;
    }
}
```
