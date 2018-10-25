package com.penghaohuan.excel.handler;

import com.penghaohuan.excel.annotation.ExcelDesc;
import com.penghaohuan.excel.exception.ExcelTemplateException;
import com.penghaohuan.excel.exception.ExcelValidateException;
import com.penghaohuan.excel.model.CellPosition;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * Excel 导入工具.
 * 使用注解@ExcelDesc，通过预定义数据类型的方式，导入Excel文件，解析行数据为预定义的数据类型.
 * 导入时读取excel,得到的结果是一个list<T>.T是自己定义的对象
 *
 * <p>支持合并的单元格识别</p>
 *
 * <p>
 *     支持数据校验：
 *     1. 支持是否为空校验
 *     2. 支持正则表达式校验
 *     3. 支持自定义方法校验
 * </p>
 * <p>
 *     支持导入文件校验：
 *     1. 根据Excel-VO实体标注的列，校验导入文件中是否包含这些列。
 *     2. 对于Date类型，只支持Excel中的单元格格式为日期格式，会进行日期格式校验
 * </p>
 *
 * Excel校验会全内容校验完毕后再返回异常信息，
 * 每一条异常信息以换行符(\r\n)连接，作为Exception中的message返回
 * @see ExcelDesc
 * @param <T> 对应Excel行数据的数据类型
 */
public class ExcelImporter<T> {

    /**
     * 日志.
     */
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelImporter.class);

    /**
     * 在验证数据有效性时，如果数据有效，加上此正确的符号.
     */
    public static final String CORRECT_SYMBOL = "%c";

    /**
     * 实体类型.
     */
    private Class<T> clazz;

    /**
     * 单元格值位置映射.
     */
    private Map<Integer, Map<Integer, CellPosition>> cellValuePositionMap;

    /**
     * 属性名-校验类映射.
     */
    private Map<String, Object> validatorMap;

    /**
     * 类校验Map-key.
     */
    private static final String CLASS_VALIDATOR_KEY = "clazz";

    /**
     * 构造.
     * @param clazz 实体类型
     */
    public ExcelImporter(final Class<T> clazz) {
        this.clazz = clazz;
    }

    /**
     * 导入excel.
     *
     * 读取第一个sheet.
     * @param fis   文件流 如：new FileInputStream(new File("D:\\test.xlsx"))
     * @param headRowNumbers 表格头行数
     * @return T类型的实体列表
     * @throws ExcelValidateException Excel校验异常
     * @throws ExcelTemplateException Excel模板异常
     */
    public List<T> importExcel(InputStream fis, Integer headRowNumbers) throws ExcelValidateException, ExcelTemplateException {
        final List<T> list = new ArrayList<>();
        try {
            final Sheet sheet = getSheet(fis);
            if (sheet == null) {
                return list;
            }
            initCellPosition(sheet);
            final int rows = sheet.getLastRowNum(); // 得到数据的行数

            if (rows > 0) {
                final Map<Integer, Field> fieldsMap = buildFieldOrder(sheet, headRowNumbers); // 从最后一行表头解析列名
                initValidator();
                final List<String> validateMassages = new LinkedList<>();
                final ExcelDesc classDesc = clazz.getAnnotation(ExcelDesc.class);
                for (int rowNum = headRowNumbers; rowNum <= rows; rowNum++) {
                    T entity = null;
                    boolean keyAttrEmpty = false;

                    for (Map.Entry<Integer, Field> entry : fieldsMap.entrySet()) {
                        final Integer column = entry.getKey();
                        final Field field = entry.getValue();

                        final Cell c = getCell(sheet, rowNum, column);
                        entity = entity == null ? clazz.newInstance() : entity;
                        final ExcelDesc fieldDesc = field.getAnnotation(ExcelDesc.class);
                        final Class<?> fieldType = field.getType();
                        final String exceptionMsg = "第" + (rowNum + 1) + "行【" + fieldDesc.name() + "】列";

                        // 日期格式校验
                        if (Date.class == fieldType && !validateDateCell(c)) {
                            validateMassages.add(exceptionMsg + "日期格式错误");
                            continue;
                        }
                        final String cellValue = getCellValue(c, fieldDesc.dateFormat());
                        final String validateData = validateData(cellValue, exceptionMsg, field, fieldDesc, classDesc);

                        if (validateData.contains(CORRECT_SYMBOL)) {
                            final Object valueFormat = typeFormat(fieldType, validateData.replace(CORRECT_SYMBOL, ""), fieldDesc.dateFormat());
                            field.setAccessible(true);
                            field.set(entity, valueFormat);
                        } else {
                            validateMassages.add(validateData);
                        }

                        if (fieldDesc.keyAttr()) {
                            final Object value = field.get(entity);
                            if (value == null || StringUtils.isEmpty(String.valueOf(value))) {
                                LOGGER.error("Key field " + field.getName() + " is empty, row num: " + rowNum);
                                keyAttrEmpty = true;
                                break;
                            }
                        }
                    }

                    if (!keyAttrEmpty && entity != null) {
                        if (StringUtils.isNotBlank(classDesc.function())) { // 行数据校验
                            final String validateResult = validateRow(entity, "第" + (rowNum + 1) + "行", classDesc);
                            if (!validateResult.contains(CORRECT_SYMBOL)) {
                                validateMassages.add(validateResult);
                            }
                        }

                        list.add(entity);
                    }
                }

                if (validateMassages.size() > 0) {
                    final StringBuilder throwExceptionMsg = new StringBuilder();
                    for (final String msg : validateMassages) {
                        throwExceptionMsg.append(msg).append("\r\n");
                    }
                    throw new ExcelValidateException("\r\n" + throwExceptionMsg.toString());
                }
            }
        } catch (final ExcelValidateException | ExcelTemplateException e) {
            throw e;
        } catch (final Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelTemplateException("文件模板错误");
        }
        return list;
    }


    /**
     * 生成sheet.
     *
     * 读取第一个sheet
     * @param fis 输入流
     * @return sheet
     */
    private Sheet getSheet(final InputStream fis) {
        try {
            final Workbook book = WorkbookFactory.create(fis);
            return book.getSheetAt(0);
        } catch (final Exception e) {
            LOGGER.error(e.getMessage(), e);
        }
        return null;
    }

    /**
     * 对合并的单元格进行单元格值读取位置的初始化，根据该位置，可以读取到单元格的值.
     * @param sheet Excel Sheet
     */
    private void initCellPosition(final Sheet sheet) {
        this.cellValuePositionMap = new HashMap<>();
        final int sheetMergeCount = sheet.getNumMergedRegions();
        for(int i = 0; i < sheetMergeCount; i++) {
            final CellRangeAddress ca = sheet.getMergedRegion(i);
            final int firstRow = ca.getFirstRow();
            final int firstColumn = ca.getFirstColumn();
            final int lastRow = ca.getLastRow();
            initCellPosition(firstRow, firstColumn, lastRow);
        }
    }

    /**
     * 初始化读取位置.
     * @param firstRow 第一行
     * @param firstCol 第一列
     * @param lastRow 最后一行
     */
    private void initCellPosition(final int firstRow, final int firstCol, final int lastRow) {
        for (int rowPos = firstRow; rowPos <= lastRow; ++rowPos) {
            cellValuePositionMap.computeIfAbsent(rowPos, key -> new HashMap<>()).put(firstCol, new CellPosition(firstRow, firstCol));
        }
    }

    /**
     * 获取列名对应的列索引.
     * @param columnName 列名
     * @param columnList 表头列表
     * @return 列索引
     * @throws ExcelTemplateException Excel模板异常
     */
    private int getExcelCol(final String columnName, final List<String> columnList) throws ExcelTemplateException {
        int pos = 0;
        for (; pos < columnList.size(); ++pos) {
            if (columnList.get(pos).equals(columnName)) {
                break;
            }
        }
        if (pos == columnList.size()) {
            LOGGER.error("Can't find column name in excel file, column:" + columnName);
            throw new ExcelTemplateException("文件模板错误，缺少列：" + columnName);
        }
        return pos;
    }

    /**
     * 构建属性与列的对应关系.
     * @param columnList 表头列表
     * @return map
     * @throws ExcelTemplateException Excel模板异常
     */
    private Map<Integer, Field> buildFieldOrder(final List<String> columnList) throws ExcelTemplateException {
        if (columnList == null || columnList.isEmpty()) {
            return new HashMap<>();
        }
        final Field[] allFields = clazz.getDeclaredFields(); // 得到类的所有field.
        final Map<Integer, Field> fieldsMap = new HashMap<>(); // 定义一个map用于存放列的序号和field.
        for (Field field : allFields) {
            // 将有注解的field存放到map中.
            if (field.isAnnotationPresent(ExcelDesc.class)) {
                final ExcelDesc attr = field.getAnnotation(ExcelDesc.class);
                final int pos = getExcelCol(attr.name().trim(), columnList);
                if (pos != -1) {
                    field.setAccessible(true); // 设置类的私有字段属性可访问.
                    fieldsMap.put(pos, field);
                }
            }
        }
        return fieldsMap;
    }

    /**
     * 构建属性与列的对应关系.
     * @param sheet Excel表
     * @param headerNum 表头行数
     * @return map
     * @throws ExcelTemplateException Excel模板异常
     */
    private Map<Integer, Field> buildFieldOrder(final Sheet sheet, final Integer headerNum) throws ExcelTemplateException {
        final Row row = sheet.getRow(headerNum - 1); //获取列头
        final List<String> columnList = new ArrayList<>();
        final int columnCount = row.getLastCellNum();
        for (int pos = 0; pos < columnCount; ++pos) {
            final Cell cell = row.getCell(pos);
            if(cell != null) {
                final String cellValue = cell.getStringCellValue().trim();
                if (StringUtils.isEmpty(cellValue)) {
                    columnList.add(""); // 保留空表头、避免构建属性到表头的映射时下标错位
                    continue;
                }
                columnList.add(cellValue);
            }
        }
        return buildFieldOrder(columnList);
    }

    /**
     * 初始化校验类.
     */
    private void initValidator() {
        final Field[] declaredFields = clazz.getDeclaredFields();
        validatorMap = new HashMap<>(declaredFields.length);
        final ExcelDesc classDesc = clazz.getAnnotation(ExcelDesc.class);
        if (StringUtils.isNotBlank(classDesc.function())) {
            try {
                validatorMap.put(CLASS_VALIDATOR_KEY, getFieldValidateClass(clazz, null, classDesc).newInstance());
            } catch (final IllegalAccessException | InstantiationException e) {
                LOGGER.warn("初始化实体校验类映射异常，Entity: {}, e:{}.", clazz.getName(), e);
            }
        }

        for (Field field : declaredFields) {
            if (field.isAnnotationPresent(ExcelDesc.class)) {
                final ExcelDesc fieldDesc = field.getAnnotation(ExcelDesc.class);
                if (!"".equals(fieldDesc.function())) {
                    try {
                        validatorMap.put(field.getName(), getFieldValidateClass(clazz, fieldDesc, classDesc).newInstance());
                    } catch (final IllegalAccessException | InstantiationException e) {
                        LOGGER.warn("初始化属性校验类映射异常，filed: {}, e:{}.", field.getName(), e);
                    }
                }
            }
        }
    }

    /**
     * 获取该属性的校验类.
     * @param entityClass 实体类
     * @param fieldDesc 属性注解
     * @param classDesc 实体注解
     * @return 校验类
     */
    private Class getFieldValidateClass(final Class entityClass, final ExcelDesc fieldDesc, final ExcelDesc classDesc) {
        Class checkClazz;
        if (fieldDesc != null && !fieldDesc.clazz().equals(ExcelDesc.NoValidateClass.class)) {
            checkClazz = fieldDesc.clazz();
        } else if (classDesc != null && !classDesc.clazz().equals(ExcelDesc.NoValidateClass.class)) {
            checkClazz = classDesc.clazz();
        } else {
            checkClazz = entityClass;
        }
        return checkClazz;
    }

    /**
     * 获取单元格.
     * @param sheet Excel表
     * @param row 行
     * @param column 列
     * @return 单元格
     */
    private Cell getCell(final Sheet sheet, int row, int column) {
        final Map<Integer, CellPosition> rowCellPositions = cellValuePositionMap.get(row);
        CellPosition position = null;
        if (rowCellPositions != null) {
            position = rowCellPositions.get(column);
        }
        if (position != null) {
            row = position.getRow();
            column = position.getColumn();
        }
        final Row rowRecord = sheet.getRow(row);
        if (rowRecord == null) {
            return null;
        }
        return rowRecord.getCell(column);
    }

    /**
     * 校验日期格式单元格.
     * @param cell 单元格
     * @return 格式是否正确
     */
    private boolean validateDateCell(final Cell cell) {
        if (cell == null || CellType.BLANK == cell.getCellType()) {
            return true;
        }
        return CellType.NUMERIC == cell.getCellType() && HSSFDateUtil.isCellDateFormatted(cell);
    }

    /**
     * 获取单元格的内容.
     * @param cell 单元格
     * @param dateFormat a format string for date value
     * @return value of cell
     */
    private String getCellValue(Cell cell, String dateFormat) {
        if (cell == null) {
            return null;
        }
        String value;
        switch (cell.getCellType()) {
            case STRING:
                value = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) { // a date cell
                    SimpleDateFormat sdf = new SimpleDateFormat(dateFormat);
                    value = sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
                } else {
                    value = getNumericCellValue(cell);
                }
                break;
            case FORMULA:
                value = cell.getCellFormula() + "";
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue() + "";
                break;
            case BLANK:
                value = "";
                break;
            case ERROR:
                value = "非法字符";
                break;
            default:
                value = cell.toString();
                break;
        }
        return value;
    }

    /**
     * 获取单元格数值类型值.
     * @param cell 单元格
     * @return 数值类型值
     */
    private String getNumericCellValue(final Cell cell) {
        final double value = cell.getNumericCellValue();
        final CellStyle style = cell.getCellStyle();
        final DecimalFormat format = new DecimalFormat();
        final String temp = style.getDataFormatString();
        // 单元格设置成常规
        if ("General".equals(temp)) {
            format.applyPattern("#.#########");
        }
        return format.format(value);
    }

    /**
     * 验证数据有效性.
     *
     * @param value         数据值
     * @param exceptionMsg 固定异常信息内容
     * @param field         属性
     * @param fieldDesc    导入实体的属性注解
     * @param classDesc  导入实体的类注解
     * @return 返回带有 %c 的数据表示验证通过的数据；否则返回错误信息
     */
    private String validateData(String value, String exceptionMsg, Field field, ExcelDesc fieldDesc, ExcelDesc classDesc) {
        String validateData = StringUtils.isBlank(value) ? CORRECT_SYMBOL : value + CORRECT_SYMBOL;
        if (fieldDesc.isCheckNull()) {
            validateData = checkNull(exceptionMsg, value);
        }
        if (StringUtils.isNotBlank(fieldDesc.regularExpression()) && StringUtils.isNotBlank(value)) {
            validateData = checkRegularExpression(exceptionMsg, value, fieldDesc.regularExpression(), fieldDesc.regularExpressionTip());
        }
        if (StringUtils.isNotBlank(fieldDesc.function())) {
            validateData = checkFunction(exceptionMsg, value, field, fieldDesc, classDesc);
        }
        return validateData;
    }
    /**
     * 验证是否为空.
     *
     * @param exceptionMsg 固定异常信息内容
     * @param value        数据值
     * @return 返回带有 %c 的数据表示验证通过的数据；否则返回错误信息
     */
    private String checkNull(String exceptionMsg, String value) {
        return StringUtils.isBlank(value) ? exceptionMsg + "为空！" : value + CORRECT_SYMBOL;
    }
    /**
     * 验证正则表达式的有效性.
     *
     * @param exceptionMsg 固定异常信息内容
     * @param value        数据值
     * @param regularExpression     正则表达式
     * @param regularExpressionTip 提示信息
     * @return 返回带有 %c 的数据表示验证通过的数据；否则返回错误信息
     */
    private String checkRegularExpression(String exceptionMsg, String value, String regularExpression, String regularExpressionTip) {
        return !Pattern.matches(regularExpression, value.trim()) ? exceptionMsg
                + (StringUtils.isBlank(regularExpressionTip) ? "数据格式错误" : regularExpressionTip) : value
                + CORRECT_SYMBOL;
    }

    /**
     * 方法验证有效性.
     * @param exceptionMsg 固定异常信息内容
     * @param value 数据值
     * @param filed 属性
     * @param fieldDesc 导入属性注解
     * @param classDesc 导入类注解
     * @return 返回带有 %c 的数据表示验证通过的数据；否则返回错误信息
     */
    private String checkFunction(String exceptionMsg, String value, Field filed, ExcelDesc fieldDesc, ExcelDesc classDesc){
        final String function = fieldDesc.function();
        final Class<?> checkClazz = getFieldValidateClass(clazz, fieldDesc, classDesc);
        try {
            final Method method = checkClazz.getMethod(function, String.class, String.class);
            Object validator = validatorMap.get(filed.getName());
            if (validator == null) {
                validator = checkClazz.newInstance();
                validatorMap.put(filed.getName(), validator);
            }
            final Object res = method.invoke(validator, value, exceptionMsg);
            if (res instanceof String) {
                return (String) res;
            }
        } catch (InstantiationException | IllegalAccessException e) {
            LOGGER.warn("Can't find the method of validator！methodName：{}，clazzName：{}", function, checkClazz.getName());
        } catch (NoSuchMethodException | InvocationTargetException e) {
            LOGGER.warn("Can't instant a validator of {}！", checkClazz.getName());
        }
        return exceptionMsg + "方法校验错误！";
    }

    /**
     * 行数据校验.
     * @param entity 实体--行数据
     * @param exceptionMsg 异常信息
     * @param clazzDesc VO类注解
     * @return 校验结果
     */
    private String validateRow(final T entity, final String exceptionMsg, final ExcelDesc clazzDesc) {
        final String function = clazzDesc.function();
        final Class<?> checkClass = getFieldValidateClass(clazz, null, clazzDesc);
        try {
            final Method method = checkClass.getMethod(function, clazz, String.class);
            Object validator = validatorMap.get(CLASS_VALIDATOR_KEY);
            if (validator == null) {
                validator = checkClass.newInstance();
                validatorMap.put(CLASS_VALIDATOR_KEY, validator);
            }
            final Object res = method.invoke(validator, entity, exceptionMsg);
            if (res instanceof String) {
                return (String) res;
            }
        } catch (InstantiationException | IllegalAccessException e) {
            LOGGER.warn("Can't find the method of validator！methodName：{}，clazzName：{}", function, checkClass.getName());
        } catch (NoSuchMethodException | InvocationTargetException e) {
            LOGGER.warn("Can't instant a validator of {}！", checkClass.getName());
        }
        return exceptionMsg + "方法校验错误！";
    }

    /**
     * 按照模板的属性类型进行转换.
     *
     * @param fieldType 模板的属性类型
     * @param cellValue 需要转换的值
     * @return 转换以后的类型对象
     * @throws Exception 通用异常
     */
    private Object typeFormat(Class<?> fieldType, String cellValue, String dateFormat) throws Exception {
        if (StringUtils.isBlank(cellValue)) {
            return null;
        } else if (Integer.TYPE == fieldType || Integer.class == fieldType) {
            return getCellValueForInteger(cellValue);
        } else if (String.class == fieldType) {
            return cellValue;
        } else if (Long.TYPE == fieldType || Long.class == fieldType) {
            return getCellValueForLong(cellValue);
        } else if (Float.TYPE == fieldType || Float.class == fieldType) {
            return getCellValueForFloat(cellValue);
        } else if (Short.TYPE == fieldType || Short.class == fieldType) {
            return getCellValueForShort(cellValue);
        } else if (Double.TYPE == fieldType || Double.class == fieldType) {
            return getCellValueForDouble(cellValue);
        } else if (BigDecimal.class == fieldType) {
            return BigDecimal.valueOf(getCellValueForDouble(cellValue));
        } else if (Date.class == fieldType) {
            final SimpleDateFormat sdf = new SimpleDateFormat(dateFormat);
            return sdf.parse(cellValue);
        } else {
            throw new Exception("导入模板的属性类型【" + fieldType.getName() + "】没有对应的转换程序，请添加！");
        }
    }

    /**
     * 单元格double值转换.
     * @param cellValue 单元格值
     * @return 转换后的值
     */
    private Double getCellValueForDouble(String cellValue) {
        if (cellValue == null) {
            return null;
        }
        if (cellValue.contains(",")) {
            cellValue = cellValue.replaceAll(",", "");
        }
        return Double.parseDouble(cellValue);
    }

    /**
     * 单元格float值转换.
     * @param cellValue 单元格值
     * @return 转换后的值
     */
    private Float getCellValueForFloat(String cellValue) {
        if (cellValue == null) {
            return null;
        }
        if (cellValue.contains(",")) {
            cellValue = cellValue.replaceAll(",", "");
        }
        return Float.parseFloat(cellValue);
    }

    /**
     * 单元格Int值转换.
     * @param cellValue 单元格值
     * @return 转换后的值
     */
    private Integer getCellValueForInteger(String cellValue) {
        if (cellValue == null) {
            return null;
        }
        if (cellValue.contains(",")) {
            cellValue = cellValue.replaceAll(",", "");
        }
        return Integer.parseInt(cellValue);
    }

    /**
     * 单元格short值转换.
     * @param cellValue 单元格值
     * @return 转换后的值
     */
    private Short getCellValueForShort(String cellValue) {
        if (cellValue == null) {
            return null;
        }
        if (cellValue.contains(",")) {
            cellValue = cellValue.replaceAll(",", "");
        }
        return Short.parseShort(cellValue);
    }

    /**
     * 单元格long值转换.
     * @param cellValue 单元格值
     * @return 转换后的值
     */
    private Long getCellValueForLong(String cellValue) {
        if (cellValue == null) {
            return null;
        }
        if (cellValue.contains(",")) {
            cellValue = cellValue.replaceAll(",", "");
        }
        return Long.parseLong(cellValue);
    }

}
