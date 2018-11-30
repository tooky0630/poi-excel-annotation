package com.penghaohuan.excel.handler;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import com.penghaohuan.excel.annotation.ExportExcelDesc;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public final class ExcelExporter<T> {

    /**
     * 日志.
     */
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelExporter.class);

    /**
     * 实体类型.
     */
    private Class<T> clazz;

    /**
     * 构造.
     * @param clazz 实体类型
     */
    public ExcelExporter(final Class<T> clazz) {
        this.clazz = clazz;
    }


    /**
     * 对list数据源将其里面的数据导入到excel表单.
     * @param list 实体列表
     * @param sheetName 工作表的名称
     * @param sheetSize 每个sheet中数据的行数,此数值必须小于65536
     * @param output java输出流
     * @throws IOException 响应流输出异常
     */
    public void exportExcel(final List<T> list, final String sheetName, int sheetSize, final OutputStream output) throws IOException {

        final Field[] allFields = clazz.getDeclaredFields();
        final List<Field> fields = new ArrayList<>(allFields.length);
        for (Field field : allFields) {
            if (field.isAnnotationPresent(ExportExcelDesc.class)) {
                fields.add(field);
            }
        }

        final SXSSFWorkbook workbook = new SXSSFWorkbook();

        // excel2003中每个sheet中最多有65536行,为避免产生错误所以加这个逻辑.
        if (sheetSize > 65536 || sheetSize < 1) {
            sheetSize = 65536;
        }
        final double sheetNo = Math.ceil((double) Math.max(list.size(), 1) / sheetSize); // 取出一共有多少个sheet.
        for (int index = 0; index < sheetNo; index++) {
            final SXSSFSheet sheet = workbook.createSheet();
            workbook.setSheetName(index, sheetName + index);

            final SXSSFRow headRow = sheet.createRow(0);
            // 写入各个字段的列头名称
            for (int col = 0; col < fields.size(); col++) {
                final Field field = fields.get(col);
                final ExportExcelDesc attr = field.getAnnotation(ExportExcelDesc.class);
                final SXSSFCell cell = headRow.createCell(col);
                cell.setCellType(CellType.STRING);
                cell.setCellValue(attr.name());
            }

            final int startNo = index * sheetSize;
            final int endNo = Math.min(startNo + sheetSize, list.size());
            // 写入各条记录,每条记录对应excel表中的一行
            for (int i = startNo; i < endNo; i++) {
                final SXSSFRow row = sheet.createRow(i + 1 - startNo);
                final T vo = list.get(i);
                for (int j = 0; j < fields.size(); j++) {
                    final Field field = fields.get(j);
                    field.setAccessible(true);
                    try {
                        final SXSSFCell cell = row.createCell(j);
                        cell.setCellType(CellType.STRING);
                        final Object fieldVal = field.get(vo);
                        cell.setCellValue(fieldVal == null ? "" : String.valueOf(fieldVal));
                    } catch (final IllegalAccessException | IllegalArgumentException e) {
                        LOGGER.error(e.getMessage(), e);
                    }
                }
            }

            // 必须在单元格设值以后进行
            // 设置为根据内容自动调整列宽
            for (int k = 0; k < fields .size(); k++) {
                sheet.autoSizeColumn(k);
            }
        }
        output.flush();
        workbook.write(output);
        output.close();

    }

}
