package com.penghaohuan.excel.util;

import com.penghaohuan.excel.handler.ExcelExporter;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.List;

/**
 * HTTP响应Excel文件工具.
 * @author penghaohuan
 */
public class ExcelWebViewUtil {

    /**
     * Sheet页大小.
     */
    private static final Integer SHEET_NUMBER = 60000;

    /**
     * web响应导出Excel.
     * @param list 实体列表
     * @param request http请求
     * @param response http响应
     * @param fileName 文件名称
     * @param sheetName sheet名称
     * @param clazz 导出实体类型
     * @param <T> 类型
     * @throws IOException e
     */
    public static <T> void exportExcel(final Class<T> clazz, final List<T> list, final HttpServletRequest request,
                                       final HttpServletResponse response, String fileName, final String sheetName) throws IOException {
        exportExcel(clazz, list, request, response, fileName, sheetName, SHEET_NUMBER);
    }

    /**
     * web响应导出Excel.
     * @param list 实体列表
     * @param request http请求
     * @param response http响应
     * @param fileName 文件名称
     * @param sheetName sheet名称
     * @param clazz 导出实体类型
     * @param sheetNumber sheet大小
     * @param <T> 类型
     * @throws IOException e
     */
    public static <T> void exportExcel(final Class<T> clazz, final List<T> list, final HttpServletRequest request,
                                       final HttpServletResponse response, String fileName, final String sheetName,
                                       final int sheetNumber) throws IOException {
        if (request.getHeader("User-Agent").toUpperCase().indexOf("MSIE") > 0) {
            fileName = URLEncoder.encode(fileName, "UTF-8");
        } else {
            fileName = new String(fileName.getBytes("UTF-8"), "ISO8859-1");
        }
        response.reset();
        response.setHeader("Content-Disposition", "attachment;fileName=\"" + fileName + "\"");
        response.setContentType("application/ms-excel");
        final ExcelExporter<T> util = new ExcelExporter<>(clazz);
        util.exportExcel(list, sheetName, sheetNumber, response.getOutputStream());
    }
}
