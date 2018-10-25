package com.penghaohuan.excel.exception;

/**
 * Excel模板异常.
 * @author penghaohuan
 * @version 1.0.1.0
 */
public class ExcelTemplateException extends Exception {

    public ExcelTemplateException() {
    }

    public ExcelTemplateException(final String msg) {
        super(msg);
    }

    public ExcelTemplateException(final Throwable cause) {
        super(cause);
    }
}
