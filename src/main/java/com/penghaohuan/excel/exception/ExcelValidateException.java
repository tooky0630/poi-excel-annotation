package com.penghaohuan.excel.exception;

/**
 * Excel校验异常.
 * @author penghaohuan
 * @version 1.0.1.0
 */
public class ExcelValidateException extends Exception {

    public ExcelValidateException() {
    }

    public ExcelValidateException(final String msg) {
        super(msg);
    }

    public ExcelValidateException(final Throwable cause) {
        super(cause);
    }
}
