package com.tx.excel.template.exception;

/** Wraps I/O failures and invalid template markup (unbalanced block markers, etc.). */
public class ExcelTemplateException extends RuntimeException {

    public ExcelTemplateException(String message) {
        super(message);
    }

    public ExcelTemplateException(String message, Throwable cause) {
        super(message, cause);
    }
}
