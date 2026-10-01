package com.tx.excel.template;

public class ExcelTemplateException extends RuntimeException {

    public ExcelTemplateException(String message) {
        super(message);
    }

    public ExcelTemplateException(String message, Throwable cause) {
        super(message, cause);
    }
}
