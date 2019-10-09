package com.sitech.crm.bcc.common.excelparser.exception;

public class ExcelParsingException extends RuntimeException {

    public ExcelParsingException(String message) {
        super(message);
    }

    public ExcelParsingException(String message, Exception exception) {
        super(message, exception);
    }

}
