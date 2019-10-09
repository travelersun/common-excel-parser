package com.sitech.crm.bcc.common.excelparser.exception;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Victor.Ikoro on 4/14/2016.
 */
public class ExcelInvalidCellValuesException extends ExcelParsingException {
    List<ExcelInvalidCell> invalidCells;
    public ExcelInvalidCellValuesException(String message) {
        super(message);
        invalidCells = new ArrayList<ExcelInvalidCell>();
    }

    public ExcelInvalidCellValuesException(String message, Exception exception) {
        super(message, exception);
        invalidCells = new ArrayList<ExcelInvalidCell>();
    }

    public List<ExcelInvalidCell> getInvalidCells() {
        return invalidCells;
    }

    public void setInvalidCells(List<ExcelInvalidCell> invalidCells) {
        this.invalidCells = invalidCells;
    }

    public void addInvalidCell(ExcelInvalidCell excelInvalidCell)
    {
        invalidCells.add(excelInvalidCell);
    }
}
