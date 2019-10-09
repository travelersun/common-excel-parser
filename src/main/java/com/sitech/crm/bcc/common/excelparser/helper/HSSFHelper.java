package com.sitech.crm.bcc.common.excelparser.helper;

import com.sitech.crm.bcc.common.excelparser.exception.ExcelParsingException;
import com.sitech.crm.bcc.common.excelparser.interfaces.ParserExceptionConsumer;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;

import java.math.BigDecimal;
import java.util.Date;
import java.util.Iterator;

import static java.text.MessageFormat.format;

public class HSSFHelper {
	
	private static DataFormatter formatter = new DataFormatter();

    @SuppressWarnings("unchecked")
    public static <T> T getCellValue(Sheet sheet, Class<T> type, Integer row, Integer col, boolean zeroIfNull, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        Cell cell = getCell(sheet, row, col);

        return validateAndParseValue(cell, sheet.getSheetName(), type, row, col, zeroIfNull, errorHandler);
    }

    @SuppressWarnings("unchecked")
    public static <T> T setCellValue(Sheet sheet, Class<T> type, Object object,Integer row, Integer col, boolean zeroIfNull, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        Cell cell = getCell(sheet, row, col);

        if(cell == null){
            createCell(sheet, row, col);
        }

        cell = getCell(sheet, row, col);

        if(null != cell && object != null){
            if (type.equals(String.class)) {
               cell.setCellValue((String)object);
            }

            if (type.equals(Date.class)) {
                cell.setCellValue((Date)object);
            }

            if (type.equals(Integer.class)) {
                cell.setCellValue((Integer)object);
            }

            if (type.equals(Double.class)) {
                cell.setCellValue((Double)object);
            }

            if (type.equals(Long.class)) {
                cell.setCellValue((Long)object);
            }

            if (type.equals(BigDecimal.class)) {
                cell.setCellValue(((BigDecimal) object).doubleValue());
            }
        }
        return validateAndParseValue(cell, sheet.getSheetName(), type, row, col, zeroIfNull, errorHandler);
    }


    public static <T> T getCellValue(Row row, String sheetName, Class<T> type, Integer rowIndex, Integer col, boolean zeroIfNull, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        Cell cell = row.getCell(col - 1);

        return validateAndParseValue(cell, sheetName, type, rowIndex, col, zeroIfNull, errorHandler);
    }

    @SuppressWarnings("unchecked")
    private static <T> T validateAndParseValue(Cell cell, String sheetName, Class<T> type, Integer row, Integer col, boolean zeroIfNull, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        if (type.equals(String.class)) {
            return (T) getStringCell(cell, errorHandler);
        }

        if (type.equals(Date.class)) {
            return cell == null ? null : (T) getDateCell(cell, new Locator(sheetName, row, col), errorHandler);
        }

        if (type.equals(Integer.class)) {
            return (T) getIntegerCell(cell, zeroIfNull, new Locator(sheetName, row, col), errorHandler);
        }

        if (type.equals(Double.class)) {
            return (T) getDoubleCell(cell, zeroIfNull, new Locator(sheetName, row, col), errorHandler);
        }

        if (type.equals(Long.class)) {
            return (T) getLongCell(cell, zeroIfNull, new Locator(sheetName, row, col), errorHandler);
        }

        if (type.equals(BigDecimal.class)) {
            return (T) getBigDecimalCell(cell, zeroIfNull, new Locator(sheetName, row, col), errorHandler);
        }

        errorHandler.accept(new ExcelParsingException(format("{0} data type not supported for parsing", type.getName())));
        return null;
    }


    private static BigDecimal getBigDecimalCell(Cell cell, boolean zeroIfNull, Locator locator, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        String val = getStringCell(cell, errorHandler);
        if(val == null || val.trim().equals("")) {
            if(zeroIfNull) {
                return BigDecimal.ZERO;
            }
            return null;
        }
        try {
            return new BigDecimal(val);
        } catch (NumberFormatException e) {
            errorHandler.accept(new ExcelParsingException(format("Invalid number found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol())));
        }

        if (zeroIfNull) {
            return BigDecimal.ZERO;
        }
        return null;
    }

    static Cell getCell(Sheet sheet, int rowNumber, int columnNumber) {
        Row row = sheet.getRow(rowNumber - 1);
        return row == null ? null : row.getCell(columnNumber - 1);
    }

    static Cell createCell(Sheet sheet, int rowNumber, int columnNumber) {

        Row row = sheet.getRow(rowNumber - 1);
        if(row == null){
            sheet.createRow(rowNumber - 1);
        }

        row = sheet.getRow(rowNumber - 1);

        Cell cell = row.getCell(columnNumber - 1);

        if(cell == null ){
            row.createCell(columnNumber - 1);
        }



        return row == null ? null : row.getCell(columnNumber - 1);
    }

    public static Row getRow(Iterator<Row> iterator, int rowNumber) {
        Row row;
        while (iterator.hasNext()) {
            row = iterator.next();
            if (row.getRowNum() == rowNumber - 1) {
                return row;
            }
        }
        throw new RuntimeException("No Row with index: " + rowNumber + " was found");
    }

    static String getStringCell(Cell cell, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        if (cell == null) {
            return null;
        }

        if (cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
            int type = cell.getCachedFormulaResultType();

            if (type == HSSFCell.CELL_TYPE_NUMERIC) {
            	FormulaEvaluator fe = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            	return formatter.formatCellValue(cell, fe);
            }

            if (type == HSSFCell.CELL_TYPE_ERROR) {
                return "";
            }

            if (type == HSSFCell.CELL_TYPE_STRING) {
                return cell.getRichStringCellValue().getString().trim();
            }

            if (type == HSSFCell.CELL_TYPE_BOOLEAN) {
                return "" + cell.getBooleanCellValue();
            }

        } else if (cell.getCellType() != HSSFCell.CELL_TYPE_NUMERIC) {
            return cell.getRichStringCellValue().getString().trim();
        }

        return formatter.formatCellValue(cell);
    }

    static Date getDateCell(Cell cell, Locator locator, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        try {
            if (!HSSFDateUtil.isCellDateFormatted(cell)) {
                errorHandler.accept(new ExcelParsingException(format("Invalid date found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol())));
            }
            return HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
        } catch (IllegalStateException illegalStateException) {
            errorHandler.accept(new ExcelParsingException(format("Invalid date found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol())));
        }
        return null;
    }

    static Double getDoubleCell(Cell cell, boolean zeroIfNull, Locator locator, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        if (cell == null) {
            return zeroIfNull ? 0d : null;
        }

        if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC || cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
            return cell.getNumericCellValue();
        }

        if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
            return zeroIfNull ? 0d : null;
        }

        errorHandler.accept(new ExcelParsingException(format("Invalid number found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol())));
        return null;
    }

    static Long getLongCell(Cell cell, boolean zeroIfNull, Locator locator, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        Double doubleValue = getNumberWithoutDecimals(cell, zeroIfNull, locator, errorHandler);
        return doubleValue == null ? null : doubleValue.longValue();
    }

    static Integer getIntegerCell(Cell cell, boolean zeroIfNull, Locator locator, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        Double doubleValue = getNumberWithoutDecimals(cell, zeroIfNull, locator, errorHandler);
        return doubleValue == null ? null : doubleValue.intValue();
    }

    private static Double getNumberWithoutDecimals(Cell cell, boolean zeroIfNull, Locator locator, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        Double doubleValue = getDoubleCell(cell, zeroIfNull, locator, errorHandler);
        if (doubleValue != null && doubleValue % 1 != 0) {
            errorHandler.accept(new ExcelParsingException(format("Invalid number found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol())));
        }
        return doubleValue;
    }

}
