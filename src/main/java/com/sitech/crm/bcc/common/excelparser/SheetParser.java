package com.sitech.crm.bcc.common.excelparser;

import com.sitech.crm.bcc.common.excelparser.annotations.ExcelField;
import com.sitech.crm.bcc.common.excelparser.annotations.ExcelObject;
import com.sitech.crm.bcc.common.excelparser.annotations.MappedExcelObject;
import com.sitech.crm.bcc.common.excelparser.annotations.ParseType;
import com.sitech.crm.bcc.common.excelparser.exception.ExcelInvalidCell;
import com.sitech.crm.bcc.common.excelparser.exception.ExcelInvalidCellValuesException;
import com.sitech.crm.bcc.common.excelparser.exception.ExcelParsingException;
import com.sitech.crm.bcc.common.excelparser.helper.HSSFHelper;
import com.sitech.crm.bcc.common.excelparser.interfaces.ParserExceptionConsumer;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static com.sitech.crm.bcc.common.excelparser.helper.HSSFHelper.getRow;


public class SheetParser {
    List<ExcelInvalidCell> excelInvalidCells;

    public SheetParser() {
        excelInvalidCells = new ArrayList<ExcelInvalidCell>();
    }

    public <T> List<T> createEntity(Sheet sheet, Class<T> clazz, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        List<T> list = new ArrayList<T>();
        ExcelObject excelObject = getExcelObject(clazz, errorHandler);
        if (excelObject.start() <= 0 || excelObject.end() < 0) {
            return list;
        }
        int end = getEnd(sheet, clazz, excelObject);

        for (int currentLocation = excelObject.start(); currentLocation <= end; currentLocation++) {
            T object = getNewInstance(sheet, clazz, excelObject.parseType(), currentLocation, excelObject.zeroIfNull(),
                errorHandler);
            List<Field> mappedExcelFields = getMappedExcelObjects(clazz);
            for (Field mappedField : mappedExcelFields) {
                Class<?> fieldType = mappedField.getType();
                Class<?> clazz1 = fieldType.equals(List.class) ? getFieldType(mappedField) : fieldType;
                List<?> fieldValue = createEntity(sheet, clazz1, errorHandler);
                if (fieldType.equals(List.class)) {
                    setFieldValue(mappedField, object, fieldValue);
                } else if (!fieldValue.isEmpty()) {
                    setFieldValue(mappedField, object, fieldValue.get(0));
                }
            }
            if(excelObject.ignoreAllZerosOrNullRows()){
                boolean isAllFieldNull = false;
                try {
                    isAllFieldNull = isAllFieldNull(object);
                } catch (Exception e) {
                    e.printStackTrace();
                }
                if(!isAllFieldNull){
                    list.add(object);
                }
            }else {
                list.add(object);
            }
        }
        return list;
    }

    public <T,V> List<T> createSheet(Sheet sheet, Class<T> clazz,List<T> data, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        List<T> list = new ArrayList<T>();
        ExcelObject excelObject = getExcelObject(clazz, errorHandler);
        if (excelObject == null || excelObject.start() <= 0 || excelObject.end() < 0 || data == null || data.size() == 0) {
            return list;
        }
        int end = data.size();
        if(excelObject.end() > 0){
            if(excelObject.end() - excelObject.start() < 0 ){
                throw new ExcelParsingException("end must >= start when end > 0 !");
            }
             end =  end > (excelObject.end() - excelObject.start()  +1)?(excelObject.end() - excelObject.start()  +1):end;
        }

        for (int currentLocation = 0; currentLocation < end; currentLocation++) {
            T object = data.get(currentLocation);
            //处理顶层属性字段
            writeNewInstance(sheet, clazz,object, excelObject.parseType(), excelObject.start()+currentLocation, excelObject.zeroIfNull(),errorHandler);

            //处理嵌套属性
            List<Field> mappedExcelFields = getMappedExcelObjects(clazz);
            for (Field mappedField : mappedExcelFields) {
                Class<?> fieldType = mappedField.getType();
                Class<?> clazz1 = fieldType.equals(List.class) ? getFieldType(mappedField) : fieldType;
                if (fieldType.equals(List.class)) {
                    List<Object> fieldValue = null;
                    mappedField.setAccessible(true);
                    try {
                        fieldValue = (List<Object>) mappedField.get(object);
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                    createSheet(sheet,(Class<Object>) clazz1,fieldValue,errorHandler);
                }
            }
            list.add(object);
        }
        return list;
    }

    public <T> List<T> createEntityWithIterator(Sheet sheet, Class<T> clazz, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        List<T> list = new ArrayList<T>();

        ExcelObject excelObject = getExcelObject(clazz, errorHandler);
        if (excelObject.start() <= 0 || excelObject.end() < 0) {
            return list;
        }
        int end = getEnd(sheet, clazz, excelObject);
        for (int currentLocation = excelObject.start(); currentLocation <= end; currentLocation++) {
            T object = getNewInstance(sheet.iterator(), sheet.getSheetName(), clazz, excelObject.parseType(), currentLocation, excelObject.zeroIfNull(),
                errorHandler);
            List<Field> mappedExcelFields = getMappedExcelObjects(clazz);
            for (Field mappedField : mappedExcelFields) {
                Class<?> fieldType = mappedField.getType();
                Class<?> clazz1 = fieldType.equals(List.class) ? getFieldType(mappedField) : fieldType;
                List<?> fieldValue = createEntityWithIterator(sheet, clazz1, errorHandler);
                if (fieldType.equals(List.class)) {
                    setFieldValue(mappedField, object, fieldValue);
                } else if (!fieldValue.isEmpty()) {
                    setFieldValue(mappedField, object, fieldValue.get(0));
                }
            }
            list.add(object);
        }

        return list;
    }

    private <T> int getEnd(Sheet sheet, Class<T> clazz, ExcelObject excelObject) {
        int end = excelObject.end();
        if (end > 0) {
            return end;
        }
        return getRowOrColumnEnd(sheet, clazz);
    }

    /**
     * @deprecated Pass an error handler lambda instead (see other signature)
     */
    @Deprecated
    public <T> List<T> createEntity(Sheet sheet, String sheetName, Class<T> clazz) {
        return createEntity(sheet, clazz, new ParserExceptionConsumer<ExcelParsingException>() {
            //@Override
            public void accept(ExcelParsingException e) {
                throw e;
            }
        });
    }

    public <T> int getRowOrColumnEnd(Sheet sheet, Class<T> clazz) {
        ExcelObject excelObject = getExcelObject(clazz, new ParserExceptionConsumer<ExcelParsingException>() {
            //@Override
            public void accept(ExcelParsingException e) {
                throw e;
            }
        });
        ParseType parseType = excelObject.parseType();
        if (parseType == ParseType.ROW) {
            return sheet.getLastRowNum() + 1;
        }

        Set<Integer> positions = getExcelFieldPositionMap(clazz).keySet();

        int max = Collections.max(positions);
        int min = Collections.min(positions);

        int maxCellNumber = 0;
        for (int i = min; i < max; i++) {
            int cellsNumber = sheet.getRow(i).getLastCellNum();
            if (maxCellNumber < cellsNumber) {
                maxCellNumber = cellsNumber;
            }
        }
        return maxCellNumber;
    }

    private Class<?> getFieldType(Field field) {
        Type type = field.getGenericType();
        if (type instanceof ParameterizedType) {
            ParameterizedType pt = (ParameterizedType) type;
            return (Class<?>) pt.getActualTypeArguments()[0];
        }

        return null;
    }

    private <T> List<Field> getMappedExcelObjects(Class<T> clazz) {
        List<Field> fieldList = new ArrayList<Field>();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            MappedExcelObject mappedExcelObject = field.getAnnotation(MappedExcelObject.class);
            if (mappedExcelObject != null) {
                field.setAccessible(true);
                fieldList.add(field);
            }
        }
        return fieldList;
    }

    private <T> ExcelObject getExcelObject(Class<T> clazz, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        ExcelObject excelObject = clazz.getAnnotation(ExcelObject.class);
        if (excelObject == null) {
            errorHandler.accept(new ExcelParsingException("Invalid class configuration - ExcelObject annotation missing - " + clazz.getSimpleName()));
        }
        return excelObject;
    }

    private <T> T getNewInstance(Sheet sheet, Class<T> clazz, ParseType parseType, Integer currentLocation, boolean zeroIfNull, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        T object = getInstance(clazz, errorHandler);
        Map<Integer, Field> excelPositionMap = getExcelFieldPositionMap(clazz);
        for (Integer position : excelPositionMap.keySet()) {
            Field field = excelPositionMap.get(position);
            Object cellValue;
            Object cellValueString;
            if (ParseType.ROW == parseType) {
                cellValue = HSSFHelper.getCellValue(sheet, field.getType(), currentLocation, position, zeroIfNull, errorHandler);
                cellValueString = HSSFHelper.getCellValue(sheet, String.class, currentLocation, position, zeroIfNull, errorHandler);
            } else {
                cellValue = HSSFHelper.getCellValue(sheet, field.getType(), position, currentLocation, zeroIfNull, errorHandler);
                cellValueString = HSSFHelper.getCellValue(sheet, String.class, position, currentLocation, zeroIfNull, errorHandler);
            }
            validateAnnotation(field, cellValueString, position, currentLocation,errorHandler);
            setFieldValue(field, object, cellValue);
        }

        return object;
    }

    private <T> T writeNewInstance(Sheet sheet, Class<T> clazz,T object, ParseType parseType, Integer currentLocation, boolean zeroIfNull, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        Map<Integer, Field> excelPositionMap = getExcelFieldPositionMap(clazz);
        for (Integer position : excelPositionMap.keySet()) {
            Field field = excelPositionMap.get(position);
            Object cellValue = null;
            try {
                cellValue = field.get(object);
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }

            if (ParseType.ROW == parseType) {
                HSSFHelper.setCellValue(sheet, field.getType(),cellValue ,currentLocation, position, zeroIfNull, errorHandler);
            } else {
                HSSFHelper.setCellValue(sheet, field.getType(),cellValue, position, currentLocation, zeroIfNull, errorHandler);
            }
            validateAnnotation(field, cellValue, position, currentLocation,errorHandler);
        }

        return object;
    }

    private <T> T getNewInstance(Iterator<Row> rowIterator, String sheetName, Class<T> clazz, ParseType parseType, Integer currentLocation, boolean zeroIfNull, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        T object = getInstance(clazz, errorHandler);
        Map<Integer, Field> excelPositionMap = getSortedExcelFieldPositionMap(clazz);
        Row row = null;
        for (Integer position : excelPositionMap.keySet()) {
            Field field = excelPositionMap.get(position);
            Object cellValue;
            Object cellValueString;

            if (ParseType.ROW == parseType) {
                if (null == row || row.getRowNum() + 1 != currentLocation) {
                    row = getRow(rowIterator, currentLocation);
                }
                cellValue = HSSFHelper.getCellValue(row, sheetName, field.getType(), currentLocation, position, zeroIfNull, errorHandler);
                cellValueString = HSSFHelper.getCellValue(row, sheetName, String.class, currentLocation, position, zeroIfNull, errorHandler);
            } else {
                if (null == row || row.getRowNum() + 1 != position) {
                    row = getRow(rowIterator, position);
                }
                cellValue = HSSFHelper.getCellValue(row, sheetName, field.getType(), position, currentLocation, zeroIfNull, errorHandler);
                cellValueString = HSSFHelper.getCellValue(row, sheetName, String.class, position, currentLocation, zeroIfNull, errorHandler);
            }
            validateAnnotation(field, cellValueString, position, currentLocation,errorHandler);
            setFieldValue(field, object, cellValue);
        }

        return object;
    }

    private void validateAnnotation(Field field, Object cellValueString, int position, int currentLocation,ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        ExcelField annotation = field.getAnnotation(ExcelField.class);
        if (annotation.validate()) {
            Pattern pattern = Pattern.compile(annotation.regex());
            cellValueString = cellValueString != null ? cellValueString.toString() : "";
            Matcher matcher = pattern.matcher((String) cellValueString);
            if (!matcher.matches()) {
                ExcelInvalidCell excelInvalidCell = new ExcelInvalidCell(position, currentLocation, (String) cellValueString,"value not matche regex: "+annotation.regex());
                excelInvalidCells.add(excelInvalidCell);
                errorHandler.accept(new ExcelInvalidCellValuesException("Invalid cell value at [" + currentLocation + ", " + position + "] in the sheet. value not matche regex: "+annotation.regex()));
                if (annotation.validationType() == ExcelField.ValidationType.HARD) {
                    throw new ExcelInvalidCellValuesException("Invalid cell value at [" + currentLocation + ", " + position + "] in the sheet. This exception can be suppressed by setting 'validationType' in @ExcelField to 'ValidationType.SOFT");
                }
            }
        }
    }

    private <T> T getInstance(Class<T> clazz, ParserExceptionConsumer<ExcelParsingException> errorHandler) {
        T object;
        try {
            Constructor<T> constructor = clazz.getDeclaredConstructor();
            constructor.setAccessible(true);
            object = constructor.newInstance();
        } catch (Exception e) {
            errorHandler.accept(new ExcelParsingException("Exception occurred while instantiating the class " + clazz.getName(), e));
            return null;
        }
        return object;
    }

    private <T> void setFieldValue(Field field, T object, Object cellValue) {
        try {
            field.set(object, cellValue);
        } catch (IllegalArgumentException e) {
            throw new ExcelParsingException("Exception occurred while setting field value ", e);
        }catch (IllegalAccessException e) {
            throw new ExcelParsingException("Exception occurred while setting field value ", e);
        }
    }

    private <T> Map<Integer, Field> getExcelFieldPositionMap(Class<T> clazz) {
        Map<Integer, Field> fieldMap = new HashMap<Integer, Field>();
        return fillMap(clazz, fieldMap);
    }

    private <T> Map<Integer, Field> getSortedExcelFieldPositionMap(Class<T> clazz) {
        Map<Integer, Field> fieldMap = new TreeMap<Integer, Field>();
        return fillMap(clazz, fieldMap);
    }

    private <T> Map<Integer, Field> fillMap(Class<T> clazz, Map<Integer, Field> fieldMap) {
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (excelField != null) {
                field.setAccessible(true);
                fieldMap.put(excelField.position(), field);
            }
        }
        return fieldMap;
    }

    //判断该对象是否: 返回ture表示所有属性为null  返回false表示不是所有属性都是null
    public static boolean isAllFieldNull(Object obj) throws Exception {
        Class stuCla = obj.getClass();
        Field[] fs = stuCla.getDeclaredFields();
        boolean flag = true;
        Field[] arr$ = fs;
        int len$ = fs.length;

        for (int i$ = 0; i$ < len$; ++i$) {
            Field f = arr$[i$];
            f.setAccessible(true);
            Object val = f.get(obj);
            if (val != null && val instanceof String) {
                if (isNotBlank((String) val)) {
                    flag = false;
                    break;
                }
            } else if (val != null) {
                flag = false;
                break;
            }

        }

        return flag;
    }

    public static boolean isBlank(String str) {
        int strLen;
        if (str != null && (strLen = str.length()) != 0) {
            for(int i = 0; i < strLen; ++i) {
                if (!Character.isWhitespace(str.charAt(i))) {
                    return false;
                }
            }

            return true;
        } else {
            return true;
        }
    }

    public static boolean isNotBlank(String str) {
        return !isBlank(str);
    }

}
