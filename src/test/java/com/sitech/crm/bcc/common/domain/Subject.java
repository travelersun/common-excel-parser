package com.sitech.crm.bcc.common.domain;


import com.sitech.crm.bcc.common.excelparser.annotations.ExcelField;
import com.sitech.crm.bcc.common.excelparser.annotations.ExcelObject;
import com.sitech.crm.bcc.common.excelparser.annotations.ParseType;

@ExcelObject(parseType = ParseType.ROW, start = 2)
public class Subject {

    @ExcelField(position = 1)
    String code;

    @ExcelField(position = 2)
    String name;

    @ExcelField(position = 3)
    Integer volume;

    @SuppressWarnings("UnusedDeclaration")
    private Subject() {
	this(null, null, null);
    }

    public Subject(String code, String name, Integer volume) {
        this.code = code;
        this.name = name;
        this.volume = volume;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getVolume() {
        return volume;
    }

    public void setVolume(Integer volume) {
        this.volume = volume;
    }
}
