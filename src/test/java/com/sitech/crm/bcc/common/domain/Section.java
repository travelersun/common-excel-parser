package com.sitech.crm.bcc.common.domain;





import com.sitech.crm.bcc.common.excelparser.annotations.ExcelField;
import com.sitech.crm.bcc.common.excelparser.annotations.ExcelObject;
import com.sitech.crm.bcc.common.excelparser.annotations.MappedExcelObject;
import com.sitech.crm.bcc.common.excelparser.annotations.ParseType;

import java.util.List;


@ExcelObject(parseType = ParseType.COLUMN, start = 2, end = 2)
public class Section {

    @ExcelField(position = 2)
    String year;

    @ExcelField(position = 3)
    String section;

    @MappedExcelObject
    List<Student> students;

    @SuppressWarnings("UnusedDeclaration")
    private Section() {
        this(null, null, null);
    }

    public Section(String year, String section, List<Student> students) {
        this.year = year;
        this.section = section;
        this.students = students;
    }

    public String getYear() {
        return year;
    }

    public void setYear(String year) {
        this.year = year;
    }

    public String getSection() {
        return section;
    }

    public void setSection(String section) {
        this.section = section;
    }

    public List<Student> getStudents() {
        return students;
    }

    public void setStudents(List<Student> students) {
        this.students = students;
    }
}
