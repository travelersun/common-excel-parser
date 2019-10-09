package com.sitech.crm.bcc.common.domain;




import com.sitech.crm.bcc.common.excelparser.annotations.ExcelField;
import com.sitech.crm.bcc.common.excelparser.annotations.ExcelObject;
import com.sitech.crm.bcc.common.excelparser.annotations.ParseType;

import java.math.BigDecimal;
import java.util.Date;


@ExcelObject(parseType = ParseType.ROW, start = 6, end = 8)
public class Student {

    @ExcelField(position = 2, validationType = ExcelField.ValidationType.HARD, regex = "[0-2][0-9][0-9][0-9]")
    Long roleNumber;

    @ExcelField(position = 3)
    String name;

    @ExcelField(position = 4)
    Date dateOfBirth;

    @ExcelField(position = 5)
    String fatherName;

    @ExcelField(position = 6)
    String motherName;

    @ExcelField(position = 7)
    String address;

    @ExcelField(position = 8)
    BigDecimal totalScore;

    @ExcelField(position = 9)
    Date admissionDate;

    @ExcelField(position = 10)
    Date admissionDateTime;

    @SuppressWarnings("UnusedDeclaration")
    private Student() {
        this(null, null, null, null, null, null, null, null, null);
    }

    public Student(Long roleNumber, String name, Date dateOfBirth, String fatherName, String motherName, String address, BigDecimal totalScore, Date admissionDate, Date admissionDateTime) {
        this.roleNumber = roleNumber;
        this.name = name;
        this.dateOfBirth = dateOfBirth;
        this.fatherName = fatherName;
        this.motherName = motherName;
        this.address = address;
        this.totalScore = totalScore;
        this.admissionDate = admissionDate;
        this.admissionDateTime = admissionDateTime;
    }

    public Long getRoleNumber() {
        return roleNumber;
    }

    public void setRoleNumber(Long roleNumber) {
        this.roleNumber = roleNumber;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Date getDateOfBirth() {
        return dateOfBirth;
    }

    public void setDateOfBirth(Date dateOfBirth) {
        this.dateOfBirth = dateOfBirth;
    }

    public String getFatherName() {
        return fatherName;
    }

    public void setFatherName(String fatherName) {
        this.fatherName = fatherName;
    }

    public String getMotherName() {
        return motherName;
    }

    public void setMotherName(String motherName) {
        this.motherName = motherName;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public BigDecimal getTotalScore() {
        return totalScore;
    }

    public void setTotalScore(BigDecimal totalScore) {
        this.totalScore = totalScore;
    }

    public Date getAdmissionDate() {
        return admissionDate;
    }

    public void setAdmissionDate(Date admissionDate) {
        this.admissionDate = admissionDate;
    }

    public Date getAdmissionDateTime() {
        return admissionDateTime;
    }

    public void setAdmissionDateTime(Date admissionDateTime) {
        this.admissionDateTime = admissionDateTime;
    }
}
