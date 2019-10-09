package com.sitech.crm.bcc.common.domain;


import com.sitech.crm.bcc.common.excelparser.annotations.ExcelField;
import com.sitech.crm.bcc.common.excelparser.annotations.ExcelObject;
import com.sitech.crm.bcc.common.excelparser.annotations.ParseType;
import com.sitech.crm.bcc.common.excelparser.utils.ValidatorUtil;

/**
 * @Author: liaoyq
 * @Desrition: 集团统付账户成员 Excel解析领域对象
 * @Created: 2018/9/5 10:19
 * @Modified: 2018/9/5 10:19
 * @Modified By: liaoyq
 */

@ExcelObject(parseType = ParseType.ROW, start = 2,zeroIfNull = false , ignoreAllZerosOrNullRows = true)
public class GroupAccountMember {

    @ExcelField(position = 1, validationType = ExcelField.ValidationType.SOFT, validate = true ,regex = ValidatorUtil.REGEX_MOBILE)
    String memtelNo; //手机号

    @ExcelField(position = 2)
    String payType; //付费类型

    @ExcelField(position = 3)
    String payexpenseType; //报销方式

    @ExcelField(position = 4)
    String isTransBill; //是否转移账单

    @ExcelField(position = 5)
    String effectType; //生效类型


    @ExcelField(position = 6)
    Double maxQuota; //限额

    public GroupAccountMember() {
    }

    public GroupAccountMember(String memtelNo, String payType, String payexpenseType, String effectType, String isTransBill, Double maxQuota) {
        this.memtelNo = memtelNo;
        this.payType = payType;
        this.payexpenseType = payexpenseType;
        this.effectType = effectType;
        this.isTransBill = isTransBill;
        this.maxQuota = maxQuota;
    }

    public String getMemtelNo() {
        return memtelNo;
    }

    public void setMemtelNo(String memtelNo) {
        this.memtelNo = memtelNo;
    }

    public String getPayType() {
        return payType;
    }

    public void setPayType(String payType) {
        this.payType = payType;
    }

    public String getPayexpenseType() {
        return payexpenseType;
    }

    public void setPayexpenseType(String payexpenseType) {
        this.payexpenseType = payexpenseType;
    }

    public String getEffectType() {
        return effectType;
    }

    public void setEffectType(String effectType) {
        this.effectType = effectType;
    }

    public String getIsTransBill() {
        return isTransBill;
    }

    public void setIsTransBill(String isTransBill) {
        this.isTransBill = isTransBill;
    }

    public Double getMaxQuota() {
        return maxQuota;
    }

    public void setMaxQuota(Double maxQuota) {
        this.maxQuota = maxQuota;
    }
}
