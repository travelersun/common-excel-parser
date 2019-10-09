package com.sitech.crm.bcc.common.domain;


import com.sitech.crm.bcc.common.excelparser.annotations.ExcelField;
import com.sitech.crm.bcc.common.excelparser.annotations.ExcelObject;
import com.sitech.crm.bcc.common.excelparser.annotations.ParseType;
import com.sitech.crm.bcc.common.excelparser.utils.ValidatorUtil;

/**
 * @Author: liaoyq
 * @Desrition: 开户产品 集团短号 Excel解析领域对象
 * @Created: 2018/9/5 10:16
 * @Modified: 2018/9/5 10:16
 * @Modified By: liaoyq
 */

@ExcelObject(parseType = ParseType.ROW, start = 2 ,zeroIfNull = false , ignoreAllZerosOrNullRows = true)
public class GroupShort {

    @ExcelField(position = 1, validationType = ExcelField.ValidationType.SOFT, validate = true ,regex = ValidatorUtil.REGEX_MOBILE)
    String memtelNo; //手机号

    @ExcelField(position = 2, validationType = ExcelField.ValidationType.SOFT, validate = true ,regex = ValidatorUtil.REGEX_SHORT_NUM)
    String shortNo; //短号

    @ExcelField(position = 3)
    String fav; //优惠

    @ExcelField(position = 4)
    String effectDate; //生效日期

    public GroupShort() {
    }

    public GroupShort(String memtelNo, String shortNo, String fav, String effectDate) {
        this.memtelNo = memtelNo;
        this.shortNo = shortNo;
        this.fav = fav;
        this.effectDate = effectDate;
    }

    public String getMemtelNo() {
        return memtelNo;
    }

    public void setMemtelNo(String memtelNo) {
        this.memtelNo = memtelNo;
    }

    public String getShortNo() {
        return shortNo;
    }

    public void setShortNo(String shortNo) {
        this.shortNo = shortNo;
    }

    public String getFav() {
        return fav;
    }

    public void setFav(String fav) {
        this.fav = fav;
    }

    public String getEffectDate() {
        return effectDate;
    }

    public void setEffectDate(String effectDate) {
        this.effectDate = effectDate;
    }
}
