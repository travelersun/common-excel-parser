package com.sitech.crm.bcc.common.domain;


import com.sitech.crm.bcc.common.excelparser.annotations.ExcelField;
import com.sitech.crm.bcc.common.excelparser.annotations.ExcelObject;
import com.sitech.crm.bcc.common.excelparser.annotations.ParseType;
import com.sitech.crm.bcc.common.excelparser.utils.ValidatorUtil;

/**
 * @Author: liaoyq
 * @Desrition: 开户产品 集团彩铃 Excel解析领域对象
 * @Created: 2018/9/5 10:17
 * @Modified: 2018/9/5 10:17
 * @Modified By: liaoyq
 */

@ExcelObject(parseType = ParseType.ROW, start = 2 ,zeroIfNull = false , ignoreAllZerosOrNullRows = true)
public class GroupColorBell {

    @ExcelField(position = 1, validationType = ExcelField.ValidationType.SOFT, validate = true ,regex = ValidatorUtil.REGEX_MOBILE)
    String memtelNo; //手机号

    @ExcelField(position = 2)
    String fav; //优惠

    @ExcelField(position = 3)
    String effectDate; //生效日期

    public GroupColorBell() {
    }

    public GroupColorBell(String memtelNo, String fav, String effectDate) {
        this.memtelNo = memtelNo;
        this.fav = fav;
        this.effectDate = effectDate;
    }

    public String getMemtelNo() {
        return memtelNo;
    }

    public void setMemtelNo(String memtelNo) {
        this.memtelNo = memtelNo;
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
