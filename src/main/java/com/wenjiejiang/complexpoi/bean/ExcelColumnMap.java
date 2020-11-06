package com.wenjiejiang.complexpoi.bean;

import java.lang.reflect.Field;

/**
 * @author wenjiejiang
 * @date 2020/10/15 11:53
 * @description excel表格的列描述
 * @since 1.0
 */
public class ExcelColumnMap {
    private String cloName;
    private Field value;

    public ExcelColumnMap() {
    }

    public ExcelColumnMap(String cloName, Field value) {
        this.cloName = cloName;
        this.value = value;
    }

    public String getCloName() {
        return cloName;
    }

    public void setCloName(String cloName) {
        this.cloName = cloName;
    }

    public Field getValue() {
        return value;
    }

    public void setValue(Field value) {
        this.value = value;
    }
}
