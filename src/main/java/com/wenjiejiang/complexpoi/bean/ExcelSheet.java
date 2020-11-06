package com.wenjiejiang.complexpoi.bean;

import java.util.List;

/**
 * @author wenjiejiang
 * @date 2020/10/14 16:55
 * @description 代表Excel中的一张工作簿数据
 * @since 1.0
 */
public class ExcelSheet {
    private String name;
    private List<ExcelTable> excelTableList;

    public ExcelSheet(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<ExcelTable> getExcelTableList() {
        return excelTableList;
    }

    public void setExcelTableList(List<ExcelTable> excelTableList) {
        this.excelTableList = excelTableList;
    }
}
