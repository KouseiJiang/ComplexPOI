package com.wenjiejiang.complexpoi.bean;

import java.lang.reflect.Field;
import java.util.List;

/**
 * @author wenjiejiang
 * @date 2020/10/14 16:57
 * @description excel中的一张表
 * 下面是一个例子
 * | 业务数据概览 |  |  |
 * | :---: | --- | --- |
 * | 日期 | 进件订单数 | 订单通过数 |
 * | 2020/01/02 | 100 | 99 |
 * | 2020/01/03 | 101 | 101 |
 * title:业务数据概览
 * List<ExcelColumnMap>:日期，进件订单数，订单通过数
 * @since 1.0
 */
public class ExcelTable<T> {
    /**
     * 表的标题
     */
    private String title;
    /**
     * 表的列标题
     */
    private List<ExcelColumnMap> colNameMap;
    /**
     * 表中的数据
     */
    private List<T> tableDataList;
    /**
     * 表格的数据图
     */
    private List<TableChart> tableChartList;

    private ExcelElementCoord location;

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public List<ExcelColumnMap> getColNameMap() {
        return colNameMap;
    }

    public void setColNameMap(List<ExcelColumnMap> colNameMap) {
        this.colNameMap = colNameMap;
    }

    public List<T> getTableDataList() {
        return tableDataList;
    }

    public void setTableDataList(List<T> tableDataList) {
        this.tableDataList = tableDataList;
    }

    public List<TableChart> getTableChartList() {
        return tableChartList;
    }

    public void setTableChartList(List<TableChart> tableChartList) {
        this.tableChartList = tableChartList;
    }

    public ExcelElementCoord getLocation() {
        return location;
    }

    public void setLocation(ExcelElementCoord location) {
        this.location = location;
    }

    public ExcelColumnMap getColMap(Field field) {
        for (ExcelColumnMap excelColumnMap : colNameMap) {
            excelColumnMap.getValue().equals(field);
            return excelColumnMap;
        }
        return null;
    }
}
