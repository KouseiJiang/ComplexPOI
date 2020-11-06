package com.wenjiejiang.complexpoi.bean;

import com.wenjiejiang.complexpoi.bean.constant.EnumChartType;

import java.lang.reflect.Field;
import java.util.List;

/**
 * @author wenjiejiang
 * @date 2020/10/14 16:59
 * @description Excel表里的一张统计图
 * @since 1.0
 */
public class TableChart {
    /**
     * 表的标题
     */
    private String title;
    /**
     * 表的长度
     */
    private Integer length;
    /**
     * 表的宽度
     */
    private Integer width;

    /**
     * x线的数据区域
     */
    private Field chartXSeries;
    /**
     * Y线的数据区域
     */
    private List<Field> chartYSeries;

    /**
     * 统计图的类型
     */
    private EnumChartType type;

    final int defaultLength = 6;

    final int defaultWidth = 11;

    public TableChart() {
    }

    public TableChart(String title, EnumChartType type) {
        this(title, null, null, type);
    }

    public TableChart(String title, Integer length, Integer width, EnumChartType type) {
        this.title = title;
        this.length = length;
        this.width = width;
        this.type = type;
        if (null == length) {
            this.length = defaultLength;
        }
        if (null == width) {
            this.width = defaultWidth;
        }
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public Integer getLength() {
        return length;
    }

    public void setLength(Integer length) {
        this.length = length;
    }

    public Integer getWidth() {
        return width;
    }

    public void setWidth(Integer width) {
        this.width = width;
    }

    public List<Field> getChartYSeries() {
        return chartYSeries;
    }

    public void setChartYSeries(List<Field> chartYSeries) {
        this.chartYSeries = chartYSeries;
    }

    public EnumChartType getType() {
        return type;
    }

    public void setType(EnumChartType type) {
        this.type = type;
    }

    public Field getChartXSeries() {
        return chartXSeries;
    }

    public void setChartXSeries(Field chartXSeries) {
        this.chartXSeries = chartXSeries;
    }
}
