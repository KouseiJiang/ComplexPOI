package com.wenjiejiang.complexpoi.bean.constant;

/**
 * @author wenjiejiang
 * @date 2020/10/14 18:00
 * @description Excel图的类型
 * @since 1.0
 */
public enum EnumChartType {
    LINE(1,"折线图"),
    BAR(2,"柱状图"),
    RADAR(3,"雷达图"),
    PIE(4,"饼图");

    private Integer code;
    private String label;

    public EnumChartType getLabelByCode(Integer code){
        for (EnumChartType chartType: EnumChartType.values()) {
            if (code.equals(chartType.getCode())){
                return chartType;
            }
        }
        return null;
    }

    EnumChartType(Integer code, String label) {
        this.code = code;
        this.label = label;
    }

    public Integer getCode() {
        return code;
    }

    public void setCode(Integer code) {
        this.code = code;
    }

    public String getLabel() {
        return label;
    }

    public void setLabel(String label) {
        this.label = label;
    }
}
