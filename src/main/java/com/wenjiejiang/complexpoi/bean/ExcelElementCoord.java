package com.wenjiejiang.complexpoi.bean;

/**
 * @author wenjiejiang
 * @date 2020/10/15 10:58
 * @description excel图表中一个元素(表格, 图)的位置信息
 * @since 1.0
 */
public class ExcelElementCoord {
    /**
     * 左上角点的x坐标
     */
    private Integer startX;
    /**
     * 左上角点的y坐标
     */
    private Integer startY;
    /**
     * 右下角点的x坐标
     */
    private Integer endX;
    /**
     * 右下角点的y坐标
     */
    private Integer endY;

    public ExcelElementCoord() {
    }

    public ExcelElementCoord(Integer startX, Integer startY, Integer endX, Integer endY) {
        this.startX = startX;
        this.startY = startY;
        this.endX = endX;
        this.endY = endY;
    }

    public Integer getStartX() {
        return startX;
    }

    public void setStartX(Integer startX) {
        this.startX = startX;
    }

    public Integer getStartY() {
        return startY;
    }

    public void setStartY(Integer startY) {
        this.startY = startY;
    }

    public Integer getEndX() {
        return endX;
    }

    public void setEndX(Integer endX) {
        this.endX = endX;
    }

    public Integer getEndY() {
        return endY;
    }

    public void setEndY(Integer endY) {
        this.endY = endY;
    }

    public void updateCoord(Integer startX, Integer startY, Integer endX, Integer endY) {
        this.startX = startX;
        this.startY = startY;
        this.endX = endX;
        this.endY = endY;
    }
}
