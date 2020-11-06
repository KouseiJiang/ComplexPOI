package com.wenjiejiang.complexpoi.util;


import org.apache.poi.ss.usermodel.*;

/**
 * @author wenjiejiang
 * @date 2020/10/15 11:53
 * @description excel工具类
 * @since 1.0
 */
public class ExcelUtils {

    /**
     * 设置单元格边框（包裹方式-全包、粗细-细、线的颜色-黑色）和对其方式（上下左右居中）
     * @param style 单元格风格
     */
    public static void setUpCellBorderAndAlignment(CellStyle style) {
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
    }
}
