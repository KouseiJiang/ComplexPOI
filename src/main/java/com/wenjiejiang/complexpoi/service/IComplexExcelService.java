package com.wenjiejiang.complexpoi.service;

import com.wenjiejiang.complexpoi.bean.ExcelSheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * @author wenjiejiang
 * @date 2020/10/15 11:53
 * @description excel表格的列描述
 * @since 1.0
 */
public interface IComplexExcelService {

    Workbook createExcel(List<ExcelSheet> excelSheets);
}
