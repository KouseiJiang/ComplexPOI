package com.wenjiejiang.complexpoi.service.impl;

import com.google.common.base.Preconditions;
import com.google.common.base.Strings;
import com.wenjiejiang.complexpoi.bean.*;
import com.wenjiejiang.complexpoi.bean.constant.EnumChartType;
import com.wenjiejiang.complexpoi.service.IComplexExcelService;
import com.wenjiejiang.complexpoi.util.ExcelUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;

import java.lang.reflect.Field;
import java.util.*;

/**
 * @author wenjiejiang
 * @date 2020/10/15 11:53
 * @description excel服务
 * @since 1.0
 */
public class ComplexExcelServiceImpl implements IComplexExcelService {

    @Override
    public Workbook createExcel(List<ExcelSheet> excelSheets) {
        Preconditions.checkArgument(null != excelSheets && !excelSheets.isEmpty(), "excelSheet不能为null");
        Workbook wb = new XSSFWorkbook();
        for (ExcelSheet excelSheet : excelSheets) {
            createSheet(excelSheet, wb);
        }
        return wb;
    }

    private Sheet createSheet(ExcelSheet excelSheet, Workbook wb) {
        Preconditions.checkArgument(!Strings.isNullOrEmpty(excelSheet.getName()), "sheet的name不能为null");
        XSSFSheet sheet = (XSSFSheet) wb.createSheet(excelSheet.getName());
        //设置默认的行高
        sheet.setDefaultRowHeight((short) (27 * 20));
        List<ExcelTable> excelTables = excelSheet.getExcelTableList();
        Preconditions.checkArgument(null != excelTables && !excelTables.isEmpty(), "不允许建立一个空的sheet");
        //创建2个点描述一个Table所占位置，为chart位置的计算和下一个table的位置的计算提供数据
        ExcelElementCoord prevCoord = new ExcelElementCoord(0, 0, 0, 0);
        //用来统计每一列的宽度，key为列索引，value为该列最长字符串的宽度
        Map<Integer, Integer> colSizeMap = new HashMap<>();
        for (int i = 0; i < excelTables.size(); i++) {
            Map<Integer, Integer> colSizeMapItem = createSheetTable(sheet, excelTables.get(i), prevCoord, wb);
            colSizeMap.putAll(colSizeMapItem);
        }
        setColWidth(sheet, colSizeMap);
        return sheet;
    }

    // 设置每一列的宽度
    private void setColWidth(XSSFSheet sheet, Map<Integer, Integer> colSizeMap) {
        for (int i = 0; i <= colSizeMap.keySet().stream().max(Comparator.comparing(indexKey -> indexKey)).get(); i++) {
            Integer colSize = colSizeMap.get(i);
            if (null != colSize) {
                sheet.setColumnWidth(i, (colSize) * 256);
            }
        }
    }

    /**
     * 创建一个excel中的一个sheet中的一张表格
     *
     * @param sheet     工作簿
     * @param table     表中的数据
     * @param prevCoord 已知table的位置信息
     * @return key表格的列索引，value列宽度
     */
    private Map<Integer, Integer> createSheetTable(XSSFSheet sheet, ExcelTable table, ExcelElementCoord prevCoord, Workbook wb) {
        String title = table.getTitle();
        List<ExcelColumnMap> colNameMap = table.getColNameMap();
        Preconditions.checkArgument(!Strings.isNullOrEmpty(title), "表的标题不能为空");
        Preconditions.checkArgument(null != colNameMap && !colNameMap.isEmpty(), "表的列不能为空");
        List<Object> tableData = table.getTableDataList();
        //定义元素之间的间隔,除了第一张表格之外,间隔为2
        final int NUM_OF_INTERVAL = prevCoord.getEndX().equals(0) ? 0 : 2;
        //表格的标题和列名固定占两行
        final int NUM_OF_TITLE = 2;
        //获取表格里数据的行数
        final int NUM_OF_DATA = null == tableData ? 0 : tableData.size();
        //改表的行起始索引
        int rowStartIndex = 0;
        //该表的列起始索引
        int colStartIndex = prevCoord.getEndX() + NUM_OF_INTERVAL;
        //表格的行最大索引
        final int MAX_INDEX_OF_ROWS = rowStartIndex + NUM_OF_TITLE + NUM_OF_DATA;
        //表格的列最大索引
        final int MAX_INDEX_OF_COLUMNS = colStartIndex + colNameMap.size();
        //更新最大坐标信息
        prevCoord.updateCoord(colStartIndex, rowStartIndex, MAX_INDEX_OF_COLUMNS, MAX_INDEX_OF_ROWS);
        //记录一张表里列对应的数据,列名，行开始索引，行终止索引，列索引
        Map<Field, String[]> colData = new HashMap<>(colNameMap.size());
        Map<Integer, Integer> colSizeMap = fillTableData(sheet, table.getTitle(), colNameMap, tableData, rowStartIndex, colStartIndex, MAX_INDEX_OF_ROWS, MAX_INDEX_OF_COLUMNS, wb, colData);
        List<TableChart> tableCharts = table.getTableChartList();
        // 如果表格没有数据，不需要绘图
        if (null != tableCharts && !tableCharts.isEmpty() && null != tableData && !tableData.isEmpty()) {
            for (int i = 0; i < tableCharts.size(); i++) {
                createTableChart(sheet, tableCharts.get(i), prevCoord, tableData, colData);
            }
        }
        return colSizeMap;
    }

    /**
     * 创建一张图
     *
     * @param sheet      工作簿
     * @param tableChart 图的描述数据
     * @param location   表格或前一张图的位置
     * @param tableData  表格数据，只有在生成雷达图时用于计算坐标轴的单位长度
     * @param colData    列数据
     */
    private void createTableChart(XSSFSheet sheet, TableChart tableChart, ExcelElementCoord location, List<Object> tableData, Map<Field, String[]> colData) {
        //定义元素之间的间隔,除了第一张表格之外,间隔为2
        final int NUM_OF_INTERVAL = 2;
        int rowStartIndex = location.getEndY() + NUM_OF_INTERVAL;
        int colStartIndex = location.getStartX().equals(0) ? 0 : location.getStartX();
        //图的行最大索引
        final int MAX_INDEX_OF_ROWS = rowStartIndex + tableChart.getWidth();
        //图的列最大索引
        final int MAX_INDEX_OF_COLUMNS = colStartIndex + tableChart.getLength();
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, colStartIndex, rowStartIndex, MAX_INDEX_OF_COLUMNS, MAX_INDEX_OF_ROWS);
        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText(tableChart.getTitle());
        //设置图例
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);
        // 设置横纵坐标描述
        XDDFCategoryAxis bottomAxis = new XDDFCategoryAxis(null);
        XDDFValueAxis leftAxis = new XDDFValueAxis(null);
        if (!EnumChartType.PIE.equals(tableChart.getType())) {
            bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
        }
        // 设置数据源
        String[] colXInfo = colData.get(tableChart.getChartXSeries());
        XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(Integer.valueOf(colXInfo[1]), Integer.valueOf(colXInfo[2]) - 1, Integer.valueOf(colXInfo[3]), Integer.valueOf(colXInfo[3])));
        setChartInfoByType(sheet, tableChart, tableData, colData, chart, bottomAxis, leftAxis, xs);
        int endX = location.getEndX() > MAX_INDEX_OF_COLUMNS ? location.getEndX() : MAX_INDEX_OF_COLUMNS;
        location.updateCoord(colStartIndex, rowStartIndex, endX, MAX_INDEX_OF_ROWS);
    }

    /**
     * 根据图类型设置图的信息
     *
     * @param sheet      工作簿
     * @param tableChart 图的描述数据
     * @param tableData  表格数据，只有在生成雷达图时用于计算坐标轴的单位长度
     * @param colData    列数据
     * @param chart      代表Excel中的一张图
     * @param bottomAxis 横坐标
     * @param leftAxis   纵坐标
     * @param xs         横坐标的数据
     */
    private void setChartInfoByType(XSSFSheet sheet, TableChart tableChart, List<Object> tableData, Map<Field, String[]> colData, XSSFChart chart, XDDFCategoryAxis bottomAxis, XDDFValueAxis leftAxis, XDDFDataSource<String> xs) {
        switch (tableChart.getType()) {
            case LINE: {
                XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
                for (Field field : tableChart.getChartYSeries()) {
                    String[] colYInfo = colData.get(field);
                    XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(Integer.valueOf(colYInfo[1]), Integer.valueOf(colYInfo[2]) - 1, Integer.valueOf(colYInfo[3]), Integer.valueOf(colYInfo[3])));
                    XDDFLineChartData.Series series1 = (XDDFLineChartData.Series) data.addSeries(xs, ys);
                    series1.setSmooth(false);
                    series1.setMarkerStyle(MarkerStyle.NONE);
                    series1.setTitle(colYInfo[0], null);
                }
                chart.plot(data);
                break;
            }
            case BAR: {
                XDDFBarChartData data = (XDDFBarChartData) chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
                data.setBarDirection(BarDirection.COL);
                for (Field field : tableChart.getChartYSeries()) {
                    String[] colYInfo = colData.get(field);
                    XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(Integer.valueOf(colYInfo[1]), Integer.valueOf(colYInfo[2]) - 1, Integer.valueOf(colYInfo[3]), Integer.valueOf(colYInfo[3])));
                    XDDFBarChartData.Series series1 = (XDDFBarChartData.Series) data.addSeries(xs, ys);
                    series1.setTitle(colYInfo[0], null);
                }
                chart.plot(data);
                break;
            }
            case RADAR: {
                //计算雷达图的坐标轴单位长度
                int majorUnit = calMajorUnit(tableChart.getChartYSeries(), tableData);
                //雷达图需要画网格线
                XDDFShapeProperties bottomAxisGridLinesShapeProperties = bottomAxis.getOrAddMajorGridProperties();
                bottomAxisGridLinesShapeProperties.setLineProperties(new XDDFLineProperties(
                        new XDDFSolidFillProperties(XDDFColor.from(PresetColor.BLACK))));
                leftAxis.setMajorUnit(majorUnit);
                XDDFShapeProperties leftAxisShapeProperties = leftAxis.getOrAddShapeProperties();
                leftAxisShapeProperties.setLineProperties(new XDDFLineProperties(
                        new XDDFSolidFillProperties(XDDFColor.from(PresetColor.BLACK))));
                XDDFShapeProperties leftAxisGridLinesShapeProperties = leftAxis.getOrAddMajorGridProperties();
                leftAxisGridLinesShapeProperties.setLineProperties(new XDDFLineProperties(
                        new XDDFSolidFillProperties(XDDFColor.from(PresetColor.GRAY))));
                leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
                leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

                XDDFRadarChartData data = (XDDFRadarChartData) chart.createData(ChartTypes.RADAR, bottomAxis, leftAxis);
                data.setVaryColors(false);
                for (Field field : tableChart.getChartYSeries()) {
                    String[] colYInfo = colData.get(field);
                    XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(Integer.valueOf(colYInfo[1]), Integer.valueOf(colYInfo[2]) - 1, Integer.valueOf(colYInfo[3]), Integer.valueOf(colYInfo[3])));
                    XDDFRadarChartData.Series series1 = (XDDFRadarChartData.Series) data.addSeries(xs, ys);
                    series1.setTitle(colYInfo[0], null);
                }
                chart.plot(data);
                data.setStyle(RadarStyle.MARKER);
                break;
            }
            case PIE: {
                XDDFPieChartData data = (XDDFPieChartData) chart.createData(ChartTypes.PIE, null, null);
                for (Field field : tableChart.getChartYSeries()) {
                    String[] colYInfo = colData.get(field);
                    XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(Integer.valueOf(colYInfo[1]), Integer.valueOf(colYInfo[2]) - 1, Integer.valueOf(colYInfo[3]), Integer.valueOf(colYInfo[3])));
                    XDDFPieChartData.Series series1 = (XDDFPieChartData.Series) data.addSeries(xs, ys);
                    series1.setTitle(colYInfo[0], null);
                }
                chart.plot(data);
                if (!chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).isSetDLbls()) {
                    chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewDLbls();
                }
                CTDLbls ctdLbls = chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls();
                ctdLbls.addNewShowLegendKey().setVal(false);
                ctdLbls.addNewShowPercent().setVal(true);
                ctdLbls.addNewShowVal().setVal(false);
                ctdLbls.addNewShowCatName().setVal(false);
                ctdLbls.addNewShowSerName().setVal(false);
                break;
            }
        }
    }

    private int calMajorUnit(List<Field> chartYSeries, List<Object> tableData) {
        int min = 0;
        int max = 0;
        for (Field field : chartYSeries) {
            for (Object item : tableData) {
                field.setAccessible(true);
                try {
                    Object value = field.get(item);
                    Integer valueInt = null;
                    if (value instanceof String) {
                        valueInt = Integer.valueOf((String) value);
                    } else if (value instanceof Double) {
                        valueInt = ((Double) value).intValue();
                    } else if (value instanceof Integer) {
                        valueInt = (Integer) value;
                    } else if (value instanceof Long) {
                        valueInt = ((Long) value).intValue();
                    } else if (Objects.isNull(value)) {
                        valueInt = 0;
                    } else {
                        throw new IllegalArgumentException("不能使用String,Double,Date,Boolean,Integer以外的类型");
                    }
                    if (valueInt < min) {
                        min = valueInt;
                    }
                    if (valueInt > max) {
                        max = valueInt;
                    }
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        }

        return (max - min) / tableData.size();
    }

    /**
     * 填充表格中每一个单元格的数据
     *
     * @param sheet                工作簿，用于获取表格中的一行
     * @param tableTitle           表格的名字
     * @param colNameMap           表格每一列的标题
     * @param tableData            表格中的数据
     * @param rowStartIndex        表格的起始行索引
     * @param colStartIndex        表格的起始列索引
     * @param MAX_INDEX_OF_ROWS    表格的终止行索引
     * @param MAX_INDEX_OF_COLUMNS 表格的终止列索引
     * @param wb                   excel文件引用，用于构建样式
     * @param colData              记录一张表里列对应的数据,列名，行开始索引，行终止索引，列索引，用于设置图的联动区域
     * @return key表格的列索引，value列宽度
     */
    private Map<Integer, Integer> fillTableData(Sheet sheet, String tableTitle, List<ExcelColumnMap> colNameMap, List<Object> tableData, int rowStartIndex,
                                                int colStartIndex, int MAX_INDEX_OF_ROWS, int MAX_INDEX_OF_COLUMNS, Workbook wb, Map<Field, String[]> colData) {
        Map<Integer, Integer> colSizeMap = new HashMap<>(colNameMap.size());
        for (int rowIndex = rowStartIndex; rowIndex < MAX_INDEX_OF_ROWS; rowIndex++) {
            // 这里要注意,创建第n(n>1)张表的时候，因为行已经存在，所以要读取旧的行，否则数据会被覆盖
            Row oldRow = sheet.getRow(rowIndex);
            Row row = null == oldRow ? sheet.createRow(rowIndex) : oldRow;
            for (int colIndex = colStartIndex; colIndex < MAX_INDEX_OF_COLUMNS; colIndex++) {
                int size = colSizeMap.get(colIndex) == null ? 10 : colSizeMap.get(colIndex);
                Cell cell = row.createCell(colIndex);
                if (rowIndex == rowStartIndex) {
                    //设置标题行
                    if (colIndex == colStartIndex) {
                        //填充标题
                        cell.setCellValue(tableTitle);
                        //合并标题的单元格
                        sheet.addMergedRegion(new CellRangeAddress(rowStartIndex, rowStartIndex, colStartIndex, MAX_INDEX_OF_COLUMNS - 1));
                    }
                    //设置单元格风格
                    CellStyle style = wb.createCellStyle();
                    style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    ExcelUtils.setUpCellBorderAndAlignment(style);
                    //设置字体
                    Font font = wb.createFont();
                    font.setFontName("微软雅黑");
                    font.setBold(true);
                    font.setColor(IndexedColors.BLACK.getIndex());
                    font.setFontHeightInPoints((short) 14);
                    style.setFont(font);
                    cell.setCellStyle(style);
                }
                if (rowIndex == rowStartIndex + 1) {
                    //设置列名所在行
                    String colName = colNameMap.get(colIndex - colStartIndex).getCloName();
                    //填充列
                    cell.setCellValue(colName);
                    //设置单元格风格
                    CellStyle style = wb.createCellStyle();
                    ExcelUtils.setUpCellBorderAndAlignment(style);
                    //设置字体
                    Font font = wb.createFont();
                    font.setFontName("微软雅黑");
                    font.setColor(IndexedColors.BLACK.getIndex());
                    font.setFontHeightInPoints((short) 13);
                    style.setFont(font);
                    cell.setCellStyle(style);
                    size = calSize(size, colName.getBytes().length);
                }
                if (rowIndex >= rowStartIndex + 2 && null != tableData && !tableData.isEmpty()) {
                    //填充每一列的数据
                    Object obj = tableData.get(rowIndex - rowStartIndex - 2);
                    Field field = colNameMap.get(colIndex - colStartIndex).getValue();
                    int colDataSize = setCellValueByObj(cell, field, obj);
                    //设置单元格风格
                    CellStyle style = wb.createCellStyle();
                    ExcelUtils.setUpCellBorderAndAlignment(style);
                    Font font = wb.createFont();
                    font.setFontName("微软雅黑");
                    style.setFont(font);
                    cell.setCellStyle(style);
                    size = calSize(size, colDataSize);
                }
                colSizeMap.put(colIndex, size);
            }
        }
        for (int i = 0; i < colNameMap.size(); i++) {
            ExcelColumnMap colInfo = colNameMap.get(i);
            colData.put(colInfo.getValue(), new String[]{colInfo.getCloName(), String.valueOf(rowStartIndex + 2), String.valueOf(MAX_INDEX_OF_ROWS), String.valueOf(colStartIndex + i)});
        }
        return colSizeMap;
    }

    private int calSize(int source, int target) {
        if (target > source) {
            return target;
        }
        return source;
    }

    /**
     * 通过反射设置单元格的值
     *
     * @param cell  单元格
     * @param field 列
     * @param obj   数据对象
     * @return 如果是字符串，返回字符串的字节长度
     * 否则返回0
     */
    private int setCellValueByObj(Cell cell, Field field, Object obj) {
        try {
            field.setAccessible(true);
            Object value = field.get(obj);
            if (value instanceof String) {
                cell.setCellValue((String) value);
                return ((String) value).getBytes().length;
            } else if (value instanceof Double) {
                cell.setCellValue((Double) value);
            } else if (value instanceof Date) {
                cell.setCellValue((Date) value);
            } else if (value instanceof Boolean) {
                cell.setCellValue((Boolean) value);
            } else if (value instanceof Integer) {
                cell.setCellValue((Integer) value);
            } else if (value instanceof Long) {
                cell.setCellValue((Long) value);
            } else if (Objects.isNull(value)) {
                cell.setCellValue(0);
            } else {
                throw new IllegalArgumentException("不能使用String,Double,Date,Boolean,Integer以外的类型");
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return 0;
    }

}
