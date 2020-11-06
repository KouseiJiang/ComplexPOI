package com.wenjiejiang.complexpoi.service;

import com.wenjiejiang.complexpoi.bean.ExcelColumnMap;
import com.wenjiejiang.complexpoi.bean.ExcelSheet;
import com.wenjiejiang.complexpoi.bean.ExcelTable;
import com.wenjiejiang.complexpoi.bean.TableChart;
import com.wenjiejiang.complexpoi.bean.constant.EnumChartType;
import com.wenjiejiang.complexpoi.service.data.CityData;
import com.wenjiejiang.complexpoi.service.impl.ComplexExcelServiceImpl;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import org.springframework.util.ReflectionUtils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class IComplexExcelServiceTest {

    @Test
    public void createExcel() throws IOException {
        IComplexExcelService complexExcelService = new ComplexExcelServiceImpl();
        List<ExcelSheet> excelSheets = createExcelSheets();
        Workbook workbook = complexExcelService.createExcel(excelSheets);
        FileOutputStream fileOut = new FileOutputStream("test.xlsx");
        workbook.write(fileOut);
        fileOut.close();
    }

    private List<ExcelSheet> createExcelSheets() {
        List<CityData> cityDataList= createCityDataList();
        Field cityNameField= ReflectionUtils.findField(CityData.class,"cityName");
        Field totalGDPField= ReflectionUtils.findField(CityData.class,"totalGDP");
        Field populationNumberField= ReflectionUtils.findField(CityData.class,"populationNumber");

        // 测试折线图
        TableChart lineChart=new TableChart("折线图", EnumChartType.LINE);
        lineChart.setChartXSeries(cityNameField);
        lineChart.setChartYSeries(Arrays.asList(totalGDPField,populationNumberField));

        // 测试柱状图
        TableChart barChart=new TableChart("柱状图", EnumChartType.BAR);
        barChart.setChartXSeries(cityNameField);
        barChart.setChartYSeries(Arrays.asList(totalGDPField,populationNumberField));

        // 测试饼图
        TableChart pieChart=new TableChart("饼图", EnumChartType.PIE);
        pieChart.setChartXSeries(cityNameField);
        // 饼图的Y系是唯一的
        pieChart.setChartYSeries(Arrays.asList(totalGDPField));

        // 测试折线图
        TableChart radarChart=new TableChart("折线图", EnumChartType.RADAR);
        radarChart.setChartXSeries(cityNameField);
        radarChart.setChartYSeries(Arrays.asList(totalGDPField,populationNumberField));

        List<TableChart> tableCharts=Arrays.asList(lineChart,barChart,pieChart,radarChart);

        // 构建Excel表格
        ExcelTable<CityData> excelTable=new ExcelTable<>();
        // 设置表的标题
        excelTable.setTitle("城市人口和GDP分布表格");
        // 设置表的列名
        ExcelColumnMap cityNameColumn=new ExcelColumnMap("城市名",cityNameField);
        ExcelColumnMap totalGDPColumn=new ExcelColumnMap("GDP",totalGDPField);
        ExcelColumnMap populationNumberColumn=new ExcelColumnMap("人口数量",populationNumberField);
        excelTable.setColNameMap(Arrays.asList(cityNameColumn,totalGDPColumn,populationNumberColumn));
        // 设置数据和图表
        excelTable.setTableDataList(cityDataList);
        excelTable.setTableChartList(tableCharts);
        List<ExcelTable> excelTables=new ArrayList<>(Arrays.asList(excelTable));

        //构建工作簿
        ExcelSheet excelSheet=new ExcelSheet("城市人口和GDP分布");
        excelSheet.setExcelTableList(excelTables);
        List<ExcelSheet> excelSheets = Arrays.asList(excelSheet);
        return excelSheets;
    }

    public List<CityData> createCityDataList(){
        CityData cityData1=new CityData("南京市",80000.58,10000.00);
        CityData cityData2=new CityData("北京市",90000.58,12000.00);
        CityData cityData3=new CityData("天津市",70000.58,9000.00);
        CityData cityData4=new CityData("上海市",92000.58,13000.00);
        CityData cityData5=new CityData("武汉市",65000.58,8000.00);
        List<CityData> cityData=new ArrayList<>(Arrays.asList(cityData1,cityData2,cityData3,cityData4,cityData5));
        return cityData;
    }
}
