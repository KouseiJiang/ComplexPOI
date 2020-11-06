# ComplexPOI
墙内可参考 https://www.yuque.com/kousei/iuq7uq/cbpgc8
# 例子
## 什么是ComplexPOI

1. 是基于POI的封装
1. 提供了简单生成excel以及图(折线图、柱状图、雷达图、饼图)的方式
1. 提供了默认的Excel表格样式(表标题加粗、蓝底，列名加粗，单元格边线，宽度自适应(支持汉字))
1. 提供了默认的Excel图样式(标题、系列、大小、饼图的%上浮)
1. 提供了表格、图之间的相对位置(2单元格)，图依次排列在表格下，当一个工作簿有多张表格时，依次向右排列
## 效果图
![image.png](https://cdn.nlark.com/yuque/0/2020/png/450073/1604648788152-3bccfa49-9e88-4d44-bd68-6c4f6f324aed.png#align=left&display=inline&height=808&margin=%5Bobject%20Object%5D&name=image.png&originHeight=808&originWidth=1897&size=146368&status=done&style=none&width=1897)
注:这里为了方便我调整了图的位置，实际上图会按顺序排在表格下面，如果有多张表格则会依次向右排列
## 测试代码
```java
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

```
# 一分钟入门
![image.png](https://cdn.nlark.com/yuque/0/2020/png/450073/1604649840857-ec649644-d270-4902-93dd-08f43755ff07.png#align=left&display=inline&height=535&margin=%5Bobject%20Object%5D&name=image.png&originHeight=535&originWidth=477&size=33014&status=done&style=none&width=477)

| 术语 | 含义 | 使用 |
| --- | --- | --- |
| ExcelSheet | 代表一个Excel工作簿 | 一个Excel文件可以包含多个Excel工作簿 |
| ExcelTable | 代表Excel工作簿中的一张表格 | 一个Excel工作簿可以包含多张表格 |
| ExcelCloumnMap | 代表一张表格里的标题列 | 一张表格包含多个标题列，如Demo中的"城市名","GDP","人口数量" |
| TableChart | 代表一张表格生成的图 | 一张表格可以包含多张图 |

# FAQ
## 什么是ChartXSeries和ChartYSeries？
针对一张图来说，每个数据都是由(X,Y)的点构成，所以ChartXSeries代表横坐标轴，ChartYSeries代表纵坐标轴(针对折线图和柱状图)；更通俗一点来说，ChartXSeries代表的是系或者类别(Excel术语)，ChartYSeries则代表数据
# 杂谈
ComplexPOI还属于新生阶段，欢迎大家给我提issue，我会尽快反馈(一般来说1周内，快的话2-3天)。
当然也欢迎大家给我提POI相关的需求，我会尽快按需求提供相关功能。


