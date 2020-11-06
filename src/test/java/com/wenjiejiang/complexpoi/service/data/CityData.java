package com.wenjiejiang.complexpoi.service.data;

/**
 * @author wenjiejiang
 * @date 2020/11/6 14:56
 * @description 生成一张表的测试数据
 * @since 0.1
 */
public class CityData {
    /**
     * 城市名
     */
    private String cityName;
    /**
     * 城市GDP
     */
    private Double totalGDP;
    /**
     * 城市人口数量
     */
    private Double populationNumber;

    public CityData(String cityName, Double totalGDP, Double populationNumber) {
        this.cityName = cityName;
        this.totalGDP = totalGDP;
        this.populationNumber = populationNumber;
    }

    public String getCityName() {
        return cityName;
    }

    public void setCityName(String cityName) {
        this.cityName = cityName;
    }

    public Double getTotalGDP() {
        return totalGDP;
    }

    public void setTotalGDP(Double totalGDP) {
        this.totalGDP = totalGDP;
    }

    public Double getPopulationNumber() {
        return populationNumber;
    }

    public void setPopulationNumber(Double populationNumber) {
        this.populationNumber = populationNumber;
    }
}
