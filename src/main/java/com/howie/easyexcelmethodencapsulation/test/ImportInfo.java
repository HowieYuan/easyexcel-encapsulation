package com.howie.easyexcelmethodencapsulation.test;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

/**
 * Created with IntelliJ IDEA
 *
 * @Author yuanhaoyue swithaoy@gmail.com
 * @Description 映射实体类，Excel 模型
 * @Date 2018-06-05
 * @Time 21:41
 */
public class ImportInfo extends BaseRowModel {
    @ExcelProperty(index = 0)
    private String station;

    @ExcelProperty(index = 1)
    private String stationsubstation;

    @ExcelProperty(index = 2)
    private String cycle;

    /*
        作为 excel 的模型映射，需要 setter 方法
     */
    public String getStation() {
        return station;
    }

    public void setStation(String station) {
        this.station = station;
    }

    public String getStationsubstation() {
        return stationsubstation;
    }

    public void setStationsubstation(String stationsubstation) {
        this.stationsubstation = stationsubstation;
    }

    public String getCycle() {
        return cycle;
    }

    public void setCycle(String cycle) {
        this.cycle = cycle;
    }

    @Override
    public String toString() {
        return "Info{" +
                "station='" + station + '\'' +
                ", stationsubstation='" + stationsubstation + '\'' +
                ", cycle='" + cycle + '\'' +
                '}';
    }
}
