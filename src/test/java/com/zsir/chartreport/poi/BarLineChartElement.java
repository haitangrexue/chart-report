package com.zsir.chartreport.poi;

import lombok.Data;

import java.util.List;

/***
 * @Description 柱状+折线元素
 * @Author zjj
 * @Date 2025/7/12 10:00
 */
@Data
public class BarLineChartElement extends ChartElement {

    /**
     * X轴数据
     */
    private List<String> xData;
    /**
     * 柱状数据 支持多组
     */
    private List<AxisYVal> yBarData;
    /**
     * 折线数据 支持多组
     */
    private List<AxisYVal> yLineData;
}

