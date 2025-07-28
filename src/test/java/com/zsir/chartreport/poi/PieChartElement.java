package com.zsir.chartreport.poi;

import lombok.Data;

import java.util.List;

/***
 * @Description 饼图元素
 * @Author zjj
 * @Date 2025/7/12 10:00
 */
@Data
public class PieChartElement extends ChartElement {

    /**
     * X轴 图例
     */
    private List<String> xData;

    /**
     * 饼状数据
     */
    private AxisYVal yData;
}

