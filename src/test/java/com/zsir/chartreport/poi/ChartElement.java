package com.zsir.chartreport.poi;

import lombok.Data;

/***
 * @Description 图表公共属性
 * @Author zjj
 * @Date 2025/7/12 10:00
 */
@Data
public class ChartElement {

    /**
     * 图表标题
     */
    private String title;
    /**
     * X轴标题
     */
    private String xTitle;
    /**
     * Y轴标题
     */
    private String yTitle;
}

