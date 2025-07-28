package com.zsir.chartreport.poi;

import lombok.Data;
import org.apache.poi.xddf.usermodel.XDDFColor;

/**
 * @description: Y周坐标值
 * @author: zjj
 * @create: 2025-07-12 10:00
 **/
@Data
public class AxisYVal {

    /**
     * Y轴数据、与X轴数据数量保持一致
     */
    private Number[] val;
    /**
     * Y轴数据 图例
     */
    private String title;
    /**
     * Y轴颜色
     */
    private XDDFColor color;
}
