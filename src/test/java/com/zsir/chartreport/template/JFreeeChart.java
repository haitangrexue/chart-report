package com.zsir.chartreport.template;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.CategoryAxis;
import org.jfree.chart.axis.CategoryLabelPositions;
import org.jfree.chart.labels.ItemLabelAnchor;
import org.jfree.chart.labels.ItemLabelPosition;
import org.jfree.chart.labels.StandardCategoryItemLabelGenerator;
import org.jfree.chart.labels.StandardPieSectionLabelGenerator;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PieLabelLinkStyle;
import org.jfree.chart.plot.PiePlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.LineAndShapeRenderer;
import org.jfree.chart.title.LegendTitle;
import org.jfree.chart.ui.RectangleEdge;
import org.jfree.chart.ui.TextAnchor;
import org.jfree.chart.util.Rotation;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DatasetUtils;
import org.jfree.data.general.DefaultPieDataset;
import org.junit.jupiter.api.Test;

import java.awt.*;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.Base64;
import java.util.Random;

/***
 * @Description JFreeeChart
 * @Author zjj
 * @Date 2025/7/12 10:00
 */
public class JFreeeChart {

    /***
     * @Description 创建饼图
     * @Author zjj
     * @Date 2025/7/12 10:00
     */
    @Test
    public void pieChart() {
        try {
            /*创建饼图数据集*/
            DefaultPieDataset ds = new DefaultPieDataset();
            ds.setValue("高效率", 2);
            ds.setValue("一般", 28);
            ds.setValue("效率低下", 5);
            /*设置统一字体*/
            Font font = new Font("雅黑", Font.BOLD, 15);
            /*创建图表对象*/
            //JFreeChart chart = ChartFactory.createPieChart3D("您认为公司目前的组织机构有效性如何？", ds, true, false, false);
            JFreeChart chart = ChartFactory.createPieChart(
                    "您认为公司目前的组织机构有效性如何？",
                    ds,
                    false,
                    false,
                    false
            );
            /*设置标题*/
            chart.getTitle().setFont(font);

            /*设置图表的背景图片*/
            //chart.setBackgroundImage(ImageIO.read(new File("")));

            /*设置图例   需要将createPieChart的第3个参数置为false*/
            LegendTitle legendTitle = new LegendTitle(chart.getPlot()); //创建图例
            legendTitle.setPosition(RectangleEdge.BOTTOM); //设置图例的位置
            legendTitle.setMargin(0, 0, 0, 5);
            legendTitle.setItemFont(font);
            chart.addLegend(legendTitle);

            /*得到绘图区*/
            PiePlot plot = (PiePlot) chart.getPlot();
            // 饼图的0°起点在3点钟方向，设置为180°是从左边开始计算旋转角度
            plot.setStartAngle(180);
            // 扇形的旋转方向
            plot.setDirection(Rotation.CLOCKWISE);
            /*设置标签字体*/
            plot.setLabelFont(font);
            /*设置阴影颜色 null-取消阴影*/
            plot.setShadowPaint(Color.YELLOW);

            plot.setSectionOutlinesVisible(false);

            /*设置分类颜色*/
            //plot.setSectionPaint("高效率", Color.ORANGE);
            //plot.setSectionPaint("一般", Color.YELLOW);
            //plot.setSectionPaint("效率低下", new Color(220,100,200));

            /*设置绘图的背景图片*/
            //plot.setBackgroundImage(ImageIO.read(new File("")));

            /*设置分离效果*/
            plot.setExplodePercent("高效率", 0.1f);
            plot.setExplodePercent("一般", 0.1f);
            plot.setExplodePercent("效率低下", 0.1f);

            /*3D效果设置前景透明  0-完全透明  1-完全不透明   同时分离效果消失*/
            plot.setForegroundAlpha(0.7f);

            // 饼图的背景全透明
            plot.setBackgroundAlpha(0.0f);

            // 去除背景边框线
            plot.setOutlinePaint(null);

            /*定制标签  0-key  1-value  2-百分比 3-总和*/
            plot.setLabelGenerator(
                    new StandardPieSectionLabelGenerator("{0} {1}   (占比:{2})")
            );

            /*设置提示线为直线*/
            plot.setLabelLinkStyle(PieLabelLinkStyle.STANDARD);

            /*是否显示连线*/
            plot.setLabelLinksVisible(true);

            out(chart, 800, 400);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /***
     * @Description 创建柱状图
     * @Author zjj
     * @Date 2025/7/12 10:00
     */
    @Test
    public void barChart() {
        try {
            /*创建列别数据集*/
            DefaultCategoryDataset ds = new DefaultCategoryDataset();
            /*设置统一字体*/
            ds.setValue(3400, "IBM", "一季度");
            ds.setValue(3100, "ORACLE", "一季度");
            ds.setValue(5900, "用友", "一季度");

            ds.setValue(3000, "IBM", "二季度");
            ds.setValue(2800, "ORACLE", "二季度");
            ds.setValue(2100, "用友", "二季度");

            ds.setValue(6000, "IBM", "三季度");
            ds.setValue(4300, "ORACLE", "三季度");
            ds.setValue(1000, "用友", "三季度");

            Font font = new Font("雅黑", Font.BOLD, 15);

            /*创建图表对象*/
            JFreeChart chart = ChartFactory.createBarChart(
                    "季度统计表",
                    "季度",
                    "销量(单位：万)",
                    ds,
                    PlotOrientation.VERTICAL,
                    true,
                    false,
                    false
            );
            //JFreeChart chart = ChartFactory.createLineChart3D("季度统计表", "季度", "销量(单位：万)", ds, PlotOrientation.VERTICAL, true, false, false);
            chart.getTitle().setFont(font);
            chart.getLegend().setItemFont(font);

            /*拿到绘图区*/
            CategoryPlot plot = (CategoryPlot) chart.getPlot();
            /*X轴*/
            plot.getDomainAxis().setLabelFont(font);
            /*X轴--小标签*/
            plot.getDomainAxis().setTickLabelFont(font);
            /*Y轴*/
            plot.getRangeAxis().setLabelFont(font);
            BarRenderer customBarRenderer = (BarRenderer) plot.getRenderer();
            customBarRenderer.setDefaultItemLabelGenerator(
                    new StandardCategoryItemLabelGenerator()
            ); //显示每个柱的数值
            customBarRenderer.setDefaultItemLabelsVisible(true);
            //注意：此句很关键，若无此句，那数字的显示会被覆盖，给人数字没有显示出来的问题
            customBarRenderer.setDefaultPositiveItemLabelPosition(
                    new ItemLabelPosition(
                            ItemLabelAnchor.INSIDE12,
                            TextAnchor.BASELINE_CENTER
                    )
            );
            customBarRenderer.setItemLabelAnchorOffset(15D); // 设置柱形图上的文字偏离值

            out(chart, 800, 600);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /***
     * @Description 创建折线图
     * @Author zjj
     * @Date 2025/7/12 10:00
     */
    @Test
    public void lineChart() {
        try {
            String[] rowKeys = {"A站点", "B站点"};
            String[] colKeys = {"0:00", "1:00", "2:00", "7:00", "8:00", "9:00"};
            double[][] data = {{4, 3, 1, 1, 1, 1}, {5, 3, 6, 8, 6, 3}};

            //此数据源可以替代 DatasetUtils.createCategoryDataset(rowKeys, colKeys, data)
      /*DefaultCategoryDataset ds = new DefaultCategoryDataset();
		设置统一字体
		ds.addValue(1, "IBM", "上午");
		ds.addValue(4, "IBM", "中午");
		ds.addValue(3, "IBM", "晚上");

		ds.addValue(3, "ORACLE", "上午");
		ds.addValue(2, "ORACLE", "中午");
		ds.addValue(1, "ORACLE", "晚上");

		ds.addValue(2, "用友", "上午");
		ds.addValue(3, "用友", "中午");
		ds.addValue(4, "用友", "晚上");*/

            Font font = new Font("雅黑", Font.BOLD, 15);
            // 创建JFreeChart对象：ChartFactory.createLineChart
            JFreeChart chart = ChartFactory.createLineChart(
                    "不同类别按小时计算拆线图",
                    "年分",
                    "数量",
                    DatasetUtils.createCategoryDataset(rowKeys, colKeys, data),
                    PlotOrientation.VERTICAL,
                    true, // legend
                    false, // tooltips
                    false
            ); // URLs
            chart.getTitle().setFont(font);
            chart.getLegend().setItemFont(font);
            // 使用CategoryPlot设置各种参数。以下设置可以省略。
            CategoryPlot plot = (CategoryPlot) chart.getPlot();
            /*X轴*/
            plot.getDomainAxis().setLabelFont(font);
            /*X轴--小标签*/
            CategoryAxis categoryAxis = plot.getDomainAxis();
            categoryAxis.setTickLabelFont(font);
            categoryAxis.setCategoryLabelPositions(CategoryLabelPositions.UP_45);
            /*Y轴*/
            plot.getRangeAxis().setLabelFont(font);
            // 背景色 透明度
            plot.setBackgroundAlpha(0.5f);
            // 前景色 透明度
            plot.setForegroundAlpha(0.5f);
            // 其他设置 参考 CategoryPlot类
            LineAndShapeRenderer renderer = (LineAndShapeRenderer) plot.getRenderer();
            renderer.setDefaultShapesVisible(true); // series 点（即数据点）可见
            renderer.setDefaultLinesVisible(true); // series 点（即数据点）间有连线可见
            renderer.setDefaultItemLabelGenerator(
                    new StandardCategoryItemLabelGenerator()
            );
            renderer.setDefaultItemLabelsVisible(true);

            out(chart, 800, 600);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void out(JFreeChart chart, int width, int height) throws IOException {
        //保存至目录文件
        Random random = new Random();
        char randomChar = (char) (random.nextInt(26) + 'A');
        ChartUtils.saveChartAsPNG(
                new File("D:/" + randomChar + ".png"),
                chart,
                width,
                height
        );

        //写入到输出流
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        ChartUtils.writeChartAsPNG(outputStream, chart, 800, 600);

        System.out.println(
                "base64编码：" +
                        new String(
                                Base64.getEncoder().encode(outputStream.toByteArray()),
                                "UTF-8"
                        )
        );
    }

}
