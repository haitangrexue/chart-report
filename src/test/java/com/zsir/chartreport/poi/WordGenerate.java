package com.zsir.chartreport.poi;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;

/**
 * @ClassName WordGenerate
 * @description:
 * @author: zjj
 * @create: 2025-07-12 10:00
 **/
public class WordGenerate {

    @Test
    public void create() throws Exception {
        XWPFDocument doc = new XWPFDocument();
        WordPoiUtils.createDocTitle(doc, "分析报告", false);
        WordPoiUtils.createParagraphTitle(doc, "一、段落标题", false);
        WordPoiUtils.createParagraph(
                doc,
                ParagraphAlignment.LEFT,
                "段落内容",
                false,
                10
        );

        //x轴数据
        List<String> xData = new ArrayList<>();
        xData.add("2025-07-01");
        xData.add("2025-07-02");
        xData.add("2025-07-03");
        xData.add("2025-07-04");
        xData.add("2025-07-05");
        xData.add("2025-07-06");

        //y轴数据
        AxisYVal yData1 = new AxisYVal();
        yData1.setTitle("用户数");
        List<Integer> y1 = new ArrayList<>();
        y1.add(20);
        y1.add(23);
        y1.add(15);
        y1.add(30);
        y1.add(11);
        y1.add(2);
        yData1.setVal(y1.toArray(new Integer[] {}));

        AxisYVal yData2 = new AxisYVal();
        yData2.setTitle("访问量");
        List<Integer> y2 = new ArrayList<>();
        y2.add(10);
        y2.add(20);
        y2.add(15);
        y2.add(12);
        y2.add(31);
        y2.add(9);
        yData2.setVal(y2.toArray(new Integer[] {}));

        List<AxisYVal> yData = new ArrayList<>();
        yData.add(yData1);
        yData.add(yData2);

        AxisYVal yData3 = new AxisYVal();
        yData3.setTitle("趋势");
        List<Double> y3 = new ArrayList<>();
        y3.add(3d);
        y3.add(19d);
        y3.add(21d);
        y3.add(32d);
        y3.add(42d);
        y3.add(31d);
        yData3.setVal(y3.toArray(new Double[] {}));
        List<AxisYVal> yLineData = new ArrayList<>();
        yLineData.add(yData3);

        BarLineChartElement chartForm = new BarLineChartElement();
        chartForm.setTitle("柱状+折线图");
        chartForm.setXTitle("日期");
        chartForm.setYTitle("数量");
        chartForm.setXData(xData);
        chartForm.setYBarData(yData);
        chartForm.setYLineData(yLineData);
        WordPoiUtils.createBarLineCharts(doc, chartForm);

        chartForm.setTitle("柱状图");
        chartForm.setYBarData(yData);
        chartForm.setYLineData(null);
        WordPoiUtils.createBarCharts(doc, chartForm);

        chartForm.setTitle("折线图");
        chartForm.setYBarData(null);
        chartForm.setYLineData(yLineData);
        WordPoiUtils.createBarLineCharts(doc, chartForm);

        //拼图
        PieChartElement pieChartElement = new PieChartElement();
        pieChartElement.setTitle("饼图标题");
        pieChartElement.setXData(xData);
        pieChartElement.setYData(yData2);
        WordPoiUtils.createPieChart(doc, pieChartElement);

        //表格
        TableElement tableElement = new TableElement();
        List<TableElement.CellData> headerData = new ArrayList<>();
        headerData.add(new TableElement.CellData("一级菜单"));
        headerData.add(new TableElement.CellData("点击人数"));
        headerData.add(new TableElement.CellData("占比"));
        tableElement.setHeaderList(headerData);

        List<List<TableElement.CellData>> rowData = new ArrayList<>();
        List<TableElement.CellData> callData = new ArrayList<>();
        callData.add(new TableElement.CellData("文章一"));
        callData.add(new TableElement.CellData("20"));
        callData.add(new TableElement.CellData("30%"));
        rowData.add(callData);
        rowData.add(callData);
        tableElement.setRowDataList(rowData);
        WordPoiUtils.createTable(doc, tableElement);

        WordPoiUtils.save(doc, "poi操作word.docx", "d:/");
    }
}

