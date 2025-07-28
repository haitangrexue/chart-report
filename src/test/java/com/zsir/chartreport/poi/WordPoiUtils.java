package com.zsir.chartreport.poi;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChartSpace;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.springframework.web.context.request.RequestAttributes;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.math.BigInteger;
import java.net.URLEncoder;
import java.util.List;

/***
 * @Description POI操作word工具类
 * @Author zjj
 * @Date 2025/7/12 10:00
 */
@Slf4j
public class WordPoiUtils {

    /**
     * 创建文件标题
     *
     * @param document 文档
     * @param docTitle 完档标题
     */
    public static void createDocTitle(
            XWPFDocument document,
            String docTitle,
            boolean isBreak
    ) {
        //新建一个标题段落对象（就是一段文字）
        XWPFParagraph titleParagraph = document.createParagraph();
        //样式居中
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);
        //创建文本对象
        XWPFRun titleFun = titleParagraph.createRun();
        //设置标题的名字
        titleFun.setText(docTitle);
        //加粗
        titleFun.setBold(true);
        //设置颜色
        titleFun.setColor("000000");
        //字体大小
        titleFun.setFontSize(18);
        if (isBreak) {
            //换行
            titleFun.addBreak();
        }
    }

    /***
     * 创建段落标题
     * @param document 文档
     * @param title 段落标题
     * @param isBreak 是否换行
     */
    public static void createParagraphTitle(
            XWPFDocument document,
            String title,
            boolean isBreak
    ) {
        XWPFParagraph titleParagraph = document.createParagraph();
        XWPFRun titleFun = titleParagraph.createRun();
        titleFun.setFontFamily("宋体(中文)");
        titleParagraph.setAlignment(ParagraphAlignment.LEFT);
        titleFun.setText(title);
        titleFun.setColor("000000");
        titleFun.setFontSize(16);
        titleFun.setBold(true);
        //换行
        if (isBreak) {
            titleFun.addBreak();
        }
    }

    /**
     * 创建word文档段落
     *
     * @param document 文档
     * @param align ParagraphAlignment.LEFT 居左
     * @param content 内容
     * @param isBreak 是否换行
     */
    public static void createParagraph(
            XWPFDocument document,
            ParagraphAlignment align,
            String content,
            boolean isBreak,
            int firstLintIndent
    ) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(align);
        XWPFRun titleFun = paragraph.createRun();
        titleFun.setFontSize(14);
        titleFun.addTab();
        titleFun.setFontFamily("宋体(中文)");
        titleFun.setText(content);
        titleFun.setColor("000000");
        titleFun.setFontSize(16);
        paragraph.setFirstLineIndent(firstLintIndent); //480
        if (isBreak) titleFun.addBreak(); //换行
    }

    /**
     * 创建饼图
     *
     * @param document 文档
     * @param element 数据
     * @throws Exception
     */
    public static void createPieChart(
            XWPFDocument document,
            PieChartElement element
    ) throws Exception {
        XWPFChart chart = document.createChart(
                15 * Units.EMU_PER_CENTIMETER,
                10 * Units.EMU_PER_CENTIMETER
        );
        // 设置图例位置:上下左右
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.BOTTOM);
        legend.setOverlay(false);
        chart.setTitleText(element.getTitle());
        chart.setTitleOverlay(false);
        // 禁用圆角
        CTChartSpace chartSpace = chart.getCTChartSpace();
        if (!chartSpace.isSetRoundedCorners()) {
            chartSpace.addNewRoundedCorners();
        }

        chartSpace.getRoundedCorners().setVal(false);
        int numOfPoints = element.getXData().size();
        String[] categories = element.getXData().toArray(new String[] {});
        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(
                categories
        );
        String range = chart.formatRange(
                new CellRangeAddress(1, numOfPoints, 0, 0)
        );
        AxisYVal yData = element.getYData();
        XDDFNumericalDataSource<?> source = XDDFDataSourcesFactory.fromArray(
                yData.getVal(),
                range
        );
        XDDFPieChartData pieChart = (XDDFPieChartData) chart.createData(
                ChartTypes.PIE,
                null,
                null
        );
        pieChart.setVaryColors(true);
        XDDFPieChartData.Series series =
                (XDDFPieChartData.Series) pieChart.addSeries(categoriesData, source);
        series.setTitle(
                yData.getTitle(),
                setTitleInDataSheet(chart, yData.getTitle(), 0)
        );
        chart.plot(pieChart);
    }

    /**
     * 创建表格
     *
     * @param document 文档
     * @param element 表格数据
     */
    public static void createTable(XWPFDocument document, TableElement element) {
        XWPFTable table = document.createTable();
        List<TableElement.CellData> headerList = element.getHeaderList();
        //创建表头
        XWPFTableRow headerRow = getTableRow(table, 0);
        for (int i = 0; i < headerList.size(); i++) {
            TableElement.CellData cellData = headerList.get(i);
            XWPFTableCell cell = getTableRowCell(headerRow, i);
            cell.setText(cellData.getVal());
        }
        //表格数据
        List<List<TableElement.CellData>> rowDataList = element.getRowDataList();
        for (int i = 0; i < rowDataList.size(); i++) {
            int rowIndex = i + 1;
            List<TableElement.CellData> cellData = rowDataList.get(i);
            XWPFTableRow row = getTableRow(table, rowIndex);
            for (int x = 0; x < cellData.size(); x++) {
                TableElement.CellData cellD = cellData.get(x);
                XWPFTableCell cell = getTableRowCell(row, x);
                cell.setText(cellD.getVal());
            }
        }

        CTTblWidth width = table.getCTTbl().addNewTblPr().addNewTblW();
        width.setType(STTblWidth.PCT);
        width.setW(BigInteger.valueOf(5000)); // 5000=100%
    }

    private static XWPFTableCell getTableRowCell(XWPFTableRow row, int index) {
        XWPFTableCell cell = row.getCell(index);
        if (cell != null) {
            return cell;
        }
        return row.createCell();
    }

    private static XWPFTableRow getTableRow(XWPFTable table, int index) {
        XWPFTableRow row = table.getRow(index);
        if (row != null) {
            return row;
        }
        return table.createRow();
    }

    /**
     * 生成图表 柱状图+折线+ 柱状折线组合
     *
     * @param document 文档
     * @param chartForm 数据
     * @throws Exception
     */
    public static void createBarLineCharts(
            XWPFDocument document,
            BarLineChartElement chartForm
    ) throws Exception {
        if (chartForm == null) {
            throw new IllegalArgumentException("数据为空");
        }

        List<String> xData = chartForm.getXData();
        if (xData == null) {
            throw new IllegalArgumentException("X轴数据为空");
        }

        List<AxisYVal> yBarData = chartForm.getYBarData();
        List<AxisYVal> yLintData = chartForm.getYLineData();
        if (yBarData == null && yLintData == null) {
            throw new IllegalArgumentException("柱状数据、折线数据为空");
        }

        XWPFChart chart = document.createChart(
                15 * Units.EMU_PER_CENTIMETER,
                10 * Units.EMU_PER_CENTIMETER
        );
        // 设置图例位置:上下左右
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP);
        legend.setOverlay(false);
        // 设置标题
        chart.setTitleText(chartForm.getTitle());
        //标题覆盖
        chart.setTitleOverlay(false);

        // 禁用圆角
        CTChartSpace chartSpace = chart.getCTChartSpace();
        if (!chartSpace.isSetRoundedCorners()) {
            chartSpace.addNewRoundedCorners();
        }

        chartSpace.getRoundedCorners().setVal(false);

        int numOfPoints = xData.size();
        //初始化X轴数据
        String[] categories = xData.toArray(new String[] {});
        String cat = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(
                categories,
                cat,
                0
        );
        //初始化X轴
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(chartForm.getXTitle());
        //初始化Y周
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(chartForm.getYTitle());
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
        int barSize = 0;
        if (yBarData != null) { //柱状图
            barSize = yBarData.size();
            XDDFBarChartData barChart = (XDDFBarChartData) chart.createData(
                    ChartTypes.BAR,
                    bottomAxis,
                    leftAxis
            );
            barChart.setBarDirection(BarDirection.COL);
            for (int i = 0; i < yBarData.size(); i++) {
                int collIndex = i + 1;
                String range = chart.formatRange(
                        new CellRangeAddress(1, numOfPoints, collIndex, collIndex)
                );
                AxisYVal yVal = yBarData.get(i);
                Number[] number = yVal.getVal();
                XDDFNumericalDataSource<?> source = XDDFDataSourcesFactory.fromArray(
                        number,
                        range
                );
                XDDFBarChartData.Series series =
                        (XDDFBarChartData.Series) barChart.addSeries(categoriesData, source);
                series.setTitle(
                        yVal.getTitle(),
                        setTitleInDataSheet(chart, yVal.getTitle(), collIndex)
                );
            }
            if (yBarData.size() > 1) {
                barChart.setBarGrouping(BarGrouping.CLUSTERED);
                chart
                        .getCTChart()
                        .getPlotArea()
                        .getBarChartArray(0)
                        .addNewOverlap()
                        .setVal((byte) 0);
            }

            chart.plot(barChart);
        }

        if (yLintData != null) { //折线图
            XDDFLineChartData lineChart = (XDDFLineChartData) chart.createData(
                    ChartTypes.LINE,
                    bottomAxis,
                    leftAxis
            );
            lineChart.setVaryColors(false);
            for (int i = 0; i < yLintData.size(); i++) {
                int collIndex = i + 1 + barSize;
                String range = chart.formatRange(
                        new CellRangeAddress(1, numOfPoints, collIndex, collIndex)
                );
                AxisYVal yVal = yLintData.get(i);
                Number[] number = yVal.getVal();
                XDDFNumericalDataSource<?> source = XDDFDataSourcesFactory.fromArray(
                        number,
                        range
                );
                XDDFLineChartData.Series series =
                        (XDDFLineChartData.Series) lineChart.addSeries(
                                categoriesData,
                                source
                        );
                series.setTitle(
                        yVal.getTitle(),
                        setTitleInDataSheet(chart, yVal.getTitle(), collIndex)
                );
                series.setMarkerStyle(MarkerStyle.DOT);
                series.setSmooth(Boolean.FALSE);
            }

            chart.plot(lineChart);
        }
    }

    /**
     * 生成横向柱状图
     *
     * @param document word文档
     * @param chartForm 图表数据
     * @throws Exception
     */
    public static void createBarCharts(
            XWPFDocument document,
            BarLineChartElement chartForm
    ) throws Exception {
        if (chartForm == null) {
            throw new IllegalArgumentException("数据为空");
        }

        List<String> xData = chartForm.getXData();
        if (xData == null) {
            throw new IllegalArgumentException("X轴数据为空");
        }

        List<AxisYVal> yBarData = chartForm.getYBarData();
        List<AxisYVal> yLintData = chartForm.getYLineData();
        if (yBarData == null && yLintData == null) {
            throw new IllegalArgumentException("柱状数据、折线数据为空");
        }

        XWPFChart chart = document.createChart(
                15 * Units.EMU_PER_CENTIMETER,
                10 * Units.EMU_PER_CENTIMETER
        );
        // 设置图例位置:上下左右
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP);
        legend.setOverlay(false);
        // 设置标题
        chart.setTitleText(chartForm.getTitle());
        //标题覆盖
        chart.setTitleOverlay(false);

        // 禁用圆角
        CTChartSpace chartSpace = chart.getCTChartSpace();
        if (!chartSpace.isSetRoundedCorners()) {
            chartSpace.addNewRoundedCorners();
        }

        chartSpace.getRoundedCorners().setVal(false);

        int numOfPoints = xData.size();
        //初始化X轴数据
        String[] categories = xData.toArray(new String[] {});
        String cat = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(
                categories,
                cat,
                0
        );
        //初始化X轴
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(chartForm.getXTitle());
        //初始化Y周
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(chartForm.getYTitle());
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
        // 创建柱状图的类型
        XDDFBarChartData barChart = (XDDFBarChartData) chart.createData(
                ChartTypes.BAR,
                bottomAxis,
                leftAxis
        );
        barChart.setBarDirection(BarDirection.BAR);
        for (int i = 0; i < yBarData.size(); i++) {
            int collIndex = i + 1;
            String range = chart.formatRange(
                    new CellRangeAddress(1, numOfPoints, collIndex, collIndex)
            );
            AxisYVal yVal = yBarData.get(i);
            Number[] number = yVal.getVal();
            XDDFNumericalDataSource<?> source = XDDFDataSourcesFactory.fromArray(
                    number,
                    range
            );
            XDDFBarChartData.Series series =
                    (XDDFBarChartData.Series) barChart.addSeries(categoriesData, source);
            series.setTitle(
                    yVal.getTitle(),
                    setTitleInDataSheet(chart, yVal.getTitle(), collIndex)
            );
            break;
        }
        if (yBarData.size() > 1) {
            barChart.setBarGrouping(BarGrouping.CLUSTERED);
            chart
                    .getCTChart()
                    .getPlotArea()
                    .getBarChartArray(0)
                    .addNewOverlap()
                    .setVal((byte) 0);
        }

        chart.plot(barChart);
    }

    private static CellReference setTitleInDataSheet(
            XWPFChart chart,
            String title,
            int column
    ) throws Exception {
        XSSFWorkbook workbook = chart.getWorkbook();
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(0);
        if (row == null) {
            row = sheet.createRow(0);
        }

        XSSFCell cell = row.getCell(column);
        if (cell == null) {
            cell = row.createCell(column);
        }

        cell.setCellValue(title);
        return new CellReference(sheet.getSheetName(), 0, column, true, true);
    }

    /**
     * 浏览器下载
     *
     * @param document 文档
     * @param fileName 文件名
     * @throws IOException
     */
    public static void save(XWPFDocument document, String fileName)
            throws IOException {
        ServletOutputStream out = null;
        try {
            if (document == null || fileName == null) {
                throw new IllegalArgumentException("参数不能为空");
            }

            RequestAttributes requestAttributes =
                    RequestContextHolder.getRequestAttributes();
            if (requestAttributes != null) {
                ServletRequestAttributes servletRequestAttributes =
                        (ServletRequestAttributes) requestAttributes;
                HttpServletResponse response = servletRequestAttributes.getResponse();
                response.setContentType("application/octet-stream");
                out = response.getOutputStream();
                String encodedFileName;
                try {
                    encodedFileName = URLEncoder.encode(fileName, "UTF-8");
                } catch (UnsupportedEncodingException var20) {
                    throw new RuntimeException(var20);
                }

                response.setHeader(
                        "Content-Disposition",
                        "attachment;fileName=" + encodedFileName
                );
                document.write(out);
                return;
            }
            log.warn("非web请求，无法操作。");
        } finally {
            if (document != null) {
                try {
                    document.close();
                } catch (Exception var19) {}
            }
            if (out != null) {
                try {
                    out.close();
                } catch (Exception var18) {}
            }
        }
    }

    /**
     * 另存为
     *
     * @param document 文档
     * @param fileName 文件名
     * @param path 文件存储目录
     * @throws IOException
     */
    public static void save(XWPFDocument document, String fileName, String path)
            throws IOException {
        FileOutputStream out = null;
        try {
            if (document == null || fileName == null || path == null) {
                throw new IllegalArgumentException("参数不能为空");
            }

            File file = new File(path + fileName);
            out = new FileOutputStream(file);
            document.write(out);
            document.close();
        } finally {
            if (document != null) {
                try {
                    document.close();
                } catch (Exception var19) {}
            }
            if (out != null) {
                try {
                    out.close();
                } catch (Exception var18) {}
            }
        }
    }
}

