package com.zsir.chartreport.poi;

import lombok.Data;

import java.util.List;

/***
 * @Description 表格元素
 * @Author zjj
 * @Date 2025/7/12 10:00
 */
@Data
public class TableElement {

    //表头数据
    private List<CellData> headerList;

    //表格数据
    private List<List<CellData>> rowDataList;

    //表格其他属性

    @Data
    public static class CellData {

        //单元格值
        private String val;

        //单元格其它属性

        public CellData(String val) {
            this.val = val;
        }
    }
}

