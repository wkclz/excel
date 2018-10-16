package com.wkclz.util.excelRd;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelRdTest {

    public static void main(String[] args) throws IOException, ExcelRdException {
        excelRd();
    }

    private static void excelRd() throws IOException, ExcelRdException {

        String path = "/Users/wangkaicun/Desktop/test.xlsx";
        ExcelRd excelRd = new ExcelRd(path);
        excelRd.setStartRow(2);    // 指定起始行，下标从0开始计数
        excelRd.setStartCol(1);    // 指定起始列，下标从0开始计数
        ExcelRdTypeEnum[] types = {
            ExcelRdTypeEnum.INTEGER,
            ExcelRdTypeEnum.LONG,
            ExcelRdTypeEnum.DOUBLE,
            ExcelRdTypeEnum.DATETIME,
            ExcelRdTypeEnum.DATE,
            ExcelRdTypeEnum.STRING
        };
        excelRd.setTypes(types);    // 指定每列的类型
        
        List<ExcelRdRow> rows = excelRd.analysisXlsx();
        Map<String, Object>[] plans = new HashMap[rows.size()];

        int size = rows.size();
        for (int i = 0; i < size; i++) {
            
            ExcelRdRow excelRdRow = rows.get(i);
            List<Object> row = excelRdRow.getRow();
            HashMap<String, Object> plan = new HashMap<String, Object>();

            for (Object t : row) {
                System.out.println(t);
            }
            System.out.println("\n");
            
            plans[i] = plan;
        }
    }
}
