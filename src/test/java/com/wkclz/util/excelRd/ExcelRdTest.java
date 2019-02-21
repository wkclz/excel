package com.wkclz.util.excelRd;

import java.io.IOException;
import java.util.List;

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

        List<List<Object>> rows = excelRd.analysisXlsx();

        for (List<Object> row : rows) {
            for (Object cell : row) {
                System.out.println(cell);
            }
            System.out.println("\n");
        }
    }
}
