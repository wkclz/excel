package com.wkclz.util.excelRd;

import com.wkclz.util.excelRd.domain.ExcelRdSheet;

import java.io.IOException;
import java.util.List;

public class ExcelRdTest {

    public static void main(String[] args) throws ExcelRdException {
        // excelRdFirstSheetDeprecated();
        excelRdSheets();
    }


    /**
     * 只能识别一个 sheet 的旧方案【过时，不再建议使用】
     * @throws ExcelRdException
     */
    private static void excelRdFirstSheetDeprecated() throws ExcelRdException {
        String path = "/Users/wangkaicun/project/code/excel/dist/test.xlsx";
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


    /**
     * 识别多个 sheet 的新方案
     * @throws ExcelRdException
     */
    private static void excelRdSheets() throws ExcelRdException {
        String path = "/Users/wangkaicun/project/code/excel/dist/test.xlsx";
        ExcelRd excelRd = new ExcelRd(path);

        ExcelRdTypeEnum[] types = {
                ExcelRdTypeEnum.INTEGER,
                ExcelRdTypeEnum.LONG,
                ExcelRdTypeEnum.DOUBLE,
                ExcelRdTypeEnum.DATETIME,
                ExcelRdTypeEnum.DATE,
                ExcelRdTypeEnum.STRING
        };

        // 创建第一个 sheet 配置
        ExcelRdSheet sheet1 = new ExcelRdSheet();
        sheet1.setStartRow(2);    // 指定起始行，下标从0开始计数
        sheet1.setStartCol(1);    // 指定起始列，下标从0开始计数
        sheet1.setTypes(types);    // 指定每列的类型
        excelRd.addSheets(sheet1);

        // 创建第二个 sheet 配置
        ExcelRdSheet sheet2 = new ExcelRdSheet();
        sheet2.setStartRow(2);    // 指定起始行，下标从0开始计数
        sheet2.setStartCol(1);    // 指定起始列，下标从0开始计数
        sheet2.setTypes(types);    // 指定每列的类型
        excelRd.addSheets(sheet2);   // add 多个 sheet 配置

        // sheet = excelRd.analysisFirstSheet();    // 默认只取第一个
        List<ExcelRdSheet> sheets = excelRd.analysis();

        for (ExcelRdSheet s:sheets) {
            System.out.println("-------> " + s.getSheetName());
            for (List<Object> row : s.getRows()) {
                for (Object cell : row) {
                    System.out.println(cell);
                }
                System.out.println("\n");
            }
        }


    }



}
