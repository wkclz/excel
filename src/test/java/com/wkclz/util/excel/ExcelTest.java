package com.wkclz.util.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.math.BigDecimal;
import java.util.Date;

public class ExcelTest {

    private static final Logger logger = LoggerFactory.getLogger(ExcelTest.class);

    public static void main(String[] args) {
        String savePath = "/Users/wangkaicun/Desktop/test.xlsx";
        Excel excel = excel(savePath);

        // 保存到指定的目录
        try {
            excel.createXlsx();
        } catch (ExcelException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        logger.info("========> 生成excel到指定目录完成: {}", ExcelUtil.SDF_DATE_TIME.format(new Date()));
        /*
        // 保存到临时文件并输出 File
        try {
            File file = excel.createXlsxByFile();
            System.out.println(file.getPath());
        } catch (ExcelException e) {
            e.printStackTrace();
        }
        */
        // System.out.println("生成excel到临时文件完成：" + sdf.format(new Date()));
    }

    public static Excel excel(String savePath) {
        logger.info("========> 数据准备: {}", ExcelUtil.SDF_DATE_TIME.format(new Date()));

        Excel excel = new Excel();
        excel.setTitle("标题");
        excel.setCreateBy("虾米");
        excel.setDateFrom("2017-07-01");
        excel.setDateTo("2017-07-12");
        excel.setSavePath(savePath);
        String[] header = {"序号", "日期", "时间", "数字", "金钱", "row合并", "col合并1", "col合并2", "超长文字自动换行"};
        excel.setHeader(header);

        /*
        Map<String, Object> params = new HashMap<String, Object>();
        params.put("titleOff", true);
        params.put("createInfoOff", true);
        excel.setParams(params);
        */

        for (int i = 0; i < 120; i++) {

            // 多 Sheet，每 30 条数据一个Sheet 【注意，Sheet 分离时，不能有row合并，否则排版会异常】
            if (i > 1 && i % 30 == 0) {
                excel.addNewSheet();
            }
            ExcelRow row = excel.createRow();
            row.addCell(i + 1);                              // 序号
            row.addCell(new java.sql.Date(new Date().getTime()));       // 日期
            row.addCell(new Date());                                    // 时间
            row.addCell(12.1222);                            // 数字
            row.addCell(new BigDecimal("12.34"));                  // 金钱
            if (i % 3 == 0) {                                           // row合并
                row.addCell("row合并", 1, 3);
            }
            row.addCell("col合并", 2, 1);      // col合并
            //超长文字自动换行
            row.addCell("超长文字自动换行，靠左边，超长文字自动换行，靠左边，超长文字自动换行，超长文字自动换行，靠左边，超长文字自动换行，靠左边，超长文字自动换行，靠左边，超长文字自动换行，靠左边", ExcelUtil.ALIGN_LEFT);
        }
        logger.info("========> 数据准备完成，准备生成excel: {}", ExcelUtil.SDF_DATE_TIME.format(new Date()));
        return excel;
    }
}
