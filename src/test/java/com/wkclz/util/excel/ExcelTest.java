package com.wkclz.util.excel;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelTest {
	
	private static SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

	public static void main(String[] args){

	    String savePath = "/Users/wangkaicun/Desktop/test.xlsx";
		Excel excel = excel(savePath);


		// 保存到指定的目录
        try {
            excel.CreateXlsx();
        } catch (ExcelException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("生成excel到指定目录完成：" + sdf.format(new Date()));

        /*
        // 保存到临时文件并输出 File
        try {
            File file = excel.CreateXlsxByFile();
            System.out.println(file.getPath());
        } catch (ExcelException e) {
            e.printStackTrace();
        }
        */

        System.out.println("生成excel到临时文件完成：" + sdf.format(new Date()));


	}

    public static Excel excel(String savePath) {

        System.out.println("数据准备：" + sdf.format(new Date()));

        Excel excel = new Excel();
        excel.setTitle("标题");
        excel.setCreateBy("虾米");
        excel.setDateFrom("2017-07-01");
        excel.setDateTo("2017-07-12");
        excel.setSavePath(savePath);
        String[] header = {"序号","日期","时间","数字","row合并","col合并1","col合并2","超长文字自动换行"};
        excel.setHeader(header);

        for (int i = 0; i < 120000; i++) {
            ExcelRow row = excel.createRow();
            row.addCell(i+1);			                    // 序号
            row.addCell(new java.sql.Date(new Date().getTime()));	// 日期
            row.addCell(new Date());	                            // 时间
            row.addCell(12.1222);		                    // 数字
            if(i%3==0)					                            // row合并
                row.addCell("row合并",1,3);
            row.addCell("col合并",2,1);	        // col合并
            //超长文字自动换行
            row.addCell("超长文字自动换行，靠左边，超长文字自动换行，靠左边，超长文字自动换行，超长文字自动换行，靠左边，超长文字自动换行，靠左边，超长文字自动换行，靠左边，超长文字自动换行，靠左边",ExcelUtil.ALIGN_LEFT);
        }

        System.out.println("数据准备完成，准备生成excel：" + sdf.format(new Date()));

        return excel;
    }

}
