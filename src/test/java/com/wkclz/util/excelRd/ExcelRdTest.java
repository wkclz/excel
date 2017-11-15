package com.wkclz.util.excelRd;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;

public class ExcelRdTest {

	public static void main(String[] args) throws IOException, ExcelReadException {
		excelRd();
	}

	public static void excelRd() throws IOException, ExcelReadException {

		String path = "/Users/wangkaicun/Desktop/test.xlsx";
		ExcelRd excelRd = new ExcelRd(path);
		excelRd.setStartRow(1);	// 指定起始行，从0开始
		excelRd.setStartCol(1);	// 指定起始列，从0开始
        ExcelRdTypeEnum[] types = {
			ExcelRdTypeEnum.INTEGER,
            ExcelRdTypeEnum.DOUBLE,
            ExcelRdTypeEnum.DATETIME,
            ExcelRdTypeEnum.DATE,
            ExcelRdTypeEnum.STRING
		};
		excelRd.setTypes(types);	// 指定每列的类型
		
		List<ExcelRdRow> rows = excelRd.analysisXlsx();
		HashMap<String, Object>[] plans = new HashMap[rows.size()];

		int size = rows.size();
		for (int i = 0; i < size; i++) {
			
			ExcelRdRow excelRdRow = rows.get(i);
			List<Object> row = excelRdRow.getRow();
			HashMap<String, Object> plan = new HashMap<String, Object>();
			
			Object t1 = row.get(0);
			Object t2 = row.get(1);
			Object t3 = row.get(2);
			Object t4 = row.get(3);
			Object t5 = row.get(4);
			
			System.out.println(t1);
			System.out.println(t2);
			System.out.println(t3);
			System.out.println(t4);
			System.out.println(t5);
			
			plans[i] = plan;
		}

	}

}
