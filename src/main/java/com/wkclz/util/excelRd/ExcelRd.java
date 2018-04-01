package com.wkclz.util.excelRd;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRd extends ExcelRdContent {
	
	private boolean xls = false;

	private static final String DOT_XLS = ".xls";
	private static final String DOT_XLSX = ".xlsx";

	public ExcelRd(String xlsxPath) {
		super();
		this.file = new File(xlsxPath);
	}
	public ExcelRd(File file) {
		super();
		this.file = file;
	}
	
	public List<ExcelRdRow> analysisXlsx() throws ExcelRdException, IOException {
		String xlsxPath = file.getPath();
		
		if(!(xlsxPath.endsWith(DOT_XLS)||xlsxPath.endsWith(DOT_XLSX))) {
			throw new ExcelRdException("Excel can only be xlsx or xls!");
		}
		
		// 03版本的excel要特别标明
		if(xlsxPath.endsWith(DOT_XLS)) {
			xls = true;
		}
		
		if(!file.exists()) {
			throw new ExcelRdException("Excel path is not correct");
		}
		if(!file.isFile()) {
			throw new ExcelRdException("Excel path is not a file");
		}
		
		List<ExcelRdTypeEnum> types = getTypes();
		if(types==null||types.size()==0) {
			throw new ExcelRdException("Types of the data must be set");
		}
		
		is = new FileInputStream(file);
		
		if(xls) {
			workbook03 = new HSSFWorkbook(is);
		} else {
			workbook = new XSSFWorkbook(is);
		}
				
		
		// 当前只考虑识别一个 sheet
		if(xls) {
			sheet03 = workbook03.getSheetAt(0);
		} else {
			sheet = workbook.getSheetAt(0);
		}
		
		// 循环所有【右边的边界】
		int right = getStartCol() + types.size();
		int rowThreshold = 0;	// 阈值【当连续取到三个空行，或者连续取到 3 * size 个空 cell 时，将会退出检测】
		int colThreshold = 0;	// 阈值【当连续取到三个空行，或者连续取到 3 * size 个空 cell 时，将会退出检测】
		
		for (int i = getStartRow();; i++) {
			
			// 阈值【当连续取到三个空行，或者连续取到 3 * size 个空 cell 时，将会退出检测】
			if(rowThreshold>=3||colThreshold>=3*types.size()) {
				break;
			}
			
			if(xls) {
				row03 = sheet03.getRow(i);
			} else {
				row = sheet.getRow(i);
			}
			
			if(row03==null&&row==null){
				rowThreshold ++;
				continue;
			}
			rowThreshold = 0;
			
			ExcelRdRow excelRdRow = new ExcelRdRow();
			for (int j = getStartCol(); j < right; j++) {
				
				if(xls) {
					cell03 = row03.getCell(j);
				} else {
					cell = row.getCell(j);
				}
				
				if(cell03==null&&cell==null){
					colThreshold ++;
					excelRdRow.addCell("");
				} else {
					colThreshold = 0;
					Object cellValue;
					
					if(xls) {
						cellValue = ExcelRdUtil.getCellValue(cell03, types.get(j - getStartCol()));
					} else {
						cellValue = ExcelRdUtil.getCellValue(cell, types.get(j - getStartCol()));
					}
					
					excelRdRow.addCell(cellValue);
				}
			}
			
			// 如果row全部为null，将不加入结果
			List<Object> rtRow = excelRdRow.getRow();
			int size = rtRow.size();
			for (Object object : rtRow) {
				if(object==null||"".equals(object.toString().trim())) {
					size--;
				}
			}
			if(size!=0) {
				addRow(excelRdRow);
			}
		}
		return getRows();
	}
}
