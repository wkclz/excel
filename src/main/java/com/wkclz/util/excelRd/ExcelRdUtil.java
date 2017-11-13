package com.wkclz.util.excelRd;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;


public class ExcelRdUtil {

	public static final String INTEGER = "INTEGER";
	public static final String DOUBLE = "DOUBLE";
	public static final String DATE = "DATE";
	public static final String DATETIME = "DATETIME";
	public static final String STRING = "STRING";

	protected static Object getCellValue(XSSFCell cell, String type){
		if(cell==null||"".equals(cell.toString().trim())) return null;
		CellType cellType = cell.getCellTypeEnum();
		if(INTEGER.equals(type) && cellType == CellType.NUMERIC){
			Double numeric = cell.getNumericCellValue();
			return numeric.intValue();
		}
		if(DOUBLE.equals(type) && cellType == CellType.NUMERIC){
			return cell.getNumericCellValue();
		}
		if(DATETIME.equals(type) && cellType == CellType.NUMERIC) {
			return cell.getDateCellValue();
		}
		if(DATE.equals(type) && cellType == CellType.NUMERIC){
			return new java.sql.Date(cell.getDateCellValue().getTime());
		}
		if(STRING.equals(type) && cellType == CellType.STRING){
			return cell.getStringCellValue();
		}
		return cell.toString();
	}
	
	protected static Object getCellValue(HSSFCell cell, String type){
		if(cell==null||"".equals(cell.toString().trim())) return null;
		CellType cellType = cell.getCellTypeEnum();
		if(INTEGER.equals(type) && cellType == CellType.NUMERIC){
			Double numeric = cell.getNumericCellValue();
			return numeric.intValue();
		}
		if(DOUBLE.equals(type) && cellType == CellType.NUMERIC){
			return cell.getNumericCellValue();
		}
		if(DATETIME.equals(type) && cellType == CellType.NUMERIC){
			return cell.getDateCellValue();
		}
		if(DATE.equals(type) && cellType == CellType.NUMERIC){
			return new java.sql.Date(cell.getDateCellValue().getTime());
		}
		if(STRING.equals(type) && cellType == CellType.STRING){
			return cell.getStringCellValue();
		}
		return cell.getStringCellValue();
	}
	

}
