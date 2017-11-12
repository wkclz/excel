package com.wkclz.util.excelRd;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public abstract class ExcelRdBase {

	/** File */
	protected File file;
	/** FileInputStream*/
	protected FileInputStream is;
	/** XSSFWorkbook */
	protected XSSFWorkbook workbook;
	/** XSSFSheet */
	protected XSSFSheet sheet;
	/** XSSFRow */
	protected XSSFRow row;
	/** XSSFCell */
	protected XSSFCell cell;
	/** rownum */
	protected int rownum = 0;
	
	// 03版本 excel
	protected HSSFWorkbook workbook03;
	protected HSSFSheet sheet03;
	protected HSSFRow row03;
	protected HSSFCell cell03;
	
	
	
	
	
}
