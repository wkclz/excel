package com.wkclz.util.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.Date;

public class ExcelUtil {
	
	/** 左对齐 */
	public static final HorizontalAlignment ALIGN_LEFT = HorizontalAlignment.LEFT;
	/** 中对齐 */
	public static final HorizontalAlignment ALIGN_CENTER = HorizontalAlignment.CENTER;
	/** 边框大小 */
	public static final BorderStyle BORDER = BorderStyle.THIN;
	/** 写 Excel 时的缓存数 */
	protected static final Integer CACHE_ROWS_IN_MEMORY = 10240;

	/**
	 * 判断字符是否为单字节字符
	 * @param c
	 * @return
	 */
	private static boolean isLetter(char c) {
		int k = 0x80;
		return c / k == 0 ? true : false;
	}

	/**
	 * 判断长度
	 * @param s
	 * @return
	 */
	private static int length(String s) {
		if (s == null) {
			return 0;
		}
		char[] c = s.toCharArray();
		int len = 0;
		for (int i = 0; i < c.length; i++) {
			len++;
			if (!isLetter(c[i])) {
				len++;
			}
		}
		return len;
	}

	// 判断宽度
	private static int width(String s){
		return (length(s)+2)*256;
	}
	
	/**
	* @Title:
	* @Description: 自动设置shell 的 col 宽度。如果超过宽度，返回true.
	* @param @param sheel
	* @param @param col
	* @param @param s
	* @param @return    设定文件 
	* @author wangkc admin@wkclz.com  
	* @date 2017年7月15日 下午9:10:16 *  
	* @throws
	 */
	public static boolean setWidth(SXSSFSheet sheel, int col, String s){
		int width = ExcelUtil.width(s);
		width = width > 72*256 ? 72*256 : width;
		if(width>sheel.getColumnWidth(col)) {
			sheel.setColumnWidth(col, width);
		}
		return width == 72*256;
	}
	
	
	/**
	 * java.util.Date 转换为 java.sql.Date
	* @Title:  
	* @Description: 使用标准时间换取年月日的时间
	* @param @param date
	* @param @return    设定文件 
	* @author wangkc admin@wkclz.com  
	* @date 2017年7月14日 下午9:18:09 *  
	* @throws
	 */
	public static java.sql.Date getSqlDate(Date date){
		if(date==null) {
			return null;
		}
		return new java.sql.Date(date.getTime());
	}
	
	
	private static void main(String[] args) {
		System.out.println(62*12*25/256);
	}




	
}
