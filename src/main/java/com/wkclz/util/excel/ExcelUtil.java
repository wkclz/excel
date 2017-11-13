package com.wkclz.util.excel;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.Date;

public class ExcelUtil {
	
	/** 左对齐 */
	public static final HorizontalAlignment ALIGN_LEFT = HorizontalAlignment.LEFT;
	/** 中对齐 */
	public static final HorizontalAlignment ALIGN_CENTER = HorizontalAlignment.CENTER;

	private static boolean isLetter(char c) {
		int k = 0x80;
		return c / k == 0 ? true : false;
	}
	
	private static int length(String s) {
		if (s == null)
			return 0;
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
	
	private static int width(String s){
		return (length(s)+2)*256;
	}
	
	/**
	 * 自动设置shell 的 col 宽度。如果超过宽度，返回true.
	* @Title:  
	* @Description: TODO(这里用一句话描述这个方法的作用) 
	* @param @param sheel
	* @param @param col
	* @param @param s
	* @param @return    设定文件 
	* @author wangkc admin@wkclz.com  
	* @date 2017年7月15日 下午9:10:16 *  
	* @throws
	 */
	protected static boolean setWidth(SXSSFSheet sheel, int col, String s){
		int width = ExcelUtil.width(s);
		width = width > 72*256 ? 72*256 : width;
		if(width>sheel.getColumnWidth(col))
			sheel.setColumnWidth(col, width);
		return width == 72*256;
	}
	
	
	/**
	 * java.util.Date 转换为 java.sql.Date
	* @Title:  
	* @Description: TODO(这里用一句话描述这个方法的作用) 
	* @param @param date
	* @param @return    设定文件 
	* @author wangkc admin@wkclz.com  
	* @date 2017年7月14日 下午9:18:09 *  
	* @throws
	 */
	public static java.sql.Date getSqlDate(Date date){
		if(date==null) return null;
		return new java.sql.Date(date.getTime());
	}
	
	
	@SuppressWarnings("unused")
	private static void main(String[] args) {
		System.out.println(62*12*25/256);
	}
	
}
