package com.wkclz.util.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class ExcelUtil {
	
	/** 左对齐 */
	protected static final HorizontalAlignment ALIGN_LEFT = HorizontalAlignment.LEFT;
	/** 中对齐 */
    protected static final HorizontalAlignment ALIGN_CENTER = HorizontalAlignment.CENTER;
	/** 边框大小 */
    protected static final BorderStyle BORDER = BorderStyle.THIN;
	/** 写 Excel 时的缓存数 */
	protected static final Integer CACHE_ROWS_IN_MEMORY = 10240;


    /**
     * 默认颜色
     */
    private static final short DEFAULT_COLOR = Font.COLOR_NORMAL;
    /**
     * 默认高度
     */
    private static final short DEFAULT_HEIGHT_IN_POINTS = 10;
    /**
     * 默认字体
     */
    private static final String DEFAULT_FONT_NAME = "宋体";


	/**
	 * 年月日，时分秒
	 */
    protected static final SimpleDateFormat SDF_DATE_TIME = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	/**
	 * 年月日
	 */
    protected static final SimpleDateFormat SDF_DATE = new SimpleDateFormat("yyyy-MM-dd");



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
    protected static boolean setWidth(SXSSFSheet sheel, int col, String s){
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





    /**
     * 构建字体
     * @param excel
     * @return
     */
    protected static Font createFont(Excel excel){
        return createFont(excel, 0, 0, null);
    }

    /**
     * 构建字体
     * @param excel
     * @param heightInPoints
     * @return
     */
    protected static Font createFont(Excel excel,int heightInPoints){
        return createFont(excel, 0, heightInPoints, null);
    }

    /**
     * 构建字体
     * @param excel
     * @param color
     * @param heightInPoints
     * @param fontName
     * @return
     */
    protected static Font createFont(Excel excel, int color, int heightInPoints, String fontName){
        SXSSFWorkbook workbook = excel.getWorkbook();
        String key = workbook.toString()+"_"+color+"_"+heightInPoints+"_"+fontName;
        Map<String, Font> fonts = excel.getWorkBookFonts();

        if (fonts == null){
            fonts = new HashMap<String, Font>();
        }
        Font font= fonts.get(key);
        if (font != null){
            return font;
        }

        font = workbook.createFont();
        font.setColor(color == 0 ? DEFAULT_COLOR : (short) color);
        font.setFontHeightInPoints(heightInPoints == 0 ? DEFAULT_HEIGHT_IN_POINTS : (short)heightInPoints);
        font.setFontName(fontName == null ? DEFAULT_FONT_NAME : fontName);

        // 缓存
        fonts.put(key,font);
        excel.setWorkBookFonts(fonts);

        return font;
    }


    protected static void setIntStrStyle(Excel excel, SXSSFCell cell, HorizontalAlignment align, boolean border){
        ExcelStyle style = excel.getStyle();
        cell.setCellStyle(style.getStyleStrCenterWithBorder(excel));
        // 边框 + 左边
        if(border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleStrLeftWithBorder(excel));
        }
        // 无边框 + 左边
        if(!border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleStrLeftNoBorder(excel));
        }
        // 无边框 + 中间
        if(!border && HorizontalAlignment.CENTER == align) {
            cell.setCellStyle(style.getStyleStrCenterNoBorder(excel));
        }
    }
    protected static void setDoubleStyle(Excel excel, SXSSFCell cell,HorizontalAlignment align, boolean border){
        ExcelStyle style = excel.getStyle();
        cell.setCellStyle(style.getStyleNumCenterWithBorder(excel));
        // 边框 + 左边
        if(border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleNumLeftWithBorder(excel));
        }
        // 无边框 + 左边
        if(!border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleNumLeftNoBorder(excel));
        }
        // 无边框 + 中间
        if(!border && HorizontalAlignment.CENTER == align) {
            cell.setCellStyle(style.getStyleNumCenterNoBorder(excel));
        }
    }
    protected static void setDateStyle(Excel excel, SXSSFCell cell,HorizontalAlignment align, boolean border){
        ExcelStyle style = excel.getStyle();
        cell.setCellStyle(style.getStyleDateCenterWithBorder(excel));
        // 边框 + 左边
        if(border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleDateLeftWithBorder(excel));
        }
        // 无边框 + 左边
        if(!border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleDateLeftNoBorder(excel));
        }
        // 无边框 + 中间
        if(!border && HorizontalAlignment.CENTER == align) {
            cell.setCellStyle(style.getStyleDateCenterNoBorder(excel));
        }
    }
    protected static void setDateTimeStyle(Excel excel, SXSSFCell cell,HorizontalAlignment align, boolean border){
        ExcelStyle style = excel.getStyle();
        cell.setCellStyle(style.getStyleDateTimeCenterWithBorder(excel));
        // 边框 + 左边
        if(border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleDateTimeLeftWithBorder(excel));
        }
        // 无边框 + 左边
        if(!border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleDateTimeLeftNoBorder(excel));
        }
        // 无边框 + 中间
        if(!border && HorizontalAlignment.CENTER == align) {
            cell.setCellStyle(style.getStyleDateTimeCenterNoBorder(excel));
        }
    }
    protected static void setWrapTextStyle(Excel excel, SXSSFCell cell,HorizontalAlignment align, boolean border){
        ExcelStyle style = excel.getStyle();
        cell.setCellStyle(style.getStyleWrapTextCenterWithBorder(excel));
        // 边框 + 左边
        if(border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleWrapTextLeftWithBorder(excel));
        }
        // 无边框 + 左边
        if(!border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleWrapTextLeftNoBorder(excel));
        }
        // 无边框 + 中间
        if(!border && HorizontalAlignment.CENTER == align) {
            cell.setCellStyle(style.getStyleWrapTextCenterNoBorder(excel));
        }
    }


    /*
    private static void main(String[] args) {
		System.out.println(62*12*25/256);
	}
    */

}
