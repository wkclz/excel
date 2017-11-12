package com.wkclz.util.excel;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public abstract class ExcelStyle extends ExcelBase {
	
	
	/*
	 * 标题和列名
	 */
	
	/** 标题样式【默认无边框】 */
	private XSSFCellStyle styleTitle;
	/** 列名样式【默认有边框】 */
	private XSSFCellStyle styleHeader;
	
	
	
	/** 标题样式【默认无边框】 */
	protected XSSFCellStyle getStyleTitle() {
		if (styleTitle==null){
			styleTitle = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleTitle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			styleTitle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 24);
			font.setFontName("宋体");
			styleTitle.setFont(font);
		}
		return styleTitle;
	}
	
	/** 列名样式【默认有边框】 */
	protected XSSFCellStyle getStyleHeader() {
		if(styleHeader==null){
			styleHeader = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleHeader.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			styleHeader.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			//边框
			styleHeader.setBorderBottom(XSSFCellStyle.BORDER_THIN); //下边框
			styleHeader.setBorderLeft(XSSFCellStyle.BORDER_THIN);//左边框
			styleHeader.setBorderTop(XSSFCellStyle.BORDER_THIN);//上边框
			styleHeader.setBorderRight(XSSFCellStyle.BORDER_THIN);//右边框
			// 字体
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 12);
			font.setFontName("宋体");
			styleHeader.setFont(font);
		}
		return styleHeader;
	}
	
	
	/*
	 * 字符串【Int也用这个】
	 */
	
	
	/** 字符串 【无边框,左边】 */
	private XSSFCellStyle styleStrLeftNoBorder;
	/** 字符串 【无边框，中间】 */
	private XSSFCellStyle styleStrCenterNoBorder;
	/** 字符串 【有边框，左边】 */
	private XSSFCellStyle styleStrLeftWithBorder;
	/** 字符串 【有边框，中间】 */
	private XSSFCellStyle styleStrCenterWithBorder;
	
	
	/** 字符串 【无边框,左边】 */
	protected XSSFCellStyle getStyleStrLeftNoBorder() {
		if(styleStrLeftNoBorder==null){
			styleStrLeftNoBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleStrLeftNoBorder.setAlignment(XSSFCellStyle.ALIGN_LEFT);
			styleStrLeftNoBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleStrLeftNoBorder.setFont(font);
		}
		return styleStrLeftNoBorder;
	}
	
	/** 字符串 【无边框，中间】 */
	protected XSSFCellStyle getStyleStrCenterNoBorder() {
		if(styleStrCenterNoBorder==null){
			styleStrCenterNoBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleStrCenterNoBorder.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			styleStrCenterNoBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleStrCenterNoBorder.setFont(font);
		}
		return styleStrCenterNoBorder;
	}
	
	/** 字符串 【有边框，左边】 */
	protected XSSFCellStyle getStyleStrLeftWithBorder() {
		if(styleStrLeftWithBorder==null){
			styleStrLeftWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleStrLeftWithBorder.setAlignment(XSSFCellStyle.ALIGN_LEFT);
			styleStrLeftWithBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			//边框
			styleStrLeftWithBorder.setBorderBottom(XSSFCellStyle.BORDER_THIN); //下边框
			styleStrLeftWithBorder.setBorderLeft(XSSFCellStyle.BORDER_THIN);//左边框
			styleStrLeftWithBorder.setBorderTop(XSSFCellStyle.BORDER_THIN);//上边框
			styleStrLeftWithBorder.setBorderRight(XSSFCellStyle.BORDER_THIN);//右边框
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleStrLeftWithBorder.setFont(font);
		}
		return styleStrLeftWithBorder;
	}
	
	/** 字符串 【有边框，中间】 */
	protected XSSFCellStyle getStyleStrCenterWithBorder() {
		if(styleStrCenterWithBorder==null){
			styleStrCenterWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleStrCenterWithBorder.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			styleStrCenterWithBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			//边框
			styleStrCenterWithBorder.setBorderBottom(XSSFCellStyle.BORDER_THIN); //下边框
			styleStrCenterWithBorder.setBorderLeft(XSSFCellStyle.BORDER_THIN);//左边框
			styleStrCenterWithBorder.setBorderTop(XSSFCellStyle.BORDER_THIN);//上边框
			styleStrCenterWithBorder.setBorderRight(XSSFCellStyle.BORDER_THIN);//右边框
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleStrCenterWithBorder.setFont(font);
		}
		return styleStrCenterWithBorder;
	}
	
	
	/*
	 * 小数
	 */
	
	
	/** 小数 【无边框,左边】 */
	private XSSFCellStyle styleNumLeftNoBorder;
	/** 小数 【无边框，中间】 */
	private XSSFCellStyle styleNumCenterNoBorder;
	/** 小数 【有边框，左边】 */
	private XSSFCellStyle styleNumLeftWithBorder;
	/** 小数 【有边框，中间】 */
	private XSSFCellStyle styleNumCenterWithBorder;
	
	
	/** 小数 【无边框,左边】 */
	protected XSSFCellStyle getStyleNumLeftNoBorder() {
		if(styleNumLeftNoBorder==null){
			styleNumLeftNoBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleNumLeftNoBorder.setAlignment(XSSFCellStyle.ALIGN_LEFT);
			styleNumLeftNoBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleNumLeftNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("0.00"));
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleNumLeftNoBorder.setFont(font);
		}
		return styleNumLeftNoBorder;
	}
	
	/** 小数 【无边框，中间】 */
	protected XSSFCellStyle getStyleNumCenterNoBorder() {
		if(styleNumCenterNoBorder==null){
			styleNumCenterNoBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleNumCenterNoBorder.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			styleNumCenterNoBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleNumCenterNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("0.00"));
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleNumCenterNoBorder.setFont(font);
		}
		return styleNumCenterNoBorder;
	}
	
	/** 小数 【有边框，左边】 */
	protected XSSFCellStyle getStyleNumLeftWithBorder() {
		if(styleNumLeftWithBorder==null){
			styleNumLeftWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleNumLeftWithBorder.setAlignment(XSSFCellStyle.ALIGN_LEFT);
			styleNumLeftWithBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleNumLeftWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("0.00"));
			//边框
			styleNumLeftWithBorder.setBorderBottom(XSSFCellStyle.BORDER_THIN); //下边框
			styleNumLeftWithBorder.setBorderLeft(XSSFCellStyle.BORDER_THIN);//左边框
			styleNumLeftWithBorder.setBorderTop(XSSFCellStyle.BORDER_THIN);//上边框
			styleNumLeftWithBorder.setBorderRight(XSSFCellStyle.BORDER_THIN);//右边框
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleNumLeftWithBorder.setFont(font);
		}
		return styleNumLeftWithBorder;
	}
	
	/** 小数 【有边框，中间】 */
	protected XSSFCellStyle getStyleNumCenterWithBorder() {
		if(styleNumCenterWithBorder==null){
			styleNumCenterWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleNumCenterWithBorder.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			styleNumCenterWithBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleNumCenterWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("0.00"));
			//边框
			styleNumCenterWithBorder.setBorderBottom(XSSFCellStyle.BORDER_THIN); //下边框
			styleNumCenterWithBorder.setBorderLeft(XSSFCellStyle.BORDER_THIN);//左边框
			styleNumCenterWithBorder.setBorderTop(XSSFCellStyle.BORDER_THIN);//上边框
			styleNumCenterWithBorder.setBorderRight(XSSFCellStyle.BORDER_THIN);//右边框
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleNumCenterWithBorder.setFont(font);
		}
		return styleNumCenterWithBorder;
	}
	
	
	/*
	 * 日期
	 */
	
	
	/** 日期 【无边框,左边】 */
	private XSSFCellStyle styleDateLeftNoBorder;
	/** 日期 【无边框，中间】 */
	private XSSFCellStyle styleDateCenterNoBorder;
	/** 日期 【有边框，左边】 */
	private XSSFCellStyle styleDateLeftWithBorder;
	/** 日期 【有边框，中间】 */
	private XSSFCellStyle styleDateCenterWithBorder;
	
	
	/** 日期 【无边框,左边】 */
	protected XSSFCellStyle getStyleDateLeftNoBorder() {
		if(styleDateLeftNoBorder==null){
			styleDateLeftNoBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleDateLeftNoBorder.setAlignment(XSSFCellStyle.ALIGN_LEFT);
			styleDateLeftNoBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleDateLeftNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd"));
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleDateLeftNoBorder.setFont(font);
		}
		return styleDateLeftNoBorder;
	}
	
	/** 日期 【无边框，中间】 */
	protected XSSFCellStyle getStyleDateCenterNoBorder() {
		if(styleDateCenterNoBorder==null){
			styleDateCenterNoBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleDateCenterNoBorder.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			styleDateCenterNoBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleDateCenterNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd"));
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleDateCenterNoBorder.setFont(font);
		}
		return styleDateCenterNoBorder;
	}
	
	/** 日期 【有边框，左边】 */
	protected XSSFCellStyle getStyleDateLeftWithBorder() {
		if(styleDateLeftWithBorder==null){
			styleDateLeftWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleDateLeftWithBorder.setAlignment(XSSFCellStyle.ALIGN_LEFT);
			styleDateLeftWithBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleDateLeftWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd"));
			//边框
			styleDateLeftWithBorder.setBorderBottom(XSSFCellStyle.BORDER_THIN); //下边框
			styleDateLeftWithBorder.setBorderLeft(XSSFCellStyle.BORDER_THIN);//左边框
			styleDateLeftWithBorder.setBorderTop(XSSFCellStyle.BORDER_THIN);//上边框
			styleDateLeftWithBorder.setBorderRight(XSSFCellStyle.BORDER_THIN);//右边框
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleDateLeftWithBorder.setFont(font);
		}
		return styleDateLeftWithBorder;
	}
	
	/** 日期 【有边框，中间】 */
	protected XSSFCellStyle getStyleDateCenterWithBorder() {
		if(styleDateCenterWithBorder==null){
			styleDateCenterWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleDateCenterWithBorder.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			styleDateCenterWithBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleDateCenterWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd"));
			//边框
			styleDateCenterWithBorder.setBorderBottom(XSSFCellStyle.BORDER_THIN); //下边框
			styleDateCenterWithBorder.setBorderLeft(XSSFCellStyle.BORDER_THIN);//左边框
			styleDateCenterWithBorder.setBorderTop(XSSFCellStyle.BORDER_THIN);//上边框
			styleDateCenterWithBorder.setBorderRight(XSSFCellStyle.BORDER_THIN);//右边框
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleDateCenterWithBorder.setFont(font);
		}
		return styleDateCenterWithBorder;
	}
	
	
	/*
	 * 日期时间
	 */
	
	
	/** 日期时间 【无边框,左边】 */
	private XSSFCellStyle styleDateTimeLeftNoBorder;
	/** 日期时间 【无边框，中间】 */
	private XSSFCellStyle styleDateTimeCenterNoBorder;
	/** 日期时间 【有边框，左边】 */
	private XSSFCellStyle styleDateTimeLeftWithBorder;
	/** 日期时间 【有边框，中间】 */
	private XSSFCellStyle styleDateTimeCenterWithBorder;
	
	
	/** 日期时间 【无边框,左边】 */
	protected XSSFCellStyle getStyleDateTimeLeftNoBorder() {
		if(styleDateTimeLeftNoBorder==null){
			styleDateTimeLeftNoBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleDateTimeLeftNoBorder.setAlignment(XSSFCellStyle.ALIGN_LEFT);
			styleDateTimeLeftNoBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleDateTimeLeftNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleDateTimeLeftNoBorder.setFont(font);
		}
		return styleDateTimeLeftNoBorder;
	}
	
	/** 日期时间 【无边框，中间】 */
	protected XSSFCellStyle getStyleDateTimeCenterNoBorder() {
		if(styleDateTimeCenterNoBorder==null){
			styleDateTimeCenterNoBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleDateTimeCenterNoBorder.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			styleDateTimeCenterNoBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleDateTimeCenterNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleDateTimeCenterNoBorder.setFont(font);
		}
		return styleDateTimeCenterNoBorder;
	}
	
	/** 日期时间 【有边框，左边】 */
	protected XSSFCellStyle getStyleDateTimeLeftWithBorder() {
		if(styleDateTimeLeftWithBorder==null){
			styleDateTimeLeftWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleDateTimeLeftWithBorder.setAlignment(XSSFCellStyle.ALIGN_LEFT);
			styleDateTimeLeftWithBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleDateTimeLeftWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
			//边框
			styleDateTimeLeftWithBorder.setBorderBottom(XSSFCellStyle.BORDER_THIN); //下边框
			styleDateTimeLeftWithBorder.setBorderLeft(XSSFCellStyle.BORDER_THIN);//左边框
			styleDateTimeLeftWithBorder.setBorderTop(XSSFCellStyle.BORDER_THIN);//上边框
			styleDateTimeLeftWithBorder.setBorderRight(XSSFCellStyle.BORDER_THIN);//右边框
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleDateTimeLeftWithBorder.setFont(font);
		}
		return styleDateTimeLeftWithBorder;
	}
	
	/** 日期时间 【有边框，中间】 */
	protected XSSFCellStyle getStyleDateTimeCenterWithBorder() {
		if(styleDateTimeCenterWithBorder==null){
			styleDateTimeCenterWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleDateTimeCenterWithBorder.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			styleDateTimeCenterWithBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleDateTimeCenterWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
			//边框
			styleDateTimeCenterWithBorder.setBorderBottom(XSSFCellStyle.BORDER_THIN); //下边框
			styleDateTimeCenterWithBorder.setBorderLeft(XSSFCellStyle.BORDER_THIN);//左边框
			styleDateTimeCenterWithBorder.setBorderTop(XSSFCellStyle.BORDER_THIN);//上边框
			styleDateTimeCenterWithBorder.setBorderRight(XSSFCellStyle.BORDER_THIN);//右边框
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleDateTimeCenterWithBorder.setFont(font);
		}
		return styleDateTimeCenterWithBorder;
	}
	
	
	/*
	 * 富文本，自动换行
	 */
	
	
	/** 富文本，自动换行 【无边框,左边】 */
	private XSSFCellStyle styleWrapTextLeftNoBorder;
	/** 富文本，自动换行 【无边框，中间】 */
	private XSSFCellStyle styleWrapTextCenterNoBorder;
	/** 富文本，自动换行 【有边框，左边】 */
	private XSSFCellStyle styleWrapTextLeftWithBorder;
	/** 富文本，自动换行 【有边框，中间】 */
	private XSSFCellStyle styleWrapTextCenterWithBorder;
	
	
	/** 富文本，自动换行 【无边框,左边】 */
	protected XSSFCellStyle getStyleWrapTextLeftNoBorder() {
		if(styleWrapTextLeftNoBorder==null){
			styleWrapTextLeftNoBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleWrapTextLeftNoBorder.setAlignment(XSSFCellStyle.ALIGN_LEFT);
			styleWrapTextLeftNoBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleWrapTextLeftNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
			styleWrapTextLeftNoBorder.setWrapText(true);
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleWrapTextLeftNoBorder.setFont(font);
		}
		return styleWrapTextLeftNoBorder;
	}
	
	/** 富文本，自动换行 【无边框，中间】 */
	protected XSSFCellStyle getStyleWrapTextCenterNoBorder() {
		if(styleWrapTextCenterNoBorder==null){
			styleWrapTextCenterNoBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleWrapTextCenterNoBorder.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			styleWrapTextCenterNoBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleWrapTextCenterNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
			styleWrapTextCenterNoBorder.setWrapText(true);
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleWrapTextCenterNoBorder.setFont(font);
		}
		return styleWrapTextCenterNoBorder;
	}
	
	/** 富文本，自动换行 【有边框，左边】 */
	protected XSSFCellStyle getStyleWrapTextLeftWithBorder() {
		if(styleWrapTextLeftWithBorder==null){
			styleWrapTextLeftWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleWrapTextLeftWithBorder.setAlignment(XSSFCellStyle.ALIGN_LEFT);
			styleWrapTextLeftWithBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleWrapTextLeftWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
			styleWrapTextLeftWithBorder.setWrapText(true);
			//边框
			styleWrapTextLeftWithBorder.setBorderBottom(XSSFCellStyle.BORDER_THIN); //下边框
			styleWrapTextLeftWithBorder.setBorderLeft(XSSFCellStyle.BORDER_THIN);//左边框
			styleWrapTextLeftWithBorder.setBorderTop(XSSFCellStyle.BORDER_THIN);//上边框
			styleWrapTextLeftWithBorder.setBorderRight(XSSFCellStyle.BORDER_THIN);//右边框
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleWrapTextLeftWithBorder.setFont(font);
		}
		return styleWrapTextLeftWithBorder;
	}
	
	/** 富文本，自动换行 【有边框，中间】 */
	protected XSSFCellStyle getStyleWrapTextCenterWithBorder() {
		if(styleWrapTextCenterWithBorder==null){
			styleWrapTextCenterWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleWrapTextCenterWithBorder.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			styleWrapTextCenterWithBorder.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			styleWrapTextCenterWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
			styleWrapTextCenterWithBorder.setWrapText(true);
			//边框
			styleWrapTextCenterWithBorder.setBorderBottom(XSSFCellStyle.BORDER_THIN); //下边框
			styleWrapTextCenterWithBorder.setBorderLeft(XSSFCellStyle.BORDER_THIN);//左边框
			styleWrapTextCenterWithBorder.setBorderTop(XSSFCellStyle.BORDER_THIN);//上边框
			styleWrapTextCenterWithBorder.setBorderRight(XSSFCellStyle.BORDER_THIN);//右边框
			Font font = getWorkbook().createFont();
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 10);
			font.setFontName("宋体");
			styleWrapTextCenterWithBorder.setFont(font);
		}
		return styleWrapTextCenterWithBorder;
	}
	
}
