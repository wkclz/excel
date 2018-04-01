package com.wkclz.util.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public abstract class ExcelStyle extends ExcelBase {

    private static final short DEFAULT_COLOR = Font.COLOR_NORMAL;
    private static final short DEFAULT_HEIGHT_IN_POINTS = 10;
    private static final String DEFAULT_FONT_NAME = "宋体";
	
	
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
			styleTitle.setAlignment(HorizontalAlignment.CENTER);
			styleTitle.setVerticalAlignment(VerticalAlignment.CENTER);
			styleTitle.setFont(createFont(getWorkbook(),24));
		}
		return styleTitle;
	}
	
	/** 列名样式【默认有边框】 */
	protected XSSFCellStyle getStyleHeader() {
		if(styleHeader==null){
			styleHeader = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleHeader.setAlignment(HorizontalAlignment.CENTER);
			styleHeader.setVerticalAlignment(VerticalAlignment.CENTER);

			//边框
			styleHeader = addBorder(styleHeader);
			styleHeader.setFont(createFont(getWorkbook(),12));
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
			styleStrLeftNoBorder.setAlignment(HorizontalAlignment.LEFT);
			styleStrLeftNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);

			styleStrLeftNoBorder.setFont(createFont(getWorkbook()));
		}
		return styleStrLeftNoBorder;
	}
	
	/** 字符串 【无边框，中间】 */
	protected XSSFCellStyle getStyleStrCenterNoBorder() {
		if(styleStrCenterNoBorder==null){
			styleStrCenterNoBorder = (XSSFCellStyle) getWorkbook().createCellStyle();

			styleStrCenterNoBorder.setAlignment(HorizontalAlignment.CENTER);
			styleStrCenterNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);

			styleStrCenterNoBorder.setFont(createFont(getWorkbook()));
		}
		return styleStrCenterNoBorder;
	}
	
	/** 字符串 【有边框，左边】 */
	protected XSSFCellStyle getStyleStrLeftWithBorder() {
		if(styleStrLeftWithBorder==null){
			styleStrLeftWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleStrLeftWithBorder.setAlignment(HorizontalAlignment.LEFT);
			styleStrLeftWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);

			//边框
			styleStrLeftWithBorder = addBorder(styleStrLeftWithBorder);
			styleStrLeftWithBorder.setFont(createFont(getWorkbook()));
		}
		return styleStrLeftWithBorder;
	}
	
	/** 字符串 【有边框，中间】 */
	protected XSSFCellStyle getStyleStrCenterWithBorder() {
		if(styleStrCenterWithBorder==null){
			styleStrCenterWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleStrCenterWithBorder.setAlignment(HorizontalAlignment.CENTER);
			styleStrCenterWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);

			//边框
			styleStrCenterWithBorder = addBorder(styleStrCenterWithBorder);
			styleStrCenterWithBorder.setFont(createFont(getWorkbook()));
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
			styleNumLeftNoBorder.setAlignment(HorizontalAlignment.LEFT);
			styleNumLeftNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleNumLeftNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("0.00"));
			styleNumLeftNoBorder.setFont(createFont(getWorkbook()));
		}
		return styleNumLeftNoBorder;
	}
	
	/** 小数 【无边框，中间】 */
	protected XSSFCellStyle getStyleNumCenterNoBorder() {
		if(styleNumCenterNoBorder==null){
			styleNumCenterNoBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleNumCenterNoBorder.setAlignment(HorizontalAlignment.CENTER);
			styleNumCenterNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleNumCenterNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("0.00"));
			styleNumCenterNoBorder.setFont(createFont(getWorkbook()));
		}
		return styleNumCenterNoBorder;
	}
	
	/** 小数 【有边框，左边】 */
	protected XSSFCellStyle getStyleNumLeftWithBorder() {
		if(styleNumLeftWithBorder==null){
			styleNumLeftWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleNumLeftWithBorder.setAlignment(HorizontalAlignment.LEFT);
			styleNumLeftWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleNumLeftWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("0.00"));

			//边框
			styleNumLeftWithBorder = addBorder(styleNumLeftWithBorder);
			styleNumLeftWithBorder.setFont(createFont(getWorkbook()));
		}
		return styleNumLeftWithBorder;
	}
	
	/** 小数 【有边框，中间】 */
	protected XSSFCellStyle getStyleNumCenterWithBorder() {
		if(styleNumCenterWithBorder==null){
			styleNumCenterWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleNumCenterWithBorder.setAlignment(HorizontalAlignment.CENTER);
			styleNumCenterWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleNumCenterWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("0.00"));

			//边框
			styleNumCenterWithBorder = addBorder(styleNumCenterWithBorder);
			styleNumCenterWithBorder.setFont(createFont(getWorkbook()));
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
			styleDateLeftNoBorder.setAlignment(HorizontalAlignment.LEFT);
			styleDateLeftNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleDateLeftNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd"));
			styleDateLeftNoBorder.setFont(createFont(getWorkbook()));
		}
		return styleDateLeftNoBorder;
	}
	
	/** 日期 【无边框，中间】 */
	protected XSSFCellStyle getStyleDateCenterNoBorder() {
		if(styleDateCenterNoBorder==null){
			styleDateCenterNoBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleDateCenterNoBorder.setAlignment(HorizontalAlignment.CENTER);
			styleDateCenterNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleDateCenterNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd"));
			styleDateCenterNoBorder.setFont(createFont(getWorkbook()));
		}
		return styleDateCenterNoBorder;
	}
	
	/** 日期 【有边框，左边】 */
	protected XSSFCellStyle getStyleDateLeftWithBorder() {
		if(styleDateLeftWithBorder==null){
			styleDateLeftWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleDateLeftWithBorder.setAlignment(HorizontalAlignment.LEFT);
			styleDateLeftWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleDateLeftWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd"));

			//边框
			styleDateLeftWithBorder = addBorder(styleDateLeftWithBorder);
			styleDateLeftWithBorder.setFont(createFont(getWorkbook()));
		}
		return styleDateLeftWithBorder;
	}
	
	/** 日期 【有边框，中间】 */
	protected XSSFCellStyle getStyleDateCenterWithBorder() {
		if(styleDateCenterWithBorder==null){
			styleDateCenterWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleDateCenterWithBorder.setAlignment(HorizontalAlignment.CENTER);
			styleDateCenterWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleDateCenterWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd"));

			//边框
			styleDateCenterWithBorder = addBorder(styleDateCenterWithBorder);
			styleDateCenterWithBorder.setFont(createFont(getWorkbook()));
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
			styleDateTimeLeftNoBorder.setAlignment(HorizontalAlignment.LEFT);
			styleDateTimeLeftNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleDateTimeLeftNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
			styleDateTimeLeftNoBorder.setFont(createFont(getWorkbook()));
		}
		return styleDateTimeLeftNoBorder;
	}
	
	/** 日期时间 【无边框，中间】 */
	protected XSSFCellStyle getStyleDateTimeCenterNoBorder() {
		if(styleDateTimeCenterNoBorder==null){
			styleDateTimeCenterNoBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleDateTimeCenterNoBorder.setAlignment(HorizontalAlignment.CENTER);
			styleDateTimeCenterNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleDateTimeCenterNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
			styleDateTimeCenterNoBorder.setFont(createFont(getWorkbook()));
		}
		return styleDateTimeCenterNoBorder;
	}
	
	/** 日期时间 【有边框，左边】 */
	protected XSSFCellStyle getStyleDateTimeLeftWithBorder() {
		if(styleDateTimeLeftWithBorder==null){
			styleDateTimeLeftWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleDateTimeLeftWithBorder.setAlignment(HorizontalAlignment.LEFT);
			styleDateTimeLeftWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleDateTimeLeftWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));

			//边框
			styleDateTimeLeftWithBorder = addBorder(styleDateTimeLeftWithBorder);
			styleDateTimeLeftWithBorder.setFont(createFont(getWorkbook()));
		}
		return styleDateTimeLeftWithBorder;
	}
	
	/** 日期时间 【有边框，中间】 */
	protected XSSFCellStyle getStyleDateTimeCenterWithBorder() {
		if(styleDateTimeCenterWithBorder==null){
			styleDateTimeCenterWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleDateTimeCenterWithBorder.setAlignment(HorizontalAlignment.CENTER);
			styleDateTimeCenterWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleDateTimeCenterWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));

			//边框
			styleDateTimeCenterWithBorder = addBorder(styleDateTimeCenterWithBorder);
			styleDateTimeCenterWithBorder.setFont(createFont(getWorkbook()));
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
			styleWrapTextLeftNoBorder.setAlignment(HorizontalAlignment.LEFT);
			styleWrapTextLeftNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleWrapTextLeftNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
			styleWrapTextLeftNoBorder.setWrapText(true);
			styleWrapTextLeftNoBorder.setFont(createFont(getWorkbook()));
		}
		return styleWrapTextLeftNoBorder;
	}
	
	/** 富文本，自动换行 【无边框，中间】 */
	protected XSSFCellStyle getStyleWrapTextCenterNoBorder() {
		if(styleWrapTextCenterNoBorder==null){
			styleWrapTextCenterNoBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleWrapTextCenterNoBorder.setAlignment(HorizontalAlignment.CENTER);
			styleWrapTextCenterNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleWrapTextCenterNoBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
			styleWrapTextCenterNoBorder.setWrapText(true);
			styleWrapTextCenterNoBorder.setFont(createFont(getWorkbook()));
		}
		return styleWrapTextCenterNoBorder;
	}
	
	/** 富文本，自动换行 【有边框，左边】 */
	protected XSSFCellStyle getStyleWrapTextLeftWithBorder() {
		if(styleWrapTextLeftWithBorder==null){
			styleWrapTextLeftWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleWrapTextLeftWithBorder.setAlignment(HorizontalAlignment.LEFT);
			styleWrapTextLeftWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleWrapTextLeftWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
			styleWrapTextLeftWithBorder.setWrapText(true);

			//边框
			styleWrapTextLeftWithBorder = addBorder(styleWrapTextLeftWithBorder);
			styleWrapTextLeftWithBorder.setFont(createFont(getWorkbook()));
		}
		return styleWrapTextLeftWithBorder;
	}
	
	/** 富文本，自动换行 【有边框，中间】 */
	protected XSSFCellStyle getStyleWrapTextCenterWithBorder() {
		if(styleWrapTextCenterWithBorder==null){
			styleWrapTextCenterWithBorder = (XSSFCellStyle) getWorkbook().createCellStyle();
			styleWrapTextCenterWithBorder.setAlignment(HorizontalAlignment.CENTER);
			styleWrapTextCenterWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
			styleWrapTextCenterWithBorder.setDataFormat(getWorkbook().createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
			styleWrapTextCenterWithBorder.setWrapText(true);
			//边框
			styleWrapTextCenterWithBorder = addBorder(styleWrapTextCenterWithBorder);

			styleWrapTextCenterWithBorder.setFont(createFont(getWorkbook()));
		}
		return styleWrapTextCenterWithBorder;
	}


	/**
	 * 添加边框
	 * @param xssfCellStyle
	 * @return
	 */
	private static XSSFCellStyle addBorder(XSSFCellStyle xssfCellStyle){
		//左边框
		xssfCellStyle.setBorderLeft(ExcelUtil.BORDER);
		//右边框
		xssfCellStyle.setBorderRight(ExcelUtil.BORDER);
		//上边框
		xssfCellStyle.setBorderTop(ExcelUtil.BORDER);
		//下边框
		xssfCellStyle.setBorderBottom(ExcelUtil.BORDER);
		return xssfCellStyle;
	}


    /**
     * 构建字体
     * @param workbook
     * @return
     */
    private static Font createFont(SXSSFWorkbook workbook){
        return createFont(workbook, 0, 0, null);
    }

    /**
     * 构建字体
     * @param workbook
     * @param heightInPoints
     * @return
     */
    private static Font createFont(SXSSFWorkbook workbook,int heightInPoints){
        return createFont(workbook, 0, heightInPoints, null);
    }

    /**
     * 构建字体
     * @param workbook
     * @param color
     * @param heightInPoints
     * @param fontName
     * @return
     */
	private static Font createFont(SXSSFWorkbook workbook, int color, int heightInPoints, String fontName){
        Font font = workbook.createFont();
        font.setColor(color == 0 ? DEFAULT_COLOR : (short) color);
        font.setFontHeightInPoints(heightInPoints == 0 ? DEFAULT_HEIGHT_IN_POINTS : (short)heightInPoints);
        font.setFontName(fontName == null ? DEFAULT_FONT_NAME : fontName);
        return font;
    }
	
}
