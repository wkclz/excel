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

    /**
     * 左对齐
     */
    public static final HorizontalAlignment ALIGN_LEFT = HorizontalAlignment.LEFT;
    /**
     * 中对齐
     */
    public static final HorizontalAlignment ALIGN_CENTER = HorizontalAlignment.CENTER;
    /**
     * 右对齐
     */
    public static final HorizontalAlignment ALIGN_RIGHT = HorizontalAlignment.RIGHT;
    /**
     * 边框大小
     */
    public static final BorderStyle BORDER = BorderStyle.THIN;


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


    private static final int DEFAULT_WIDTH = 72 * 256;


    /**
     * Description 判断字符是否为单字节字符
     *
     * @param c c
     * @return int
     */
    private static boolean isLetter(char c) {
        int k = 0x80;
        return c / k == 0 ? true : false;
    }

    /**
     * Description 判断长度
     *
     * @param s s
     * @return int
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
    private static int width(String s) {
        return (length(s) + 2) * 256;
    }

    /**
     * Description 自动设置shell的col宽度。如果超过宽度，返回真
     * create @ 2017-07-15 21:10:16
     *
     * @param sheel sheel
     * @param col   col
     * @param s     s
     * @return boolean
     * @author wangkc admin@wkclz.com
     */
    protected static boolean setWidth(SXSSFSheet sheel, int col, String s) {
        int width = ExcelUtil.width(s);
        width = width > DEFAULT_WIDTH ? DEFAULT_WIDTH : width;
        if (width > sheel.getColumnWidth(col)) {
            sheel.setColumnWidth(col, width);
        }
        return width == DEFAULT_WIDTH;
    }


    /**
     * Description 使用标准时间换取年月日的时间
     * create @ 2017-07-14 21:18:09
     *
     * @param date date
     * @author wangkc admin@wkclz.com
     */
    public static java.sql.Date getSqlDate(Date date) {
        if (date == null) {
            return null;
        }
        return new java.sql.Date(date.getTime());
    }


    /**
     * Description 构建字体
     *
     * @param excel excel
     * @return Font
     */
    protected static Font createFont(Excel excel) {
        return createFont(excel, 0, 0, null);
    }

    /**
     * Description 构建字体
     *
     * @param excel          excel
     * @param heightInPoints heightInPoints
     * @return Font
     */
    protected static Font createFont(Excel excel, int heightInPoints) {
        return createFont(excel, 0, heightInPoints, null);
    }

    /**
     * Description 构建字体
     *
     * @param excel          excel
     * @param color          color
     * @param heightInPoints heightInPoints
     * @param fontName       fontName
     * @return Font
     */
    protected static Font createFont(Excel excel, int color, int heightInPoints, String fontName) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        String key = workbook.toString() + "_" + color + "_" + heightInPoints + "_" + fontName;
        Map<String, Font> fonts = excel.getWorkBookFonts();

        if (fonts == null) {
            fonts = new HashMap<String, Font>();
        }
        Font font = fonts.get(key);
        if (font != null) {
            return font;
        }

        font = workbook.createFont();
        font.setColor(color == 0 ? DEFAULT_COLOR : (short) color);
        font.setFontHeightInPoints(heightInPoints == 0 ? DEFAULT_HEIGHT_IN_POINTS : (short) heightInPoints);
        font.setFontName(fontName == null ? DEFAULT_FONT_NAME : fontName);

        // 缓存
        fonts.put(key, font);
        excel.setWorkBookFonts(fonts);

        return font;
    }


    /**
     * @param excel  excel
     * @param cell   cell
     * @param align  align
     * @param border border
     */
    protected static void setIntStrStyle(Excel excel, SXSSFCell cell, HorizontalAlignment align, boolean border) {
        ExcelStyle style = excel.getStyle();
        cell.setCellStyle(style.getStyleStrCenterWithBorder(excel));
        // 边框 + 左边
        if (border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleStrLeftWithBorder(excel));
        }
        // 无边框 + 左边
        if (!border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleStrLeftNoBorder(excel));
        }
        // 无边框 + 中间
        if (!border && HorizontalAlignment.CENTER == align) {
            cell.setCellStyle(style.getStyleStrCenterNoBorder(excel));
        }
    }

    /**
     * @param excel  excel
     * @param cell   cell
     * @param align  align
     * @param border border
     */
    protected static void setDoubleStyle(Excel excel, SXSSFCell cell, HorizontalAlignment align, boolean border) {
        ExcelStyle style = excel.getStyle();
        cell.setCellStyle(style.getStyleNumCenterWithBorder(excel));
        // 边框 + 左边
        if (border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleNumLeftWithBorder(excel));
        }
        // 无边框 + 左边
        if (!border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleNumLeftNoBorder(excel));
        }
        // 无边框 + 中间
        if (!border && HorizontalAlignment.CENTER == align) {
            cell.setCellStyle(style.getStyleNumCenterNoBorder(excel));
        }
    }

    /**
     * @param excel  excel
     * @param cell   cell
     * @param align  align
     * @param border border
     */
    protected static void setDateStyle(Excel excel, SXSSFCell cell, HorizontalAlignment align, boolean border) {
        ExcelStyle style = excel.getStyle();
        cell.setCellStyle(style.getStyleDateCenterWithBorder(excel));
        // 边框 + 左边
        if (border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleDateLeftWithBorder(excel));
        }
        // 无边框 + 左边
        if (!border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleDateLeftNoBorder(excel));
        }
        // 无边框 + 中间
        if (!border && HorizontalAlignment.CENTER == align) {
            cell.setCellStyle(style.getStyleDateCenterNoBorder(excel));
        }
    }

    /**
     * @param excel  excel
     * @param cell   cell
     * @param align  align
     * @param border border
     */
    protected static void setDateTimeStyle(Excel excel, SXSSFCell cell, HorizontalAlignment align, boolean border) {
        ExcelStyle style = excel.getStyle();
        cell.setCellStyle(style.getStyleDateTimeCenterWithBorder(excel));
        // 边框 + 左边
        if (border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleDateTimeLeftWithBorder(excel));
        }
        // 无边框 + 左边
        if (!border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleDateTimeLeftNoBorder(excel));
        }
        // 无边框 + 中间
        if (!border && HorizontalAlignment.CENTER == align) {
            cell.setCellStyle(style.getStyleDateTimeCenterNoBorder(excel));
        }
    }

    /**
     * @param excel  excel
     * @param cell   cell
     * @param align  align
     * @param border border
     */
    protected static void setWrapTextStyle(Excel excel, SXSSFCell cell, HorizontalAlignment align, boolean border) {
        ExcelStyle style = excel.getStyle();
        cell.setCellStyle(style.getStyleWrapTextCenterWithBorder(excel));
        // 边框 + 左边
        if (border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleWrapTextLeftWithBorder(excel));
        }
        // 无边框 + 左边
        if (!border && HorizontalAlignment.LEFT == align) {
            cell.setCellStyle(style.getStyleWrapTextLeftNoBorder(excel));
        }
        // 无边框 + 中间
        if (!border && HorizontalAlignment.CENTER == align) {
            cell.setCellStyle(style.getStyleWrapTextCenterNoBorder(excel));
        }
    }


    /*
    private static void main(String[] args) {
		System.out.println(62*12*25/256);
	}
    */

}
