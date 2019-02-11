package com.wkclz.util.excel;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public class ExcelStyle {


    /*
     * 标题和列名
     */

    /**
     * 标题样式【默认无边框】
     */
    private XSSFCellStyle styleTitle;
    /**
     * 列名样式【默认有边框】
     */
    private XSSFCellStyle styleHeader;


    /**
     * 标题样式【默认无边框】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleTitle(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleTitle == null) {
            styleTitle = (XSSFCellStyle) workbook.createCellStyle();
            styleTitle.setAlignment(HorizontalAlignment.CENTER);
            styleTitle.setVerticalAlignment(VerticalAlignment.CENTER);
            styleTitle.setFont(ExcelUtil.createFont(excel, 24));
        }
        return styleTitle;
    }

    /**
     * 列名样式【默认有边框】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleHeader(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleHeader == null) {
            styleHeader = (XSSFCellStyle) workbook.createCellStyle();
            styleHeader.setAlignment(HorizontalAlignment.CENTER);
            styleHeader.setVerticalAlignment(VerticalAlignment.CENTER);

            //边框
            styleHeader = addBorder(styleHeader);
            styleHeader.setFont(ExcelUtil.createFont(excel, 12));
        }
        return styleHeader;
    }


    /*
     * 字符串【Int也用这个】
     */


    /**
     * 字符串 【无边框,左边】
     */
    private XSSFCellStyle styleStrLeftNoBorder;
    /**
     * 字符串 【无边框，中间】
     */
    private XSSFCellStyle styleStrCenterNoBorder;
    /**
     * 字符串 【有边框，左边】
     */
    private XSSFCellStyle styleStrLeftWithBorder;
    /**
     * 字符串 【有边框，中间】
     */
    private XSSFCellStyle styleStrCenterWithBorder;


    /**
     * 字符串【无边框,左边】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleStrLeftNoBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleStrLeftNoBorder == null) {
            styleStrLeftNoBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleStrLeftNoBorder.setAlignment(HorizontalAlignment.LEFT);
            styleStrLeftNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleStrLeftNoBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleStrLeftNoBorder;
    }

    /**
     * 字符串【无边框，中间】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleStrCenterNoBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleStrCenterNoBorder == null) {
            styleStrCenterNoBorder = (XSSFCellStyle) workbook.createCellStyle();

            styleStrCenterNoBorder.setAlignment(HorizontalAlignment.CENTER);
            styleStrCenterNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);

            styleStrCenterNoBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleStrCenterNoBorder;
    }

    /**
     * 字符串【有边框，左边】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleStrLeftWithBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleStrLeftWithBorder == null) {
            styleStrLeftWithBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleStrLeftWithBorder.setAlignment(HorizontalAlignment.LEFT);
            styleStrLeftWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);

            //边框
            styleStrLeftWithBorder = addBorder(styleStrLeftWithBorder);
            styleStrLeftWithBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleStrLeftWithBorder;
    }

    /**
     * 字符串【有边框，中间】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleStrCenterWithBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleStrCenterWithBorder == null) {
            styleStrCenterWithBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleStrCenterWithBorder.setAlignment(HorizontalAlignment.CENTER);
            styleStrCenterWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);

            //边框
            styleStrCenterWithBorder = addBorder(styleStrCenterWithBorder);
            styleStrCenterWithBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleStrCenterWithBorder;
    }


    /*
     * 小数
     */


    /**
     * 小数 【无边框,左边】
     */
    private XSSFCellStyle styleNumLeftNoBorder;
    /**
     * 小数 【无边框，中间】
     */
    private XSSFCellStyle styleNumCenterNoBorder;
    /**
     * 小数 【有边框，左边】
     */
    private XSSFCellStyle styleNumLeftWithBorder;
    /**
     * 小数 【有边框，中间】
     */
    private XSSFCellStyle styleNumCenterWithBorder;


    /**
     * 小数 【无边框,左边】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleNumLeftNoBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleNumLeftNoBorder == null) {
            styleNumLeftNoBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleNumLeftNoBorder.setAlignment(HorizontalAlignment.LEFT);
            styleNumLeftNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleNumLeftNoBorder.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
            styleNumLeftNoBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleNumLeftNoBorder;
    }

    /**
     * 小数 【无边框，中间】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleNumCenterNoBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleNumCenterNoBorder == null) {
            styleNumCenterNoBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleNumCenterNoBorder.setAlignment(HorizontalAlignment.CENTER);
            styleNumCenterNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleNumCenterNoBorder.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
            styleNumCenterNoBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleNumCenterNoBorder;
    }

    /**
     * 小数 【有边框，左边】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleNumLeftWithBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleNumLeftWithBorder == null) {
            styleNumLeftWithBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleNumLeftWithBorder.setAlignment(HorizontalAlignment.LEFT);
            styleNumLeftWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleNumLeftWithBorder.setDataFormat(workbook.createDataFormat().getFormat("0.00"));

            //边框
            styleNumLeftWithBorder = addBorder(styleNumLeftWithBorder);
            styleNumLeftWithBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleNumLeftWithBorder;
    }

    /**
     * 小数 【有边框，中间】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleNumCenterWithBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleNumCenterWithBorder == null) {
            styleNumCenterWithBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleNumCenterWithBorder.setAlignment(HorizontalAlignment.CENTER);
            styleNumCenterWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleNumCenterWithBorder.setDataFormat(workbook.createDataFormat().getFormat("0.00"));

            //边框
            styleNumCenterWithBorder = addBorder(styleNumCenterWithBorder);
            styleNumCenterWithBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleNumCenterWithBorder;
    }


    /*
     * 日期
     */


    /**
     * 日期 【无边框,左边】
     */
    private XSSFCellStyle styleDateLeftNoBorder;
    /**
     * 日期 【无边框，中间】
     */
    private XSSFCellStyle styleDateCenterNoBorder;
    /**
     * 日期 【有边框，左边】
     */
    private XSSFCellStyle styleDateLeftWithBorder;
    /**
     * 日期 【有边框，中间】
     */
    private XSSFCellStyle styleDateCenterWithBorder;


    /**
     * 日期 【无边框,左边】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleDateLeftNoBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleDateLeftNoBorder == null) {
            styleDateLeftNoBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleDateLeftNoBorder.setAlignment(HorizontalAlignment.LEFT);
            styleDateLeftNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleDateLeftNoBorder.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd"));
            styleDateLeftNoBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleDateLeftNoBorder;
    }

    /**
     * 日期 【无边框，中间】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleDateCenterNoBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleDateCenterNoBorder == null) {
            styleDateCenterNoBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleDateCenterNoBorder.setAlignment(HorizontalAlignment.CENTER);
            styleDateCenterNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleDateCenterNoBorder.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd"));
            styleDateCenterNoBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleDateCenterNoBorder;
    }

    /**
     * 日期 【有边框，左边】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleDateLeftWithBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleDateLeftWithBorder == null) {
            styleDateLeftWithBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleDateLeftWithBorder.setAlignment(HorizontalAlignment.LEFT);
            styleDateLeftWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleDateLeftWithBorder.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd"));

            //边框
            styleDateLeftWithBorder = addBorder(styleDateLeftWithBorder);
            styleDateLeftWithBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleDateLeftWithBorder;
    }

    /**
     * 日期 【有边框，中间】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleDateCenterWithBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleDateCenterWithBorder == null) {
            styleDateCenterWithBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleDateCenterWithBorder.setAlignment(HorizontalAlignment.CENTER);
            styleDateCenterWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleDateCenterWithBorder.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd"));

            //边框
            styleDateCenterWithBorder = addBorder(styleDateCenterWithBorder);
            styleDateCenterWithBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleDateCenterWithBorder;
    }


    /*
     * 日期时间
     */


    /**
     * 日期时间 【无边框,左边】
     */
    private XSSFCellStyle styleDateTimeLeftNoBorder;
    /**
     * 日期时间 【无边框，中间】
     */
    private XSSFCellStyle styleDateTimeCenterNoBorder;
    /**
     * 日期时间 【有边框，左边】
     */
    private XSSFCellStyle styleDateTimeLeftWithBorder;
    /**
     * 日期时间 【有边框，中间】
     */
    private XSSFCellStyle styleDateTimeCenterWithBorder;


    /**
     * 日期时间 【无边框,左边】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleDateTimeLeftNoBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleDateTimeLeftNoBorder == null) {
            styleDateTimeLeftNoBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleDateTimeLeftNoBorder.setAlignment(HorizontalAlignment.LEFT);
            styleDateTimeLeftNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleDateTimeLeftNoBorder.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
            styleDateTimeLeftNoBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleDateTimeLeftNoBorder;
    }

    /**
     * 日期时间 【无边框，中间】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleDateTimeCenterNoBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleDateTimeCenterNoBorder == null) {
            styleDateTimeCenterNoBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleDateTimeCenterNoBorder.setAlignment(HorizontalAlignment.CENTER);
            styleDateTimeCenterNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleDateTimeCenterNoBorder.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
            styleDateTimeCenterNoBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleDateTimeCenterNoBorder;
    }

    /**
     * 日期时间 【有边框，左边】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleDateTimeLeftWithBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleDateTimeLeftWithBorder == null) {
            styleDateTimeLeftWithBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleDateTimeLeftWithBorder.setAlignment(HorizontalAlignment.LEFT);
            styleDateTimeLeftWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleDateTimeLeftWithBorder.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));

            //边框
            styleDateTimeLeftWithBorder = addBorder(styleDateTimeLeftWithBorder);
            styleDateTimeLeftWithBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleDateTimeLeftWithBorder;
    }

    /**
     * 日期时间 【有边框，中间】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleDateTimeCenterWithBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleDateTimeCenterWithBorder == null) {
            styleDateTimeCenterWithBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleDateTimeCenterWithBorder.setAlignment(HorizontalAlignment.CENTER);
            styleDateTimeCenterWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleDateTimeCenterWithBorder.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));

            //边框
            styleDateTimeCenterWithBorder = addBorder(styleDateTimeCenterWithBorder);
            styleDateTimeCenterWithBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleDateTimeCenterWithBorder;
    }


    /*
     * 富文本，自动换行
     */


    /**
     * 富文本，自动换行 【无边框,左边】
     */
    private XSSFCellStyle styleWrapTextLeftNoBorder;
    /**
     * 富文本，自动换行 【无边框，中间】
     */
    private XSSFCellStyle styleWrapTextCenterNoBorder;
    /**
     * 富文本，自动换行 【有边框，左边】
     */
    private XSSFCellStyle styleWrapTextLeftWithBorder;
    /**
     * 富文本，自动换行 【有边框，中间】
     */
    private XSSFCellStyle styleWrapTextCenterWithBorder;


    /**
     * 富文本，自动换行 【无边框,左边】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleWrapTextLeftNoBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleWrapTextLeftNoBorder == null) {
            styleWrapTextLeftNoBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleWrapTextLeftNoBorder.setAlignment(HorizontalAlignment.LEFT);
            styleWrapTextLeftNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleWrapTextLeftNoBorder.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
            styleWrapTextLeftNoBorder.setWrapText(true);
            styleWrapTextLeftNoBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleWrapTextLeftNoBorder;
    }

    /**
     * 富文本，自动换行 【无边框，中间】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleWrapTextCenterNoBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleWrapTextCenterNoBorder == null) {
            styleWrapTextCenterNoBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleWrapTextCenterNoBorder.setAlignment(HorizontalAlignment.CENTER);
            styleWrapTextCenterNoBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleWrapTextCenterNoBorder.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
            styleWrapTextCenterNoBorder.setWrapText(true);
            styleWrapTextCenterNoBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleWrapTextCenterNoBorder;
    }

    /**
     * 富文本，自动换行 【有边框，左边】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleWrapTextLeftWithBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleWrapTextLeftWithBorder == null) {
            styleWrapTextLeftWithBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleWrapTextLeftWithBorder.setAlignment(HorizontalAlignment.LEFT);
            styleWrapTextLeftWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleWrapTextLeftWithBorder.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
            styleWrapTextLeftWithBorder.setWrapText(true);

            //边框
            styleWrapTextLeftWithBorder = addBorder(styleWrapTextLeftWithBorder);
            styleWrapTextLeftWithBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleWrapTextLeftWithBorder;
    }

    /**
     * Descripition 富文本，自动换行【有边框，中间】
     *
     * @param excel excel
     * @return XSSFCellStyle
     */
    protected XSSFCellStyle getStyleWrapTextCenterWithBorder(Excel excel) {
        SXSSFWorkbook workbook = excel.getWorkbook();
        if (styleWrapTextCenterWithBorder == null) {
            styleWrapTextCenterWithBorder = (XSSFCellStyle) workbook.createCellStyle();
            styleWrapTextCenterWithBorder.setAlignment(HorizontalAlignment.CENTER);
            styleWrapTextCenterWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            styleWrapTextCenterWithBorder.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
            styleWrapTextCenterWithBorder.setWrapText(true);
            //边框
            styleWrapTextCenterWithBorder = addBorder(styleWrapTextCenterWithBorder);

            styleWrapTextCenterWithBorder.setFont(ExcelUtil.createFont(excel));
        }
        return styleWrapTextCenterWithBorder;
    }


    /**
     * Description 添加边框
     *
     * @param xssfCellStyle nxssfCellStyle
     * @return XSSFCellStyle
     */
    private static XSSFCellStyle addBorder(XSSFCellStyle xssfCellStyle) {
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

}
