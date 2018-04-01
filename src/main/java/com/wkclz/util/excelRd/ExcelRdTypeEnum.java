package com.wkclz.util.excelRd;

/**
 * Description:
 * Created: wangkaicun @ 2017-11-15 下午10:42
 */
public enum ExcelRdTypeEnum {

    /**
     * Excel 所支持读取的类型
     */
    INTEGER("整形"),
    DOUBLE("双精浮点型"),
    DATE("日期型"),
    DATETIME("日期时间型"),
    STRING("字符型");

    private String value;
    private ExcelRdTypeEnum(String value) { this.value = value; }
    public String getValue() { return value; }

}
