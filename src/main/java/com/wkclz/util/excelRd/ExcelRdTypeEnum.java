package com.wkclz.util.excelRd;

/**
 * Description enum
 * create @: 2017-11-15 下午10:42
 *
 * @author wangkaicun wkclz@qq.com
 */
public enum ExcelRdTypeEnum {

    /**
     * Excel 所支持读取的类型
     */
    INTEGER("整形"),
    LONG("长整形"),
    DOUBLE("双精浮点型"),
    DATE("日期型"),
    DATETIME("日期时间型"),
    STRING("字符型");

    private String value;

    private ExcelRdTypeEnum(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }

}
