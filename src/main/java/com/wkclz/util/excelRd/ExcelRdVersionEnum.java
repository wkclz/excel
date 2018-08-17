package com.wkclz.util.excelRd;

/**
 * Description enum
 * create @: 2017-11-15 22:42:00
 * @author wangkaicun wkclz@qq.com
 */
public enum ExcelRdVersionEnum {

    /**
     * Excel 版本
     */
    XLS("03版本Excel"),
    XLSX("07+版本Excel");

    private String value;
    private ExcelRdVersionEnum(String value) { this.value = value; }
    public String getValue() { return value; }

}
