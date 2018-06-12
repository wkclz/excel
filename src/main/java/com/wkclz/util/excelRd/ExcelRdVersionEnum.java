package com.wkclz.util.excelRd;

/**
 * Description:
 * Created: wangkaicun @ 2017-11-15 下午10:42
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
