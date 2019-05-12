package com.wkclz.util.excelRd.domain;

import com.wkclz.util.excelRd.ExcelRdTypeEnum;

import java.util.ArrayList;
import java.util.List;

public class ExcelRdSheet {

    /**
     * Sheet
     */
    private Integer sheet;

    /**
     * sheet 名称
     */
    private String sheetName;

    /**
     * 指定起始行，下标从0开始计数
     */
    private Integer startRow = 0;

    /**
     * 指定起始列，下标从0开始计数
     */
    private Integer startCol = 0;

    /**
     * 指定每列的类型
     */
    private List<ExcelRdTypeEnum> types;

    /**
     * 行对象
     */
    private List<List<Object>> rows;



    public Integer getSheet() {
        return sheet;
    }

    public void setSheet(Integer sheet) {
        this.sheet = sheet;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Integer getStartRow() {
        return startRow;
    }

    public void setStartRow(Integer startRow) {
        this.startRow = startRow;
    }

    public Integer getStartCol() {
        return startCol;
    }

    public void setStartCol(Integer startCol) {
        this.startCol = startCol;
    }

    public List<ExcelRdTypeEnum> getTypes() {
        return types;
    }

    public void setTypes(List<ExcelRdTypeEnum> types) {
        this.types = types;
    }

    public void setTypes(ExcelRdTypeEnum[] types) {
        if (this.types == null) {
            this.types = new ArrayList<ExcelRdTypeEnum>();
        }
        for (ExcelRdTypeEnum type : types) {
            this.types.add(type);
        }
    }

    public List<List<Object>> getRows() {
        return rows;
    }

    private void setRows(List<List<Object>> rows) {
        this.rows = rows;
    }

    public void addRow(List<Object> row) {
        if (rows == null) {
            rows = new ArrayList<List<Object>>();
        }
        rows.add(row);
    }

}
