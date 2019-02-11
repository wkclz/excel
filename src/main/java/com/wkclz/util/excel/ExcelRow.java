package com.wkclz.util.excel;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.util.ArrayList;
import java.util.List;

public class ExcelRow {

    private Integer sheetNum;

    private List<ExcelCell> row;

    protected Integer getSheetNum() {
        return sheetNum;
    }

    protected void setSheetNum(Integer sheetNum) {
        this.sheetNum = sheetNum;
    }

    protected List<ExcelCell> getRow() {
        return row;
    }

    protected void setRow(List<ExcelCell> row) {
        this.row = row;
    }

    public void addCell(Object cellContent) {
        addCell(cellContent, true, ExcelUtil.ALIGN_CENTER);
    }

    public void addCell(Object cellContent, boolean border) {
        addCell(cellContent, border, ExcelUtil.ALIGN_CENTER);
    }

    public void addCell(Object cellContent, HorizontalAlignment align) {
        addCell(cellContent, true, align);
    }

    public void addCell(Object cellContent, boolean border, HorizontalAlignment align) {
        addCell(cellContent, border, align, 1, 1);
    }

    public void addCell(Object cellContent, int col, int row) {
        addCell(cellContent, true, ExcelUtil.ALIGN_CENTER, col, row);
    }

    public void addCell(Object cellContent, boolean border, int col, int row) {
        addCell(cellContent, border, ExcelUtil.ALIGN_CENTER, col, row);
    }

    public void addCell(Object cellContent, HorizontalAlignment align, int col, int row) {
        addCell(cellContent, true, align, col, row);
    }

    public void addCell(Object cellContent, boolean border, HorizontalAlignment align, int col, int row) {
        ExcelCell excelCell = new ExcelCell(cellContent, border, align, col, row);
        if (this.row == null) {
            this.row = new ArrayList<ExcelCell>();
        }
        this.row.add(excelCell);
    }

    protected int size() {
        if (row == null) {
            return 0;
        }
        return row.size();
    }

    protected ExcelCell get(int i) {
        if (row == null) {
            return null;
        }
        return row.get(i);
    }
}
