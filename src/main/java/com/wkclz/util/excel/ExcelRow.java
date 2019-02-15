package com.wkclz.util.excel;

import org.apache.poi.ss.usermodel.Font;
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
        addCell(cellContent, border, align, 1, 1, null);
    }

    public void addCell(Object cellContent, int col, int row) {
        addCell(cellContent, true, ExcelUtil.ALIGN_CENTER, col, row, null);
    }

    public void addCell(Object cellContent, boolean border, int col, int row) {
        addCell(cellContent, border, ExcelUtil.ALIGN_CENTER, col, row, null);
    }

    public void addCell(Object cellContent, HorizontalAlignment align, int col, int row) {
        addCell(cellContent, true, align, col, row, null);
    }


    /**
     * 以下有包含 Font
     * @param cellContent
     * @param font
     */
    public void addCell(Object cellContent, Font font) {
        addCell(cellContent, true, ExcelUtil.ALIGN_CENTER, font);
    }

    public void addCell(Object cellContent, boolean border, Font font) {
        addCell(cellContent, border, ExcelUtil.ALIGN_CENTER, font);
    }

    public void addCell(Object cellContent, HorizontalAlignment align, Font font) {
        addCell(cellContent, true, align, font);
    }

    public void addCell(Object cellContent, boolean border, HorizontalAlignment align, Font font) {
        addCell(cellContent, border, align, 1, 1, font);
    }

    public void addCell(Object cellContent, int col, int row, Font font) {
        addCell(cellContent, true, ExcelUtil.ALIGN_CENTER, col, row, font);
    }

    public void addCell(Object cellContent, boolean border, int col, int row, Font font) {
        addCell(cellContent, border, ExcelUtil.ALIGN_CENTER, col, row, font);
    }

    public void addCell(Object cellContent, HorizontalAlignment align, int col, int row, Font font) {
        addCell(cellContent, true, align, col, row, font);
    }


    public void addCell(Object cellContent, boolean border, HorizontalAlignment align, int col, int row, Font font) {
        ExcelCell excelCell = new ExcelCell(cellContent, border, align, col, row, font);
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
