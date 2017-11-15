package com.wkclz.util.excelRd;

import java.util.ArrayList;
import java.util.List;

public abstract class ExcelRdContent extends ExcelRdBase {

	/** 起始行 */
	private int startSheet;
	/** 起始行 */
	private int startRow;
	/** 起始列 */
	private int startCol;
	/** 列类型 */
	private List<ExcelRdTypeEnum> types;
	/** 行对象 */
	private List<ExcelRdRow> rows;
	
	
	protected int getStartSheet() {
		return startSheet;
	}
	public void setStartSheet(int startSheet) {
		this.startSheet = startSheet;
	}
	protected int getStartRow() {
		return startRow;
	}
	public void setStartRow(int startRow) {
		this.startRow = startRow;
	}
	protected int getStartCol() {
		return startCol;
	}
	public void setStartCol(int startCol) {
		this.startCol = startCol;
	}
	protected List<ExcelRdTypeEnum> getTypes() {
		return types;
	}
	public void setTypes(List<ExcelRdTypeEnum> types) {
		this.types = types;
	}
	public void setTypes(ExcelRdTypeEnum[] types) {
		if(this.types==null)
			this.types = new ArrayList<ExcelRdTypeEnum>();
		for (ExcelRdTypeEnum type : types) {
			this.types.add(type);
		}
	}
	protected List<ExcelRdRow> getRows() {
		return rows;
	}
	protected void setRows(List<ExcelRdRow> rows) {
		this.rows = rows;
	}
	protected void addRow(ExcelRdRow row) {
		if(rows==null)
			rows = new ArrayList<ExcelRdRow>();
		rows.add(row);
	}
}
