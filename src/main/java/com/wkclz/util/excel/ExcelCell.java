package com.wkclz.util.excel;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

public class ExcelCell {
	
	/** cell内容 */
	private Object cellContent;
	/** 是否有边框 */
	private boolean border;
	/** 对齐方式，只有左和中，默认中对齐，从ExcelUtil 取值*/
	private HorizontalAlignment align;
	/** 合并列数【宽度】 */
	private int col;
	/**	合并行数【高度】 */
	private int row;

	/**
	 * 创建 cell
	 * @param cellContent cellContent
	 * @param border border
	 * @param align align
	 * @param col col
	 * @param row row
	 */
	protected ExcelCell(Object cellContent, boolean border, HorizontalAlignment align, int col, int row) {
		super();
		this.cellContent = cellContent;
		this.border = border;
		this.align = align;
		this.col = col;
		this.row = row;
	}
	
	/** cell内容 */
	protected Object getCellContent() {
		return cellContent;
	}
	/** cell内容 */
	protected void setCellContent(Object cellContent) {
		this.cellContent = cellContent;
	}
	/** 是否有边框【默认有边框】 */
	protected boolean getBorder() {
		return border;
	}
	/** 是否有边框【默认有边框】 */
	protected void setBorder(boolean border) {
		this.border = border;
	}
	/** 对齐方式，只有左和中，默认中对齐，从ExcelUtil 取值*/
	protected HorizontalAlignment getAlign() {
		return align;
	}
	/** 对齐方式，只有左和中，默认中对齐，从ExcelUtil 取值*/
	protected void setAlign(HorizontalAlignment align) {
		this.align = align;
	}
	/** 合并列数【宽度】 */
	protected int getCol() {
		return col;
	}
	/** 合并列数【宽度】 */
	protected void setCol(int col) {
		this.col = col;
	}
	/**	合并行数【高度】 */
	protected int getRow() {
		return row;
	}
	/**	合并行数【高度】 */
	protected void setRow(int row) {
		this.row = row;
	}
}
