package com.wkclz.util.excelRd;

import java.util.ArrayList;
import java.util.List;

public class ExcelRdRow {
	
	private List<Object> row;
	
	public List<Object> getRow() {
		return row;
	}
	protected void addCell(Object cellContent) {
		if(this.row==null) {
            this.row = new ArrayList<Object>();
        }
		this.row.add(cellContent);
	}
	
}
