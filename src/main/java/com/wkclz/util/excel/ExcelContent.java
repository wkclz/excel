package com.wkclz.util.excel;

import java.util.ArrayList;
import java.util.List;

public abstract class ExcelContent extends ExcelStyle {
	
	/** 标题 */
	private String title;
	/** 创建人 */
	private String createBy;
	/** 时间从 */
	private String dateFrom;
	/** 时间从 */
	private String dateTo;
	/** 保存路径 */
	private String savePath;
	/** 表格列名 */
	private List<String> header;
	/** Excel 宽度【用于在没有title的情况下定义标题合并】 */
	private Integer width;
	/** 行对象 */
	private List<ExcelRow> rows;

	private Integer cacheRowsInMemory;

	public String getTitle() {
		return title;
	}
	public void setTitle(String title) {
		this.title = title;
	}
	public String getCreateBy() {
		return createBy;
	}
	public void setCreateBy(String createBy) {
		if(createBy==null)
			createBy = "";
		this.createBy = createBy;
	}
	public void setCreateBy(Object createBy) {
		if(createBy==null)
			createBy = "";
		this.createBy = createBy.toString();
	}
	public String getDateFrom() {
		return dateFrom;
	}
	public void setDateFrom(String dateFrom) {
		this.dateFrom = dateFrom;
	}
	public String getDateTo() {
		return dateTo;
	}
	public void setDateTo(String dateTo) {
		this.dateTo = dateTo;
	}
	public String getSavePath() {
		return savePath;
	}
	public void setSavePath(String savePath) {
		this.savePath = savePath;
	}
	public List<String> getHeader() {
		return header;
	}
	/** 使用List<String>初始化列名 */
	public void setHeader(List<String> header) {
		if(header!=null)
			this.width = header.size();
		this.header = header;
	}
	/** 使用String[] 初始化列名 */
	public void setHeader(String[] header) {
		if(header!=null)
			this.width = header.length;
		this.header = new ArrayList<String>(); 
		for (String h : header)
			this.header.add(h);
	}
	public Integer getWidth() {
		return width;
	}
	public void setWidth(Integer width) {
		this.width = width;
	}
	public void setRows(List<ExcelRow> rows) {
		this.rows = rows;
	}
	protected List<ExcelRow> getRows() {
		if(rows==null)
			rows = new ArrayList<ExcelRow>();
		return rows;
	}
	/** 使用setLines初始化内容行 */
	protected void addRow(ExcelRow row) {
		if(this.rows==null)
			this.rows = new ArrayList<ExcelRow>();
		this.rows.add(row);
	}

    public Integer getCacheRowsInMemory() {
        return cacheRowsInMemory;
    }

    public void setCacheRowsInMemory(Integer cacheRowsInMemory) {
        this.cacheRowsInMemory = cacheRowsInMemory;
    }

    /**
     * 此方法为了兼容旧版，
     * @param createBy
     */
    @Deprecated
    public void setCreate_by(String createBy) {
        if(createBy==null)
            createBy = "";
        this.createBy = createBy;
    }

    /**
     * 此方法为了兼容旧版，
     * @param createBy
     */
    @Deprecated
    public void setCreate_by(Object createBy) {
        if(createBy==null)
            createBy = "";
        this.createBy = createBy.toString();
    }


    /**
     * 此方法为了兼容旧版，
     * @return
     */
    @Deprecated
    public String getCreate_by() {
        return createBy;
    }
	
}
