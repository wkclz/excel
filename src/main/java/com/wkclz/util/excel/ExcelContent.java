package com.wkclz.util.excel;


/*
                            _ooOoo_
                           o8888888o
                           88" . "88
                           (| -_- |)
                            O\ = /O
                        ____/`---'\____
                      .   ' \\| | `.
                       / \\||| : ||| \
                     / _||||| -:- |||||- \
                       | | \\\ - / | |
                     | \_| ''\---/'' | |
                      \ .-\__ `-` ___/-. /
                   ___`. .' /--.--\ `. . __
                ."" '< `.___\_<|>_/___.' >'"".
               | | : `- \`.;`\ _ /`;.`/ - ` : | |
                 \ \ `-. \_ __\ /__ _/ .-` / /
         ======`-.____`-.___\_____/___.-`____.-'======
                            `=---='

         .............................................
                  佛祖保佑             永无BUG
          佛曰:
                  写字楼里写字间，写字间里程序员；
                  程序人员写程序，又拿程序换酒钱。
                  酒醒只在网上坐，酒醉还来网下眠；
                  酒醉酒醒日复日，网上网下年复年。
                  但愿老死电脑间，不愿鞠躬老板前；
                  奔驰宝马贵者趣，公交自行程序员。
                  别人笑我忒疯癫，我笑自己命太贱；
                  不见满街漂亮妹，哪个归得程序员？
*/

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.w3c.dom.ls.LSException;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public abstract class ExcelContent {

    /** excel 工作簿 */
    protected SXSSFWorkbook workbook;
    /** sheet */
    /** 多 sheet 支持 */
    protected SXSSFSheet sheet;
    protected List<SXSSFSheet> sheets;
    /** sheet号 */
    protected Integer sheetNum = 0;
    /** 行号 */
    protected Integer rowNum = 0;
    /** 基本信息后的第一行。 */

	/** 标题 */
    protected String title;
	/** 创建人 */
    protected String createBy;
	/** 时间从 */
    protected String dateFrom;
	/** 时间从 */
    protected String dateTo;
	/** 保存路径 */
    protected String savePath;
	/** 表格列名 */
    protected List<String> header;
	/** Excel 宽度【用于在没有title的情况下定义标题合并】 */
    protected Integer width;
	/** 行对象 */
    protected List<ExcelRow> rows;

	/** 字体缓存重用 */
    protected Map<String, Font> workBookFonts;

    /** 所有样式 */
    protected ExcelStyle style;

    protected Integer cacheRowsInMemory = 10240;

    protected List<SXSSFSheet> getSheets() {
        if(this.sheets==null) {
            this.sheets = new ArrayList<SXSSFSheet>();
        }
        return this.sheets;
    }
    protected void addSheet(SXSSFSheet sheet) {
        if(this.sheets==null) {
            this.sheets = new ArrayList<SXSSFSheet>();
        }

        // 不显示风格线
        sheet.setDisplayGridlines(false);
        PrintSetup ps = sheet.getPrintSetup();
        // 打印方向，true：横向，false：纵向(默认)
        ps.setLandscape(false);
        // A4纸
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        this.sheets.add(sheet);
    }

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
		if(createBy==null) {
			createBy = "";
		}
		this.createBy = createBy;
	}
	public void setCreateBy(Object createBy) {
		if(createBy==null) {
			createBy = "";
		}
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
		if(header!=null) {
			this.width = header.size();
		}
		this.header = header;
	}
	/** 使用String[] 初始化列名 */
	public void setHeader(String[] header) {
		if(header!=null) {
			this.width = header.length;
		}
		this.header = new ArrayList<String>();
		for (String h : header) {
			this.header.add(h);
		}
	}
	public Integer getWidth() {
		return width;
	}
	/** nothing */
	public void setWidth(Integer width) {
		this.width = width;
	}
	/** nothing */
	public void setRows(List<ExcelRow> rows) {
		this.rows = rows;
	}
	/** nothing */
	protected List<ExcelRow> getRows() {
		if(rows==null) {
			rows = new ArrayList<ExcelRow>();
		}
		return rows;
	}
	/** nothing */
	protected void addRow(ExcelRow row) {
		if(this.rows==null) {
			this.rows = new ArrayList<ExcelRow>();
		}
		this.rows.add(row);
	}

	/** nothing */
    public Integer getCacheRowsInMemory() {
        return cacheRowsInMemory;
    }

	/** nothing */
    public void setCacheRowsInMemory(Integer cacheRowsInMemory) {
        this.cacheRowsInMemory = cacheRowsInMemory;
    }

	/** nothing */
    public Map<String, Font> getWorkBookFonts() {
        return workBookFonts;
    }

	/** nothing */
    public void setWorkBookFonts(Map<String, Font> workBookFonts) {
        this.workBookFonts = workBookFonts;
    }

	/** nothing */
    public ExcelStyle getStyle() {
        return style;
    }

	/** nothing */
    public void setStyle(ExcelStyle style) {
        this.style = style;
    }

	/** nothing */
    public SXSSFWorkbook getWorkbook() {
        return workbook;
    }
	/** nothing */
    public void setWorkbook(SXSSFWorkbook workbook) {
        this.workbook = workbook;
    }

}
