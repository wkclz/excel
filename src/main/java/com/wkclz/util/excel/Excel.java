package com.wkclz.util.excel;

/**
 * ┌───┐   ┌───┬───┬───┬───┐ ┌───┬───┬───┬───┐ ┌───┬───┬───┬───┐ ┌───┬───┬───┐
 * │Esc│   │ F1│ F2│ F3│ F4│ │ F5│ F6│ F7│ F8│ │ F9│F10│F11│F12│ │P/S│S L│P/B│  ┌┐    ┌┐    ┌┐
 * └───┘   └───┴───┴───┴───┘ └───┴───┴───┴───┘ └───┴───┴───┴───┘ └───┴───┴───┘  └┘    └┘    └┘
 * ┌───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───────┐ ┌───┬───┬───┐ ┌───┬───┬───┬───┐
 * │~ `│! 1│@ 2│# 3│$ 4│% 5│^ 6│& 7│* 8│( 9│) 0│_ -│+ =│ BacSp │ │Ins│Hom│PUp│ │N L│ / │ * │ - │
 * ├───┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─────┤ ├───┼───┼───┤ ├───┼───┼───┼───┤
 * │ Tab │ Q │ W │ E │ R │ T │ Y │ U │ I │ O │ P │{ [│} ]│ | \ │ │Del│End│PDn│ │ 7 │ 8 │ 9 │   │
 * ├─────┴┬──┴┬──┴┬──┴┬──┴┬──┴┬──┴┬──┴┬──┴┬──┴┬──┴┬──┴┬──┴─────┤ └───┴───┴───┘ ├───┼───┼───┤ + │
 * │ Caps │ A │ S │ D │ F │ G │ H │ J │ K │ L │: ;│" '│ Enter  │               │ 4 │ 5 │ 6 │   │
 * ├──────┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴────────┤     ┌───┐     ├───┼───┼───┼───┤
 * │ Shift  │ Z │ X │ C │ V │ B │ N │ M │< ,│> .│? /│  Shift   │     │ ↑ │     │ 1 │ 2 │ 3 │   │
 * ├─────┬──┴─┬─┴──┬┴───┴───┴───┴───┴───┴──┬┴───┼───┴┬────┬────┤ ┌───┼───┼───┐ ├───┴───┼───┤ E││
 * │ Ctrl│    │Alt │         Space         │ Alt│    │    │Ctrl│ │ ← │ ↓ │ → │ │   0   │ . │←─┘│
 * └─────┴────┴────┴───────────────────────┴────┴────┴────┴────┘ └───┴───┴───┘ └───────┴───┴───┘
 */

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Excel extends ExcelContent {
	
	private ExcelRow row;
	/** 创建一行 row */
	public ExcelRow createRow() {
		// 只有第一次 createRow 的时候才能为空。之后的 createRow过程都需要把前一次的 row 添加进去
		if(row!=null) {
			addRow(row);
		}
		row = new ExcelRow();
		return row;
	}
	
	/** 在缓存里面创建 row【需要在适当时候，手动加入到 excel中】 */
	public ExcelRow createRowInCache() {
		return new ExcelRow();
	}
	
	/** 在缓存里面的 row 加入到 excel中【会将前一次row 提前写入excel并置空】 */
	public void addRowFromCache(ExcelRow row) {
		// 还是要将缓存的row 给设置到excel里面
		if(this.row!=null){
			addRow(this.row);
			this.row = null;
		}
		addRow(row);
	}


    /**
     * @Title:
     * @Description: 生成 Excel 到指定目录
     * @param @return    设定文件
     * @author wangkc admin@wkclz.com
     * @date 2017年7月16日 上午12:57:41 *
     * @throws
     */
    public String CreateXlsx() throws ExcelException, IOException {
        // 把最后一次的数据加进去
        if(row!=null) {
			addRow(row);
		}

        String path = getSavePath();
        if(path==null||"".equals(path.trim())) {
			throw new ExcelException("savePath cannot be null or empty!");
		}

        create();   // 生成的过程

        // 导出到文件
        FileOutputStream outputStream = new FileOutputStream(getSavePath());
        getWorkbook().write(outputStream);
        outputStream.flush();
        outputStream.close();
        return path;
    }

    /**
     * @Title:
     * @Description: 生成 Excel 到输出流
     * @param @return    设定文件
     * @author wangkc admin@wkclz.com
     * @date 2017年12月09日 上午10:31:09 *
     * @throws
     */
    public File CreateXlsxByFile() throws ExcelException {
        // 把最后一次的数据加进去
        if(row!=null) {
			addRow(row);
		}

        create();   // 生成的过程

        File file = null;
        try {

            file = File.createTempFile("temp", ".xlsx");
            FileOutputStream stream = new FileOutputStream(file);
            getWorkbook().write(stream);
            stream.flush();
            stream.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
        return file;

    }


    /**
     * 生成 excel 的具体过程
     */
	private void create() throws ExcelException {

	    String title = getTitle();
		if(title==null||"".equals(title.trim())) {
			throw new ExcelException("title cannot be null or empty!");
		}

		// 找出不允许的 str
		String[] notAllowdStrs = {":","：","/","?","？","\\","*","[","]"};
		List<String> existStr = new ArrayList<String>();
        for (String notAllowdStr: notAllowdStrs) {
            if (title.contains(notAllowdStr)) {
				existStr.add(notAllowdStr);
			}
        }

        if (!existStr.isEmpty()){
            String rt = "";
            for (String s: existStr) {
                rt = s+",";
            }
            rt = rt.substring(0,rt.length()-1);
            throw new ExcelException("title contains this chars: \""+rt + "\" is not allowd!");
        }


		if((getHeader()==null||getHeader().size()==0)&&getWidth()==null) {
			throw new ExcelException("header or width cannot be null or empty!");
		}


        Integer cacheRowsInMemory = getCacheRowsInMemory() == null ? ExcelUtil.CACHE_ROWS_IN_MEMORY : getCacheRowsInMemory();
		setWorkbook(new SXSSFWorkbook(cacheRowsInMemory));	// 内存保留 10240 行数据，多余的刷新到固化存储
		setSheet(getWorkbook().createSheet(title));

		// title
		if( title!=null && !"".equals(title.trim()) ){
			getSheet().addMergedRegion(new CellRangeAddress(rownum, rownum, 0, getWidth()-1));
			SXSSFRow row_title = getSheet().createRow(rownum++);
			SXSSFCell cell_title = row_title.createCell(0);
			cell_title.setCellStyle(getStyleTitle());
			cell_title.setCellValue(title);
		}

		// infomation of this excel
		getSheet().addMergedRegion(new CellRangeAddress(rownum, rownum, 0, 2));
		SXSSFRow row_info = getSheet().createRow(rownum);
		SXSSFCell cell_time = row_info.createCell(0);
		cell_time.setCellStyle(getStyleStrLeftNoBorder());
		cell_time.setCellValue("创建时间："+sdf_dateTime.format(new Date()));

		// create_by
		if(getCreateBy()!=null){
			getSheet().addMergedRegion(new CellRangeAddress(rownum, rownum, 3, 5));
			SXSSFCell cell_create_by = row_info.createCell(3);
			cell_create_by.setCellStyle(getStyleStrLeftNoBorder());
			cell_create_by.setCellValue("创建人："+getCreateBy());
		}

		rownum++;

		// date_from_to
		String date_info = "";
		if(getDateFrom()!=null&&getDateTo()!=null) {
			date_info = "时间范围：从"+getDateFrom()+"到"+getDateTo();
		}
		if(getDateFrom()!=null&&getDateTo()==null) {
			date_info = "时间："+getDateFrom();
		}
		if(getDateFrom()==null&&getDateTo()!=null) {
			date_info = "时间："+getDateTo();
		}

		if(!"".equals(date_info)){
			getSheet().addMergedRegion(new CellRangeAddress(rownum, rownum, 0, 5));
			SXSSFRow row_date_info = getSheet().createRow(rownum);
			SXSSFCell cell_date_info = row_date_info.createCell(0);
			cell_date_info.setCellStyle(getStyleStrLeftNoBorder());
			cell_date_info.setCellValue(date_info);
			rownum++;
		}


		// 写 列名称
		List<String> hs = getHeader();
		if(hs!=null&&hs.size()>0){
			SXSSFRow row_header = getSheet().createRow(rownum++);
			for (int i = 0; i < hs.size(); i++) {
				SXSSFCell cell_header = row_header.createCell(i);
				cell_header.setCellStyle(getStyleHeader());
				cell_header.setCellValue(hs.get(i));
				ExcelUtil.setWidth(getSheet(), i, hs.get(i));
			}
		}

		// 写数据
		/**
		 * 要考虑的情况：
		 * 1、类型，
		 * 2、对齐，
		 * 3、边框，
		 * 4、合并
		 */
		ExcelCell excelCell;	// cell 对象
		Object content;			// cell 内容
		HorizontalAlignment align;			// cell 对齐方式【默认居中】
		boolean border;			// cell 边框【默认有边框】
		List<ExcelRow> lines = getRows();

		// 对所有行对象进行循环
		for (ExcelRow line : lines) {

			int col_num = 0;	// 列号
			SXSSFRow row = getSheet().getRow(rownum);
			if(row==null) {
				row = getSheet().createRow(rownum);
			}

			// 对所有的cell 对象进行循环【在设置表格的时候，若有合并的cell，会自动跳过】
			int size = line.size();
			for (int j = 0; j < size; j++) {

				int now_cell = col_num;	// 当前单元格，只用于cell的宽度设定。col_num 将在使用完后就指定下一cell

				excelCell = line.get(j);
				content = excelCell.getCellContent();
				align = excelCell.getAlign();
				border = excelCell.getBorder();
				int col_merge = excelCell.getCol();
				if(col_merge<1) {
					col_merge = 1;
				}
				int row_merge = excelCell.getRow();
				if(row_merge<1) {
					row_merge = 1;
				}

				// 若有创建cell 直接获取，否则，新建【新建cell ，合并，设置边框，这些都将只有在新建的时候进行操作，之后只是跳过相应的cell】
				SXSSFCell cell = row.getCell(col_num);
				if(cell==null){
					cell = row.createCell(col_num, CellType.NUMERIC);
					// 合并单元格
					mergeCell(col_merge, row_merge, col_num, border);
					col_num += col_merge;	 // 列号向前
				} else {
					// 如果cell已经有了找到下一个空的cell
					col_num = getCell(row,col_num);
					// 当前列号需要更新
					now_cell = col_num;
					cell = row.createCell(col_num, CellType.NUMERIC);
					// 合并单元格
					mergeCell(col_merge, row_merge, col_num, border);
					col_num ++;
				}

				// 空
				if (content == null) {
					content = "";
				}

				// Integer
				if(content instanceof Integer){
					setIntStrStyle(cell, align, border);
					cell.setCellValue((Integer) content);
					continue;
				}

				// Double
				if(content instanceof Double){
					setDoubleStyle(cell, align, border);
					cell.setCellValue((Double) content);
					// 列不合并才自动宽度
					if(col_merge==1) {
						ExcelUtil.setWidth(getSheet(), now_cell, content.toString());
					}
					continue;
				}

				// 时间【不能把 java.sql.Date 的时间算在内】：java.util.Date
				if( !(content instanceof java.sql.Date) && (content instanceof Date) ){
					content = sdf_dateTime.format(content);
					setDateTimeStyle(cell, align, border);
					try {
						cell.setCellValue(sdf_dateTime.parse(content.toString()));
					} catch (ParseException e) {
						e.printStackTrace();
					}
					// 列不合并才自动宽度
					if(col_merge==1) {
						ExcelUtil.setWidth(getSheet(), now_cell, content.toString());
					}
					continue;
				}

				// java.sql.Date
				if( (content instanceof Date) ){
					content = sdf_date.format(content);
					setDateStyle(cell, align, border);
					try {
						cell.setCellValue(sdf_date.parse(content.toString()));
					} catch (ParseException e) {
						e.printStackTrace();
					}
					// 列不合并才自动宽度
					if(col_merge==1) {
						ExcelUtil.setWidth(getSheet(), now_cell, content.toString());
					}
					continue;
				}

				// 最后，当字符串处理
				cell = row.createCell(now_cell, CellType.STRING);
				setIntStrStyle(cell, align, border);
				cell.setCellValue(content.toString());

				// 如果文字内容太长了，设置为富文本类型
				// 列不合并才自动宽度
				if(col_merge==1){
					boolean too_long = ExcelUtil.setWidth(getSheet(), now_cell, content.toString());
					if(too_long) {
						setWrapTextStyle(cell, align, border);
					}
				}

			}
			rownum++;
		}
	}




	
	private void setIntStrStyle(SXSSFCell cell, HorizontalAlignment align, boolean border){
		cell.setCellStyle(getStyleStrCenterWithBorder());
		// 边框 + 左边
		if(border && HorizontalAlignment.LEFT == align) {
			cell.setCellStyle(getStyleStrLeftWithBorder());
		}
		// 无边框 + 左边
		if(!border && HorizontalAlignment.LEFT == align) {
			cell.setCellStyle(getStyleStrLeftNoBorder());
		}
		// 无边框 + 中间
		if(!border && HorizontalAlignment.CENTER == align) {
			cell.setCellStyle(getStyleStrCenterNoBorder());
		}
	}
	private void setDoubleStyle(SXSSFCell cell,HorizontalAlignment align, boolean border){
		cell.setCellStyle(getStyleNumCenterWithBorder());
		// 边框 + 左边
		if(border && HorizontalAlignment.LEFT == align) {
			cell.setCellStyle(getStyleNumLeftWithBorder());
		}
		// 无边框 + 左边
		if(!border && HorizontalAlignment.LEFT == align) {
			cell.setCellStyle(getStyleNumLeftNoBorder());
		}
		// 无边框 + 中间
		if(!border && HorizontalAlignment.CENTER == align) {
			cell.setCellStyle(getStyleNumCenterNoBorder());
		}
	}
	private void setDateStyle(SXSSFCell cell,HorizontalAlignment align, boolean border){
		cell.setCellStyle(getStyleDateCenterWithBorder());
		// 边框 + 左边
		if(border && HorizontalAlignment.LEFT == align) {
			cell.setCellStyle(getStyleDateLeftWithBorder());
		}
		// 无边框 + 左边
		if(!border && HorizontalAlignment.LEFT == align) {
			cell.setCellStyle(getStyleDateLeftNoBorder());
		}
		// 无边框 + 中间
		if(!border && HorizontalAlignment.CENTER == align) {
			cell.setCellStyle(getStyleDateCenterNoBorder());
		}
	}
	private void setDateTimeStyle(SXSSFCell cell,HorizontalAlignment align, boolean border){
		cell.setCellStyle(getStyleDateTimeCenterWithBorder());
		// 边框 + 左边
		if(border && HorizontalAlignment.LEFT == align) {
			cell.setCellStyle(getStyleDateTimeLeftWithBorder());
		}
		// 无边框 + 左边
		if(!border && HorizontalAlignment.LEFT == align) {
			cell.setCellStyle(getStyleDateTimeLeftNoBorder());
		}
		// 无边框 + 中间
		if(!border && HorizontalAlignment.CENTER == align) {
			cell.setCellStyle(getStyleDateTimeCenterNoBorder());
		}
	}
	private void setWrapTextStyle(SXSSFCell cell,HorizontalAlignment align, boolean border){
		cell.setCellStyle(getStyleWrapTextCenterWithBorder());
		// 边框 + 左边
		if(border && HorizontalAlignment.LEFT == align) {
			cell.setCellStyle(getStyleWrapTextLeftWithBorder());
		}
		// 无边框 + 左边
		if(!border && HorizontalAlignment.LEFT == align) {
			cell.setCellStyle(getStyleWrapTextLeftNoBorder());
		}
		// 无边框 + 中间
		if(!border && HorizontalAlignment.CENTER == align) {
			cell.setCellStyle(getStyleWrapTextCenterNoBorder());
		}
	}
	
	/**
	 * 递归找到空的cell
	* @Title:  
	* @Description: TODO(这里用一句话描述这个方法的作用) 
	* @param @param row
	* @param @param col_num
	* @param @return    设定文件 
	* @author wangkc admin@wkclz.com  
	* @date 2017年7月15日 下午9:09:00 *  
	* @throws
	 */
	private int getCell(SXSSFRow row,int col_num){
		if (row.getCell(col_num)!=null) {
			return getCell(row, col_num+1);
		}
		return col_num;
	}
	
	private void mergeCell(int col_merge, int row_merge, int col_num, boolean border){
		// 检查是否需要合并单元格
		if(col_merge>1||row_merge>1){
			getSheet().addMergedRegion(new CellRangeAddress( rownum, rownum+row_merge-1, col_num, col_num+col_merge-1));
			
			//预设内容
			for (int x=rownum; x < rownum+row_merge; x++) {
				for (int y = col_num; y < col_num+col_merge; y++) {
					SXSSFRow r = getSheet().getRow(x);
					if(r==null) {
						r = getSheet().createRow(x);
					}
					SXSSFCell c = r.getCell(y);
					if(c==null) {
						c = r.createCell(y);
					}
					
					// 是否需要设置边框
					if(border) {
						c.setCellStyle(getStyleNumCenterWithBorder());
					} else {
						c.setCellStyle(getStyleNumCenterNoBorder());
					}
				}
			}
		}
	}
	
}
