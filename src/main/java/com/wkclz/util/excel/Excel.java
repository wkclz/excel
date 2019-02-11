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
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Excel extends ExcelContent {

    private ExcelRow row;

    /**
     * 创建一行 row
     */
    public ExcelRow createRow() {
        // 只有第一次 createRow 的时候才能为空。之后的 createRow过程都需要把前一次的 row 添加进去
        if (this.row != null) {
            this.row.setSheetNum(sheetNum);
            addRow(this.row);
        }
        this.row = new ExcelRow();
        return this.row;
    }

    /**
     * 在缓存里面创建 row【需要在适当时候，手动加入到 excel中】
     */
    public ExcelRow createRowInCache() {
        return new ExcelRow();
    }

    /**
     * 在缓存里面的 row 加入到 excel中【会将前一次row 提前写入excel并置空】
     */
    public void addRowFromCache(ExcelRow row) {
        // 还是要将缓存的row 给设置到excel里面
        if (this.row != null) {
            this.row.setSheetNum(sheetNum);
            addRow(this.row);
            this.row = null;
        }
        addRow(row);
    }

    /**
     * 创建一个 Sheet
     */
    public void addNewSheet() {
        // 将 row 给添加完。
        if (this.row != null) {
            addRow(this.row);
            this.row.setSheetNum(sheetNum);
            this.row = null;
        }
        sheetNum++;
    }


    /**
     * Description 生成Excel到指定目录
     * create @ 2017-07-16 12:57:41
     *
     * @throws ExcelException,IOException
     * @author wangkc admin@wkclz.com
     */
    public String createXlsx() throws ExcelException, IOException {
        // 把最后一次的数据加进去
        if (this.row != null) {
            this.row.setSheetNum(sheetNum);
            addRow(this.row);
            this.row = null;
        }

        if (savePath == null || "".equals(savePath.trim())) {
            throw new ExcelException("savePath cannot be null or empty!");
        }

        create();   // 生成的过程

        // 导出到文件
        FileOutputStream outputStream = new FileOutputStream(savePath);
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
        return savePath;
    }

    /**
     * Description 生成Excel到输出流
     * create @ 2017-12-09 10:31:09
     *
     * @throws ExcelException
     * @author wangkc admin@wkclz.com
     */
    public File createXlsxByFile() throws ExcelException {
        // 把最后一次的数据加进去
        if (this.row != null) {
            this.row.setSheetNum(sheetNum);
            addRow(this.row);
            this.row = null;
        }

        this.create();   // 生成的过程

        File file = null;
        try {

            file = File.createTempFile("temp", ".xlsx");
            FileOutputStream stream = new FileOutputStream(file);
            workbook.write(stream);
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

        String title = this.title;
        if (title == null || "".equals(title.trim())) {
            throw new ExcelException("title cannot be null or empty!");
        }

        // 找出不允许的 str
        String[] notAllowdStrs = {":", "：", "/", "?", "？", "\\", "*", "[", "]"};
        List<String> existStr = new ArrayList<String>();
        for (String notAllowdStr : notAllowdStrs) {
            if (title.contains(notAllowdStr)) {
                existStr.add(notAllowdStr);
            }
        }

        if (!existStr.isEmpty()) {
            String rt = "";
            for (String s : existStr) {
                rt = s + ",";
            }
            rt = rt.substring(0, rt.length() - 1);
            throw new ExcelException("title contains this chars: \"" + rt + "\" is not allowd!");
        }

        boolean headerError = (header == null || header.size() == 0) && width == null;
        if (headerError) {
            throw new ExcelException("header or width cannot be null or empty!");
        }

        // 内存保留 10240 行数据，多余的刷新到固化存储
        workbook = new SXSSFWorkbook(cacheRowsInMemory);
        sheetNum = 0;
        sheet = newSheet();

        // 写数据
        /**
         * 要考虑的情况：
         * 1、类型，
         * 2、对齐，
         * 3、边框，
         * 4、合并
         */
        ExcelCell excelCell;    // cell 对象
        Object content;            // cell 内容
        HorizontalAlignment align;            // cell 对齐方式【默认居中】
        boolean border;            // cell 边框【默认有边框】

        // 对所有行对象进行循环
        for (ExcelRow excelRow : rows) {

            // sheet 号不一样的时候，需要添加一页
            if (!excelRow.getSheetNum().equals(sheetNum)) {
                sheetNum = excelRow.getSheetNum();
                rowNum = 0;
                sheet = newSheet();
            }

            // 列号
            int colNum = 0;
            SXSSFRow row = sheet.getRow(rowNum);
            if (row == null) {
                row = sheet.createRow(rowNum);
            }

            // 对所有的cell 对象进行循环【在设置表格的时候，若有合并的cell，会自动跳过】
            int size = excelRow.size();
            for (int j = 0; j < size; j++) {

                // 当前单元格，只用于cell的宽度设定。col_num 将在使用完后就指定下一cell
                int nowCell = colNum;

                excelCell = excelRow.get(j);
                content = excelCell.getCellContent();
                align = excelCell.getAlign();
                border = excelCell.getBorder();
                int colMerge = excelCell.getCol();
                if (colMerge < 1) {
                    colMerge = 1;
                }
                int rowMerge = excelCell.getRow();
                if (rowMerge < 1) {
                    rowMerge = 1;
                }

                // 若有创建cell 直接获取，否则，新建【新建cell ，合并，设置边框，这些都将只有在新建的时候进行操作，之后只是跳过相应的cell】
                SXSSFCell cell = row.getCell(colNum);
                if (cell == null) {
                    cell = row.createCell(colNum, CellType.NUMERIC);
                    // 合并单元格
                    mergeCell(this, colMerge, rowMerge, colNum, border);
                    // 列号向前
                    colNum += colMerge;
                } else {
                    // 如果cell已经有数据了，找到下一个空的cell，需要在空的 cell 继续创建数据
                    colNum = this.getCell(row, colNum);
                    // 当前列号需要更新
                    nowCell = colNum;
                    cell = row.createCell(colNum, CellType.NUMERIC);
                    // 合并单元格
                    mergeCell(this, colMerge, rowMerge, colNum, border);
                    colNum++;
                }

                // 空
                if (content == null) {
                    content = "";
                }

                // Integer
                if (content instanceof Integer) {
                    ExcelUtil.setIntStrStyle(this, cell, align, border);
                    cell.setCellValue((Integer) content);
                    continue;
                }

                // Double
                if (content instanceof Double) {
                    ExcelUtil.setDoubleStyle(this, cell, align, border);
                    cell.setCellValue((Double) content);
                    // 列不合并才自动宽度
                    if (colMerge == 1) {
                        ExcelUtil.setWidth(sheet, nowCell, content.toString());
                    }
                    continue;
                }

                // 时间【不能把 java.sql.Date 的时间算在内】：java.util.Date
                if (!(content instanceof java.sql.Date) && (content instanceof Date)) {
                    content = ExcelUtil.SDF_DATE_TIME.format(content);
                    ExcelUtil.setDateTimeStyle(this, cell, align, border);
                    try {
                        cell.setCellValue(ExcelUtil.SDF_DATE_TIME.parse(content.toString()));
                    } catch (ParseException e) {
                        e.printStackTrace();
                    }
                    // 列不合并才自动宽度
                    if (colMerge == 1) {
                        ExcelUtil.setWidth(sheet, nowCell, content.toString());
                    }
                    continue;
                }

                // java.sql.Date
                if ((content instanceof Date)) {
                    content = ExcelUtil.SDF_DATE.format(content);
                    ExcelUtil.setDateStyle(this, cell, align, border);
                    try {
                        cell.setCellValue(ExcelUtil.SDF_DATE.parse(content.toString()));
                    } catch (ParseException e) {
                        e.printStackTrace();
                    }
                    // 列不合并才自动宽度
                    if (colMerge == 1) {
                        ExcelUtil.setWidth(sheet, nowCell, content.toString());
                    }
                    continue;
                }

                // 最后，当字符串处理
                cell = row.createCell(nowCell, CellType.STRING);
                ExcelUtil.setIntStrStyle(this, cell, align, border);
                cell.setCellValue(content.toString());

                // 如果文字内容太长了，设置为富文本类型
                // 列不合并才自动宽度
                if (colMerge == 1) {
                    boolean tooLong = ExcelUtil.setWidth(sheet, nowCell, content.toString());
                    if (tooLong) {
                        ExcelUtil.setWrapTextStyle(this, cell, align, border);
                    }
                }

            }
            rowNum++;
        }
    }


    private SXSSFSheet newSheet() {

        String title = this.title;
        if (sheetNum != 0) {
            title = title + "_" + (sheetNum + 1);
        }
        SXSSFSheet sheet = workbook.createSheet(title);
        style = new ExcelStyle();

        // title
        if (title != null && !"".equals(title.trim())) {
            sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 0, width - 1));
            SXSSFRow rowTitle = sheet.createRow(rowNum++);
            SXSSFCell cellTitle = rowTitle.createCell(0);
            cellTitle.setCellStyle(style.getStyleTitle(this));
            cellTitle.setCellValue(title);
        }

        // infomation of this excel
        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 0, 2));
        SXSSFRow rowInfo = sheet.createRow(rowNum);
        SXSSFCell cellTime = rowInfo.createCell(0);
        cellTime.setCellStyle(style.getStyleStrLeftNoBorder(this));
        cellTime.setCellValue("创建时间：" + ExcelUtil.SDF_DATE_TIME.format(new Date()));

        // create_by
        if (createBy != null) {
            sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 3, 5));
            SXSSFCell cellCreateBy = rowInfo.createCell(3);
            cellCreateBy.setCellStyle(style.getStyleStrLeftNoBorder(this));
            cellCreateBy.setCellValue("创建人：" + createBy);
        }

        rowNum++;

        // date_from_to
        String dateInfo = "";
        if (dateFrom != null && dateTo != null) {
            dateInfo = "时间范围：从" + dateFrom + "到" + dateTo;
        }
        if (dateFrom != null && dateTo == null) {
            dateInfo = "时间：" + dateFrom;
        }
        if (dateFrom == null && dateTo != null) {
            dateInfo = "时间：" + dateTo;
        }

        if (!"".equals(dateInfo)) {
            sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 0, 5));
            SXSSFRow rowDateInfo = sheet.createRow(rowNum);
            SXSSFCell cellDateInfo = rowDateInfo.createCell(0);
            cellDateInfo.setCellStyle(style.getStyleStrLeftNoBorder(this));
            cellDateInfo.setCellValue(dateInfo);
            rowNum++;
        }


        // 写 列名称
        if (header != null && header.size() > 0) {
            SXSSFRow rowHeader = sheet.createRow(rowNum++);
            for (int i = 0; i < header.size(); i++) {
                SXSSFCell cellHeader = rowHeader.createCell(i);
                cellHeader.setCellStyle(style.getStyleHeader(this));
                cellHeader.setCellValue(header.get(i));
                ExcelUtil.setWidth(sheet, i, header.get(i));
            }
        }

        this.addSheet(sheet);
        return sheet;
    }


    /**
     * Description 递归找到空的cell
     * create @ 2017-07-15 21:09:00
     *
     * @param row    row
     * @param colNum colNum
     * @author wangkc admin@wkclz.com
     */
    private int getCell(SXSSFRow row, int colNum) {
        if (row.getCell(colNum) != null) {
            return this.getCell(row, colNum + 1);
        }
        return colNum;
    }

    private void mergeCell(Excel excel, int colMerge, int rowMerge, int colNum, boolean border) {
        // 检查是否需要合并单元格
        if (colMerge > 1 || rowMerge > 1) {
            sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum + rowMerge - 1, colNum, colNum + colMerge - 1));

            //预设内容
            for (int x = rowNum; x < rowNum + rowMerge; x++) {
                for (int y = colNum; y < colNum + colMerge; y++) {
                    SXSSFRow r = sheet.getRow(x);
                    if (r == null) {
                        r = sheet.createRow(x);
                    }
                    SXSSFCell c = r.getCell(y);
                    if (c == null) {
                        c = r.createCell(y);
                    }

                    // 是否需要设置边框
                    if (border) {
                        c.setCellStyle(excel.getStyle().getStyleNumCenterWithBorder(excel));
                    } else {
                        c.setCellStyle(excel.getStyle().getStyleNumCenterNoBorder(excel));
                    }
                }
            }
        }
    }

}
