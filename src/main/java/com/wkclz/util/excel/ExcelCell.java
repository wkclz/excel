package com.wkclz.util.excel;

import lombok.Data;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

@Data
public class ExcelCell {

    /**
     * cell内容
     */
    private Object cellContent;
    /**
     * 是否有边框
     */
    private Boolean border;
    /**
     * 对齐方式，只有左和中，默认中对齐，从ExcelUtil 取值
     */
    private HorizontalAlignment align;
    /**
     * 合并列数【宽度】
     */
    private int col;
    /**
     * 合并行数【高度】
     */
    private int row;

    /**
     * 字体
     */
    private Font font;

    /**
     * 创建 cell
     *
     * @param cellContent cellContent
     * @param border      border
     * @param align       align
     * @param col         col
     * @param row         row
     */
    protected ExcelCell(Object cellContent, boolean border, HorizontalAlignment align, int col, int row, Font font) {
        super();
        this.cellContent = cellContent;
        this.border = border;
        this.align = align;
        this.col = col;
        this.row = row;
    }

    public Boolean getBorder() {
        if (border == null) {
            border = true;
        }
        return border;
    }

}
