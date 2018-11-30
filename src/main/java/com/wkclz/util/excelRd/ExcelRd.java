package com.wkclz.util.excelRd;

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


import java.io.*;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRd extends ExcelRdContent {

    private boolean xls = false;

    private static final String DOT_XLS = ".xls";
    private static final String DOT_XLSX = ".xlsx";

    public ExcelRd(String xlsxPath) {
        super();
        // 03版本的excel要特别标明
        if (xlsxPath.endsWith(DOT_XLS)) {
            xls = true;
        }
        try {
            this.is = new FileInputStream(xlsxPath);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    public ExcelRd(File file) {
        super();
        // 03版本的excel要特别标明
        if (file.getAbsolutePath().endsWith(DOT_XLS)) {
            xls = true;
        }
        try {
            this.is = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    public ExcelRd(InputStream ins, ExcelRdVersionEnum version) {
        super();
        // 03版本的excel要特别标明
        if (version == ExcelRdVersionEnum.XLS) {
            xls = true;
        }
        this.is = (FileInputStream) ins;
    }

    public ExcelRd(FileInputStream fins, ExcelRdVersionEnum version) {
        super();
        // 03版本的excel要特别标明
        if (version == ExcelRdVersionEnum.XLS) {
            xls = true;
        }
        this.is = fins;
    }


    public List<ExcelRdRow> analysisXlsx() throws ExcelRdException, IOException {

        List<ExcelRdTypeEnum> types = getTypes();
        if (types == null || types.size() == 0) {
            throw new ExcelRdException("Types of the data must be set");
        }

        if (xls) {
            this.workbook03 = new HSSFWorkbook(this.is);
        } else {
            this.workbook07 = new XSSFWorkbook(this.is);
        }


        // 当前只考虑识别一个 sheet
        if (xls) {
            this.sheet03 = this.workbook03.getSheetAt(0);
        } else {
            this.sheet07 = this.workbook07.getSheetAt(0);
        }

        // 循环所有【右边的边界】
        int right = getStartCol() + types.size();
        // 阈值【当连续取到三个空行，或者连续取到 3 * size 个空 cell 时，将会退出检测】
        int rowThreshold = 0;
        // 阈值【当连续取到三个空行，或者连续取到 3 * size 个空 cell 时，将会退出检测】
        int colThreshold = 0;

        for (int i = getStartRow(); ; i++) {

            // 阈值【当连续取到三个空行，或者连续取到 3 * size 个空 cell 时，将会退出检测】
            if (rowThreshold >= 3 || colThreshold >= 3 * types.size()) {
                break;
            }

            if (xls) {
                this.row03 = this.sheet03.getRow(i);
            } else {
                this.row07 = this.sheet07.getRow(i);
            }

            if (this.row03 == null && this.row07 == null) {
                rowThreshold++;
                continue;
            }
            rowThreshold = 0;

            ExcelRdRow excelRdRow = new ExcelRdRow();
            for (int j = getStartCol(); j < right; j++) {

                if (xls) {
                    this.cell03 = this.row03.getCell(j);
                } else {
                    this.cell07 = this.row07.getCell(j);
                }

                if (this.cell03 == null && this.cell07 == null) {
                    colThreshold++;
                    excelRdRow.addCell("");
                } else {
                    colThreshold = 0;
                    Object cellValue;

                    if (xls) {
                        cellValue = ExcelRdUtil.getCellValue(this.cell03, types.get(j - getStartCol()));
                    } else {
                        cellValue = ExcelRdUtil.getCellValue(this.cell07, types.get(j - getStartCol()));
                    }

                    excelRdRow.addCell(cellValue);
                }
            }

            // 如果row全部为null，将不加入结果
            List<Object> rtRow = excelRdRow.getRow();
            int size = rtRow.size();
            for (Object object : rtRow) {
                if (object == null || "".equals(object.toString().trim())) {
                    size--;
                }
            }
            if (size != 0) {
                addRow(excelRdRow);
            }
        }
        return getRows();
    }
}
