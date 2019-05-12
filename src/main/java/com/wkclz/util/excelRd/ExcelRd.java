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


import com.wkclz.util.excelRd.domain.ExcelRdSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

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

    /**
     * 第一个 sheet 识别
     * @return
     * @throws ExcelRdException
     */
    @Deprecated
    public List<List<Object>> analysisXlsx() throws ExcelRdException {
        initWorkbook();
        ExcelRdSheet excelRdSheet = analysisXlsxSheet(this, sheets.get(0));
        return excelRdSheet.getRows();
    }

    /**
     * 第一个 sheet 识别
     * @return
     * @throws ExcelRdException
     */
    public ExcelRdSheet analysisFirstSheet() throws ExcelRdException {
        initWorkbook();
        ExcelRdSheet excelRdSheet = analysisXlsxSheet(this, sheets.get(0));
        return excelRdSheet;
    }


    /**
     * 识别excel 多 sheet
     * @return
     * @throws ExcelRdException
     */
    public List<ExcelRdSheet> analysis() throws ExcelRdException {
        initWorkbook();
        List<ExcelRdSheet> result = new ArrayList<ExcelRdSheet>();
        for (ExcelRdSheet sheet: sheets) {
            ExcelRdSheet excelRdSheet = analysisXlsxSheet(this, sheet);
            result.add(excelRdSheet);
        }
        return result;
    }


    private static ExcelRdSheet analysisXlsxSheet(ExcelRd excelRd, ExcelRdSheet sheet) throws ExcelRdException {
        List<ExcelRdTypeEnum> types = sheet.getTypes();
        if (types == null || types.size() == 0) {
            throw new ExcelRdException("Types of the data must be set @ sheet " + sheet.getSheet() + ", " + sheet.getSheetName());
        }

        // 当前只考虑识别一个 sheet
        if (excelRd.xls) {
            excelRd.sheet03 = excelRd.workbook03.getSheetAt(sheet.getSheet());
            sheet.setSheetName(excelRd.sheet03.getSheetName());
        } else {
            excelRd.sheet07 = excelRd.workbook07.getSheetAt(sheet.getSheet());
            sheet.setSheetName(excelRd.sheet07.getSheetName());
        }

        // 循环所有【右边的边界】
        int right = sheet.getStartCol() + types.size();
        // 阈值【当连续取到三个空行，或者连续取到 3 * size 个空 cell 时，将会退出检测】
        int rowThreshold = 0;
        // 阈值【当连续取到三个空行，或者连续取到 3 * size 个空 cell 时，将会退出检测】
        int colThreshold = 0;

        for (int i = sheet.getStartRow(); ; i++) {

            // 阈值【当连续取到三个空行，或者连续取到 3 * size 个空 cell 时，将会退出检测】
            if (rowThreshold >= 3 || colThreshold >= 3 * types.size()) {
                break;
            }

            if (excelRd.xls) {
                excelRd.row03 = excelRd.sheet03.getRow(i);
            } else {
                excelRd.row07 = excelRd.sheet07.getRow(i);
            }

            if (excelRd.row03 == null && excelRd.row07 == null) {
                rowThreshold++;
                continue;
            }
            rowThreshold = 0;

            List<Object> excelRdRow = new ArrayList<Object>();
            for (int j = sheet.getStartCol(); j < right; j++) {

                if (excelRd.xls) {
                    excelRd.cell03 = excelRd.row03.getCell(j);
                } else {
                    excelRd.cell07 = excelRd.row07.getCell(j);
                }

                if (excelRd.cell03 == null && excelRd.cell07 == null) {
                    colThreshold++;
                    excelRdRow.add("");
                } else {
                    colThreshold = 0;
                    Object cellValue;

                    if (excelRd.xls) {
                        cellValue = ExcelRdUtil.getCellValue(excelRd.cell03, types.get(j - sheet.getStartCol()));
                    } else {
                        cellValue = ExcelRdUtil.getCellValue(excelRd.cell07, types.get(j - sheet.getStartCol()));
                    }

                    excelRdRow.add(cellValue);
                }
            }

            // 如果row全部为null，将不加入结果
            int size = excelRdRow.size();
            for (Object object : excelRdRow) {
                if (object == null || "".equals(object.toString().trim())) {
                    size--;
                }
            }
            if (size != 0) {
                sheet.addRow(excelRdRow);
            }
        }

        return sheet;
    }

    /**
     * 初始化 Workbook
     */
    private void initWorkbook() throws ExcelRdException {
        int numberOfSheets = 0;
        try {
            if (this.xls) {
                this.workbook03 = new HSSFWorkbook(this.is);
                numberOfSheets = this.workbook03.getNumberOfSheets();
            } else {
                this.workbook07 = new XSSFWorkbook(this.is);
                numberOfSheets = this.workbook07.getNumberOfSheets();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        if (this.sheets == null || this.sheets.size() == 0){
            throw new ExcelRdException("Sheets of the data must be set");
        }
        if (this.sheets.size() > numberOfSheets){
            throw new ExcelRdException(numberOfSheets + " Sheets was found in excel, but you set " + this.sheets.size() + " Sheets for analysis!");
        }
    }

}
