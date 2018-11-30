package com.wkclz.util.excelRd;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;


public class ExcelRdUtil {


    protected static Object getCellValue(XSSFCell cell, ExcelRdTypeEnum type) {
        if (cell == null || "".equals(cell.toString().trim())) {
            return null;
        }
        CellType cellType = cell.getCellTypeEnum();
        if (type == ExcelRdTypeEnum.INTEGER && cellType == CellType.NUMERIC) {
            Double numeric = cell.getNumericCellValue();
            return numeric.intValue();
        }
        if (type == ExcelRdTypeEnum.LONG && cellType == CellType.NUMERIC) {
            Double numeric = cell.getNumericCellValue();
            return numeric.longValue();
        }
        if (type == ExcelRdTypeEnum.DOUBLE && cellType == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        if (type == ExcelRdTypeEnum.DATETIME && cellType == CellType.NUMERIC) {
            return cell.getDateCellValue();
        }
        if (type == ExcelRdTypeEnum.DATE && cellType == CellType.NUMERIC) {
            return new java.sql.Date(cell.getDateCellValue().getTime());
        }
        if (type == ExcelRdTypeEnum.STRING && cellType == CellType.STRING) {
            return cell.getStringCellValue();
        }
        return cell.toString();
    }

    protected static Object getCellValue(HSSFCell cell, ExcelRdTypeEnum type) {
        if (cell == null || "".equals(cell.toString().trim())) {
            return null;
        }
        CellType cellType = cell.getCellTypeEnum();
        if (type == ExcelRdTypeEnum.INTEGER && cellType == CellType.NUMERIC) {
            Double numeric = cell.getNumericCellValue();
            return numeric.intValue();
        }
        if (type == ExcelRdTypeEnum.LONG && cellType == CellType.NUMERIC) {
            Double numeric = cell.getNumericCellValue();
            return numeric.longValue();
        }
        if (type == ExcelRdTypeEnum.DOUBLE && cellType == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        if (type == ExcelRdTypeEnum.DATETIME && cellType == CellType.NUMERIC) {
            return cell.getDateCellValue();
        }
        if (type == ExcelRdTypeEnum.DATE && cellType == CellType.NUMERIC) {
            return new java.sql.Date(cell.getDateCellValue().getTime());
        }
        if (type == ExcelRdTypeEnum.STRING && cellType == CellType.STRING) {
            return cell.getStringCellValue();
        }
        return cell.getStringCellValue();
    }


}
