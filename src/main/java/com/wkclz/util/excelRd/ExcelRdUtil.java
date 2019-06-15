package com.wkclz.util.excelRd;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;


public class ExcelRdUtil {

    protected static Object getCellValue(Cell cell, ExcelRdTypeEnum type) {

        if (cell == null || "".equals(cell.toString().trim())) {
            return null;
        }

        // 可预测的类型
        CellType cellType = cell.getCellType();
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

        // 不可预测的类型
        if (cellType == CellType.STRING) {
            return cell.getStringCellValue();
        }
        if (cellType == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        if (cellType == CellType.FORMULA) {
            return cell.getNumericCellValue();
        }
        if (cellType == CellType.BOOLEAN) {
            return cell.getBooleanCellValue();
        }
        if (cellType == CellType.ERROR) {
            return cell.getErrorCellValue();
        }
        if (cellType == CellType.BLANK) {
            return null;
        }
        if (cellType == CellType._NONE) {
            return null;
        }

        return cell.toString();
    }

}
