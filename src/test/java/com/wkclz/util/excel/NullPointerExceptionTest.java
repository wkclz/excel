package com.wkclz.util.excel;

import java.util.List;

public class NullPointerExceptionTest {

    public static void main(String[] args) {
        Excel excel = new Excel();
        List<ExcelRow> rows = excel.rows;
        System.out.println(rows.size());

        /**
         * 9行断点，则 10 行输出 0
         * 9行取消断点，则 10 行 NullPointerException
         */


    }

}
