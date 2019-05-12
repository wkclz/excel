package com.wkclz.util.excelRd;


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


import com.wkclz.util.excelRd.domain.ExcelRdSheet;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public abstract class ExcelRdContent {


    /**
     *  Excek 读取过程使用到的全局变量，统一管理
     */
    /**
     * FileInputStream
     */
    protected FileInputStream is;

    /**
     * 07版本 excel
     */
    protected XSSFWorkbook workbook07;
    protected XSSFSheet sheet07;
    protected XSSFRow row07;
    protected XSSFCell cell07;

    /**
     * 03版本 excel
     */
    protected HSSFWorkbook workbook03;
    protected HSSFSheet sheet03;
    protected HSSFRow row03;
    protected HSSFCell cell03;


    /**
     * sheet 配置
     */
    protected List<ExcelRdSheet> sheets;



    public void setSheets(List<ExcelRdSheet> sheets) {
        for(int i = 0; i< sheets.size(); i++){sheets.get(i).setSheet(i);}
        this.sheets = sheets;
    }
    public void addSheets(ExcelRdSheet sheet) {
        if (this.sheets == null){
            this.sheets = new ArrayList<ExcelRdSheet>();
        }
        sheet.setSheet(this.sheets.size());
        this.sheets.add(sheet);
    }

    @Deprecated
    public void setStartRow(Integer startRow) {
        ExcelRdSheet rdSheet = init();
        rdSheet.setStartRow(startRow);
    }

    @Deprecated
    public void setStartCol(int startCol) {
        ExcelRdSheet rdSheet = init();
        rdSheet.setStartCol(startCol);
    }

    @Deprecated
    public void setTypes(List<ExcelRdTypeEnum> types) {
        ExcelRdSheet rdSheet = init();
        rdSheet.setTypes(types);
    }

    @Deprecated
    public void setTypes(ExcelRdTypeEnum[] types) {
        ExcelRdSheet rdSheet = init();
        if (rdSheet.getTypes() == null) {
            rdSheet.setTypes(new ArrayList<ExcelRdTypeEnum>());
        }
        for (ExcelRdTypeEnum type : types) {
            rdSheet.getTypes().add(type);
        }
    }



    // 初始化 config
    private ExcelRdSheet init(){
        if (this.sheets == null){
            this.sheets = new ArrayList<ExcelRdSheet>();
            ExcelRdSheet sheet = new ExcelRdSheet();
            sheet.setSheet(0);
            this.sheets.add(sheet);
        }
        ExcelRdSheet sheet = this.sheets.get(0);
        return sheet;
    }
}
