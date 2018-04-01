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

import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public abstract class ExcelBase {

	protected static SimpleDateFormat sdf_dateTime = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	protected static SimpleDateFormat sdf_date = new SimpleDateFormat("yyyy-MM-dd");
	
	/** excel 工作簿 */
	private SXSSFWorkbook workbook;
	/** sheet */
	private SXSSFSheet sheet;
	/** 行号 */
	protected int rownum = 0;
	
	protected SXSSFWorkbook getWorkbook() {
		return workbook;
	}
	protected void setWorkbook(SXSSFWorkbook workbook) {
		this.workbook = workbook;
	}
	protected SXSSFSheet getSheet() {
		return sheet;
	}
	protected void setSheet(SXSSFSheet sheet) {
		// 不显示风格线
		sheet.setDisplayGridlines(false);
		PrintSetup ps = sheet.getPrintSetup();
		// 打印方向，true：横向，false：纵向(默认)
		ps.setLandscape(false);
		// A4纸
		ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
		this.sheet = sheet;
	}
}
