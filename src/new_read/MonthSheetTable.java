package new_read;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.xb.xsdschema.Public;


public class MonthSheetTable extends OutputTable{
	public static String tableHeadString = "中国建筑项目管理表格";
	public static String imagePath = "photo path";
	public static String tableNumNameString = "表格编号";
	public static String tableNum = "CSCEC83-ME-B30902";
	public static String tableName = "物资采购月度统计表";
	public static String acquireCompanyString = "需方";
	public static String acquireCompany = "中建八局三公司徐州分公司";
	public static String fromCompanyString = "供方";
	public static String fromCompany = "根据原始表按需获取";//根据原始表获取
	public static String hetonNumString = "合同编号";//暂定
	public static String hetongNum = "合同编号待定...";
	public static String proNameAndNum = "项目名称及编码";
	public static String proNameAndNumString = "金融集聚区一期A9-3项目  编号1220117669";//暂定
	public static String number = "编号";
	public static String dateString = "日期";
	public static String dateString2 = "2001....待定";
	public static String unitPriceString = "单价:元";
	
	public static String str[] = {
			"序号",
			"材料名称",
			"规格型号",
			"单位",
			"供货数量",
			"供货时间",
			"不含税单价",
			"不含税金额",
			"税额",
			"税率",
			"含税金额",
			"备注"
	};
	public static XSSFWorkbook createSheetTable() {
		return OutputTable.createTable("1", 30, 12, str,12);
	}

	public static XSSFWorkbook createSheetTableWithOther(String fromCom,String hetongNum,int rowNum,int colNum) {
		
		XSSFWorkbook wbWorkbook = new XSSFWorkbook();
		XSSFSheet sheet = wbWorkbook.createSheet(tableName);
		
		sheet.setColumnWidth(3, 4000);
//		XSSFCellStyle cellStyle = wbWorkbook.createCellStyle();
//        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
//        cellStyle.setAlignment(HorizontalAlignment.CENTER);
//        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        
      
        
   
		for(int i = 0;i<rowNum ;i++) {
			Row row = sheet.createRow(i);//表头
			for (int j = 0;j<colNum;j++) {
			
				Cell cell = row.createCell(j);//使用样式
				//cell.setCellStyle(cellStyle);
				if(i==0 && j==0) {
					cell.setCellValue(imagePath);//图片
				}
				if(i==0 && j== 2) {
					cell.setCellValue(tableHeadString);//中国建筑项目管理表格
					row.setHeightInPoints(20);
					OutputTable.fontStyle(cell, wbWorkbook, "黑体", 10);
					
				}
				if(i==1 && j==2) {
					cell.setCellValue(tableName);//物资采购月度统计表
				}
				if(i==1 && j==10) {
					cell.setCellValue(tableNumNameString);//表格编号"
				}
				if(i==2 && j==10) {
					cell.setCellValue(tableNum);//CSCEC83-ME-B30902
				}
				if(i==3 && j==0) {
					cell.setCellValue(acquireCompanyString);//需方
				}
				if(i==3 && j==2) {
					cell.setCellValue(acquireCompany);//中建八局三公司徐州分公司
				}
				if(i==4 && j==0) {
					cell.setCellValue(fromCompanyString);//供方
				}
				if(i==4 && j==2) {
					cell.setCellValue(fromCompany);
				}
				if(i==3 && j==7) {
					cell.setCellValue(dateString);//日期
				}
				if(i==3 && j==10) {
					cell.setCellValue(dateString2);
				}
				if(i==3 && j==11) {
					cell.setCellValue(unitPriceString);//单价：元
				}
				if(i==5 && j==0) {
					cell.setCellValue(proNameAndNum);
				}
				if(i==5 && j==2) {
					cell.setCellValue(proNameAndNumString);
				}
				if(i == 6) {
					cell.setCellValue(str[j]);
				}
								
			}
		}
	    
		OutputTable.merged(sheet, 0, 2, 0, 1);//图
		OutputTable.merged(sheet, 0, 0, 2, 11);//中建表头
		OutputTable.merged(sheet, 1, 2, 2, 9);//采购表头
		OutputTable.merged(sheet, 1, 1, 10, 11);//表格编号
		OutputTable.merged(sheet, 2, 2, 10, 11);//编号
		OutputTable.merged(sheet, 3, 3, 0, 1);//需方
		OutputTable.merged(sheet, 3, 3, 2, 6);//具体需方
		OutputTable.merged(sheet, 3, 3, 7, 9);//日期
		OutputTable.merged(sheet, 4, 4, 0, 1);//供方
		OutputTable.merged(sheet, 4, 4, 2, 5);//具体供方
		OutputTable.merged(sheet, 4, 4, 7, 11);//合同编号
		OutputTable.merged(sheet, 5, 5, 0, 1);//项目名称及编码
		OutputTable.merged(sheet, 5, 5, 2, 8);//金融集聚区一期A9-3项目  编号1220117669
		
		
		return wbWorkbook;	
	}

}
