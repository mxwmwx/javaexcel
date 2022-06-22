package new_read;

import java.util.Date;

import jxl.Sheet;
import jxl.Workbook;

public class RowItems extends GetItems{
		
	public static String proName;//项目名称0
	public Date proDate;//日期1
	public int invoiceNum;//单据编号2
	public String invoiceType;//单据类型3
	public String itemType;//物资类别4
	public String itemName;//物资名称5
	public String itemSize;//规格型号6
	public String itemUnit;//单位7
	public Float actualAcceptNum;//实际收取数量8
	public Float singleMoney;//不含税单价9
	public Float totalMoney;//不含税金额10
	public Float itemRate;//税率11
	public Float itemRateMoney;//税额12
	public Float itemMoneyWithRateMney;//含税金额13
	public String itemLogo;//物资品牌14
	public String itemSource;//物资来源15
	public String itemTakeGroup;//领用班组16
	public String itemUseArea;//使用部位17
	public String workType;//施工类别18
	public String costDivision;//费用划分19
	public String financialAccountTime;//入账时间20
	public Float caijiduizhangjine;//集采对账金额21
	public String others;//备注22
	
	public String informationString;
	
	//getItems newItem = new getItems(); 
	public String getItemInformation(int row, int col) {
		this.informationString = GetItems.getSheet().getCell(row, col).getContents();
		return  informationString;
	}
	public  static Workbook getRowItemsWorkbook() {
		return GetItems.getFile();
	}
	public static Sheet getRowItemsSheet() {
	 	return GetItems.getSheet();	 
	}
}
