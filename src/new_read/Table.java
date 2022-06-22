package new_read;



import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import jxl.Sheet;
import jxl.Workbook;

public class Table extends RowItems{
	
    public  Workbook getTableWorkbook() {
    	return RowItems.getRowItemsWorkbook();//excel
    }
    public Sheet getTableSheet() {
		return RowItems.getRowItemsSheet();//sheet
	}
    
    public static int getRowcount() {
    	return  RowItems.getRowItemsSheet().getRows();
    	//行数
    }
    public static int getColCount() {
    	return RowItems.getRowItemsSheet().getColumns();
    	//列数
    }
}
