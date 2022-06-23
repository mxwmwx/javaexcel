package new_read;

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
    public static String getMessage(int row,int col) {
    	return RowItems.getItemInformation(row,col);
    }
}
