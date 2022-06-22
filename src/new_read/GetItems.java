package new_read;

import jxl.Sheet;
import jxl.Workbook;

public class GetItems extends FileRead{
	
	public static Workbook getWorkbook() {
		return  FileRead.getFile();	
	}
	public static Sheet getFixedSheet() {
		return FileRead.getSheet();		
	}
}
