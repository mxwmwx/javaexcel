package new_read;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jxl.Workbook;

public class main extends MonthSheetTable{

	public static void main(String[] args) {
		XSSFWorkbook wbWorkbook = MonthSheetTable.createSheetTableWithOther("1", "1", 32, 12);
		FileOutputStream fout=null;
		double ram = Math.random();
		try {
            fout = new FileOutputStream("D:\\mx\\uu"+ram+".xlsx");
            wbWorkbook.write(fout);
            fout.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
	}
}
