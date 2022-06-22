package new_read;

import java.io.File;
import java.io.IOException;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
//文件载入类
public class FileRead {
	 File file = new File("C:\\Users\\mx\\Desktop\\data.xls");
	 //固定读取
	 public static Workbook getFile() {
		 File file1 = new File("C:\\Users\\mx\\Desktop\\data.xls");
		 try {
			return Workbook.getWorkbook(file1);
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	 }
	 
	 public static Sheet getSheet() {
		 //固定读取第一个表格
		 File file1 = new File("C:\\Users\\mx\\Desktop\\data.xls");
		 try {
			return Workbook.getWorkbook(file1).getSheet(0);
		} catch (IndexOutOfBoundsException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	 }
	 public static Workbook getFile(File file) {
		 //自定义路径读取
		 try {
			return Workbook.getWorkbook(file);
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	 }
    public static Sheet getSheet(Workbook book,int i) {
    	//读取自定义sheet
		return book.getSheet(i);	 
    }
}
