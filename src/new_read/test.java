package new_read;

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class test extends Table{
	

	public static void main(String[] args) {
		// TODO Auto-generated method stub
        int i;
        Sheet sheet0;
        Workbook book;
        Cell cell1 = null ;
		Cell cell2 = null ,cell16 = null ;
        
        File file = new File("C:\\Users\\mx\\Desktop\\data.xls");
        //获取行数和列数

        	 try {
				book = Workbook.getWorkbook(file);
				
				sheet0 = book.getSheet(0);
				
				 i = 0;
				 //获取行列数
				 int row = sheet0.getRows();
				 int cell = sheet0.getColumns();
				 System.out.println(row);
				 System.out.println(cell);
				 
				 cell1 = sheet0.getCell(0,i);
				 cell2 = sheet0.getCell(1,i);
	        	 cell16 = sheet0.getCell(15,i);
//	        		 cell1 = sheet.getCell(0,i);
//	        		 Cell cell2 = sheet.getCell(1, i);
//	        		 System.out.println(cell1.getContents().toString());
//	        		 System.out.println(cell2.getContents());
	        		
	        		 
	        	
			} catch (BiffException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
        	 
         XSSFWorkbook wb = new XSSFWorkbook();//创建excel表格
         XSSFSheet sheet1 = wb.createSheet("sheet1");//创建sheet
         XSSFCellStyle cellStyle = wb.createCellStyle();
         cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
         
         Row row = ((XSSFSheet) sheet1).createRow(0);//创建行
         //4、创建格
         org.apache.poi.ss.usermodel.Cell cell = row.createCell(0);
         cell.setCellValue(cell2.getContents());
         
//         sheet1.addMergedRegion(new CellRangeAddress(0,2,2,4));
         //Table.merged(sheet1,1 , 2, 1, 3);
         cell = row.createCell(1);
         cell.setCellValue(cell16.getContents());
         if(cell16.getContents().equals("物资来源")) {
        	 System.out.println("get it");
        	 //System.out.println(sheet0.getMergedCells().toString());
         }
        
         cell = row.createCell(2);
         cell.setCellValue("时间");
         sheet1.getRow(0).getCell(2).setCellStyle(cellStyle);
         FileOutputStream fout=null;
         
         try {
             fout = new FileOutputStream("D:\\mx\\uu2212.xlsx");
             wb.write(fout);
             fout.close();
         } catch (IOException e) {
             e.printStackTrace();
         }

            	 
         
	}

}
