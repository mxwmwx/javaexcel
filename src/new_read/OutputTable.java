package new_read;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hslf.dev.SlideAndNotesAtomListing;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import jxl.Sheet;
import jxl.Workbook;

public class OutputTable extends Table{
	//获取表格信息
	public static String[][] getExcelMessage(){
		int row = Table.getRowcount();
		int col = Table.getColCount();
		String arr[][] = new String[row][col];
		for(int i = 0; i< row;i++) {
			for(int j = 0; j< col; j++) {
				arr[i][j] = Table.getMessage(i, j);
			}
		}
		return arr;
		
	}
	
	public static void createTable() {
		//创建表固定的
		XSSFWorkbook wb = new XSSFWorkbook();//创建excel表格
	    XSSFSheet sheet1 = wb.createSheet("sheet1");//创建sheet
	    Row row = ((XSSFSheet) sheet1).createRow(0);//创建行
		
		Cell cell = row.createCell(0);
		cell.setCellValue("1");//第一行第一列
		cell = row.createCell(1);//第一行第二列
		cell.setCellValue("2");
		
		FileOutputStream fout=null;
		double ram = Math.random();
		try {
            fout = new FileOutputStream("D:\\mx\\uu"+ram+".xlsx");
            wb.write(fout);
            fout.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
		
	}
	public static XSSFWorkbook createTable(String sheetName,int rowTotal,int colTotal,String[] str,int startRow) {
		//带参数的创建表
		XSSFWorkbook wbWorkbook = new XSSFWorkbook();
		XSSFSheet sheet = wbWorkbook.createSheet(sheetName);
		Row row = sheet.createRow(startRow);//表头
		for (int i = 0;i<str.length;i++) {
			Cell cell = row.createCell(i);
			cell.setCellValue(str[i]);
		}
		FileOutputStream fout=null;
		double ram = Math.random();
		try {
            fout = new FileOutputStream("D:\\mx\\uu"+ram+".xlsx");
            wbWorkbook.write(fout);
            fout.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
		return wbWorkbook;
	}	
	//合并单元格
    public static int merged(XSSFSheet sheet,int startRow,int endRow,int startCol,int endCol) {
    	return  sheet.addMergedRegion(new CellRangeAddress(startRow,endRow,startCol,endCol));
    	//合并单元格
    }
    public static XSSFWorkbook fontStyle( Cell cell,XSSFWorkbook wbWorkbook,String wordsKind,double wordSize) {
    	//字体样式
    	
    	XSSFCellStyle cellStyle2 = wbWorkbook.createCellStyle();
    	cellStyle2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        cellStyle2.setAlignment(HorizontalAlignment.CENTER);
        cellStyle2.setVerticalAlignment(VerticalAlignment.CENTER);
        
    	XSSFFont font = wbWorkbook.createFont();
    	font.setFontName(wordsKind);
		font.setFontHeight(wordSize);
		cellStyle2.setFont(font);
		cell.setCellStyle(cellStyle2);
		return wbWorkbook;
	}
    public static String getElemenString(int row,int col) {
    	return Table.getMessage(row, col);
    }
}
