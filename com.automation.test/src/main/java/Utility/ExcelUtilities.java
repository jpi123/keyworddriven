package Utility;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.commons.lang3.ObjectUtils.Null;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ExecutionEngine.DriverScript;

public class ExcelUtilities {

	
	private static XSSFSheet ExcelWSheet;
	private static XSSFWorkbook ExcelWBook;
	private static org.apache.poi.ss.usermodel.Cell Cell;


	public static void setExcelFile(String Path, String SheetName) throws Exception {
		FileInputStream ExcelFile = new FileInputStream(Path);
		ExcelWBook = new XSSFWorkbook(ExcelFile);
		ExcelWSheet = ExcelWBook.getSheet(SheetName);
	}


	public static String getCellData(int RowNum, int ColNum) throws Exception {
		Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
		String CellData = Cell.getStringCellValue();
		return CellData;

	}
	
	
	public static void setCellData(String abc,int RowNum, int ColNum, String SheetName) throws Exception {
		try {
		ExcelWSheet = ExcelWBook.getSheet(SheetName);
		Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
		XSSFRow Row = ExcelWSheet.getRow(RowNum);
		Cell = Row.getCell(ColNum,Row.RETURN_BLANK_AS_NULL);
		
		if(Cell == null)
		{
			Cell = Row.createCell(ColNum);
			Cell.setCellValue(abc);
		}
		else 
		{
			Cell.setCellValue(abc);
		}
		 FileOutputStream fileOut = new FileOutputStream(DriverScript.sPath);
		 ExcelWBook.write(fileOut);
		
		 fileOut.close();
		 ExcelWBook = new XSSFWorkbook(new FileInputStream(DriverScript.sPath));
		 System.out.println("End1");
		 }
		catch(Exception e){
		 DriverScript.flag = false;
		
		}
	

	}
}
