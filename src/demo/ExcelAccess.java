package demo;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelAccess 
{
	
	static XSSFSheet ExcelWSheet;	 
	static XSSFWorkbook ExcelWBook;
	static XSSFCell Cell;
	static XSSFRow Row;


public static void setExcelFile(String Path,String SheetName) throws Exception 
{

		try 
		{

		FileInputStream ExcelFile = new FileInputStream(Path);
		ExcelWBook = new XSSFWorkbook(ExcelFile);
		ExcelWSheet = ExcelWBook.getSheet(SheetName);

		} 
		catch (Exception e)
		{
			throw (e);
		}
}

public static String getCellData(int RowNum, int ColNum) throws Exception
{
        try
		{
			Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
			String CellData = Cell.toString();
			//System.out.println(CellData);
			return CellData;
		}
        
		catch (Exception e)
		{
		throw (e);
		}
}

public static void setCellData(String Path, String Result,  int RowNum, int ColNum) throws Exception	
{

		try
		{

			Row  = ExcelWSheet.getRow(RowNum);
		Cell = Row.getCell(ColNum);

		if (Cell == null) 
		{
			Cell = Row.createCell(ColNum);
			Cell.setCellValue(Result);

		} 
		else 
		{
			Cell.setCellValue(Result);
		}

		FileOutputStream fileOut = new FileOutputStream(Path);
				ExcelWBook.write(fileOut);
				fileOut.flush();
				fileOut.close();
		}
		catch(Exception e)
		{
				throw (e);
		}

	}

}

