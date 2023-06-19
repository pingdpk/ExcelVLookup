package org.testng;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	public String excelpath;
	public String sheetname1;
	public String sheetname2;
	public XSSFWorkbook workbook;
	public XSSFSheet sheet1 = null;
	public XSSFSheet sheet2 = null;
	public XSSFRow row   =null;
	public XSSFCell cell = null;

	// Parameterized constructor
	public ExcelReader(String excelpath) {
		try {
			workbook = new XSSFWorkbook(new FileInputStream(excelpath));
		}
		catch(IOException e) {
			e.printStackTrace();
		}
	}

	// Get the total row count of the excel sheet
	public int getRowCount() {
		int rowCount =0;
		try {
			rowCount = sheet1.getPhysicalNumberOfRows();
			//System.out.println("Total number of rows =" + rowCount);			
		}
		catch(Exception e)
		{
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}			
		return rowCount;		
	}

	// Get the total column count of the excel sheet
	public int getColumnCount() {
		int colCount = 0;
		try {
			/* To reduce the complexity , finding the column count for 1st row in the excel sheet
			 * Here the count starts from Zero(0)
			 */
			colCount = sheet1.getRow(0).getPhysicalNumberOfCells();
			//System.out.println("Total number of columns =" + colCount);			
		}
		catch(Exception e)
		{
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
		return colCount;

	}

	// returns the data from a cell
	public String getCellData(int rowNum,int colNum){
		//cell = null ;
		try{			
			row = sheet1.getRow(rowNum);
			cell = row.getCell(colNum);		

			switch(cell.getCellType())
			{
			case STRING:
				return cell.getStringCellValue();

			case NUMERIC:			
				// 600100.00 -> 600100 -> "600100"
				String cellText = String.valueOf((int)cell.getNumericCellValue());
				return cellText;

			case BOOLEAN:
				// true -> "true"
				String boolcell = String.valueOf(cell.getBooleanCellValue());
				return boolcell;

			case BLANK:
				return "";

			default: 
				return "";
			}	
		}
		catch(Exception e){
			e.printStackTrace();	
			return "row "+rowNum+" or column "+colNum +" does not exist  in xlsx";
		}

	}	

}