package org.testng;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import org.util.ConfigReader;
import org.util.ResultWriter;

import java.awt.print.Book;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class GetExcelData {

	@BeforeSuite
	public void setup(){
		ConfigReader.getAndSetConfigs();
	}

	@AfterSuite
	public void printResult(){
		if(!ResultWriter.errors.isEmpty()) {
			System.out.println("\n\n\t|||<------Errors----->|||\n");
			ResultWriter.errors.forEach((errors) -> System.out.println(errors));
		}
	}

	@Test (dataProvider="getData")
	public void patientList(String testCaseName, String patientName, String country)
	{
		System.out.println("");
	}

	//Getting excel data - passing every data as an object
	@DataProvider
	public Iterator<Object[]> getData() throws IOException {
		XSSFSheet sheet1 = null;
		XSSFSheet sheet2 = null;

		ExcelReader excel = new ExcelReader(ConfigReader.configs.getMainXLSXFilePath());
		XSSFWorkbook workbook = excel.workbook;		
		sheet1 = workbook.getSheet(ConfigReader.configs.getSheet1Name());
		sheet2 = workbook.getSheet(ConfigReader.configs.getSheet2Name());


		/**
		 * Make list of primary key under sheet 1
		 */
		List<String> primaryKeys = makePrimaryKeyList(sheet1);
		//primaryKeys.forEach(primaryKey -> System.out.println("PrimaryKey:: " + primaryKey));

		/**
		 * Check duplicates in primary key
		 * if there, program will not continue
		 */
		if(hasDuplicates(primaryKeys)){
			System.out.println("Error: " + ConfigReader.configs.getPrimaryKeyColHeader() + " in sheet 1 container duplicates.");
			System.exit(-1);
		}

		/**
		 * Make map of search key and value from sheet 2
		 */
		HashMap<String, String> searchValuesMap = makeSearchValuesMap(sheet2);

		/**
		 * Validate (like vlookup function)
		 */
		HashMap<String, String> result = doVLookUp(primaryKeys, searchValuesMap);

		/**
		 * Write result to excel
		 * Primary keys also provided for maintaining the order
		 */
		doWriteToExcel(workbook, sheet1, primaryKeys, result);

		return null;

	}

	private void doWriteToExcel(XSSFWorkbook workbook, XSSFSheet sheet, List<String> primaryKeys, HashMap<String, String> result) throws IOException {
		/**
		 * check if already data exists in given cell before write into.
		 */
		//todo

		List<String> orderedData = new ArrayList<>();
		for(String primaryKey : primaryKeys){

			for(Map.Entry<String, String> entry : result.entrySet()){
				if(primaryKey.equals(entry.getKey())){
					orderedData.add(entry.getValue());
				}
			}
		}
//-----------------------------


		int rowNum = 0;
		for(Row row : sheet) {
			XSSFCellStyle style = workbook.createCellStyle();
			style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (rowNum == 0){
				Cell cell = row.createCell(ConfigReader.configs.getSheet1ResultColNum());
				if(style != null) {
					XSSFFont headerFont = workbook.createFont();
					headerFont.setBold(true);
					style.setFont(headerFont);
					cell.setCellStyle(style);
				}

				cell.setCellValue(ConfigReader.configs.getSheet1ResultColHeader());

				rowNum++;
				continue;
			}

			//for(String value : orderedData){
				Cell cell = row.createCell(ConfigReader.configs.getSheet1ResultColNum());
				cell.setCellStyle(style);
				cell.setCellValue(orderedData.get(rowNum - 1));

			//}
			rowNum++;
		}

		FileOutputStream outputStream = new FileOutputStream(ConfigReader.configs.getMainXLSXFilePath());
		try(outputStream){
			workbook.write(outputStream);
			outputStream.close();
		}

	}

	private boolean hasDuplicates(List<String> primaryKeys) {
		return primaryKeys.size() != primaryKeys.stream().distinct().count();
	}


	private List<String> makePrimaryKeyList(XSSFSheet sheet1) {
		List<String> primaryKeys = new ArrayList<String>();
		int primaryKeyColNum = findPrimaryKeyColNumber(sheet1);
		int rowNum = 0;
		for(Row r : sheet1) {
			++rowNum;
			if (rowNum == 1){
				continue;
			}

			Cell c = r.getCell(primaryKeyColNum);
			if(c != null) {
				if(c.getCellType() == CellType.STRING) {
					primaryKeys.add(c.getStringCellValue());
				} else {
					ResultWriter.errors.add("sheet:" + sheet1.getSheetName() + ", row:" + rowNum+ " -> Primary key cell is not a string format");
				}
			}else {
				ResultWriter.errors.add("sheet:" + sheet1.getSheetName() + ", row:" + rowNum+ " -> A primary key cell is null/empty");
			}
		}
		return primaryKeys;
	}




	private HashMap<String, String> makeSearchValuesMap(XSSFSheet sheet2) {
		HashMap<String, String> searchValuesMap = new HashMap<>();
		List<String> orderedPrimaryKeys = new ArrayList<>();
		int primaryKeyColNum = findPrimaryKeyColNumber(sheet2);
		int sheet2ValueColNum = findSheet2ValueColNumber(sheet2);

		int rowNum = 0;
		for(Row r : sheet2) {
			++rowNum;
			if (rowNum == 1){
				continue;
			}

			Cell keyColCell = r.getCell(primaryKeyColNum);
			Cell valColCell = r.getCell(sheet2ValueColNum);

			if(keyColCell != null && valColCell != null) {
				if(keyColCell.getCellType() == CellType.STRING && valColCell.getCellType() == CellType.STRING) {
					searchValuesMap.put(keyColCell.getStringCellValue(), valColCell.getStringCellValue());
					orderedPrimaryKeys.add(keyColCell.getStringCellValue());
				} else {
					ResultWriter.errors.add("sheet:" + sheet2.getSheetName() + ", row:" + rowNum+ " -> key/value cell is not a string format");
				}
			}else {
				ResultWriter.errors.add("sheet:" + sheet2.getSheetName() + ", row:" + rowNum+ " -> key/value cell is null/empty");
			}
		}

		if(hasDuplicates(orderedPrimaryKeys)){
			System.out.println("Error: " + ConfigReader.configs.getPrimaryKeyColHeader() + " in sheet 2 container duplicates.");
			System.exit(-1);
		}

		return searchValuesMap;
	}

	private int findPrimaryKeyColNumber(XSSFSheet sheet) {
		int rowNum = 0;
		for(Cell cell : sheet.getRow(0)) {
			++rowNum;
			if(cell.getStringCellValue().equals(ConfigReader.configs.getPrimaryKeyColHeader())){
				return rowNum - 1;
			}
		}
		return -1;
	}
	private int findSheet2ValueColNumber(XSSFSheet sheet) {
		int rowNum = 0;
		for(Cell cell : sheet.getRow(0)) {
			++rowNum;
			if(cell.getStringCellValue().equals(ConfigReader.configs.getSheet2ValueColHeader())){
				return rowNum - 1;
			}
		}
		return -1;
	}


	private HashMap<String, String> doVLookUp(List<String> primaryKeys, HashMap<String, String> searchValuesMap) {
		HashMap<String, String> resultsAndComments = new HashMap<>();
		if(ConfigReader.configs.isPrintOutput()){
			System.out.println("Note: Print result in console would be as in order of sheet 2");
			System.out.println("\n" + ConfigReader.configs.getPrimaryKeyColHeader() + "\t\t\t|\t\t\t" + ConfigReader.configs.getSheet2ValueColHeader());
			System.out.println("=============================================================");
		}
		for(String primaryKey : primaryKeys){

				for(Map.Entry<String, String> entry : searchValuesMap.entrySet()){
					if(primaryKey.equals(entry.getKey())){
						resultsAndComments.put(entry.getKey(), entry.getValue());
						if(ConfigReader.configs.isPrintOutput()){
							System.out.println(entry.getKey() + "\t\t\t|\t\t\t" + entry.getValue());
						}
					}
				}

		}
		return resultsAndComments;
	}
}
