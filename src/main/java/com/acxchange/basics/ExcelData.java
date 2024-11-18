package com.acxchange.basics;

import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelData {

	public static void main(String[] args) throws Exception {
		//Asssignig the rertuned data from excel to 2d String Array
		String[][] exceldata = readExcelData("./data/testdataexcel.xlsx");
		//Printing the 1st cell value(1st row is skipped and Value Reading will start from 1)
		 System.out.println(exceldata[0][0]);
	}
	
	public static String[][] readExcelData(String filePath) throws IOException {
		//Initializing workbook object to null
		XSSFWorkbook wbk = null;
		//To get the workbook in the specified filepath
		wbk = new XSSFWorkbook(filePath);
		//To get the 1st sheet
		XSSFSheet shtat = wbk.getSheetAt(0);
		//To get the total row in 1st sheet
		int rowCount = shtat.getLastRowNum();
		//To get the column count from the 1st Row(Index Row) with help of Last Used Cell Number
		short colCount=shtat.getRow(0).getLastCellNum();
		//Two Dimensional String array to store the excel values from the sheet with row and columnn
		String[][] data = new String[rowCount][colCount];
		//Loop to iterate row
		for (int i = 1; i <=rowCount; i++) {
			//Loop to iterate Column
			for(int j=0;j<colCount;j++) {
				//Retreiving the excel content and Storing in 2d array
				String stringCellValue = shtat.getRow(i).getCell(j).getStringCellValue();
				data[i-1][j]=stringCellValue;
			}
		}
		//Closing the opened workbook
		wbk.close();
		//Returning the 2d array upon calling the function
		return data;
	}
}
