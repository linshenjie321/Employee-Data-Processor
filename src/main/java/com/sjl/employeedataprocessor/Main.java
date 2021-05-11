package com.sjl.employeedataprocessor;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
	
	public static void main(String[] args) {
		String folderLocation = "C:\\Users\\Cometstrike\\jee-2021-03-workspace\\Employee-Data-Processor\\test-data";
		String fileName = "(MODIFIED) GOOGLE RESPONSES - ESL_LINC Summer Bidding Application 2021 - Copy.xlsx";
		String fileLocation = folderLocation + "\\" + fileName;
		
		readExcelSheet(fileLocation, 0);
		
	}
	
	private static void readExcelSheet(String fileLocation, int sheetIndex) {
		try {
			FileInputStream file = new FileInputStream(new File(fileLocation));
			Workbook workbook = new XSSFWorkbook(file);
			
			Sheet sheet = workbook.getSheetAt(sheetIndex);
			
			SummerBiddingApplication summerBiddingApplication = new SummerBiddingApplication();
			summerBiddingApplication.initiateColumnHeader(sheet.getRow(0));
			
			int i = 0;
			for(Row row : sheet) {
				if(i == 0) {
					// skipping the header column
					i++;
					continue;
				}
				Map<Integer, String> columnHeader = summerBiddingApplication.getColumnHeader();
				for(int j = 0; j < columnHeader.size(); j++) {
					Cell cell = row.getCell(j);
					//we assume the cell column header is string type
					readCellValue(cell, columnHeader.get(j).toString());
				}
				i++;
			}
			
			workbook.close();
			file.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	private static void readCellValue(Cell cell, String columnHeader) {
		if(cell == null) {
			System.out.println(columnHeader +  " = ");
			return;
		}
		try {
			System.out.println(columnHeader + " = " + cell.getRichStringCellValue());
		} catch (IllegalStateException ex) {
			System.out.println(columnHeader + " = " + cell.getNumericCellValue());
		}
	}

}
