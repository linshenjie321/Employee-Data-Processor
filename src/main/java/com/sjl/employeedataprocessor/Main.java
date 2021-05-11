package com.sjl.employeedataprocessor;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
	
	static String folderLocation = "C:\\Users\\Cometstrike\\jee-2021-03-workspace\\Employee-Data-Processor\\test-data";
	static String outputFileName = "(MODIFIED) GOOGLE RESPONSES - ESL_LINC Summer Bidding Application 2021 - Processed.xlsx";
	static String outputFileLocation = folderLocation + "\\" + outputFileName;
	
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
				supplyOptionProcessor(row);
				i++;
			}

			FileOutputStream outputStream = new FileOutputStream(outputFileLocation);
			workbook.write(outputStream);
			workbook.close();
			file.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	private static void supplyOptionProcessor(Row row) {
		Cell supplyOption = row.getCell(0);
		if("I am ONLY interested in supply work for the summer.".equalsIgnoreCase(supplyOption.getStringCellValue())) {
			row.getCell(1).setCellValue("X");
		} else if("I would like to supply for the summer if I do not get an assignment.".equalsIgnoreCase(supplyOption.getStringCellValue())) {
			row.getCell(2).setCellValue("X");
		} else if("If I obtain a summer assignment, I also wish to supply for the summer outside of my assignment hours.".equalsIgnoreCase(supplyOption.getStringCellValue())) {
			row.getCell(3).setCellValue("X");
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
