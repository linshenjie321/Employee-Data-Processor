package com.sjl.employeedataprocessor;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

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
				jobPreferenceProcessor(row);
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
	
	private static void jobPreferenceProcessor(Row row) {
		Cell cell = row.getCell(15);
		String jobPreference = readCellValueAsString(cell);
		List<String> jobPreferenceList = Arrays.asList(jobPreference.split(","));
		System.out.println(jobPreference + " || can process = " + shouldBeProcessed(jobPreferenceList.get(0)) + " || " + identifyDuplicateNumbers(jobPreferenceList));
		if(shouldBeProcessed(jobPreferenceList.get(0))) {
			int i = 1;
			for(String jobPreferenceOption : jobPreferenceList) {
				Cell optionCell = row.getCell(15 + i);
				if(optionCell == null) {
					optionCell = row.createCell(15 + i);
				}
				optionCell.setCellValue(jobPreferenceOption.trim());
				i++;
			}
		}
	}
	
	private static boolean shouldBeProcessed(String input) {
		try {
			Integer.parseInt(input);
			return true;
		} catch (NumberFormatException e) {
			return false;
		}
		
	}
	
	private static String identifyDuplicateNumbers(List<String> jobPreferenceList) {
		if(!shouldBeProcessed(jobPreferenceList.get(0))) {
			return "Not Applicable";
		}
		String result = "";
		List<String> jobPreferenceOptionChecker = new ArrayList<>();
		for(String jobPreferenceOption : jobPreferenceList) {
			if(jobPreferenceOptionChecker.contains(jobPreferenceOption.trim())) {
				result = result + jobPreferenceOption.trim() + " ; ";
			}else {
				jobPreferenceOptionChecker.add(jobPreferenceOption);
			}
		}
		
		if (result.isBlank()) {
			return "No duplicate found";
		}else {
			return "DUP FOUND - " + result;
		}
	}
	
	private static String readCellValueAsString(Cell cell) {
		if(cell == null) {
			return "";
		}
		try {
			return cell.getStringCellValue();
		} catch (IllegalStateException ex) {
			double cellValue = cell.getNumericCellValue();
			return String.valueOf(Double.valueOf(cellValue).intValue());
		}
	}

}
