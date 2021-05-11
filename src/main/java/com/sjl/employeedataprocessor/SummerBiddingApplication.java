package com.sjl.employeedataprocessor;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class SummerBiddingApplication {
	
	private Map<Integer, String> columnHeader;
	
	public void initiateColumnHeader(Row row) {
		this.columnHeader = new HashMap<>();
		int i = 0;
		for(Cell cell : row) {
			columnHeader.put(i, cell.getStringCellValue());
			i++;
		}
	}

	public Map<Integer, String> getColumnHeader() {
		return columnHeader;
	}

	public void setColumnHeader(Map<Integer, String> columnHeader) {
		this.columnHeader = columnHeader;
	}

}
