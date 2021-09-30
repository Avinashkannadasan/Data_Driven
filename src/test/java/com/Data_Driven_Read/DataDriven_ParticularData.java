package com.Data_Driven_Read;

import java.io.File;
import java.io.FileInputStream;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven_ParticularData {

	
	
	public static void particular_Data() throws IOException {

		File file = new File("C:\\Users\\user\\eclipse-workspace\\Data_Driven\\Excel_Data\\Data_Read.xlsx");

		FileInputStream fis = new FileInputStream(file);

		Workbook w = new XSSFWorkbook(fis); // ------------------------------------> Up Casting

		Sheet sheetAt = w.getSheetAt(0);

		Row row = sheetAt.getRow(2);

		Cell cell = row.getCell(0);

		CellType cellType = cell.getCellType();

		
		
		
		if (cellType.equals(cellType.STRING)) {

			String stringCellValue = cell.getStringCellValue();
			System.out.println(stringCellValue);
		}

		
		
		
		
		else if (cellType.equals(cellType.NUMERIC)) {

			double numericCellValue = cell.getNumericCellValue();

			int value = (int) numericCellValue; // ---------------------------------> Narrowing type Casting

			System.out.println(value);

		}

	}

	
	
	
	public static void main(String[] args) throws Throwable {
		particular_Data();

	}
}
