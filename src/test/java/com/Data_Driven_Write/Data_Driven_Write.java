package com.Data_Driven_Write;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Data_Driven_Write {

	public static void write_Data() throws Throwable {

		File file = new File("C:\\Users\\user\\Desktop\\Data_Write.xlsx");

		FileInputStream fis = new FileInputStream(file);

		Workbook w = new XSSFWorkbook(fis);

		Sheet createSheet = w.createSheet("UserDetails");

		Row createRow = createSheet.createRow(0);

		Cell createCell = createRow.createCell(0);

		createCell.setCellValue("User Name");

		w.getSheet("UserDetails").getRow(0).createCell(1).setCellValue("Password");

		w.getSheet("UserDetails").createRow(1).createCell(0).setCellValue("Avinash");
		
		w.getSheet("UserDetails").getRow(1).createCell(1).setCellValue("12345");

		FileOutputStream fos = new FileOutputStream(file);

		w.write(fos);

		w.close();

		System.out.println("Successfully");

	}

	public static void main(String[] args) throws Throwable {
		write_Data();
	}

}
