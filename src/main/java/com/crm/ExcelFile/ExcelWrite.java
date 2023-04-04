package com.crm.ExcelFile;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {
	public static void main(String[] args) {
		try {
			Workbook workbook = new XSSFWorkbook();

			Sheet sheet = workbook.createSheet("Sheet1");

			Row row = sheet.createRow(0);

			Cell cell = row.createCell(0);

			cell.setCellValue("Hello, World!");

			FileOutputStream fileOut = new FileOutputStream("..\\FileReadAndWrite\\ExcelFiles\\workbook.xlsx");
			workbook.write(fileOut);
			fileOut.close();
			workbook.close();

			System.out.println("Excel file created successfully!");

		} catch (Exception ex) {
			System.out.println(ex.getMessage());
		}
	}
}
