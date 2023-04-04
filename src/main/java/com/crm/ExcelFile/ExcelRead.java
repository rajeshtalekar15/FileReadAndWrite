package com.crm.ExcelFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
	public static void main(String[] args) throws IOException {
		FileInputStream inputStream = new FileInputStream(new File("..\\FileReadAndWrite\\ExcelFiles\\Book1.xlsx"));
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheetAt(0);

		for (Row row : sheet) {
			for (Cell cell : row) {
				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue() + "\t");
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue() + "\t");
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue() + "\t");
					break;
				case FORMULA:
					System.out.print(cell.getCellFormula() + "\t");
					break;
				default:
					System.out.print("\t");
				}
			}
			System.out.println();
		}

		workbook.close();
		inputStream.close();
	}
}
