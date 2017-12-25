package com.k2js.excellearning.practice;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelLearning {

	private static FileInputStream fis = null;
	private static Workbook wb = null;
	private static Sheet s = null;

	static // it will execute only once(needs to execute below code only once)
	{
		try {
			fis = new FileInputStream(".\\TestData\\NTData.xlsx");
			wb = WorkbookFactory.create(fis);// complete excel sheet
			s = wb.getSheet("sheet1");// sheet in a excelsheet
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		// int rc = ExcelLearning.getTotalRow();
		// System.out.println(rc);
		// int rcc=ExcelLearning.getRow1CellCount();
		// System.out.println(rcc);
		// String data = ExcelLearning.getRowCellData();
		// System.out.println(data);
			//int rc1=ExcelLearning.getAllCellCount();
			//System.out.println(rc1);
		//String rc2 = ExcelLearning.getAllCellData();
		//System.out.println(rc2);
		String rc3=ExcelLearning.getsfData();
		System.out.println(rc3);

	}

	public static int getTotalRow() throws EncryptedDocumentException, InvalidFormatException, IOException {

		int rowcount = s.getPhysicalNumberOfRows();// getting no of rows
		return rowcount;
	}

	private static int getRow1CellCount()// no of cells in Row0
	{
		Row r = s.getRow(0);
		int cellcount = r.getPhysicalNumberOfCells();
		return cellcount;
	}

	private static String getRowCellData() {
		Row r = s.getRow(0);
		Cell c = r.getCell(0);
		String celldata = c.getStringCellValue();
		return celldata;
	}

	public static int getAllCellCount() throws EncryptedDocumentException, InvalidFormatException, IOException {
		for (int i = 0; i < ExcelLearning.getTotalRow(); i++) {
			Row r1 = s.getRow(i);
			int cellcount1 = r1.getPhysicalNumberOfCells();
			System.out.println(cellcount1);

		}
		return 0;

	}
	public static String getAllCellData() throws EncryptedDocumentException, InvalidFormatException, IOException
	{
		for(int j=0;j<ExcelLearning.getTotalRow();j++)
		{
			Row r2 = s.getRow(j);
			Cell c1 = r2.getCell(j);
			String celldata1 = c1.getStringCellValue();
			System.out.println(celldata1);
		}
		return "";
	}
	
	public static String getsfData() throws EncryptedDocumentException, InvalidFormatException, IOException
	{
		for(int k=0;k<ExcelLearning.getTotalRow();k++)
		{
			Row r3=s.getRow(k);
			Cell c3=r3.getCell(1);
			System.out.println(c3);
			Cell c4 = r3.getCell(2);
			System.out.println(c4);
			
		}
		return null;
		
	}

}
