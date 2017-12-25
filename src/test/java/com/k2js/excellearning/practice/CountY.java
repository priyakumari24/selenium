package com.k2js.excellearning.practice;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class CountY {
	private static FileInputStream fis = null;
	private static Workbook wb = null;
	private static Sheet s = null;
	static {
		try {
			fis = new FileInputStream(".\\TestData\\NTData.xlsx");
			wb = WorkbookFactory.create(fis);// complete excel sheet
			s = wb.getSheet("sheet1");// sheet in a excelsheet
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		int rc = CountY.getcountY();
		

	}


	public static int getcountY()  {
		int rc=s.getPhysicalNumberOfRows();
		int counter=0;
		for (int i = 0; i < rc; i++) {
			Row r = s.getRow(i);
			Cell c = r.getCell(2);
			String celldata = c.getStringCellValue();
			System.out.println(celldata);
			if(celldata.equalsIgnoreCase("Y"))
			{
				counter++;
			}
			

		}
		System.out.println(counter);
		return counter;
		
	}

}
