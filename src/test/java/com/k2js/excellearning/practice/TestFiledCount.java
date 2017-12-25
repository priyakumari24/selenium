package com.k2js.excellearning.practice;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class TestFiledCount {
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
	
	public static void main(String[] args) {
		//System.out.println(TestFiledCount.getDataFiledCount("verifyHomePage"));;

		
	}
	
	/*public static int getDataFiledCount(String tcname)
	{
		
			
	}*/
		
		
		

}
