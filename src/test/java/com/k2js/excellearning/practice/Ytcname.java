//for each testname, count of Y
package com.k2js.excellearning.practice;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Ytcname {
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
		Ytcname.tcY("verifyRegistrationProcess");
		
	}
	public static int tcY(String tcname)
	{
		int rc=s.getPhysicalNumberOfRows();
		int counter=0;
		for (int i = 0; i < rc; i++) {
			Row r = s.getRow(i);
			Cell c = r.getCell(2);
			Cell c1= r.getCell(1);
			String celldata = c.getStringCellValue();
			String celldata1=c1.getStringCellValue();
			System.out.println(celldata);
			if(celldata1.equals(tcname) && celldata.equalsIgnoreCase("Y"))
			{
				counter++;
			}
			

		}
		System.out.println(counter);
		return counter;
	}

}
