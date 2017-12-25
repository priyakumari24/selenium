package com.k2js.excellearning.practice;

import java.io.FileInputStream;
import java.lang.reflect.Method;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ExcelReadWrite {
	static FileInputStream fileInputStream = null;
	static Workbook workbook = null;
	static Sheet sheet = null;
	
	static {
		try {
			fileInputStream = new FileInputStream(".\\TestData\\NTData.xlsx");
			workbook = WorkbookFactory.create(fileInputStream);
			sheet = workbook.getSheet("Sheet1");
			System.out.println(sheet);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	static final int TOTAL_XL_ROWS = ExcelReadWrite.getTotalRows();

	private static int getTotalRows() {
		return sheet.getPhysicalNumberOfRows();
	}

	public static int getTestCaseRows(String tc) {
		int rowCount = 0;
		for (int i = 0; i < TOTAL_XL_ROWS; i++) {
			Row row = sheet.getRow(i);
			String tcn = row.getCell(1).getStringCellValue();
			String tcs = row.getCell(2).getStringCellValue();
			if (tcn.equalsIgnoreCase(tc) && tcs.equalsIgnoreCase("Y")) {
				rowCount++;
			}
		}
		return rowCount;
	}

	public static int getDataFieldsCount(String tc) {
		for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			String testCase = row.getCell(1).getStringCellValue();
			if (testCase.equalsIgnoreCase(tc)) {
				return row.getPhysicalNumberOfCells() - 3;
			}
		}
		return 0;
	}

	
	@DataProvider(name="abcd")
	public static String[][] storeTestData(Method m) { //Method m will take the method 'verifyRegistration'
		String tcn = m.getName();
		int tcrc = ExcelReadWrite.getTestCaseRows(tcn);
		int tccc = ExcelReadWrite.getDataFieldsCount(tcn);
		String[][] tcd = new String[tcrc][tccc + 1];
		int nri = 0;
		for (int i = 1; i < TOTAL_XL_ROWS; i++) {
			Row row = sheet.getRow(i);
			String testCase = row.getCell(1).getStringCellValue();
			String testCaseStatus = row.getCell(2).getStringCellValue();
			if (testCase.equalsIgnoreCase(tcn) && testCaseStatus.equalsIgnoreCase("Y")) {
				int nci = 0;
				for (int j = 3; j < row.getPhysicalNumberOfCells(); j++) {
					tcd[nri][nci++] = row.getCell(j).getStringCellValue();
				}
				tcd[nri][nci] = i + "";
				nri++;
			}
		}
		return tcd;
	}
	
	/*
	@Test(dataProvider="abcd",dataProviderClass=ExcelReadWrite.class)
	public void verifyRegistrationProcess(String... abc) {
		//String[][] testData = ExcelReadWrite.storeTestData("verifyRegistrationProcess");
		for(String t:abc)
		{
			System.out.println(t);
		}
	}
	*/
}