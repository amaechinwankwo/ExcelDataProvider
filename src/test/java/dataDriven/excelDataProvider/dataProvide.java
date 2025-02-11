package dataDriven.excelDataProvider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;

import org.testng.annotations.Test;

public class dataProvide {

	@Test(dataProvider = "driveTest")
	public void testCaseData(String greeting, String communication, String id) {
		System.out.println(greeting + communication + id);
	}

	
	DataFormatter formatter = new DataFormatter();
	@DataProvider(name = "driveTest")
	public Object[][] getData() throws IOException {
		
		// Object[][] data = { {"hello", "text", 1}, {"bye", "message", 143}, {"solo",
		// "call", 453} };
		// return data;

		FileInputStream fls = new FileInputStream("/Users/user/Desktop/Selenium/demoData.xlsx");

		XSSFWorkbook wb = new XSSFWorkbook(fls);

		XSSFSheet sheet = wb.getSheetAt(1); // to locate sheet 2

		int rowCount = sheet.getPhysicalNumberOfRows();

		XSSFRow row = sheet.getRow(0);

		int colCount = row.getLastCellNum();

		Object data[][] = new Object[rowCount - 1][colCount];

		for (int i = 0; i < rowCount - 1; i++) 
		{
			row = sheet.getRow(i + 1); // +1 To avoid getting header row

			for (int j = 0; j < colCount; j++) 
			{
				XSSFCell cell = row.getCell(j);
				
				data[i][j] = formatter.formatCellValue(cell);
			}
			
			
		}
		
		return data;

	}

}