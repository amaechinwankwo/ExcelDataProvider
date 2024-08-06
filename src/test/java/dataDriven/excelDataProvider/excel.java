package dataDriven.excelDataProvider;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class excel {

	@Test
	public void getExcel() throws IOException {

		FileInputStream fls = new FileInputStream("/Users/user/Desktop/Selenium/demoData.xlsx");

		XSSFWorkbook wb = new XSSFWorkbook(fls);

		XSSFSheet sheet = wb.getSheetAt(1); // to locate sheet 2

		int rowCount = sheet.getPhysicalNumberOfRows();

		XSSFRow row = sheet.getRow(0);

		int colCount = row.getLastCellNum();

		Object data[][] = new Object[rowCount - 1][colCount];

		for (int i = 0; i < rowCount - 1; i++) 
		{
			
			System.out.println("Outer loop Started");
			
			row = sheet.getRow(i + 1); // +1 To avoid getting header row

			for (int j = 0; j < colCount; j++) 
			{
				System.out.println(row.getCell(j));
			}
			
			System.out.println("Outer loop ended");

		}

	}

}
