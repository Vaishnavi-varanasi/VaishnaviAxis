package Axis.DataDrivenTesting;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class WriteIntoExcel {
	WebDriver driver;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	XSSFCell cell;

	@SuppressWarnings("deprecation")
	@Test
	public void FBLogin() throws IOException {
		System.setProperty("webdriver.chrome.driver", "C:\\chromedriver-win64\\chromedriver-win64/chromedriver.exe");
		driver = new ChromeDriver();
		driver.get("https://www.facebook.com/");

		driver.manage().window().maximize();

		driver.manage().timeouts().implicitlyWait(20, TimeUnit.MILLISECONDS);

		// Import Excel Sheet File Source

		File src = new File("C:\\Users\\hp\\eclipse-workspace\\DataDrivenTesting/TestData.xlsx");

		// Loading the file

		FileInputStream FIS = new FileInputStream(src);

		// Loading the Workbook

		workbook = new XSSFWorkbook(FIS);

		// Accessing the Sheet with input data
		sheet = workbook.getSheetAt(0);
		
		sheet.getRow(0).getCell(2).setCellValue("Title");

		for (int i = 1; i <= sheet.getLastRowNum(); i++) {

			/*
			 * // Import data from Username Column cell = sheet.getRow(i).getCell(0);
			 * 
			 * driver.findElement(By.xpath("//input[@name = 'email']")).clear();
			 * driver.findElement(By.xpath("//input[@name = 'email']")).sendKeys(cell.
			 * getStringCellValue());
			 * 
			 * // Import data from Password Column cell = sheet.getRow(i).getCell(1);
			 * 
			 * driver.findElement(By.xpath("//input[@name = 'pass']")).clear();
			 * driver.findElement(By.xpath("//input[@name = 'pass']")).sendKeys(cell.
			 * getStringCellValue());
			 */

			String title = driver.getTitle();
			System.out.println(title);
			// To write the data into the Excel Sheet
			FileOutputStream FOS = new FileOutputStream(src);

			// Create the cell where data needs to be entered
			sheet.getRow(i).getCell(2).setCellValue(title);
			
			workbook.write(FOS);
			FOS.close();

		}

	}
}
