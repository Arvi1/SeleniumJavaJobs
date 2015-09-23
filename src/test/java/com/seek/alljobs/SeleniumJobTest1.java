package com.seek.alljobs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.AfterMethod;

public class SeleniumJobTest1 {

	WebDriver driver;
	DateFormat dateFormat;
	Date date;
	FileInputStream fis;
	HSSFWorkbook wb;
	HSSFSheet sheet;
	FileOutputStream fos;
  @BeforeMethod
  public void beforeMethod() throws IOException {
	  driver = new FirefoxDriver();
	  driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	  driver.manage().window().maximize();
	  
	  driver.get("http://www.seek.com.au/");
	  
	  // creating a file, workbook and sheet 
	  fis = new FileInputStream(new File("D:/EclipseProjects/SeekProject/SeekDemoProject/Res.xls"));
	  wb = new HSSFWorkbook();
	  sheet = wb.getSheet("Sheet1");
	  
	// System date 
	dateFormat = new SimpleDateFormat("yyyy/mm/dd");
	date = new Date();
	System.out.println(dateFormat.format(date));
  }
  
  @Test
  public void SeleniumJobs() throws Exception {
		// Enter Job Keyword
		driver.findElement(By.id("keywords")).sendKeys("Selenium");
		
		// Enter Job Classification
		Select classification = new Select(driver.findElement(By.id("classification")));
		classification.selectByVisibleText("Information & Communication Technology");

		// Enter Location
		driver.findElement(By.id("where")).sendKeys("All Australia");
		
		// Enter Job Sub Classification
		Select subclassification = new Select(driver.findElement(By.id("subclassification")));
		subclassification.selectByVisibleText("Testing & Quality Assurance");
		
		// Click Seek Button for search jobs
		driver.findElement(By.xpath("//*[@id='search-panel']/div/div[3]/div[2]/button[2]")).click();
		
		// Enter SortBy 
		Select sortBy = new Select(driver.findElement(By.id("sortmode")));
		sortBy.selectByVisibleText("date");

		// xpaths for date and jobs in the application
		String s1 = "//*[@id='jobsListing']/div[3]/article[";
		String s2 = "]/dl/dd[2]/span[1]";		
		String s3 = "]/dl/dd[1]/h2/a";
		
		try {
		
		for (int i=1; i<sheet.getLastRowNum()+1; i++){
			Cell resultDate = sheet.getRow(i).getCell(1);
			Cell jobPostedDate = sheet.getRow(i).getCell(2);
			Cell resultCell = sheet.getRow(i).getCell(3);	
			if (driver.findElement(By.xpath(s1 + i + s2)).getText().length()==8){
				System.out.println("Today's Job details");								
				String jobPostedOn = driver.findElement(By.xpath(s1 + i + s2)).getText().toString();
				System.out.println(i + " --  " + driver.findElement(By.xpath(s1 + i + s3)).getText());
				String sResult = driver.findElement(By.xpath(s1 + i + s3)).getText().toString();							
				
				System.out.println(jobPostedOn);
				resultDate.setCellValue(dateFormat.format(date));
				jobPostedDate.setCellValue(jobPostedOn);
				resultCell.setCellValue(sResult);
				
			}else {
				System.out.println("Out dated Job details");
				String jobPostedOn = driver.findElement(By.xpath(s1 + i + s2)).getText().toString();
				System.out.println(i + " --  "  + driver.findElement(By.xpath(s1 + i + s3)).getText());								
				String sResult = driver.findElement(By.xpath(s1 + i + s3)).getText().toString();
				
				resultDate.setCellValue(dateFormat.format(date));
				jobPostedDate.setCellValue(jobPostedOn);
				resultCell.setCellValue(sResult);
			}
		}
		wb.close();
		fis.close();
		fos = new FileOutputStream(new File("D:/EclipseProjects/SeekProject/SeekDemoProject/Res2.xls"));
		wb.write(fos);
		fos.close();
		} catch (Exception e) {
			System.out.println(e.getStackTrace());
		}
		
  }

  @AfterMethod
  public void afterMethod() throws Exception {
//	  wb.write(fos);
//	  fos.close();
	  driver.quit();
  }

}
