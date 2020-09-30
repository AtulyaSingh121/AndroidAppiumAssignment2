package Mobile_Selenium_TestNG_Maven_Assignment2.AndroidAppiumAssignment;

import org.testng.annotations.Test;

import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.remote.MobileCapabilityType;

import org.testng.annotations.BeforeTest;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.Array;
import java.net.MalformedURLException;
import java.net.URL;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;
import org.testng.annotations.AfterTest;

public class DW007TC_usingXLsheel {
	AndroidDriver driver;
	  @Test
	  public void Login() throws InterruptedException, IOException {
		  int i ;
			FileInputStream fileinput=new FileInputStream("C:\\Users\\AtulyaSingh\\eclipse-workspace\\AndroidAppiumAssignment2\\Datasheet.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(fileinput);
			XSSFSheet sheet = workbook.getSheetAt(0);
			int lastRow = sheet.getLastRowNum();
			driver.get("http://demowebshop.tricentis.com/");
			  driver.findElement(By.xpath("//span[@class='icon']")).click();
		      driver.findElement(By.xpath("//a[contains(text(),'Electronics')]//following-sibling::span")).click();
		      driver.findElement(By.xpath("//ul[@class='mob-top-menu show']//li//a[contains(text(),'Camera, photo')]")).click();
		      Thread.sleep(5000);
			  for(i=1; i<=lastRow; i++){
				int display =(int)sheet.getRow(i).getCell(0).getNumericCellValue();
				System.out.println("value chosen from xl file " +display);	
				int expectedvalue = display ;
			      Select dropdown = new Select( driver.findElement(By.xpath("//select[@id=\"products-pagesize\"]")));
			      dropdown.selectByVisibleText(String.valueOf(display));
			      Thread.sleep(5000);		     
			     String actualvalue = new Select(driver.findElement(By.xpath("//select[@id=\"products-pagesize\"]"))).getFirstSelectedOption().getText().trim();
                Assert.assertEquals(Integer.parseInt(actualvalue),expectedvalue);
			    System.out.println("Value chosen is" +actualvalue);
			  } 
			    		  
			}
			
	  @BeforeTest
	  public void URLLaunch() throws MalformedURLException {
		  DesiredCapabilities capability= new DesiredCapabilities();
	      capability.setCapability("deviceName", "Mobile");
	      capability.setCapability(MobileCapabilityType.PLATFORM_NAME, "Android");
	      capability.setCapability(MobileCapabilityType.BROWSER_NAME, "Chrome");
	      driver = new AndroidDriver(new URL("http://0.0.0.0:4723/wd/hub"),capability);
	        
	  }

	  @AfterTest
	  public void LogOut() throws InterruptedException {
		  Thread.sleep(5000);
	      driver.close();

}
	  }
