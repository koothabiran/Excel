package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.*;


public class Sample {
	
	
	WebDriver driver;	
	XSSFWorkbook workbook;
    XSSFSheet sheet;
	XSSFCell cell;
	String URL =  "https://www.linkedin.com/";
	
	
	
	@BeforeMethod
	public void setup() throws InterruptedException
	{
		System.setProperty("webdriver.gecko.driver", "D:\\setup\\drivers\\geckodriver.exe");
	    driver = new FirefoxDriver(); 
		Thread.sleep(3000);
				
		driver.get(URL);
		
		
		String title = driver.getTitle();
		System.out.println(title);
		Thread.sleep(3000);
		}
	
	@SuppressWarnings("deprecation")
	@Test
	public void readdata() throws IOException, InterruptedException 
	{ 
		// Import excel sheet.
		File src=new File("C:\\Users\\Varun\\Desktop\\testdate.xlsx");
				
		// Load the file.
		FileInputStream finput = new FileInputStream(src);
		
		// Load the workbook.
		workbook = new XSSFWorkbook(finput);
		
		// Load the sheet in which data is stored.
		sheet= workbook.getSheetAt(0);
		
		int rc=sheet.getLastRowNum();
		
		for(int i=1; i<=rc; i++)
			
						
			     {
				
			         // Import data for Email.		
			         cell = sheet.getRow(i).getCell(1);
			     
			         cell.setCellType(Cell.CELL_TYPE_STRING);			
			         driver.findElement(By.id("login-email")).sendKeys(cell.getStringCellValue());
						          			
			         // Import data for password.		
			         cell = sheet.getRow(i).getCell(2);		
			         cell.setCellType(Cell.CELL_TYPE_STRING);			
			         driver.findElement(By.id("login-password")).sendKeys(cell.getStringCellValue());
			         
			         
			
			         
			         Thread.sleep(1000);
			         driver.findElement(By.id("login-submit")).click();
			         String title = driver.getTitle();
			         System.out.println(title);
			         Thread.sleep(6000);
			         
			         
			         // Write data in the excel.		         
			            @SuppressWarnings({ "unused", "resource" })
						FileOutputStream foutput=new FileOutputStream(src);
			                
			         // Specify the message needs to be written.		         
			             String message = "Pass";
			             String message1 = "Fail";
			                 
			             //Boolean aa = driver.findElements(By.xpath("html/body/nav/div/ul[1]/li[2]/a/span[1]")).size()!=0;
			             
			             
			                 
			             int aaa = driver.findElements(By.xpath("html/body/nav/div/ul[1]/li[2]/a/span[1]")).size();
			                       			                 
			              
			                 if(aaa==1)
			                 {
			                 // Create cell where data needs to be written.			         
			                 sheet.getRow(i).createCell(3).setCellValue(message);
			                 }
			                 else
			                 {
			                	 Thread.sleep(3000);
			                	 sheet.getRow(i).createCell(3).setCellValue(message1);			                	 
			                	 
			                 }
			                 
			          
			                 // Specify the file in which data needs to be written.		         
			                 FileOutputStream fileOutput = new FileOutputStream(src);
			         		         
			                 // finally write content
			                 workbook.write(fileOutput);		         
			                
			                 // close the file
			                 fileOutput.close();
			                
			              //CLicking on logout button  
			               if(aaa==1) 
			               {
			            	   
			            	   	Thread.sleep(6000);
			       	        	driver.findElement(By.xpath("//*[@id=\'nav-settings__dropdown-trigger\']/img")).click();
			             	   	Thread.sleep(3000);
			             	   		
			             	   	WebElement element = driver.findElement(By.xpath("html/body/nav/div/ul[1]/li[6]/div/ul/li[4]/ul/li/a"));	      	   		    	   		
			             	   	Actions action = new Actions(driver);      	  
			             	   	action.moveToElement(element).build().perform();    
			             	   	driver.findElement(By.xpath("html/body/nav/div/ul[1]/li[6]/div/ul/li[4]/ul/li/a")).click();
			             	   	Thread.sleep(4000);
			            	  
			               }
			                      
			              driver.get(URL);
			              Thread.sleep(3000);
			                 
			        }
		
	}
			
		@AfterMethod
		 public void closebrowser() {
		 driver.close();
		 
		  }
			
		
}


		

	
	



