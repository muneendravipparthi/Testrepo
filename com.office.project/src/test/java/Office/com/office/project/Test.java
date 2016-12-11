package Office.com.office.project;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.Select;
public class Test {

	public static void main(String[] args) throws InterruptedException, IOException {

		File file = new File("H:\\Abc.xlsx");
		FileInputStream inputfile = new FileInputStream(file);
		Workbook Project = new XSSFWorkbook(inputfile);
		Sheet sheet = Project.getSheet("Sheet1");

		int k = sheet.getLastRowNum() - sheet.getFirstRowNum();
		System.out.println(k);
		for (int i = 28; i <= k; i++) {
			Row row = sheet.getRow(i);

			String FullName = row.getCell(0).getStringCellValue();
			String Email = row.getCell(1).getStringCellValue();
			String Password = row.getCell(2).getStringCellValue();
			String MotherMaiden = row.getCell(3).getStringCellValue();
			
			long Mobile1 = (long) row.getCell(4).getNumericCellValue();
			String Mobile = Long.toString(Mobile1);
			 
			/*int Mobile1 = (int) row.getCell(4).getNumericCellValue();
			String Mobile = Integer.toString(Mobile1);*/
			int Day1 = (int) row.getCell(5).getNumericCellValue();
			String Day = Integer.toString(Day1);
			String Month = row.getCell(6).getStringCellValue();
			int Year1 = (int) row.getCell(7).getNumericCellValue();
			String Year = Integer.toString(Year1);
			String City = row.getCell(9).getStringCellValue();
			String Answer = row.getCell(10).getStringCellValue();
			System.out.println(FullName);

			// chromedriver
			System.setProperty("webdriver.chrome.driver",
					"H:\\Selenium\\chromedriver.exe");
			WebDriver driver = new ChromeDriver();

			String baseUrl = "https://register.rediff.com/";
			driver.get(baseUrl + "/register/register.php?FormName=user_details");
			driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
			driver.findElement(By.xpath("//input")).clear();
			driver.findElement(By.xpath("//input")).sendKeys(FullName);
			driver.findElement(By.xpath("//tr[7]/td[3]/input")).clear();
			driver.findElement(By.xpath("//tr[7]/td[3]/input")).sendKeys(Email);
			driver.findElement(By.xpath("//tr[9]/td[3]/input")).clear();
			driver.findElement(By.xpath("//tr[9]/td[3]/input")).sendKeys(Password);
			driver.findElement(By.xpath("//tr[11]/td[3]/input")).clear();
			driver.findElement(By.xpath("//tr[11]/td[3]/input")).sendKeys(Password);
			driver.findElement(By.xpath("//td[2]/table/tbody/tr/td/input")).click();
			new Select(driver.findElement(By.xpath("//select")))
					.selectByVisibleText("What is the name of your first school?");
			driver.findElement(By.xpath("//tr[4]/td[3]/input")).clear();
			driver.findElement(By.xpath("//tr[4]/td[3]/input")).sendKeys(Answer);
			driver.findElement(By.xpath("//tr[6]/td[3]/input")).clear();
			driver.findElement(By.xpath("//tr[6]/td[3]/input")).sendKeys(MotherMaiden);
			driver.findElement(By.xpath("//div[3]/input")).clear();
			driver.findElement(By.xpath("//div[3]/input")).sendKeys(Mobile);
			new Select(driver.findElement(By.xpath("//tr[24]/td[3]/select"))).selectByVisibleText(Day);
			new Select(driver.findElement(By.xpath("//select[2]"))).selectByVisibleText(Month);
			new Select(driver.findElement(By.xpath("//select[3]"))).selectByVisibleText(Year);
			new Select(driver.findElement(By.xpath("//tr[30]/td/div/table/tbody/tr/td[3]/select")))
					.selectByVisibleText(City);
			Thread.sleep(60000);
			driver.quit();
	}

}
}

