package config;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import ExecutionEngine.DriverScript;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class ActionKeyword {

	public static WebDriver driver;
	public String Resultp = "Pass";
	public String Resultf = "fail";

	public void openBrowser() {
		try {
			System.setProperty("webdriver.chrome.driver",
					"D:\\Jignesh QA Files\\QA Automation Setup Files\\chromedriver_win32\\chromedriver.exe");
			driver = new ChromeDriver();
		} catch (Exception e) {
			DriverScript.flag = false;
		}

	}

	public void navigate() {
		try {
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			driver.get("https://www.toolsqa.com");
		} catch (Exception e) {
			// TODO: handle exception
			DriverScript.flag = false;
		}
	}

	public void verifytitle() {
		try {
			String title = driver.getTitle();
			String expectedresult = "Free Q Automation Tools Tutorial for Beginners with Examples";
			if (title.equals(expectedresult)) {
				System.out.println("Passed");
			}

			else {
				// TODO: handle exception
				DriverScript.flag = false;
			}
		} catch (Exception e) {
			// TODO: handle exception
		}
	}

	public static void Close() {
		try {
			driver.close();
		} catch (Exception e) {
			DriverScript.flag = false;
			// TODO: handle exception
		}

	}

}
