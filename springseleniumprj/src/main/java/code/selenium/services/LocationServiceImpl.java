package code.selenium.services;

import org.openqa.selenium.By;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.springframework.stereotype.Component;

import code.selenium.entities.Location;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.logging.log4j.core.tools.picocli.CommandLine.Help.TextTable.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

import org.openqa.selenium.support.ui.ExpectedConditions;

@Component
public class LocationServiceImpl {
	public static List<Location> LocationList = new ArrayList<>();

	public WebDriver initialization() {
		// Initializing Chrome Webdriver to open Chrome and perform Operation

		WebDriver driver;
		System.setProperty("webdriver.http.factory", "jdk-http-client");
		ChromeOptions option = new ChromeOptions();
		option.addArguments("--remote-allow-origins=*");

		// Initializing chrome driver object
		driver = new ChromeDriver(option);

		// maximize window
		driver.manage().window().maximize();

		// method to go to the rent a car website
		driver.get(" https://www.acerentacar.com/Locations");
		try {
			Thread.sleep(5000);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
		return driver;

	}

	public void getCity2Locations(WebDriver driver, WebDriverWait wait) {
		// wait is used to wait the screen until the given condition is true
		wait.until(ExpectedConditions.visibilityOfElementLocated(
				By.xpath("//*[@id=\"__next\"]/div/div[2]/main/div/div[1]/div/div[2]/ul/div[2]/li")));

		// element to click on another city
		WebElement cityName2 = driver
				.findElement(By.xpath("//*[@id=\"__next\"]/div/div[2]/main/div/div[1]/div/div[2]/ul/div[2]/li"));
		cityName2.click();

		String str = "//*[@id=\"__next\"]/div/div[2]/main/div/div/div[1]/div/div/div[1]/";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(str + "p")));

		for (int i = 1; i < 6; i++) {
			str = "//*[@id=\"__next\"]/div/div[2]/main/div/div/div[" + i + "]/div/div/div[1]/";
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(str + "p")));

			String address2 = "", phNo2 = "", locTyp2 = "", locHr2 = "";
			String city2 = driver.findElement(By.xpath(str + "p")).getText();
			for (int j = 1; j < 3; j++) {
				address2 += "\n" + driver.findElement(By.xpath(str + "span[" + j + "]")).getText();
			}
			phNo2 = "\n" + driver.findElement(By.xpath(str + "span[4]")).getText();
			locTyp2 = "\n" + driver
					.findElement(By.xpath("//*[@id=\"__next\"]/div/div[2]/main/div/div/div[" + i + "]/div/div/div[2]"))
					.getText();
			locHr2 = "\n" + driver
					.findElement(By.xpath("//*[@id=\"__next\"]/div/div[2]/main/div/div/div[" + i + "]/div/div/div[3]"))
					.getText();
			LocationList.add(new Location(city2, address2, phNo2, locTyp2, locHr2));
			try {
				Thread.sleep(2000);
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		}

	}

	public void getCity3Locations(WebDriver driver, WebDriverWait wait) {
		// wait until the required condition is fulfiled
		wait.until(ExpectedConditions.visibilityOfElementLocated(
				By.xpath("//*[@id=\"__next\"]/div/div[2]/main/div/div[1]/div/div[2]/ul/div[18]/li")));

		WebElement cityName3 = driver
				.findElement(By.xpath("//*[@id=\"__next\"]/div/div[2]/main/div/div[1]/div/div[2]/ul/div[18]/li"));
		// To click on the city name
		cityName3.click();

		String str = "//*[@id=\"__next\"]/div/div[2]/main/div/div/div[1]/div/div/div[1]/";
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(str + "p")));

		for (int i = 1; i < 5; i++) {
			str = "//*[@id=\"__next\"]/div/div[2]/main/div/div/div[" + i + "]/div/div/div[1]/";
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(str + "p")));

			String address4 = "", phNo4 = "", locTyp4 = "", locHr4 = "";
			String city7 = driver.findElement(By.xpath(str + "p")).getText();
			for (int j = 1; j <= 2; j++) {
				address4 += "\n" + driver.findElement(By.xpath(str + "span[" + j + "]")).getText();
			}
			phNo4 += "\n" + driver.findElement(By.xpath(str + "span[4]")).getText();
			locTyp4 = "\n" + driver
					.findElement(By.xpath("//*[@id=\"__next\"]/div/div[2]/main/div/div/div[" + i + "]/div/div/div[2]"))
					.getText();
			locHr4 = "\n" + driver
					.findElement(By.xpath("//*[@id=\"__next\"]/div/div[2]/main/div/div/div[" + i + "]/div/div/div[3]"))
					.getText();
			LocationList.add(new Location(city7, address4, phNo4, locTyp4, locHr4));
			try {
				Thread.sleep(2000);
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		}
	}

	public List<Location> getAllLocations() {

		WebDriver driver = initialization();

		// find element and click on the city name
		WebElement cityName1 = driver
				.findElement(By.xpath("//*[@id=\"__next\"]/div/div[2]/main/div/div[1]/div/div[2]/ul/div[1]/li"));
		cityName1.click();

		// Initializing wait object for wait the screen for 30 seconds
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
		String str = "//*[@id=\"location_detail_last_location__c82el\"]/div[1]/div[1]/";

		// pause the operation until the element is present on the screen
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(str + "p")));

		// declaring local variables to temperory store data
		String address1 = "", phNo1 = "", locTyp1 = "", locHr1 = "";
		String city1 = driver.findElement(By.xpath(str + "p")).getText();
		for (int i = 1; i < 3; i++)
			address1 += driver.findElement(By.xpath(str + "span[" + i + "]")).getText() + "\n";

		phNo1 = driver.findElement(By.xpath(str + "span[4]")).getText();
		locTyp1 = driver.findElement(By.xpath("//*[@id=\"location_detail_last_location__c82el\"]/div/div[2]"))
				.getText();
		locHr1 = driver.findElement(By.xpath("//*[@id=\"location_detail_last_location__c82el\"]/div/div[3]")).getText();
		// save data to the list
		LocationList.add(new Location(city1, address1, phNo1, locTyp1, locHr1));
		// sleep method to pause the screen for loading page
		try {
			Thread.sleep(2000);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}

		// method to navigate to the previous page
		driver.navigate().back();

		// getting location of particular city
		getCity2Locations(driver, wait);

		try {
			Thread.sleep(5000);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}

		// navigate to back page
		driver.navigate().back();

		// getting location of particular city
		getCity3Locations(driver, wait);

		//writeDataInExcel();
		return LocationList;
	}

	public void addDataInExcel(XSSFSheet sheet) {
		int rowNum = 1;
		// Adding data in excel sheet in a loop
		for (Location loc : LocationList) {

			XSSFRow row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue(loc.getCity());
			row.createCell(1).setCellValue((loc.getAddress()));
			row.createCell(2).setCellValue((loc.getPhNo().substring(14)));
			row.createCell(3).setCellValue((loc.getLocType().substring(14)));
			row.createCell(4).setCellValue((loc.getLocHrs().substring(15)));
		}
	}
	public String writeDataInExcel() {

		if(LocationList.isEmpty())
			return "This List is an empty list";
		else {
		// creating workbook
		XSSFWorkbook workBook = new XSSFWorkbook();
		XSSFSheet sheet = workBook.createSheet("Location Info");
		String[] columns = { "Cities", "Address", "Phone Numbers", "Location Types", "Location Hours" };
		// Create a Font for styling header cells
		Font headerFont = workBook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.RED.getIndex());
		// Create a CellStyle with the font
		CellStyle headerCellStyle = workBook.createCellStyle();
		headerCellStyle.setFont(headerFont);
		// Create a Row
		XSSFRow headerRow = sheet.createRow(0);
		// Create cells
		for (int i = 0; i < columns.length; i++) {
			XSSFCell cell = headerRow.createCell(i);
			cell.setCellValue(columns[i]);
			cell.setCellStyle(headerCellStyle);
		}
		
		//adding data in excel sheet
		addDataInExcel(sheet);
		
		// Resize all columns to fit the content size
		for (int i = 0; i < columns.length; i++) {
			if (i == 1) {
				sheet.setColumnWidth(i, 50 * 256);
				continue;
			}
			sheet.setColumnWidth(i, 30 * 256);
		}
		String filePath = ".\\datafiles\\location.xlsx";
		FileOutputStream outstream = null;
		try {
			outstream = new FileOutputStream(filePath);
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		}
		try {
			workBook.write(outstream);
		} catch (IOException e) {
			e.printStackTrace();
		}
		// Closing the file
		try {
			outstream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		// Closing the workbook
		try {
			workBook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return "Operation on excel file has been performed";
	}
  }
}
