package GaneshBorkar.NSE;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Temp2 {

	
	
	public static void main(String[] agr) throws InterruptedException, IOException {
		System.setProperty("webdriver.chrome.driver",
				System.getProperty("user.dir") + "\\src\\main\\java\\resources\\chromedriver.exe");

		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://www.edelweiss.in/market/nse-option-chain");
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
	
		/*************************************
		 * Pop up close
		 ************************************************************/
		Thread.sleep(2000);
		if (driver.findElements(By.xpath("//*[@id=\"wzrk-cancel\"]")).size() > 0) {
			driver.findElement(By.xpath("//*[@id=\"wzrk-cancel\"]")).click();
		}
		/******************************************
		 * Frame close
		 ******************************************************/
		Thread.sleep(25000);
		
		WebElement frame = driver.findElement(By.id("wiz-iframe-intent"));
		if (frame.isDisplayed()) {
		driver.switchTo().frame(frame);
		String text = driver.findElement(By.xpath("//*[@id=\"contentDiv\"]/div/div/p/a")).getAttribute("href");
		System.out.println(text);
		driver.findElement(By.xpath("//*[@id=\'contentDiv\']/div/div/p/a/span")).click();
		}
		//driver.navigate().refresh();
		Thread.sleep(7000);
		// driver.findElement(By.xpath("//div[@id='tableContent']//a[contains(text(),'Show
		// all')]")).click();
		driver.findElement(By.xpath(
				"//*[@id=\"tableContent\"]/table[2]/tfoot/tr/td[8]/a/span"))
				.click();
		Thread.sleep(7000);
		/***************************************
		 * Create .xls file
		 ************************************************/
		String path = System.getProperty("user.dir") + "\\reports\\Report.xlsx";
		File file = new File(path);
		if (!file.exists()) {
			file.createNewFile();
		}
		FileInputStream fin = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fin);
		// workbook.createSheet("TestData2");
		
		XSSFSheet sheet = workbook.getSheetAt(0);
		workbook.setSheetName(0, "edelweiss");
		FileOutputStream fout = new FileOutputStream(file);
		/************************
		 * write column in xls
		 *************************************/

		List<WebElement> column_Name = driver
				.findElements(By.xpath("//*[@id=\"optioncontainer\"]/div[3]/table/thead/tr[2]/th"));
		int Colunm_Size = column_Name.size();
		System.out.println(Colunm_Size);
		ArrayList<String> c = new ArrayList<String>();

		for (int i = 0; i < Colunm_Size; i++) {
			WebElement z = column_Name.get(i);
			String Column_Title = z.getText().toUpperCase();
			c.add(Column_Title);

			System.out.println(Column_Title);
		}

		int c_Size = c.size();
		Row ro = sheet.createRow(0);

		for (int j = 0; j < c_Size; ++j) {
			sheet.autoSizeColumn(j);
			ro.createCell(j).setCellValue(c.get(j));
			System.out.println(c.get(j));
		}
		// ***************************************************************** Second
		// Row**********************************************/

		List<WebElement> rows = driver.findElements(By.xpath("//*[@id=\"tableContent\"]/table[1]/tbody/tr"));
		int Totle_Rows = rows.size();
		System.out.println(Totle_Rows);

		for (int k = 1; k <= Totle_Rows; k++) {
			String before_xPath = "//*[@id=\"tableContent\"]/table[1]/tbody/tr[";
			String after_xPath = "]/td";
			String actual_xPath = before_xPath + k + after_xPath;
			List<WebElement> row_Data1 = driver.findElements(By.xpath(actual_xPath));
			int row1_Size = row_Data1.size();

			System.out.println(row1_Size);
			ArrayList<String> r1 = new ArrayList<String>();

			for (int i = 0; i < row1_Size; i++) {
				WebElement row1 = row_Data1.get(i);
				String row_data = row1.getText();
				r1.add(row_data);

				System.out.println(row_data);
			}

			int r1_Size = r1.size();
			Row ro1 = sheet.createRow(k);

			for (int j = 0; j < r1_Size; ++j) {
				sheet.autoSizeColumn(j);
				ro1.createCell(j).setCellValue(r1.get(j));
				System.out.println(r1.get(j));
			}

		}

		workbook.write(fout);
		workbook.close();
		driver.quit();

	}
}
