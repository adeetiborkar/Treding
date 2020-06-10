package GaneshBorkar.NSE;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class Base {
	public  WebDriver driver;
	public Properties prop;
	
	public WebDriver initialiseDriver () throws IOException {
		
		// TODO Auto-generated method stub
		prop = new Properties();

		FileInputStream file = new FileInputStream(
				System.getProperty("user.dir") + "\\src\\main\\java\\resources\\data.properties");
		prop.load(file);
		String browserName = prop.getProperty("browser");

		System.setProperty("webdriver.chrome.driver",
				System.getProperty("user.dir") + "\\src\\main\\java\\resources\\chromedriver.exe");
		// run test case without invoking browser use ChromeOption
		ChromeOptions options = new ChromeOptions();
		if (browserName.contains("headless")) {
			options.addArguments("--headless");
		}
		driver = new ChromeDriver(options);

		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		String url = prop.getProperty("URL");
		driver.get(url);
		driver.manage().window().maximize();
		return driver;
	}
	
	public String createFile() {
		//String base_path =System.getProperty("user.dir"); 
		try {
		      File myObj = new File(System.getProperty("user.dir")+"\\reports\\Report.xlsx");
		      if (myObj.createNewFile()) {
		        System.out.println("File created: " + myObj.getName());
		      } else  {
		        System.out.println("File already exists.");
		      }
		    } catch (IOException e) {
		      System.out.println("An error occurred.");
		      e.printStackTrace();
		    }
				
		String path =System.getProperty("user.dir")+"\\reports\\Report.xlsx";
		//File file= new File(System.getProperty("user.dir")+"\\reports\\Report.xlsx");
		//return file;
		return path;
	}
	

}
