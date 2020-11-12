package utilities;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;

import org.openqa.selenium.OutputType;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.io.FileHandler;
import org.openqa.selenium.opera.OperaDriver;
import org.openqa.selenium.remote.RemoteWebDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class TestUtilityWithoutDecorations 
{
	public RemoteWebDriver driver;
	
	public TestUtilityWithoutDecorations()
	{
		driver=null;
	}
	
	//Reusable methods
	//Without cross browser
	public RemoteWebDriver launchChromeBrowser()
	{
		WebDriverManager.chromedriver().setup();
		driver=new ChromeDriver();
		return(driver);
	}
	//With cross browser
	public RemoteWebDriver launchBrowser(String bn)
	{
		if(bn.equalsIgnoreCase("chrome"))
		{
			WebDriverManager.chromedriver().setup();
			driver=new ChromeDriver();
		}
		else if(bn.equalsIgnoreCase("firefox"))
		{
			WebDriverManager.firefoxdriver().setup();
			driver=new FirefoxDriver();
		}
		else if(bn.equalsIgnoreCase("opera"))
		{
			WebDriverManager.operadriver().setup();
			driver=new OperaDriver();
		}
		else if(bn.equalsIgnoreCase("edge"))
		{
			WebDriverManager.edgedriver().setup();
			driver=new EdgeDriver();
		}
		else if(bn.equalsIgnoreCase("ie"))
		{
			WebDriverManager.iedriver().setup();
			driver=new InternetExplorerDriver();
		}
		else
		{
			System.out.println("Please enter correct browser name");
			System.exit(0);   //Stop execution forcibly
		}
		return(driver);
	}
	
	public void launchSite(String url)
	{
		//Launch Site and maximize
		driver.get(url);
		driver.manage().window().maximize();
	}
	
	public Properties accessProperties() throws Exception
	{
		//Calling Properties
		File f1=new File("E:\\Automation\\AutomationNested\\com.pom.gmail\\src\\test\\resources\\properties\\config.properties");
		FileInputStream fi1=new FileInputStream(f1);
		Properties pro=new Properties();
		pro.load(fi1);
		return(pro);
	}
	
	public void closeSite()
	{
		//Close site
		driver.close();
	}
	
	public String screenshot() throws Exception
	{
		SimpleDateFormat sf=new SimpleDateFormat("dd-MMM-yyyy-hh-mm-ss");
		Date dt=new Date();
		File src=driver.getScreenshotAs(OutputType.FILE);
		String ssname=sf.format(dt)+".png";
		File dest=new File(ssname);
		FileHandler.copy(src,dest);
		return(dest.getAbsolutePath());
	}
}
