package extras;

import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CacheTest
{
	public static void main(String[] args)
	{
		//Open browser
		WebDriverManager.chromedriver().setup();
		System.setProperty("webdriver.chrome.silentOutput","true");
		ChromeDriver driver=new ChromeDriver();
		//Create object to Page class
		GooglePage sp=new GooglePage(driver);
		//Launch site
		driver.get("http://www.google.co.in");
		//Get text of link 1000 times
		long x=System.currentTimeMillis();
		for(int i=1;i<=1000;i++)
		{
			sp.mylink1.getText();
		}
		long y=System.currentTimeMillis();
		System.out.println("Time taken in seconds Without cache " +(y-x)/1000); //more time
		//Get text of link 1000 times
		long z=System.currentTimeMillis();
		for(int i=1;i<=1000;i++)
		{
			sp.mylink2.getText();
		}
		long w=System.currentTimeMillis();
		System.out.println("Time taken in seconds With cache " +(w-z)/1000); //9
		//close site
		driver.close();
	}
}








