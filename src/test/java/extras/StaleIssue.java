package extras;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import io.github.bonigarcia.wdm.WebDriverManager;

public class StaleIssue 
{
	public static void main(String[] args) throws Exception
	{
		//Open browser
		WebDriverManager.chromedriver().setup();
		System.setProperty("webdriver.chrome.silentOutput","true");
		ChromeDriver driver=new ChromeDriver();
		//Launch site
		driver.get("http://www.gmail.com");
		driver.manage().window().maximize();
		Thread.sleep(3000);
		//Locate an element
		WebElement e=driver.findElement(By.name("identifier"));
		//Operate element and goto next page
		e.sendKeys("magnitiait",Keys.ENTER);
		Thread.sleep(3000);
		//Back to previous page
		driver.navigate().back();
		Thread.sleep(3000);
		//Operate element
		e.sendKeys("seleniumhalfstack",Keys.ENTER);
		Thread.sleep(3000);
		//Close site
		driver.close();
	}
}







