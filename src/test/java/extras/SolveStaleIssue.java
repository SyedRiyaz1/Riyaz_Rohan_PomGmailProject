package extras;

import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.openqa.selenium.support.PageFactory;

import io.github.bonigarcia.wdm.WebDriverManager;

public class SolveStaleIssue 
{
	//Properties
	public ChromeDriver driver;
	
	@FindBy(how=How.NAME,using="identifier")
	public WebElement e;
	//Constructor method
	public SolveStaleIssue(ChromeDriver driver)
	{
		this.driver=driver;
		PageFactory.initElements(driver,this);
	}
	//main method to run
	public static void main(String[] args) throws Exception
	{
		//Open browser
		WebDriverManager.chromedriver().setup();
		System.setProperty("webdriver.chrome.silentOutput","true");
		ChromeDriver driver=new ChromeDriver();
		SolveStaleIssue obj=new SolveStaleIssue(driver);
		//Launch site
		driver.get("http://www.gmail.com");
		Thread.sleep(3000);
		//Locate an element
		obj.e.sendKeys("magnitiait",Keys.ENTER);
		Thread.sleep(3000);
		driver.navigate().back();
		Thread.sleep(3000);
		obj.e.clear();
		obj.e.sendKeys("seleniumhalfstack",Keys.ENTER);
		Thread.sleep(3000);
		//Close site
		driver.close();
	}
}






