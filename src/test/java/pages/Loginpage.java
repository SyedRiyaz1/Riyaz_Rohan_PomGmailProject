package pages;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.openqa.selenium.support.PageFactory;

public class Loginpage
{
public RemoteWebDriver driver;
	
	@FindBy(how=How.NAME,using="password")
	public WebElement pwd;
	
	@FindBy(how=How.XPATH,using="//*[@class='VfPpkd-RLmnJb']")
	public WebElement pwdnext;
	
	@FindBy(how=How.XPATH,using="//*[text()='Enter a password']")
	public WebElement blankpwderr;
	
	@FindBy(how=How.XPATH,using="//*[contains(text(),'Wrong password')]")
	public WebElement invalidpwderr;
	
	public Loginpage(RemoteWebDriver driver)
	{
		this.driver=driver;
		PageFactory.initElements(driver,this);
	}
	
	public void fillPWD(String x)
	{
		pwd.sendKeys(x);
	}
	
	public void clickPWDNext()
	{
		pwdnext.click();
	}
}
