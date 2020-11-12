package pages;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.openqa.selenium.support.PageFactory;

public class Homepage
{
	public RemoteWebDriver driver;
	
	@FindBy(how=How.NAME,using="identifier")
	public WebElement uid;
	
	@FindBy(how=How.XPATH,using="//*[@class='VfPpkd-RLmnJb']")
	public WebElement uidnext;
	
	@FindBy(how=How.XPATH,using="(//*[contains(text(),'Enter an email')])[2]")
	public WebElement blankuiderr;
	
	@FindBy(how=How.XPATH,using="(//*[contains(text(),'find your Google Account')])[2]")
	public WebElement invaliduiderr;
	/*@FindBy(how=How.XPATH,using="//*[text()=\"Couldn't find your Google Account\"]")
	public WebElement invaliduiderr;*/
	
	public Homepage(RemoteWebDriver driver)
	{
		this.driver=driver;
		PageFactory.initElements(driver,this);
	}
	
	public void fillUID(String x)
	{
		uid.sendKeys(x);
	}
	
	public void clickUIDNext()
	{
		uidnext.click();
	}	
}