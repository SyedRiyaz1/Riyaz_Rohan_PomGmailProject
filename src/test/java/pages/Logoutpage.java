package pages;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.openqa.selenium.support.PageFactory;

public class Logoutpage
{
	public RemoteWebDriver driver;
	
	@FindBy(how=How.XPATH,using="//*[contains(@aria-label,'Google Account')]")
	public WebElement profilepic;
	
	@FindBy(how=How.XPATH,using="(//*[contains(text(),'Sign out')])[2]")
	public WebElement signout;
	
	public Logoutpage(RemoteWebDriver driver)
	{
		this.driver=driver;
		PageFactory.initElements(driver,this);
	}
	
	public void clickProfilePic()
	{
		profilepic.click();
	}
	
	public void clickSignout()
	{
		signout.click();
	}
}
