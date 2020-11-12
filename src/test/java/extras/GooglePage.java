package extras;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.CacheLookup;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.openqa.selenium.support.PageFactory;

public class GooglePage
{
	//Properties
	public RemoteWebDriver driver;
	
	@FindBy(how=How.LINK_TEXT,using="Gmail")
	public WebElement mylink1;
	
	@FindBy(how=How.LINK_TEXT,using="Gmail")
	@CacheLookup
	public WebElement mylink2;
	
	//Constructor method
	public GooglePage(RemoteWebDriver driver)
	{
		this.driver=driver;
		PageFactory.initElements(driver,this);
	}
}









