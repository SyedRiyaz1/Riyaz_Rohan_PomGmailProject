package tests;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import pages.Homepage;
import pages.Loginpage;
import pages.Logoutpage;
import utilities.TestUtility;

public class GmailLoginTestWithoutDecorations 
{
	public static void main(String[] args) throws Exception
	{
		//Connect to Excel file
		File f=new File("gmaillogintestwithoutdecorations.xlsx");
		FileInputStream fi=new FileInputStream(f);
		Workbook wb=WorkbookFactory.create(fi);
		Sheet sh=wb.getSheet("Sheet1");
		int nour=sh.getPhysicalNumberOfRows();
		int nouc=sh.getRow(0).getLastCellNum();
		//Give headings to results in excel file
		SimpleDateFormat sf=new SimpleDateFormat("dd-MMM-yyyy-hh-mm-ss");
		Date dt=new Date();
		String cname=sf.format(dt);
		sh.getRow(0).createCell(nouc).setCellValue("Result on "+cname);
		//Create object to utility class
		TestUtility tu=new TestUtility();
		//Data Driven from 2nd row(index=1)
		for(int i=1;i<nour;i++)
		{
			//Read data from excel
			DataFormatter df=new DataFormatter();
			String uid=df.formatCellValue(sh.getRow(i).getCell(0));
			String uidc=df.formatCellValue(sh.getRow(i).getCell(1));
			String pwd=df.formatCellValue(sh.getRow(i).getCell(2));
			String pwdc=df.formatCellValue(sh.getRow(i).getCell(3));
			
			//RemoteWebDriver driver=tu.launchChromeBrowser();
			RemoteWebDriver driver=tu.launchBrowser("chrome"); //Add column and take browser name from excel for crossbrowser
			//Activate properties file
			Properties pro=tu.accessProperties();
			//Launch site
			tu.launchSite(pro.getProperty("url"));
			//Create wait object
			int w=Integer.parseInt(pro.getProperty("maxwait"));
			WebDriverWait wait=new WebDriverWait(driver,w);
			//Create page class objects
			Homepage hp=new Homepage(driver);
			Loginpage lginp=new Loginpage(driver);
			Logoutpage lgoutp=new Logoutpage(driver);
			wait.until(ExpectedConditions.visibilityOf(hp.uid));
			hp.fillUID(uid);
			wait.until(ExpectedConditions.elementToBeClickable(hp.uidnext));
			hp.clickUIDNext();
			Thread.sleep(5000);
			//UID Validations
			try
			{
				if(uidc.equalsIgnoreCase("blank") && hp.blankuiderr.isDisplayed())
				{
					sh.getRow(i).createCell(nouc).setCellValue("Blank UID Test Passed");
				}
				else if(uidc.equalsIgnoreCase("invalid") && hp.invaliduiderr.isDisplayed())
				{
					sh.getRow(i).createCell(nouc).setCellValue("Invalid UID Test Passed");
				}
				else if(uidc.equalsIgnoreCase("valid") && lginp.pwd.isDisplayed())
				{
					sh.getRow(i).createCell(nouc).setCellValue("Valid UID Test Passed");
					//PWD Validations
					lginp.fillPWD(pwd);
					wait.until(ExpectedConditions.elementToBeClickable(lginp.pwdnext));
					lginp.clickPWDNext();
					Thread.sleep(5000);
					if(pwdc.equalsIgnoreCase("blank") && lginp.blankpwderr.isDisplayed())
					{
						sh.getRow(i).createCell(nouc).setCellValue("Blank PWD Test Passed");
					}
					else if(pwdc.equalsIgnoreCase("invalid") && lginp.invalidpwderr.isDisplayed())
					{
						sh.getRow(i).createCell(nouc).setCellValue("Invalid PWD Test Passed");
					}
					else if(pwdc.equalsIgnoreCase("valid") && lgoutp.profilepic.isDisplayed())
					{
						sh.getRow(i).createCell(nouc).setCellValue("Valid PWD Test Passed");
						//Logout
						wait.until(ExpectedConditions.elementToBeClickable(lgoutp.profilepic));
						lgoutp.clickProfilePic();
						wait.until(ExpectedConditions.elementToBeClickable(lgoutp.signout));
						lgoutp.clickSignout();
						wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[text()='Use another account']")));
					}
					else
					{
						String sspath=tu.screenshot();
						sh.getRow(i).createCell(nouc).setCellValue("Valid PWD Test Failed and refer "+sspath);
					}
				}
				else
				{
					String sspath=tu.screenshot();
					sh.getRow(i).createCell(nouc).setCellValue("Valid UID Test Failed and refer "+sspath);
				}
			}
			catch(Exception ex)
			{
				System.out.println(ex.getMessage());
			}
			
			//Close site
			tu.closeSite();	
		}
		
		sh.autoSizeColumn(nouc);
		
		//Save data back to excel
		FileOutputStream fo=new FileOutputStream(f);
		wb.write(fo);
		fi.close();
		fo.close();
		wb.close();
	}
}
