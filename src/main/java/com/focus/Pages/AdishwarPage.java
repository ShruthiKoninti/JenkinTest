package com.focus.Pages;

import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.record.SaveRecalcRecord;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.server.handler.GetCurrentUrl;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;

import com.focus.base.BaseEngine;

public class AdishwarPage  extends BaseEngine
{


	@FindBy(xpath="//select[@id='ddlCompany']")
	private static WebElement companyDropDownList;
	
    
	
	public boolean checkLogin() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		
		getDriver().navigate().to("http://175.101.7.11:5252/Focus9");
		
		LoginPage lp=new LoginPage(getDriver()); 
		
		String unamelt="su";
			      
		String pawslt="su";
			      
		lp.enterUserName(unamelt);
		
		Thread.sleep(2000);
			
		lp.enterPassword(pawslt);
		
		
		Select oSelect = new Select(companyDropDownList);
		
		oSelect.selectByValue("3060");
		
		Thread.sleep(2000);
		
		lp.clickOnSignInBtn();
		
		Thread.sleep(3000);
			           	
		String userInfo=userNameDisplay.getText();
			    	
		System.out.println("User Info : "+userInfo);
			    	
		System.out.println("User Info Capture Text :"+userNameDisplay.getText());
			    	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
			    	
		companyLogo.click();
			    	
		if(userInfo.equalsIgnoreCase("SU"))
		{
		
		System.out.println("Test Pass :Logined to Billwise Company");
		return true;
		
		}
		else
		{
		System.out.println("Test Fail :Logined to Billwise Company");
		return false;
		
		}
	}
	
	
	

@FindBy(xpath="//span[@class='hidden-xs']")
private static WebElement userNameDisplay;

@FindBy(xpath="//*[@id='ulCompanyDetails_HomePage']/li[1]")
private static WebElement companyName;

@FindBy(xpath="//*[@id='companyLogo']")
private static WebElement companyLogo;

@FindBy(xpath="//input[@id='txtSearchMenu_MainLayout']")
private static WebElement searchTxt;

@FindBy(xpath="//input[@id='id_header_1']")
private static WebElement  documentNumberTxt;


//Vouchers Home Page Ribbon Control
@FindBy(xpath="//div[@id='myNavbar']//span[@class='icon-new icon-font6']")
private static WebElement  newBtn;
@FindBy(xpath="//div[@class='theme_color font-6']")
public static WebElement errorMessage;
 
@FindBy(xpath="//span[@class='icon-reject2 theme_color']")
public static WebElement errorMessageCloseBtn;


public static String checkValidationMessage(String ExpMessage) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	try
	{
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(errorMessage));
		String actErrorMessage=errorMessage.getText();
		String expErrorMessage=ExpMessage;
		
		try
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
			errorMessageCloseBtn.click();
			
			System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);
	
	return actErrorMessage;
}
catch(Exception ee)
{
	
	System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);
		
		return actErrorMessage;
	}
}
catch(Exception e)
{
	System.err.println("Error Message NOT Found or NOT Clickable");
		System.err.println(e.getMessage());
              
		String Exception=e.getMessage();
		
		return Exception;	
	} 
}	


public boolean checkOpenVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
	searchTxt.click();
	searchTxt.sendKeys("Receipt Against Order");
	Thread.sleep(1000);
	searchTxt.sendKeys(Keys.ENTER);
	
	Thread.sleep(2000);
	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
	newBtn.click();
	
	checkValidationMessage("screen opened");
	
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END);
		documentNumberTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys("2122-1041-358");
		Thread.sleep(1200);
		documentNumberTxt.sendKeys(Keys.TAB);
		
		checkValidationMessage("Voucher loaded successfully");
		
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
		Thread.sleep(3000);
		
		int size = getDriver().findElements(By.tagName("iframe")).size();
		System.err.println("int  "+size);
		
		getDriver().switchTo().frame(size-1);
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(okbtn));
		okbtn.click();
		
		getWaitForAlert();
		getAlert().accept();
		
		Thread.sleep(1000);
		
		getWaitForAlert();
		getAlert().accept();
		
		Thread.sleep(4000);
		
		getDriver().navigate().refresh();
		
		Thread.sleep(1000);
		
		return true;
	
}

public boolean checkLoop() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	for (int i = 0; i < 20; i++) 
	{
		System.err.println(i+1);
		checkOpenVoucher();
	}
	return true;
}

@FindBy(xpath="//body/div[1]/div[5]/div[1]/button[1]")
private static WebElement  okbtn;;


@FindBy(xpath="//*[@id='id_transactionentry_save']/div[2]")
private static WebElement  saveBtn;


public AdishwarPage(WebDriver driver)
{
	
	PageFactory.initElements(driver, this);
}

	
	
	
}

