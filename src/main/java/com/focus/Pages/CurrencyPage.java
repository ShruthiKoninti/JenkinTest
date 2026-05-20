package com.focus.Pages;

import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;

import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

public class CurrencyPage extends BaseEngine 
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "CurrencyPage";	
	
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
	
	
		
		
		
	
	public static boolean checkOpenCurrencyMenu() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMenu));
		currencyMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMasterMenu));
		currencyMasterMenu.click();
		    	  
		Thread.sleep(6000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveIcon));
		
        String expected = excelReader.getCellData(xlSheetName, 8, 6);
    	
    	String actual = Boolean.toString(saveIcon.isDisplayed());
    	
    	System.out.println("Save Button is Displayed : "+actual+"  "+expected);
    	
    	excelReader.setCellData(xlfile, xlSheetName, 8, 7, actual);
    	
		    
		if(actual.equalsIgnoreCase(expected))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 8, 8, resPass);
			return true;
		}	 
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 8, 8, resFail);
			return false;
		}
 	}
	
	
	
	
	
	
	
	
	public static boolean checkUpdatingISOCurrencyCodeSLLCoinName() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ISOCurrencyCodeTxt));
		ISOCurrencyCodeTxt.click();
		ISOCurrencyCodeTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(coinsNameTxt));
		coinsNameTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		coinsNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
	
		Thread.sleep(2000);
				    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveIcon));
		saveIcon.click();
		Thread.sleep(2000);
	  
		System.out.println("*********************checkUpdatingISOCurrencyCodeSLLCoinName******************************************");
		
		String expMessage=excelReader.getCellData(xlSheetName, 9, 6);
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("actMessage: "+actMessage);
		System.out.println("expMessage: "+expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, actMessage);
		Thread.sleep(2000);
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{	
			 excelReader.setCellData(xlfile, xlSheetName, 9, 8, resPass);
			return true;
		}	 
		else
		{    
			
			 excelReader.setCellData(xlfile, xlSheetName, 9, 8, resFail);
			 return false;
		}
    }
	
	
	
	
	
	
	
	
	public static boolean checkUpdatingAgainISOCurrencyCodeSLLCoinName() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ISOCurrencyCodeTxt));
		ISOCurrencyCodeTxt.click();
		ISOCurrencyCodeTxt.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		
		Thread.sleep(2000);
		
		ISOCurrencyCodeTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(coinsNameTxt));
		coinsNameTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		coinsNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		
		Thread.sleep(2000);
				    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveIcon));
		saveIcon.click();
		Thread.sleep(2000);
	  
		System.out.println("*********************checkUpdatingAgainISOCurrencyCodeSLLCoinName******************************************");
		
		String expMessage=excelReader.getCellData(xlSheetName, 11, 6);
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("actMessage: "+actMessage);
		System.out.println("expMessage: "+expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 11, 7, actMessage);
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMasterCancelBtn));
			currencyMasterCancelBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 11, 8, resPass);
			return true;
		}	 
		else
		{    
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMasterCancelBtn));
			currencyMasterCancelBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 11, 8, resFail);

			return false;
		}
     }
	
	
	public static boolean CheckLogin() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	        Thread.sleep(3000);

	        getDriver().navigate().refresh();

	        LoginPage lp=new LoginPage(getDriver()); 

	        lp.checkLoginPageTitleByURLInputInBrowser();

	        String unamelt="su";

	        String pawslt="su";

	        lp.enterUserName(unamelt);

	        Thread.sleep(2000);
	        
	        lp.enterPassword(pawslt);

	        Thread.sleep(2000);

	        String compname = "Automation Company : 08/10/2020";

	        Select oSelect = new Select(companyDropDownList);

	        List<WebElement> elementCount = oSelect.getOptions();

	        int cqSize = elementCount.size();

	        System.out.println("CompanyDropdownList Count :" + cqSize);

	        int i;

	        for (i = 0; i < elementCount.size(); i++) {

	                elementCount.get(i).getText();

	                String optionName = elementCount.get(i).getText();
	                if (optionName.toUpperCase().startsWith(compname.toUpperCase())) {
	                        System.out.println("q" + elementCount.get(i).getText());
	                        elementCount.get(i).click();
	                }

	        }

	        Thread.sleep(2000);

	        lp.clickOnSignInBtn();

	        Thread.sleep(2000);

	        String actUserInfo1=userNameDisplay.getText();

	        System.out.println("User Info  : "+actUserInfo1);

	        System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

	      /*  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	        companyLogo.click();*/

	        String getCompanyTxt1=Company_Name.getText();
	        String getLoginCompanyName1=getCompanyTxt1.substring(0, 31);
	        System.out.println("company name  :  "+ getLoginCompanyName1);
	       // companyLogo.click();

	        String expUserInfo1           ="SU";
	        String expLoginCompanyName1   ="Automation Company : 08/10/2020";

	        System.out.println("UserInfo1             : "+actUserInfo1            +" Value Expected : "+expUserInfo1);
	        System.out.println("LoginCompanyName1     : "+getLoginCompanyName1    +" Value Expected : "+expLoginCompanyName1);

	        if(actUserInfo1.equalsIgnoreCase(expUserInfo1) && getLoginCompanyName1.contains(expLoginCompanyName1))
	        {
	                return true;
	        }
	        else
	        {
	                return false;
	        }

	}


	public boolean checkLogoutCurrencyPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		 
		 try
			{
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
			  userNameDisplayLogo.click();
			  Thread.sleep(2000);
			 
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			  logoutOption.click();
			  
			  Thread.sleep(2000);
			  
			  boolean actUserLoginPage              = username.isDisplayed() && username.isEnabled()
	                                               && password.isDisplayed() && password.isEnabled();
	                                      
			  boolean expUserLoginPage              = true;
			  
			  if(actUserLoginPage==expUserLoginPage)  
		      {
				System.out.println("***Test Pass: Login Successfull***");
				
				return true;
			  }
		      else
		      {
		  	 
				System.out.println("***Test Fail: Login Not Successfull***");
				
				return false;
			  }
			}
			catch (Exception e)
			{
			 	String exception = e.getMessage();
			 		
				return false;
			}
		}


	
	public CurrencyPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}
	
	
	
}
