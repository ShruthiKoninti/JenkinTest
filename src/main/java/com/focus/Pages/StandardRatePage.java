package com.focus.Pages;

import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;

import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

public class StandardRatePage extends BaseEngine 
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "StandardRatePage";
	
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



@FindBy(xpath="//*[@id='160']/span")
private static WebElement standardRateMenu;



	public static boolean checkStandardRateDeleteOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
		utilities.click();
		Thread.sleep(2000);
		
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", standardRateMenu);
		
		Thread.sleep(2000);
		getAction().moveToElement(standardRateMenu).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(standardRateMenu));
		standardRateMenu.click();
		

	/*	int count = utilitiesSubMenusList.size();
		System.err.println(count);

		for (int i = 0; i < count; i++) 
		{
			String data = utilitiesSubMenusList.get(i).getText();

			System.err.println(data);

			//ScrollIntoView(utilitiesSubMenusList.get(i));

			if (data.equalsIgnoreCase("Standard Rate")) 
			{
				
				utilitiesSubMenusList.get(i).click();
				break;
			}
			//ScrollIntoView(utilitiesSubMenusList.get(i));
		}*/

		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(standardRateItemCombo));
		standardRateItemCombo.click();
		standardRateItemCombo.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		Thread.sleep(3000);
		standardRateItemCombo.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(standardRateDeleteRow1));
		standardRateDeleteRow1.click();	 		

		getWaitForAlert();
		getDriver().switchTo().alert().accept();

		Thread.sleep(2000);

		String actstandardRateTableRow1Rate				 = standardRateRateTxt.getAttribute("value");

		String expstandardRateTableRow1Rate				 = excelReader.getCellData(xlSheetName, 8, 6);
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actstandardRateTableRow1Rate);

		System.out.println("*************************************checkStandardRateDeleteOption***************************************");
		System.out.println("StandardRateTableRow1Rate     : "+actstandardRateTableRow1Rate           +"  value expected  "+expstandardRateTableRow1Rate);

		if (actstandardRateTableRow1Rate.equalsIgnoreCase(expstandardRateTableRow1Rate))
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






	public static boolean checkSavingStandardRate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		standardRateRateTxt.sendKeys(Keys.END);
		standardRateRateTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		standardRateRateTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		standardRateRateTxt.sendKeys(Keys.TAB);
		standardRateAddBtn.click();

		String actstandardRateTableRow1Rate				 = standardRateTableRow1Rate.getText();
		String expstandardRateTableRow1Rate				 = excelReader.getCellData(xlSheetName, 10, 6);
		excelReader.setCellData(xlfile, xlSheetName, 10, 7, actstandardRateTableRow1Rate);

		System.out.println("*************************************checkSavingStandardRate***************************************");

		System.out.println("StandardRateTableRow1Rate            : "+actstandardRateTableRow1Rate           +"  value expected  "+expstandardRateTableRow1Rate);

		if (actstandardRateTableRow1Rate.equalsIgnoreCase(expstandardRateTableRow1Rate))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 10, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 10, 8, resFail);
			return false;
		}
	}






	public static boolean checkEditingStandardRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(standardRateEditRow1));
		standardRateEditRow1.click();

		standardRateRateTxt.sendKeys(Keys.END);
		standardRateRateTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		standardRateRateTxt.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		standardRateRateTxt.sendKeys(Keys.TAB);
		
		standardRateEditBtn.click();

		String actstandardRateTableRow1Rate				 = standardRateTableRow1Rate.getText();
		String expstandardRateTableRow1Rate				 = excelReader.getCellData(xlSheetName, 11, 6);
		excelReader.setCellData(xlfile, xlSheetName, 11, 7, actstandardRateTableRow1Rate);

		System.out.println("*************************************checkEditingStandardRate***************************************");

		System.out.println("StandardRateTableRow1Rate            : "+actstandardRateTableRow1Rate           +"  value expected  "+expstandardRateTableRow1Rate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtnInStandardRate));
		closeBtnInStandardRate.click();

		if (actstandardRateTableRow1Rate.equalsIgnoreCase(expstandardRateTableRow1Rate))
		{
			excelReader.setCellData(xlfile, xlSheetName, 11, 8, resPass);
			return true;
		}
		else
		{
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

	        /*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	        companyLogo.click();*/

	        String getCompanyTxt1=Company_Name.getText();
	        String getLoginCompanyName1=getCompanyTxt1.substring(0, 31);
	        System.out.println("company name  :  "+ getLoginCompanyName1);
	        //companyLogo.click();

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


	public boolean checkLogoutStandardRatePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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























	public StandardRatePage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
