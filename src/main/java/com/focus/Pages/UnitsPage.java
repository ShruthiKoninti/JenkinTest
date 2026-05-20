package com.focus.Pages;

import java.io.IOException;
import java.util.ArrayList;
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

public class UnitsPage extends BaseEngine
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "UnitsPage";

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



	public  boolean checkUnitMasterInMAsterMenu() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		System.out.println("***********checkUnitMasterInUnitMenu Method Excutes **********");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		homeMasterItemMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitsMenu));
		unitsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		
		String expected = excelReader.getCellData(xlSheetName, 8, 6);
    	String actual = Boolean.toString(masterNewBtn.isDisplayed());
    	
    	System.out.println("Act New Button is Displayed : "+actual);
    	System.out.println("Exp New Button is Displayed : "+expected);
    	
    	
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





	public static boolean checkSavingUnitsInUnitsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		int count = masterItemUnitsList.size();

		ArrayList<String> actItemUnitsList = new ArrayList<String>();

		for (int i = 0; i < count; i++) 
		{
			String itemUnit = masterItemUnitsList.get(i).getText();

			actItemUnitsList.add(itemUnit);
		}

		/*ArrayList<String> expItemUnitsList = new ArrayList<String>();
		expItemUnitsList.add("Pcs, Dozs");*/

		String actUnits = actItemUnitsList.toString();
		String expUnits = excelReader.getCellData(xlSheetName, 10, 6);//expItemUnitsList.toString();

		System.out.println("********************************checkSavingUnitsInUnitsMaster********************");

		System.out.println(actUnits);
		System.out.println(expUnits);
		
		excelReader.setCellData(xlfile, xlSheetName, 10, 7, actUnits);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		masterNewBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.clear();
		nameTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		nameTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.clear();
		codeTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		codeTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noOfDecimalsTxt));
		noOfDecimalsTxt.click();
		noOfDecimalsTxt.clear();
		noOfDecimalsTxt.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		noOfDecimalsTxt.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();

		String expMessage=excelReader.getCellData(xlSheetName, 9, 6);
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("actMessage: "+actMessage);
		System.out.println("expMessage: "+expMessage);

		Thread.sleep(6000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
		masterCloseBtn.click();
		
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, actMessage);

		if(actUnits.equalsIgnoreCase(expUnits) && actMessage.equalsIgnoreCase(expMessage))
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





	public static boolean checkEditingInSavedUnitsINUnitsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(4000);

		
		System.out.println("********************************checkEditingInSavedUnitsINUnitsMaster********************8");

		int count = masterItemUnitsList.size();

		for (int i = 0; i < count; i++) 
		{
			String itemUnit = masterItemUnitsList.get(i).getText();

			if (itemUnit.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 12, 5))) 
			{
				getAction().doubleClick(masterItemUnitsList.get(i)).build().perform();
				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.clear();
		nameTxt.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
		nameTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.clear();
		codeTxt.sendKeys(excelReader.getCellData(xlSheetName, 14, 5));
		codeTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();

		String expMessage=excelReader.getCellData(xlSheetName, 12, 6);
		String actMessage=checkValidationMessage(expMessage);

		System.out.println("actMessage: "+actMessage);
		System.out.println("expMessage: "+expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actMessage);

		Thread.sleep(6000);

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 8, resPass);

			return true;
		}

		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 8, resFail);
			return false;
		}
	}







	public static boolean checkUpdatedUnitsMasterDisplayInUnits() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		String actaccountNewCreationName = null;

		int count = masterItemUnitsList.size();

		for (int i = 0; i < count; i++) 
		{
			String itemUnit = masterItemUnitsList.get(i).getText();

			if (itemUnit.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 15, 5))) 
			{
				actaccountNewCreationName = itemUnit;
				break;
			}
		}

		String expaccountNewCreationName  =excelReader.getCellData(xlSheetName, 15, 6);

		System.out.println("************************checkUpdatedItemDisplayInItemMasterGridBeforeCustomization***********************");

		System.err.println("actaccountNewCreationName  : "+actaccountNewCreationName);
		System.err.println("expaccountNewCreationName  : "+expaccountNewCreationName);

		 excelReader.setCellData(xlfile, xlSheetName, 15, 7, actaccountNewCreationName);
		
		if(actaccountNewCreationName.equalsIgnoreCase(expaccountNewCreationName))
		{ 
			 excelReader.setCellData(xlfile, xlSheetName, 15, 8, resPass);

			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 15, 8, resFail);
			return false;
		}
	}



	


	public static boolean checkDeleteUnitsInUnitsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		int count = masterItemUnitsList.size();

		for (int i = 0; i < count; i++) 
		{
			String data = masterItemUnitsList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 16, 5))) 
			{
				if (masterGridBodyChkbox.get(i).isSelected()==false) 
				{
					masterGridBodyChkbox.get(i).click();
					break;
				}
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDeleteBtn));
		masterDeleteBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getMsgOnDelete));
		System.out.println(getMsgOnDelete.getText());

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelete));
		clickOnOkInDelete.click();

		String expMessage=excelReader.getCellData(xlSheetName, 16, 6);

		String actMessage=checkValidationMessage(expMessage);

		System.out.println("********************************checkDeleteItemInItemMaster***************************");

		System.out.println("actMessage: "+actMessage);
		System.out.println("expMessage: "+expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 16, 7, actMessage);


		if(actMessage.equalsIgnoreCase(expMessage))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
			closeBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 16, 8, resPass);

			return true;
		}

		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
			closeBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 16, 8, resFail);

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


	public boolean checkLogoutUnitsPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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









	public UnitsPage(WebDriver driver)
	{
		PageFactory.initElements(driver, this);

	}




}
