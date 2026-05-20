package com.focus.Pages;

import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

public class TransactionAuthorizationPage extends BaseEngine 
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "TransactionAuthorizationPage";
	
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










	public boolean checkSavingTransactionAuthorizationInPurchaseReturns() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());

		//getAction().moveToElement(settingsmenuBtn).build().perform();
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();*/
		
		ClickUsingJs(settingsmenuBtn);

		/*int count = SettingsSubMenusList.size();

		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();

			System.err.println(data);

			ScrollIntoView(SettingsSubMenusList.get(i));

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 8, 5))) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}*/

		Thread.sleep(4000);
		
		ClickUsingJs(Transactionautorization);
		
		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNameDropdown));
		documentNameDropdown.click();
		Thread.sleep(1000);
		documentNameDropdown.sendKeys(Keys.SPACE );
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesReturns));
		purchasesReturns.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transAuthNameTxt));
		//transAuthNameTxt.click();
		transAuthNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		Thread.sleep(1000);
		transAuthNameTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addLevelBtn));
		ClickUsingJs(addLevelBtn);

		Thread.sleep(5500);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
		descriptionTxt.click();
		Thread.sleep(1000);
		descriptionTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		
		//((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", conditionNotRequiredChkBox);
		
		Thread.sleep(2000);
		
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,300)","");
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
		conditionNotRequiredChkBox.click();

		Thread.sleep(2000);

		 ((JavascriptExecutor)getDriver()).executeScript("window.scrollTo(0, 0)","");
		 Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
		userSelectionTab.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
		userrow1.click();
		Select user = new Select(userDropdown);
		user.selectByVisibleText(excelReader.getCellData(xlSheetName, 11, 5));

		Thread.sleep(2000);

		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
		definitionOkBtn.click();

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthSaveBtn));
		tranAuthSaveBtn.click();

		System.out.println("*****************************************checkSavingTransactionAuthorizationInPurchaseReturns******************************************");

		String expValidationMsg=excelReader.getCellData(xlSheetName, 8, 6);
		String actValidationMsg=checkValidationMessage(expValidationMsg);
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actValidationMsg);

	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthCancelBtn));
		tranAuthCancelBtn.click();*/

		if(/*actCount==expCount && actTemplateList.equalsIgnoreCase(expTemplateList) &&*/ actValidationMsg.equalsIgnoreCase(expValidationMsg))
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








	public boolean 	checkEditingTransactionAuthorizationInPurchaseReturns() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getAction().moveToElement(settingsmenuBtn).build().perform();
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();*/
		
		ClickUsingJs(settingsmenuBtn);

		/*int count1 = SettingsSubMenusList.size();

		for (int i = 0; i < count1; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();

			System.err.println(data);

			ScrollIntoView(SettingsSubMenusList.get(i));

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 12, 5))) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}

		Thread.sleep(10000);*/
		
		Thread.sleep(4000);
		
		ClickUsingJs(Transactionautorization);
		
		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNameDropdown));
		documentNameDropdown.click();
		documentNameDropdown.sendKeys(Keys.SPACE);
		Thread.sleep(6000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesReturns1));
		purchasesReturns1.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transAuthNameTxt));
		transAuthNameTxt.click();
		transAuthNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
		Thread.sleep(3000);
		transAuthNameTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1Icon));
		getAction().doubleClick(Level1Icon).build().perform();
		
		Thread.sleep(4000);

		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)","");
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
		conditionNotRequiredChkBox.click();

		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Row1ConjunctionDropdown));
		Row1ConjunctionDropdown.click();
		Thread.sleep(2000);

		Select s1=new Select(Row1ConjunctionDropdown);

		s1.selectByValue(excelReader.getCellData(xlSheetName, 14, 5));
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Row1SelectFieldTxt));
		Row1SelectFieldTxt.click();

		Thread.sleep(6000);
		
		 ((JavascriptExecutor)getDriver()).executeScript("window.scrollTo(0, 0)","");
		 Thread.sleep(2000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Row1SelectFieldTxt));
			Row1SelectFieldTxt.click();

			Thread.sleep(6000);
			Row1SelectFieldTxt.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldVendorAC));
		selectFieldVendorAC.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldVendorACName));
		selectFieldVendorACName.click();

		Thread.sleep(2000);

		Select s2=new Select(Row1SelectOperatorDropdown);

		s2.selectByValue(excelReader.getCellData(xlSheetName, 15, 5));

		Select s3=new Select(Row1CompareWithDropdown);

		s3.selectByValue(excelReader.getCellData(xlSheetName, 16, 5));

		///Row1CompareWithDropdown.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Row1VendorValueTxt));
		Row1VendorValueTxt.click();
		Row1VendorValueTxt.sendKeys(excelReader.getCellData(xlSheetName, 17, 5));

		Thread.sleep(2000);

		Row1VendorValueTxt.sendKeys(Keys.TAB);


		// ((JavascriptExecutor)getDriver()).executeScript("window.scrollTo(0, 0)","");
		getAction().moveToElement(definitionOkBtn).build().perform();
		 Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
		definitionOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthSaveBtn));
		tranAuthSaveBtn.click();

		System.out.println("***************************************checkEditingTransactionAuthorizationInPurchaseReturns**************************************");

		String expValidationMsg=excelReader.getCellData(xlSheetName, 12, 6);

		String actValidationMsg=checkValidationMessage(expValidationMsg);
		
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actValidationMsg);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthCancelBtn));
		tranAuthCancelBtn.click();

		if(actValidationMsg.equalsIgnoreCase(expValidationMsg))
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






	public boolean 	checkDeletingTransactionAuthorizationInPurchaseReturns() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getAction().moveToElement(settingsmenuBtn).build().perform();
		Thread.sleep(2000);
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();*/
		
		ClickUsingJs(settingsmenuBtn);

		/*int count1 = SettingsSubMenusList.size();

		for (int i = 0; i < count1; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();

			System.err.println(data);

			getAction().moveToElement(SettingsSubMenusList.get(i)).build().perform();
			Thread.sleep(2000);

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 18, 5))) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);*/
		
		Thread.sleep(4000);
		
		ClickUsingJs(Transactionautorization);
		
		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNameDropdown));
		documentNameDropdown.click();
		documentNameDropdown.sendKeys(Keys.SPACE);
		Thread.sleep(6000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesReturns1));
		purchasesReturns1.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transAuthNameTxt));
		transAuthNameTxt.click();
		transAuthNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 19, 5));
		Thread.sleep(3000);
		transAuthNameTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthDeleteBtn));
		tranAuthDeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		System.out.println("***************************************checkEditingTransactionAuthorizationInPurchaseReturns**************************************");

		String expValidationMsg=excelReader.getCellData(xlSheetName, 18, 6);

		String actValidationMsg=checkValidationMessage(expValidationMsg);
		excelReader.setCellData(xlfile, xlSheetName, 18, 7, actValidationMsg);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthCancelBtn));
		tranAuthCancelBtn.click();

		if(actValidationMsg.equalsIgnoreCase(expValidationMsg))
		{
			excelReader.setCellData(xlfile, xlSheetName, 18, 8, resPass);

			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 18, 8, resFail);

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

	       /* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
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


	public boolean checkLogoutTransactionAuthorizationPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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
















	public TransactionAuthorizationPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
