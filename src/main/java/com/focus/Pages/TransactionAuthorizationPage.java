package com.focus.Pages;

import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
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

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();

		int count = SettingsSubMenusList.size();

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
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNameDropdown));
		documentNameDropdown.click();
		documentNameDropdown.sendKeys(Keys.SPACE );
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesReturns));
		purchasesReturns.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transAuthNameTxt));
		transAuthNameTxt.click();
		transAuthNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		transAuthNameTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addLevelBtn));
		addLevelBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
		descriptionTxt.click();
		descriptionTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", conditionNotRequiredChkBox);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
		conditionNotRequiredChkBox.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
		userSelectionTab.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
		userrow1.click();
		Select user = new Select(userDropdown);
		user.selectByVisibleText(excelReader.getCellData(xlSheetName, 11, 5));

		Thread.sleep(2000);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertsTab));
		alertsTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailHeaderTemplateLink));
		emailHeaderTemplateLink.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(templateTxt));
		templateTxt.click();
		templateTxt.sendKeys(Keys.SPACE);
		Thread.sleep(2000);

		ArrayList<String>  TemplatesList= new ArrayList<String>(); 

		int actCount = custTemplatesList.size();
		int expCount = 3; 

		System.err.println("Templates List : "+actCount+"  Value Expected  "+expCount);

		for(int i=0; i < actCount; i++)
		{
			String data = custTemplatesList.get(i).getText();
			TemplatesList.add(data);
		}

		templateTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		String actTemplateList = TemplatesList.toString();
		String expTemplateList = "[createBudgetEmailTemplate, createEmailTemplate, createEmailTemplate]";

		System.out.println("Actual   : "+actTemplateList);
		System.out.println("Expected : "+expTemplateList);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(custTemplateCloseBtn));
		custTemplateCloseBtn.click();

		Thread.sleep(2000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
		definitionOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthSaveBtn));
		tranAuthSaveBtn.click();

		System.out.println("*****************************************checkSavingTransactionAuthorizationInPurchaseReturns******************************************");

		String expValidationMsg=excelReader.getCellData(xlSheetName, 8, 6);
		String actValidationMsg=checkValidationMessage(expValidationMsg);
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actValidationMsg);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthCancelBtn));
		tranAuthCancelBtn.click();

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
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();

		int count1 = SettingsSubMenusList.size();

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

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNameDropdown));
		documentNameDropdown.click();
		documentNameDropdown.sendKeys(Keys.SPACE);
		Thread.sleep(2000);

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
		
		Thread.sleep(2000);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1editBtn));
		getAction().moveToElement(Level1editBtn).click().build().perform();
		
		Thread.sleep(1000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
		conditionNotRequiredChkBox.click();

		Thread.sleep(2000);

		Select s1=new Select(Row1ConjunctionDropdown);

		s1.selectByValue(excelReader.getCellData(xlSheetName, 14, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Row1SelectFieldTxt));
		Row1SelectFieldTxt.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldVendorAC));
		selectFieldVendorAC.click();

		Thread.sleep(1000);

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
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();

		int count1 = SettingsSubMenusList.size();

		for (int i = 0; i < count1; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();

			System.err.println(data);

			ScrollIntoView(SettingsSubMenusList.get(i));

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 18, 5))) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNameDropdown));
		documentNameDropdown.click();
		documentNameDropdown.sendKeys(Keys.SPACE);
		Thread.sleep(2000);

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



















	public TransactionAuthorizationPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
