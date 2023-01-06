package com.focus.Pages;

import java.io.IOException;
import java.util.ArrayList;
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

public class MasterAuthorizationPage extends BaseEngine 
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "MasterAuthorizationPage";
	
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







	public boolean 	checkSavingMasterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();
		Thread.sleep(2000);
	
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
		Thread.sleep(2000);
		
		int count = SettingsSubMenusList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			ScrollIntoView(SettingsSubMenusList.get(i));
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 8, 5))) 
			{
				ScrollIntoView(SettingsSubMenusList.get(i));
				SettingsSubMenusList.get(i).click();
				break;
			}
		}
		
		Thread.sleep(2000);
		//ScrollIntoView(masterAuthorization);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDropdown));
		masterDropdown.click();
		masterDropdown.sendKeys(Keys.SPACE);
		Thread.sleep(2000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthWarehouse));
		masterAuthWarehouse.click();

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

		System.out.println(actTemplateList);
		System.out.println(expTemplateList);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(custTemplateCloseBtn));
		custTemplateCloseBtn.click();*/

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
		definitionOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthSaveBtn));
		tranAuthSaveBtn.click();

		System.out.println("*****************************************checkSavingMasterAuthorization******************************************");

		String expValidationMsg=excelReader.getCellData(xlSheetName, 8, 6);
		String actValidationMsg=checkValidationMessage(expValidationMsg);
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actValidationMsg);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthCancelBtn));
		tranAuthCancelBtn.click();

		if(/*actCount==expCount && actTemplateList.equalsIgnoreCase(expTemplateList) && */actValidationMsg.equalsIgnoreCase(expValidationMsg))
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





	public static boolean checkCreatingWarehouseFromVoucherLevelAndValidatingWarehouseList() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		financialsTransactionsPurchaseMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
		purchaseVouchersVat.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		Thread.sleep(2000);

		getAction().moveToElement(headerWarehouse).doubleClick(headerWarehouse).build().perform();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		nameTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
		codeTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(account_saveBtn));
		account_saveBtn.click();
		Thread.sleep(2000);

		String actErrorMessageIsEmpty = Boolean.toString(validationConfirmationMessage.getText().isEmpty());
		String expErrorMessageIsEmpty = excelReader.getCellData(xlSheetName, 12, 6);
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actErrorMessageIsEmpty);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();
		enter_WarehouseTxt.sendKeys(Keys.SPACE);

		int warehousecount=warehouseBodyComboList.size();

		ArrayList WarehouseList = new ArrayList<String>();

		for(int i=0; i < warehousecount; i++)
		{
			String data=warehouseBodyComboList.get(i).getText();

			WarehouseList.add(data);
		}

		enter_WarehouseTxt.sendKeys(Keys.TAB);

		System.err.println(WarehouseList);

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();

		getWaitForAlert();
		getAlert().accept();

		Thread.sleep(2000);

		if(actErrorMessageIsEmpty.equalsIgnoreCase(expErrorMessageIsEmpty) && WarehouseList.contains(excelReader.getCellData(xlSheetName, 14, 5))==false)
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







	public static boolean checkShowAllUnAuthorizeOptionInAccountsMasterPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseMenu));
		warehouseMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterShowAllUnauthorisedRecords)); 
		masterShowAllUnauthorisedRecords.click();	

		Thread.sleep(2000);
		System.out.println(clickOnEmptyGroup.getText());

		String actAccountGroupEmptyRecords = clickOnEmptyGroup.getText().toString();

		String expAccountGroupEmptyRecords = excelReader.getCellData(xlSheetName, 15, 6);
		
		//excelReader.setCellData(xlfile, xlSheetName, 15, 7, actAccountGroupEmptyRecords);

		System.out.println("Empty Records Display Value Actual  : " + actAccountGroupEmptyRecords + "  Value Expected : " + expAccountGroupEmptyRecords);

		if (actAccountGroupEmptyRecords.equalsIgnoreCase(expAccountGroupEmptyRecords)) 
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

@FindBy(xpath="//*[@id='div_Authorize']/div/div[3]/div/input[2]")
private static WebElement AuthrzClsBtn;




	public  boolean checkValidatingAndAuthorizingWarehouseForAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseMenu));
		warehouseMenu.click();

		Thread.sleep(2000);

		ArrayList AccountsList = new ArrayList<String>();

		int count = masterWarehouseList.size();

		for (int i = 0; i < count; i++) 
		{
			String data = masterWarehouseList.get(i).getText();

			AccountsList.add(data);	
		}

		String actAccountsList = AccountsList.toString();
		String expAccountsList = excelReader.getCellData(xlSheetName, 19, 6);
	//	excelReader.setCellData(xlfile, xlSheetName, 19, 7, actAccountsList);

		System.out.println(actAccountsList);
		System.out.println(expAccountsList);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(showAllUnAuthBtn));
		showAllUnAuthBtn.click();

		Thread.sleep(2000);

		ArrayList AccountsUnAuthList = new ArrayList<String>();

		int count1 = masterWarehouseList.size();

		for (int i = 0; i < count1; i++) 
		{
			String data = masterWarehouseList.get(i).getText();

			AccountsUnAuthList.add(data);	
		}

		String actAccountsUnAuthList = AccountsUnAuthList.toString();
		String expAccountsUnAuthList = excelReader.getCellData(xlSheetName, 20, 6);
		//excelReader.setCellData(xlfile, xlSheetName, 20, 7, actAccountsUnAuthList);

		System.out.println(actAccountsUnAuthList);
		System.out.println(expAccountsUnAuthList);

		for (int i = 0; i < count1; i++) 
		{
			String data = masterWarehouseList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 19, 5))) 
			{
				masterAccountsChkboxList.get(i).click();
				System.err.println("**Warehouse for Authorization is Selected**");
				break;
			}	
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthorizeBtn));
		masterAuthorizeBtn.click();

	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reasonForAuthorizationTxt));
		reasonForAuthorizationTxt.sendKeys(excelReader.getCellData(xlSheetName, 20, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reasonPopupSaveBtn));
		reasonPopupSaveBtn.click();*/

		String expMessage = excelReader.getCellData(xlSheetName, 21, 6);
		String actMessage = checkValidationMessage(expMessage);
		//excelReader.setCellData(xlfile, xlSheetName, 21, 7, actMessage);
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AuthrzClsBtn));
		AuthrzClsBtn.click();*/
		Thread.sleep(2000);

		if(actAccountsList.equalsIgnoreCase(expAccountsList) && actAccountsUnAuthList.equalsIgnoreCase(expAccountsUnAuthList) && actMessage.equalsIgnoreCase(expMessage))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 19, 8, resPass);
			return true;
		}	 
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 19, 8, resFail);
			return false;
		}
	}




	public static boolean checkVerifyingWarehouseInPurchaseVoucherAfterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		financialsTransactionsPurchaseMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
		purchaseVouchersVat.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();
		enter_WarehouseTxt.sendKeys(Keys.SPACE);

		int warehousecount=warehouseBodyComboList.size();

		ArrayList WarehouseList = new ArrayList<String>();

		for(int i=0; i < warehousecount; i++)
		{
			String data=warehouseBodyComboList.get(i).getText();

			WarehouseList.add(data);
		}

		enter_WarehouseTxt.sendKeys(Keys.TAB);

		System.err.println(WarehouseList);

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();

		getWaitForAlert();
		getAlert().accept();

		Thread.sleep(2000);

		if(WarehouseList.contains(excelReader.getCellData(xlSheetName, 25, 5)))
		{
			excelReader.setCellData(xlfile, xlSheetName, 25, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 25, 8, resFail);
			return false;
		}
	}




	public boolean 	checkEditingMasterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();
		Thread.sleep(2000);
		
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,350)", "");
		Thread.sleep(2000);

		int count = SettingsSubMenusList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			ScrollIntoView(SettingsSubMenusList.get(i));
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 26, 5))) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}
		
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDropdown));
		masterDropdown.click();
		masterDropdown.sendKeys(Keys.SPACE);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthWarehouse1));
		masterAuthWarehouse1.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transAuthNameTxt));
		transAuthNameTxt.click();
		transAuthNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 27, 5));
		Thread.sleep(3000);
		transAuthNameTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1Icon));
		getAction().doubleClick(Level1Icon).build().perform();

		Thread.sleep(2000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
		conditionNotRequiredChkBox.click();

		Thread.sleep(2000);
		
		
		
		

		Select s1=new Select(conjunctionDropdown);

		s1.selectByValue(excelReader.getCellData(xlSheetName, 28, 5));
		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
		selectFieldTxt.click();

		Thread.sleep(2000);
		
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
		Thread.sleep(2000);
		
		getAction().moveToElement(sName).build().perform();
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sName));
		sName.click();

		Thread.sleep(2000);

		Select s2=new Select(selectOperatorDropdown);

		s2.selectByValue(excelReader.getCellData(xlSheetName, 29, 5));

		Select s3=new Select(compareWithDropdown);

		s3.selectByValue(excelReader.getCellData(xlSheetName, 30, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueTxt));
		valueTxt.click();
		valueTxt.sendKeys(excelReader.getCellData(xlSheetName, 31, 5));
		valueTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
		definitionOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthSaveBtn));
		tranAuthSaveBtn.click();

		System.out.println("***************************************checkEditingMasterAuthorization**************************************");

		String expValidationMsg=excelReader.getCellData(xlSheetName, 26, 6);

		String actValidationMsg=checkValidationMessage(expValidationMsg);
		//excelReader.setCellData(xlfile, xlSheetName, 26, 7, actValidationMsg);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthCancelBtn));
		tranAuthCancelBtn.click();

		if(actValidationMsg.equalsIgnoreCase(expValidationMsg))
		{
			
			excelReader.setCellData(xlfile, xlSheetName, 26, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 26, 8, resFail);
			return false;
		}
	}







	public  boolean checkCreatingDemoWarehouseAndTrialWarehouse() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseMenu));
		warehouseMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		masterNewBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.clear();
		nameTxt.sendKeys(excelReader.getCellData(xlSheetName, 32, 5));
		nameTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.clear();
		codeTxt.sendKeys(excelReader.getCellData(xlSheetName, 33, 5));
		codeTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();

		String expMessage=excelReader.getCellData(xlSheetName, 32, 6);

		String actMessage=checkValidationMessage(expMessage);
		//excelReader.setCellData(xlfile, xlSheetName, 32, 7, actMessage);

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.clear();
		nameTxt.sendKeys(excelReader.getCellData(xlSheetName, 34, 5));
		nameTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.clear();
		codeTxt.sendKeys(excelReader.getCellData(xlSheetName, 35, 5));
		codeTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		String expMessage1=excelReader.getCellData(xlSheetName, 33, 6);
		String actMessage1=checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actMessage1);

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
		masterCloseBtn.click();

		Thread.sleep(2000);

		ArrayList AccountsList = new ArrayList<String>();

		int count = masterWarehouseList.size();

		for (int i = 0; i < count; i++) 
		{
			String data = masterWarehouseList.get(i).getText();

			AccountsList.add(data);	
		}

		System.err.println(AccountsList);

		if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1)
				&& AccountsList.contains(excelReader.getCellData(xlSheetName, 36, 5)) && AccountsList.contains(excelReader.getCellData(xlSheetName, 37, 5)))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 32, 8, resPass);
			return true;
		}	 
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 32, 8, resFail);
			return false;
		}
	}






	public static boolean checkVerifyingWarehouseCreated() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		financialsTransactionsPurchaseMenu.click();

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
		purchaseVouchersVat.click();

		Thread.sleep(2000);

		/*checkDeleteLinkStatus();

			Thread.sleep(2000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();
		enter_WarehouseTxt.sendKeys(Keys.SPACE);

		int warehousecount=warehouseBodyComboList.size();

		ArrayList WarehouseList = new ArrayList<String>();

		for(int i=0; i < warehousecount; i++)
		{
			String data=warehouseBodyComboList.get(i).getText();

			WarehouseList.add(data);
		}

		enter_WarehouseTxt.sendKeys(Keys.TAB);

		System.err.println(WarehouseList);

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();

		getWaitForAlert();
		getAlert().accept();

		Thread.sleep(2000);
		
		String actWarehouse1 = excelReader.getCellData(xlSheetName, 38, 6);
		String expWarehouse1 = Boolean.toString(WarehouseList.contains(excelReader.getCellData(xlSheetName, 38, 5))); 
		//excelReader.setCellData(xlfile, xlSheetName, 38, 7, actWarehouse1);
		
		String actWarehouse2 = excelReader.getCellData(xlSheetName, 39, 6);
		String expWarehouse2 = Boolean.toString(WarehouseList.contains(excelReader.getCellData(xlSheetName, 39, 5)));
		//excelReader.setCellData(xlfile, xlSheetName, 39, 7, actWarehouse2);
		
		
		
		if(actWarehouse1.equalsIgnoreCase(expWarehouse1) && actWarehouse2.equalsIgnoreCase(expWarehouse2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 38, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 38, 8, resFail);
			return false;
		}
	}


	@FindBy(xpath="//*[@id='ddlCompany']")
	private static WebElement companyList;





	public static boolean checkLogoutAndLoginWithUserAllOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		//getAction().moveToElement(userNameDisplay).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);
		
		//getAction().moveToElement(userNameDisplayLogo).build().perform();

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(3000);

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt=excelReader.getCellData(xlSheetName, 16, 5);

		String pawslt=excelReader.getCellData(xlSheetName, 17, 5);
		
		String compname="Automation company";

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);

		lp.clickOnSignInBtn();
		Thread.sleep(2000);
		lp.reLogin(unamelt, pawslt, compname);

		//checkRefershPopOnlogin();

		//checkPopUpWindow();

		Thread.sleep(6000);
				
		
		
		String actUserInfo=userNameUserDisplay.getText();

		System.out.println("User Info  : "+actUserInfo);

		System.out.println("User Info Capture Text  :  "+userNameUserDisplay.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String getLoginCompanyName=getCompanyTxt.substring(0, 19);
		System.out.println("company name  :  "+ getLoginCompanyName);
		companyLogo.click();


		String expUserInfo           =excelReader.getCellData(xlSheetName, 16, 6);
		excelReader.setCellData(xlfile, xlSheetName, 16, 7, actUserInfo);
		
		String expLoginCompanyName   =excelReader.getCellData(xlSheetName, 17, 6);
		excelReader.setCellData(xlfile, xlSheetName, 17, 7, getLoginCompanyName);

		String actDashboard = Boolean.toString(dashboard.isDisplayed());
		String expDashboard = excelReader.getCellData(xlSheetName, 18, 6);
		excelReader.setCellData(xlfile, xlSheetName, 18, 7, actDashboard);

		System.out.println("**********************************************checkLogoutAndLoginWithUserAllOptions*********************************************");

		System.out.println("User Info               : "+actUserInfo         +" Value Expected "+expUserInfo);
		System.out.println("Login Company           : "+getLoginCompanyName +" Value Expected "+expLoginCompanyName);
		System.out.println("Dashboard is Displaying : "+actDashboard        +" Value Expected "+expDashboard );

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName) && actDashboard.equalsIgnoreCase(expDashboard)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 16, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 16, 8, resFail);
			return false;
		}
	}












	public static boolean checkLogoutAndLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		//getAction().moveToElement(userNameDisplay).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(3000);

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt=excelReader.getCellData(xlSheetName, 22, 5);

		String pawslt=excelReader.getCellData(xlSheetName, 23, 5);
		
		String compname="Automation Company";

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);

		lp.clickOnSignInBtn();
		Thread.sleep(2000);

		//checkRefershPopOnlogin();

		//checkPopUpWindow();
		lp.reLogin(unamelt, pawslt, compname);

		Thread.sleep(6000);
		/*String unamelt1=excelReader.getCellData(xlSheetName, 16, 5);

		String pawslt1=excelReader.getCellData(xlSheetName, 17, 5);

		lp.enterUserName(unamelt1);

		lp.enterPassword(pawslt1);
		 Select s1=new Select(companyList);
	        s1.selectByValue("108");
	        Thread.sleep(2000);

		lp.clickOnSignInBtn();*/
		

		String actUserInfo=userNameDisplay.getText();

		System.out.println("User Info  : "+actUserInfo);

		System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String getLoginCompanyName=getCompanyTxt.substring(0, 19);
		System.out.println("company name  :  "+ getLoginCompanyName);
		companyLogo.click();


		String expUserInfo           =excelReader.getCellData(xlSheetName, 22, 6);
		//excelReader.setCellData(xlfile, xlSheetName, 22, 7, actUserInfo);
		
		String expLoginCompanyName   =excelReader.getCellData(xlSheetName, 23, 6);
		//excelReader.setCellData(xlfile, xlSheetName, 23, 7, getLoginCompanyName);

		String actDashboard = Boolean.toString(dashboard.isDisplayed());
		String expDashboard = excelReader.getCellData(xlSheetName, 24, 6);
		//excelReader.setCellData(xlfile, xlSheetName, 24, 7, expDashboard);

		System.out.println("**********************************************checkLogoutAndLoginWithSU*********************************************");

		System.out.println("User Info               : "+actUserInfo         +" Value Expected "+expUserInfo);
		System.out.println("Login Company           : "+getLoginCompanyName +" Value Expected "+expLoginCompanyName);
		System.out.println("Dashboard is Displaying : "+actDashboard        +" Value Expected "+expDashboard );

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName) && actDashboard.equalsIgnoreCase(expDashboard)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 22, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 22, 8, resFail);
			return false;
		}
	}



	
	
	
	public  boolean checkValidatingAndAuthorizingTrialWarehouse() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseMenu));
		warehouseMenu.click();

		Thread.sleep(2000);

		ArrayList AccountsList = new ArrayList<String>();

		int count = masterWarehouseList.size();

		for (int i = 0; i < count; i++) 
		{
			String data = masterWarehouseList.get(i).getText();

			AccountsList.add(data);	
		}

		String actAccountsList = AccountsList.toString();
		String expAccountsList = excelReader.getCellData(xlSheetName, 40, 6);
	//	excelReader.setCellData(xlfile, xlSheetName, 40, 7, actAccountsList);

		System.out.println(actAccountsList);
		System.out.println(expAccountsList);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(showAllUnAuthBtn));
		showAllUnAuthBtn.click();

		Thread.sleep(2000);

		ArrayList AccountsUnAuthList = new ArrayList<String>();

		int count1 = masterWarehouseList.size();

		for (int i = 0; i < count1; i++) 
		{
			String data = masterWarehouseList.get(i).getText();

			AccountsUnAuthList.add(data);	
		}

		String actAccountsUnAuthList = AccountsUnAuthList.toString();
		String expAccountsUnAuthList = excelReader.getCellData(xlSheetName, 41, 6);
		//excelReader.setCellData(xlfile, xlSheetName, 41, 7, actAccountsUnAuthList);

		System.out.println(actAccountsUnAuthList);
		System.out.println(expAccountsUnAuthList);

		for (int i = 0; i < count1; i++) 
		{
			String data = masterWarehouseList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 40, 5))) 
			{
				masterAccountsChkboxList.get(i).click();
				break;
			}	
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthorizeBtn));
		masterAuthorizeBtn.click();

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reasonForAuthorizationTxt));
		reasonForAuthorizationTxt.sendKeys(excelReader.getCellData(xlSheetName, 41, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reasonPopupSaveBtn));
		reasonPopupSaveBtn.click();
*/

		String expMessage = excelReader.getCellData(xlSheetName, 42, 6);
		String actMessage = checkValidationMessage(expMessage);
		//excelReader.setCellData(xlfile, xlSheetName, 42, 7, actMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AuthrzClsBtn));
		AuthrzClsBtn.click();

		if(actAccountsList.equalsIgnoreCase(expAccountsList) && actAccountsUnAuthList.equalsIgnoreCase(expAccountsUnAuthList) && actMessage.equalsIgnoreCase(expMessage))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 40, 8, resPass);

			return true;
		}	 
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 40, 8, resFail);

			return false;
		}
	}

	
	


	public static boolean checkVerifyingWarehouseCreatedAfterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		financialsTransactionsPurchaseMenu.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
		purchaseVouchersVat.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();
		enter_WarehouseTxt.sendKeys(Keys.SPACE);

		int warehousecount=warehouseBodyComboList.size();

		ArrayList WarehouseList = new ArrayList<String>();

		for(int i=0; i < warehousecount; i++)
		{
			String data=warehouseBodyComboList.get(i).getText();

			WarehouseList.add(data);
		}

		enter_WarehouseTxt.sendKeys(Keys.TAB);

		System.err.println(WarehouseList);

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();

		getWaitForAlert();
		getAlert().accept();

		Thread.sleep(2000);
		
		String actWarehouse1 = excelReader.getCellData(xlSheetName, 43, 6);
		String expWarehouse1 = Boolean.toString(WarehouseList.contains(excelReader.getCellData(xlSheetName, 43, 5))); 
		//excelReader.setCellData(xlfile, xlSheetName, 43, 7, actWarehouse1);
		
		String actWarehouse2 = excelReader.getCellData(xlSheetName, 44, 6);
		String expWarehouse2 = Boolean.toString(WarehouseList.contains(excelReader.getCellData(xlSheetName, 44, 5)));
		//excelReader.setCellData(xlfile, xlSheetName, 44, 7, actWarehouse2);
		
		
		
		if(actWarehouse1.equalsIgnoreCase(expWarehouse1) && actWarehouse2.equalsIgnoreCase(expWarehouse2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 38, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 38, 8, resFail);
			return false;
		}
		
	}






	public static boolean checkDeleteInWarehouseInWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseMenu));
		warehouseMenu.click();

		Thread.sleep(2000);

		ArrayList AccountsList = new ArrayList<String>();

		int count = masterWarehouseList.size();

		for (int i = 0; i < count; i++) 
		{
			String data = masterWarehouseList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 45, 5)) || data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 46, 5)) || data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 47, 5))) 
			{
				masterAccountsChkboxList.get(i).click();
				Thread.sleep(1000);
			}
		}


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDeleteBtn));
		masterDeleteBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getMsgOnDelete));
		System.out.println(getMsgOnDelete.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelete));
		clickOnOkInDelete.click();

		String expMessage=excelReader.getCellData(xlSheetName, 45, 6);

		String actMessage=checkValidationMessage(expMessage);
		//excelReader.setCellData(xlfile, xlSheetName, 45, 7, actMessage);


		System.out.println("********************************checkDeleteInWarehouseInWarehouseMaster***************************");

		System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 45, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 45, 8, resFail);
			return false;
		}	 
	}







	public boolean 	checkDeletingMasterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();
		
		Thread.sleep(2000);

		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
		Thread.sleep(2000);
		
		int count = SettingsSubMenusList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			ScrollIntoView(SettingsSubMenusList.get(i));
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 48, 5))) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDropdown));
		masterDropdown.click();
		masterDropdown.sendKeys(Keys.SPACE);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthWarehouse1));
		masterAuthWarehouse1.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transAuthNameTxt));
		transAuthNameTxt.click();
		transAuthNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 49, 5));
		Thread.sleep(3000);
		transAuthNameTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthDeleteBtn));
		tranAuthDeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		System.out.println("***************************************checkDeletingMasterAuthorization**************************************");

		String expValidationMsg=excelReader.getCellData(xlSheetName, 48, 6);

		String actValidationMsg=checkValidationMessage(expValidationMsg);
		//excelReader.setCellData(xlfile, xlSheetName, 48, 7, actValidationMsg);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthCancelBtn));
		tranAuthCancelBtn.click();

		if(actValidationMsg.equalsIgnoreCase(expValidationMsg))
		{
			excelReader.setCellData(xlfile, xlSheetName, 48, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 48, 8, resFail);
			return false;
		}
	}


















	public MasterAuthorizationPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
