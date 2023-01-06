package com.focus.Pages;

import java.awt.AWTException;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashSet;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
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

public class OpeningStocksPage extends BaseEngine 
{
	private static String xlfile = getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlSheetName = "OpeningStocksPage";
	
	public static String checkValidationMessage(String ExpMessage) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(2000);
		
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

	@FindBy (xpath="//div[@id='ReIndexingError_Modal']/div/div")
	private static WebElement reindexingPopup;

	@FindBy (xpath="(//div[@id='ReIndexingError_Modal']//div/button)[1]")
	private static WebElement reindexingPopupCloseBtn;

	@FindBy (xpath="//*[@id='ReIndexingError_Modal']//div[3]/button")
	private static WebElement reindexingPopupCancelBtn;






	/*public boolean checkRestoreAutomationCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt=excelReader.getCellData(xlSheetName, 8, 5);

		String pawslt=excelReader.getCellData(xlSheetName, 9, 5);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataMangementMenu));
		dataMangementMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restore));
		restore.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(folderpathExpandBtn));
		folderpathExpandBtn.click();

		Thread.sleep(3000);

		Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\RestoreAutomationCompanyMounted.exe");

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreCompanyBtn));
		restoreCompanyBtn.click();

		try
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(overRideYesBtn));
			overRideYesBtn.click();
			Thread.sleep(40000);
		}
		catch(Exception e)
		{
			System.err.println("NO OLDER COMPANY EXISTS");
		}

		Thread.sleep(30000);

		
			getWaitForAlert();

			Thread.sleep(3000);
			getAlert().accept();
			Thread.sleep(3000);
	
		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(3000);

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);

		Thread.sleep(2000);

		lp.clickOnSignInBtn();

		Thread.sleep(10000);
		
		
		if (reindexingPopup.isDisplayed()==false) 
		{
			System.out.println("Reindexing Log is Not Displaying");
		} 
		else 
		{
			System.out.println("Reindexing Log is Displaying");
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reindexingPopupCancelBtn));
			reindexingPopupCancelBtn.click();
			Thread.sleep(6000);
			lp.clickOnSignInBtn();
		}
		


		String actUserInfo1=userNameDisplay.getText();

		System.out.println("User Info  : "+actUserInfo1);

		System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt1=companyName.getText();
		String getLoginCompanyName1=getCompanyTxt1.substring(0, 19);
		System.out.println("company name  :  "+ getLoginCompanyName1);
		companyLogo.click();

		String expUserInfo1           = excelReader.getCellData(xlSheetName, 8, 6);
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actUserInfo1);
		
		String expLoginCompanyName1   = excelReader.getCellData(xlSheetName, 9, 6);
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, getLoginCompanyName1);
		

		System.out.println("UserInfo1             : "+actUserInfo1            +" Value Expected : "+expUserInfo1);
		System.out.println("LoginCompanyName1     : "+getLoginCompanyName1    +" Value Expected : "+expLoginCompanyName1);

		if(actUserInfo1.equalsIgnoreCase(expUserInfo1) && getLoginCompanyName1.contains(expLoginCompanyName1))
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




*/



	public boolean checkSavingOpeningStockVoucher1WithBatchItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksVoucher));
		openingStocksVoucher.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 14, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 15, 5));
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		Date date=new Date();

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		String presentDate = currentDate.format(cal.getTime());
		System.out.println("--------------Date Before Adding:: " + presentDate);

		cal.add(Calendar.DAY_OF_WEEK, 5); 

		String nextFiveDate=df.format(cal.getTime());

		System.out.println("--------------Date After Adding:: " + nextFiveDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.sendKeys(Keys.HOME,nextFiveDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		Thread.sleep(2000);

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 1");
		expMessage.add("Saving in background.");*/

		//String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		/*expMsg.add(excelReader.getCellData(xlSheetName, 11, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 12, 6));
		
		String expMessage = expMsg.toString();
		
		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 11, 7, actMessage);
		*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 11, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 12, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 11, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
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








	public boolean checkEditingAndVerifingDataOfSavedOpeningStocksVoucher1AddingRow2WithBatchItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 16, 6);
		excelReader.setCellData(xlfile, xlSheetName, 16, 7, actDocNo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		String actWarehouse=warehouseTxt.getAttribute("value");
		String expWarehouse=excelReader.getCellData(xlSheetName, 17, 6);
		excelReader.setCellData(xlfile, xlSheetName, 17, 7, actWarehouse);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actItem=select1stRow_1stColumn.getText();
		String expItem=excelReader.getCellData(xlSheetName, 18, 6);
		excelReader.setCellData(xlfile, xlSheetName, 18, 7, actItem);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
		String actUnits=select1stRow_2ndColumn.getText();
		String expUnits=excelReader.getCellData(xlSheetName, 19, 6);
		excelReader.setCellData(xlfile, xlSheetName, 19, 7, actUnits);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		String actQuantity=select1stRow_3rdColumn.getText();
		String expQuantity=excelReader.getCellData(xlSheetName, 20, 6);
		excelReader.setCellData(xlfile, xlSheetName, 20, 7, actQuantity);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		String actRate=select1stRow_4thColumn.getText();
		String expRate=excelReader.getCellData(xlSheetName, 21, 6);
		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actRate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		String actGross=select1stRow_5thColumn.getText();
		String expGross=excelReader.getCellData(xlSheetName, 22, 6);
		excelReader.setCellData(xlfile, xlSheetName, 22, 7, actGross);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
		String actBatch=select1stRow_6thColumn.getText();
		String expBatch=excelReader.getCellData(xlSheetName, 23, 6);
		excelReader.setCellData(xlfile, xlSheetName, 23, 7, actBatch);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		String actExpiryDate=select1stRow_8thColumn.getText();

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		Calendar cal=Calendar.getInstance();
		cal.add(Calendar.DAY_OF_WEEK, 5); 
		String expExpiryDate=df.format(cal.getTime());

		System.out.println("**********************checkEditingAndVerifingDataOfSavedOpeningStocksVoucher1AddingRow2WithBatchItem*********************");

		System.out.println("*********** Voucher No : "+actDocNo      +"  value expected  "+expDocNo);
		System.out.println("*********** Warehouse  : "+actWarehouse  +"  value expected  "+expWarehouse);
		System.out.println("*********** Item       : "+actItem       +"  value expected  "+expItem);
		System.out.println("*********** Units      : "+actUnits      +"  value expected  "+expUnits);
		System.out.println("*********** Quantity   : "+actQuantity   +"  value expected  "+expQuantity);
		System.out.println("*********** Rate       : "+actRate       +"  value expected  "+expRate);
		System.out.println("*********** Gross      : "+actGross      +"  value expected  "+expGross);
		System.out.println("*********** Batch      : "+actBatch      +"  value expected  "+expBatch);
		System.out.println("*********** ExpiryDate : "+actExpiryDate +"  value expected  "+expExpiryDate);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 17, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 18, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 19, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 20, 5));
		enter_Batch.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		Calendar cal2=Calendar.getInstance();
		cal2.add(Calendar.DAY_OF_WEEK, 7); 

		String nextSevenDate=df.format(cal2.getTime());

		System.out.println("--------------Date After Adding:: " + nextSevenDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.sendKeys(Keys.HOME,nextSevenDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 1");
		expMessage.add("Saving in background.");*/
		
	/*	String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 24, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 25, 6));
		
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actMessage);
		
		String expMessage = expMsg.toString();
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 24, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 25, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		if (actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) && actItem.equalsIgnoreCase(expItem) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate)
				&& actGross.equalsIgnoreCase(expGross) && actBatch.equalsIgnoreCase(expBatch) && actExpiryDate.equalsIgnoreCase(expExpiryDate)
				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
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








	public boolean checkEditingAndVerifingDataOfSavedOpeningStocksVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 26, 6);
		excelReader.setCellData(xlfile, xlSheetName, 26, 7, actDocNo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		String actWarehouse=warehouseTxt.getAttribute("value");
		String expWarehouse=excelReader.getCellData(xlSheetName, 27, 6);
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actWarehouse);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actItem=select1stRow_1stColumn.getText();
		String expItem=excelReader.getCellData(xlSheetName, 28, 6);
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actItem);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
		String actUnits=select1stRow_2ndColumn.getText();
		String expUnits=excelReader.getCellData(xlSheetName, 29, 6);
		excelReader.setCellData(xlfile, xlSheetName, 29, 7, actUnits);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		String actQuantity=select1stRow_3rdColumn.getText();
		String expQuantity=excelReader.getCellData(xlSheetName, 30, 6);
		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actQuantity);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		String actRate=select1stRow_4thColumn.getText();
		String expRate=excelReader.getCellData(xlSheetName, 31, 6);
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actRate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		String actGross=select1stRow_5thColumn.getText();
		String expGross=excelReader.getCellData(xlSheetName, 32, 6);
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actGross);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
		String actBatch=select1stRow_6thColumn.getText();
		String expBatch=excelReader.getCellData(xlSheetName, 33, 6);
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actBatch);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		String actExpiryDate=select1stRow_8thColumn.getText();

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		Calendar cal=Calendar.getInstance();
		cal.add(Calendar.DAY_OF_WEEK, 5); 
		String expExpiryDate=df.format(cal.getTime());


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actItem2=select2ndRow_1stColumn.getText();
		String expItem2=excelReader.getCellData(xlSheetName, 34, 6);
		excelReader.setCellData(xlfile, xlSheetName, 34, 7, actItem2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_2ndColumn));
		String actUnits2=select2ndRow_2ndColumn.getText();
		String expUnits2=excelReader.getCellData(xlSheetName, 35, 6);
		excelReader.setCellData(xlfile, xlSheetName, 35, 7, actUnits2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
		String actQuantity2=select2ndRow_3rdColumn.getText();
		String expQuantity2=excelReader.getCellData(xlSheetName, 36, 6);
		excelReader.setCellData(xlfile, xlSheetName, 36, 7, actQuantity2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
		String actRate2=select2ndRow_4thColumn.getText();
		String expRate2=excelReader.getCellData(xlSheetName, 37, 6);
		excelReader.setCellData(xlfile, xlSheetName, 37, 7, actRate2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
		String actGross2=select2ndRow_5thColumn.getText();
		String expGross2=excelReader.getCellData(xlSheetName, 38, 6);
		excelReader.setCellData(xlfile, xlSheetName, 38, 7, actGross2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_6thColumn));
		String actBatch2=select2ndRow_6thColumn.getText();
		String expBatch2=excelReader.getCellData(xlSheetName, 39, 6);
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actBatch2);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_8thColumn));
		String actExpiryDate2=select2ndRow_8thColumn.getText();

		Calendar cal2=Calendar.getInstance();
		cal2.add(Calendar.DAY_OF_WEEK, 7); 

		String expExpiryDate2=df.format(cal2.getTime());


		System.out.println("**********************checkEditingAndVerifingDataOfSavedOpeningStocksVoucher1*********************");

		System.out.println("*********** Voucher No  : "+actDocNo      +"  value expected  "+expDocNo);
		System.out.println("*********** Warehouse   : "+actWarehouse  +"  value expected  "+expWarehouse);
		System.out.println("*********** Item        : "+actItem       +"  value expected  "+expItem);
		System.out.println("*********** Units       : "+actUnits      +"  value expected  "+expUnits);
		System.out.println("*********** Quantity    : "+actQuantity   +"  value expected  "+expQuantity);
		System.out.println("*********** Rate        : "+actRate       +"  value expected  "+expRate);
		System.out.println("*********** Gross       : "+actGross      +"  value expected  "+expGross);
		System.out.println("*********** Batch       : "+actBatch      +"  value expected  "+expBatch);
		System.out.println("*********** ExpiryDate  : "+actExpiryDate +"  value expected  "+expExpiryDate);
		System.out.println("*********** Item2       : "+actItem2       +"  value expected  "+expItem2);
		System.out.println("*********** Units2      : "+actUnits2      +"  value expected  "+expUnits2);
		System.out.println("*********** Quantity2   : "+actQuantity2   +"  value expected  "+expQuantity2);
		System.out.println("*********** Rate2       : "+actRate2       +"  value expected  "+expRate2);
		System.out.println("*********** Gross2      : "+actGross2      +"  value expected  "+expGross2);
		System.out.println("*********** Batch2      : "+actBatch2      +"  value expected  "+expBatch2);
		System.out.println("*********** ExpiryDate2 : "+actExpiryDate2 +"  value expected  "+expExpiryDate2);

		if (actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) && actItem.equalsIgnoreCase(expItem) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate)
				&& actGross.equalsIgnoreCase(expGross) && actBatch.equalsIgnoreCase(expBatch) && actExpiryDate.equalsIgnoreCase(expExpiryDate)
				&& actItem2.equalsIgnoreCase(expItem2) && actUnits2.equalsIgnoreCase(expUnits2) && actQuantity2.equalsIgnoreCase(expQuantity2) 
				&& actRate2.equalsIgnoreCase(expRate2) && actGross2.equalsIgnoreCase(expGross2) && actBatch2.equalsIgnoreCase(expBatch2) 
				&& actExpiryDate2.equalsIgnoreCase(expExpiryDate2)) 
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







	public boolean checkSavingOpeningStockVoucher2WithBatchItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
		new_newBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 40, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 41, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 42, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 43, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 44, 5));
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		Date date=new Date();

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		String presentDate = currentDate.format(cal.getTime());
		System.out.println("--------------Date Before Adding:: " + presentDate);

		cal.add(Calendar.DAY_OF_WEEK, 2); 

		String nextFiveDate=df.format(cal.getTime());

		System.out.println("--------------Date After Adding:: " + nextFiveDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.sendKeys(Keys.HOME,nextFiveDate);
		Thread.sleep(1000);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Saving in background.");
		expMessage.add("Voucher saved successfully : 2");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 40, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 41, 6));
		
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 40, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 41, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		//if(actMessage.equals(expMessage))
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








	public boolean checkVerifingDataOfSavedOpeningStocksVoucher2AndDeleting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 45, 6);
		excelReader.setCellData(xlfile, xlSheetName, 45, 7, actDocNo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		String actWarehouse=warehouseTxt.getAttribute("value");
		String expWarehouse="SECUNDERABAD";
		excelReader.getCellData(xlSheetName, 46, 6);
		excelReader.setCellData(xlfile, xlSheetName, 46, 7, actWarehouse);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actItem=select1stRow_1stColumn.getText();
		String expItem="BR COGS ITEM";
		excelReader.getCellData(xlSheetName, 47, 6);
		excelReader.setCellData(xlfile, xlSheetName, 47, 7, actItem);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
		String actUnits=select1stRow_2ndColumn.getText();
		String expUnits="Dozs";
		excelReader.getCellData(xlSheetName, 48, 6);
		excelReader.setCellData(xlfile, xlSheetName, 48, 7, actUnits);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		String actQuantity=select1stRow_3rdColumn.getText();
		String expQuantity="2.00";
		excelReader.getCellData(xlSheetName, 49, 6);
		excelReader.setCellData(xlfile, xlSheetName, 49, 7, actQuantity);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		String actRate=select1stRow_4thColumn.getText();
		String expRate="10.00";
		excelReader.getCellData(xlSheetName, 50, 6);
		excelReader.setCellData(xlfile, xlSheetName, 50, 7, actRate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		String actGross=select1stRow_5thColumn.getText();
		String expGross="20.00";
		excelReader.getCellData(xlSheetName, 51, 6);
		excelReader.setCellData(xlfile, xlSheetName, 51, 7, actGross);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
		String actBatch=select1stRow_6thColumn.getText();
		String expBatch="Batch#OSV2R1";
		excelReader.getCellData(xlSheetName, 52, 6);
		excelReader.setCellData(xlfile, xlSheetName, 52, 7, actBatch);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		String actExpiryDate=select1stRow_8thColumn.getText();

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		Calendar cal=Calendar.getInstance();
		cal.add(Calendar.DAY_OF_WEEK, 2); 
		String expExpiryDate=df.format(cal.getTime());

		System.out.println("**********************checkVerifingDataOfSavedOpeningStocksVoucher2AndDeleting*********************");

		System.out.println("*********** Voucher No : "+actDocNo      +"  value expected  "+expDocNo);
		System.out.println("*********** Warehouse  : "+actWarehouse  +"  value expected  "+expWarehouse);
		System.out.println("*********** Item       : "+actItem       +"  value expected  "+expItem);
		System.out.println("*********** Units      : "+actUnits      +"  value expected  "+expUnits);
		System.out.println("*********** Quantity   : "+actQuantity   +"  value expected  "+expQuantity);
		System.out.println("*********** Rate       : "+actRate       +"  value expected  "+expRate);
		System.out.println("*********** Gross      : "+actGross      +"  value expected  "+expGross);
		System.out.println("*********** Batch      : "+actBatch      +"  value expected  "+expBatch);
		System.out.println("*********** ExpiryDate : "+actExpiryDate +"  value expected  "+expExpiryDate);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		String expMessage=excelReader.getCellData(xlSheetName, 53, 6);
		String actMessage=checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 53, 7, actMessage);

		if (actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) && actItem.equalsIgnoreCase(expItem) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate)
				&& actGross.equalsIgnoreCase(expGross) && actBatch.equalsIgnoreCase(expBatch) && actExpiryDate.equalsIgnoreCase(expExpiryDate)
				&& actMessage.equalsIgnoreCase(expMessage)) 
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








	public boolean checkSavingOpeningStockVoucher2WithStockItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 54, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 55, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 56, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 57, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 2");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 54, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 55, 6));
		
		excelReader.setCellData(xlfile, xlSheetName, 54, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actMessage.equals(expMessage))*/
		
String expMessage1 = excelReader.getCellData(xlSheetName, 54, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 55, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 54, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		{
			excelReader.setCellData(xlfile, xlSheetName, 54, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 54, 8, resFail);
			return false;
		}
	}







	public boolean checkEditingAndVerifingDataOfSavedOpeningStocksVoucher2AddingRow2WithStockItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 58, 6);
		excelReader.setCellData(xlfile, xlSheetName, 58, 7, actDocNo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		String actWarehouse=warehouseTxt.getAttribute("value");
		String expWarehouse=excelReader.getCellData(xlSheetName, 59, 6);
		excelReader.setCellData(xlfile, xlSheetName, 59, 7, actWarehouse);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actItem=select1stRow_1stColumn.getText();
		String expItem=excelReader.getCellData(xlSheetName, 60, 6);
		excelReader.setCellData(xlfile, xlSheetName, 60, 7, actItem);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
		String actUnits=select1stRow_2ndColumn.getText();
		String expUnits=excelReader.getCellData(xlSheetName, 61, 6);
		excelReader.setCellData(xlfile, xlSheetName, 61, 7, actUnits);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		String actQuantity=select1stRow_3rdColumn.getText();
		String expQuantity=excelReader.getCellData(xlSheetName, 62, 6);
		excelReader.setCellData(xlfile, xlSheetName, 62, 7, actQuantity);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		String actRate=select1stRow_4thColumn.getText();
		String expRate=excelReader.getCellData(xlSheetName, 63, 6);
		excelReader.setCellData(xlfile, xlSheetName, 63, 7, actRate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		String actGross=select1stRow_5thColumn.getText();
		String expGross=excelReader.getCellData(xlSheetName, 64, 6);
		excelReader.setCellData(xlfile, xlSheetName, 64, 7, actGross);
		
		System.out.println("**********************checkEditingAndVerifingDataOfSavedOpeningStocksVoucher2AddingRow2WithStockItem*********************");

		System.out.println("*********** Voucher No : "+actDocNo      +"  value expected  "+expDocNo);
		System.out.println("*********** Warehouse  : "+actWarehouse  +"  value expected  "+expWarehouse);
		System.out.println("*********** Item       : "+actItem       +"  value expected  "+expItem);
		System.out.println("*********** Units      : "+actUnits      +"  value expected  "+expUnits);
		System.out.println("*********** Quantity   : "+actQuantity   +"  value expected  "+expQuantity);
		System.out.println("*********** Rate       : "+actRate       +"  value expected  "+expRate);
		System.out.println("*********** Gross      : "+actGross      +"  value expected  "+expGross);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 58, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 59, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 60, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 61, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 2");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 65, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 66, 6));
		
		excelReader.setCellData(xlfile, xlSheetName, 65, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 65, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 66, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 65, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		
		if (actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) && actItem.equalsIgnoreCase(expItem) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate)
				&& actGross.equalsIgnoreCase(expGross) 
				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 58, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 58, 8, resFail);
			return false;
		}
	}




	



	public boolean checkEditingAndVerifingDataOfSavedOpeningStocksVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 67, 6);
		excelReader.setCellData(xlfile, xlSheetName, 67, 7, actDocNo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		String actWarehouse=warehouseTxt.getAttribute("value");
		String expWarehouse=excelReader.getCellData(xlSheetName, 68, 6);
		excelReader.setCellData(xlfile, xlSheetName, 68, 7, actWarehouse);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actItem=select1stRow_1stColumn.getText();
		String expItem=excelReader.getCellData(xlSheetName, 69, 6);
		excelReader.setCellData(xlfile, xlSheetName, 69, 7, actItem);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
		String actUnits=select1stRow_2ndColumn.getText();
		String expUnits=excelReader.getCellData(xlSheetName, 70, 6);
		excelReader.setCellData(xlfile, xlSheetName, 70, 7, actUnits);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		String actQuantity=select1stRow_3rdColumn.getText();
		String expQuantity=excelReader.getCellData(xlSheetName, 71, 6);
		excelReader.setCellData(xlfile, xlSheetName, 71, 7, actQuantity);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		String actRate=select1stRow_4thColumn.getText();
		String expRate=excelReader.getCellData(xlSheetName, 72, 6);
		excelReader.setCellData(xlfile, xlSheetName, 72, 7, actRate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		String actGross=select1stRow_5thColumn.getText();
		String expGross=excelReader.getCellData(xlSheetName, 73, 6);
		excelReader.setCellData(xlfile, xlSheetName, 73, 7, actGross);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actItem2=select2ndRow_1stColumn.getText();
		String expItem2=excelReader.getCellData(xlSheetName, 74, 6);
		excelReader.setCellData(xlfile, xlSheetName, 74, 7, actItem2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_2ndColumn));
		String actUnits2=select2ndRow_2ndColumn.getText();
		String expUnits2=excelReader.getCellData(xlSheetName, 75, 6);
		excelReader.setCellData(xlfile, xlSheetName, 75, 7, actUnits2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
		String actQuantity2=select2ndRow_3rdColumn.getText();
		String expQuantity2=excelReader.getCellData(xlSheetName, 76, 6);
		excelReader.setCellData(xlfile, xlSheetName, 76, 7, actQuantity2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
		String actRate2=select2ndRow_4thColumn.getText();
		String expRate2=excelReader.getCellData(xlSheetName, 77, 6);
		excelReader.setCellData(xlfile, xlSheetName, 77, 7, actRate2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
		String actGross2=select2ndRow_5thColumn.getText();
		String expGross2=excelReader.getCellData(xlSheetName, 78, 6);
		excelReader.setCellData(xlfile, xlSheetName, 78, 7, actGross2);

		System.out.println("**********************checkEditingAndVerifingDataOfSavedOpeningStocksVoucher2*********************");

		System.out.println("*********** Voucher No  : "+actDocNo      +"  value expected  "+expDocNo);
		System.out.println("*********** Warehouse   : "+actWarehouse  +"  value expected  "+expWarehouse);
		System.out.println("*********** Item        : "+actItem       +"  value expected  "+expItem);
		System.out.println("*********** Units       : "+actUnits      +"  value expected  "+expUnits);
		System.out.println("*********** Quantity    : "+actQuantity   +"  value expected  "+expQuantity);
		System.out.println("*********** Rate        : "+actRate       +"  value expected  "+expRate);
		System.out.println("*********** Gross       : "+actGross      +"  value expected  "+expGross);
		System.out.println("*********** Item2       : "+actItem2       +"  value expected  "+expItem2);
		System.out.println("*********** Units2      : "+actUnits2      +"  value expected  "+expUnits2);
		System.out.println("*********** Quantity2   : "+actQuantity2   +"  value expected  "+expQuantity2);
		System.out.println("*********** Rate2       : "+actRate2       +"  value expected  "+expRate2);
		System.out.println("*********** Gross2      : "+actGross2      +"  value expected  "+expGross2);

		if (actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) && actItem.equalsIgnoreCase(expItem) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate)
				&& actGross.equalsIgnoreCase(expGross) 
				&& actItem2.equalsIgnoreCase(expItem2) && actUnits2.equalsIgnoreCase(expUnits2) && actQuantity2.equalsIgnoreCase(expQuantity2) 
				&& actRate2.equalsIgnoreCase(expRate2) && actGross2.equalsIgnoreCase(expGross2)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 67, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 67, 8, resFail);
			return false;
		}
	}







	public boolean checkSavingOpeningStockVoucher3WithStockItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
		new_newBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 79, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 80, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 81, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 82, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 3");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 79, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 80, 6));
		
		excelReader.setCellData(xlfile, xlSheetName, 79, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actMessage.equals(expMessage))*/
		
String expMessage1 = excelReader.getCellData(xlSheetName, 79, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 80, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 79, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		{
			excelReader.setCellData(xlfile, xlSheetName, 79, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 79, 8, resFail);
			return false;
		}
	}








	public boolean checkVerifingDataOfSavedOpeningStocksVoucher3AndDeleting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 83, 6);
		excelReader.setCellData(xlfile, xlSheetName, 83, 7, actDocNo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		String actWarehouse=warehouseTxt.getAttribute("value");
		String expWarehouse=excelReader.getCellData(xlSheetName, 84, 6);
		excelReader.setCellData(xlfile, xlSheetName, 84, 7, actWarehouse);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actItem=select1stRow_1stColumn.getText();
		String expItem=excelReader.getCellData(xlSheetName, 85, 6);
		excelReader.setCellData(xlfile, xlSheetName, 85, 7, actItem);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
		String actUnits=select1stRow_2ndColumn.getText();
		String expUnits=excelReader.getCellData(xlSheetName, 86, 6);
		excelReader.setCellData(xlfile, xlSheetName, 86, 7, actUnits);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		String actQuantity=select1stRow_3rdColumn.getText();
		String expQuantity=excelReader.getCellData(xlSheetName, 87, 6);
		excelReader.setCellData(xlfile, xlSheetName, 87, 7, actQuantity);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		String actRate=select1stRow_4thColumn.getText();
		String expRate=excelReader.getCellData(xlSheetName, 88, 6);
		excelReader.setCellData(xlfile, xlSheetName, 88, 7, actRate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		String actGross=select1stRow_5thColumn.getText();
		String expGross=excelReader.getCellData(xlSheetName, 89, 6);
		excelReader.setCellData(xlfile, xlSheetName, 89, 7, actGross);

		System.out.println("**********************checkVerifingDataOfSavedOpeningStocksVoucher3AndDeleting*********************");

		System.out.println("*********** Voucher No : "+actDocNo      +"  value expected  "+expDocNo);
		System.out.println("*********** Warehouse  : "+actWarehouse  +"  value expected  "+expWarehouse);
		System.out.println("*********** Item       : "+actItem       +"  value expected  "+expItem);
		System.out.println("*********** Units      : "+actUnits      +"  value expected  "+expUnits);
		System.out.println("*********** Quantity   : "+actQuantity   +"  value expected  "+expQuantity);
		System.out.println("*********** Rate       : "+actRate       +"  value expected  "+expRate);
		System.out.println("*********** Gross      : "+actGross      +"  value expected  "+expGross);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		String expMessage=excelReader.getCellData(xlSheetName, 90, 6);
		String actMessage=checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 90, 7, actMessage);

		if (actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) && actItem.equalsIgnoreCase(expItem) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate)
				&& actGross.equalsIgnoreCase(expGross) 
				&& actMessage.equalsIgnoreCase(expMessage)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 83, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 83, 8, resFail);
			return false;
		}
	}









	public boolean checkOpeningStockVoucher3Row1WithBinItemWithAutoAllocationOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();
		
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksVoucher));
		openingStocksVoucher.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 91, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 92, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 93, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 94, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		String actBinTotalQty           = binTotalQty.getAttribute("value");
		String actBinBaseUom            = binBaseUOM.getAttribute("value");
		String actBinSearchTxt          = Boolean.toString(binSearchTxt.isDisplayed());
		String actBinSearchBtn          = Boolean.toString(binSearchBtn.isDisplayed());
		String actBinAutoAllocateBtn    = Boolean.toString(binAutoAllocateBtn.isDisplayed());
		String actBinPickBtn            = Boolean.toString(binPickBtn.isDisplayed());
		String actBinCancelBtn          = Boolean.toString(binCancelBtn.isDisplayed());
		//String actBinBalanceQty         = binBalanceTxt.getAttribute("value");

		String expBinTotalQty           = excelReader.getCellData(xlSheetName, 91, 6);
		excelReader.setCellData(xlfile, xlSheetName, 91, 7, actBinTotalQty);
		
		String expBinBaseUom            = excelReader.getCellData(xlSheetName, 92, 6);
		excelReader.setCellData(xlfile, xlSheetName, 92, 7, actBinBaseUom);
		
		String expBinSearchTxt         = excelReader.getCellData(xlSheetName, 93, 6);
		excelReader.setCellData(xlfile, xlSheetName, 93, 7, actBinSearchTxt);
		
		String expBinSearchBtn         = excelReader.getCellData(xlSheetName, 94, 6);
		excelReader.setCellData(xlfile, xlSheetName, 94, 7, actBinSearchBtn);
		
		String expBinAutoAllocateBtn   = excelReader.getCellData(xlSheetName, 95, 6);
		excelReader.setCellData(xlfile, xlSheetName, 95, 7, actBinAutoAllocateBtn);
				
		String expBinPickBtn           = excelReader.getCellData(xlSheetName, 96, 6);
		excelReader.setCellData(xlfile, xlSheetName, 96, 7, actBinPickBtn);
		
		String expBinCancelBtn         = excelReader.getCellData(xlSheetName, 97, 6);
		excelReader.setCellData(xlfile, xlSheetName, 97, 7, actBinCancelBtn);
		
		//String expBinBalanceQty         = excelReader.getCellData(xlSheetName, 98, 6);
		//excelReader.setCellData(xlfile, xlSheetName, 98, 7, actBinBalanceQty);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();
		
		Thread.sleep(2000);

		System.out.println("***********************************BeforeAllocation**************************************");

		System.out.println("BinTotalQty        : "+actBinTotalQty		 +"  Value Expected  "+expBinTotalQty);
		System.out.println("BinBaseUom         : "+actBinBaseUom		 +"  Value Expected  "+expBinBaseUom);
		System.out.println("BinSearchTxt       : "+actBinSearchTxt		 +"  Value Expected  "+expBinSearchTxt);
		System.out.println("BinSearchBtn       : "+actBinSearchBtn		 +"  Value Expected  "+expBinSearchBtn);
		System.out.println("BinAutoAllocateBtn : "+actBinAutoAllocateBtn +"  Value Expected  "+expBinAutoAllocateBtn);
		System.out.println("BinPickBtn         : "+actBinPickBtn		 +"  Value Expected  "+expBinPickBtn);
		System.out.println("BinCancelBtn       : "+actBinCancelBtn		 +"  Value Expected  "+expBinCancelBtn);
		//System.out.println("BinBalanceQty      : "+actBinBalanceQty	   +"  Value Expected  "+expBinBalanceQty);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
		binAutoAllocateBtn.click();

		Thread.sleep(2000);

		String actTotalToBeAllocated2 	= binRowTotalToBeAllocated.getText();
		String actBinBalanceQty2        = binBalanceTxt.getAttribute("value");

		String expTotalToBeAllocated2 	= excelReader.getCellData(xlSheetName, 99, 6);
		excelReader.setCellData(xlfile, xlSheetName, 99, 7, actTotalToBeAllocated2);
		
		String expBinBalanceQty2        = excelReader.getCellData(xlSheetName, 100, 6);
		excelReader.setCellData(xlfile, xlSheetName, 100, 7, actBinBalanceQty2);
		
		System.out.println("***********************************AfterAllocation**************************************");

		System.out.println("TotalToBeAllocated    : "+actTotalToBeAllocated2	+"  Value Expected  "+expTotalToBeAllocated2);
		System.out.println("BinBalanceQty         : "+actBinBalanceQty2			+"  Value Expected  "+expBinBalanceQty2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		binOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
		String actBinColumn=select1stRow_7thColumn.getText();

		String expBinColumn=excelReader.getCellData(xlSheetName, 101, 6);
		excelReader.setCellData(xlfile, xlSheetName, 101, 7, actBinColumn);
		
		System.out.println("Bin Column: "+actBinColumn+"  value expected  "+expBinColumn);

		if(actBinTotalQty.equalsIgnoreCase(expBinTotalQty) && actBinBaseUom.equalsIgnoreCase(expBinBaseUom) && actBinSearchTxt.equalsIgnoreCase(expBinSearchTxt) 
				&& actBinSearchBtn.equalsIgnoreCase(expBinSearchBtn) && actBinAutoAllocateBtn.equalsIgnoreCase(expBinAutoAllocateBtn) && actBinPickBtn.equalsIgnoreCase(expBinPickBtn) 
				&& actTotalToBeAllocated2.equalsIgnoreCase(expTotalToBeAllocated2) && actBinBalanceQty2.equalsIgnoreCase(expBinBalanceQty2)
				&& actBinColumn.equalsIgnoreCase(expBinColumn))
		{
			excelReader.setCellData(xlfile, xlSheetName, 91, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 91, 8, resFail);
			return false;
		}
	}



	



	public boolean checkSavingOpeningStockVoucher3With2RowsAndBinItemDoubleClickInIndex() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 102, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 103, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 104, 5));
		Thread.sleep(2000);
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		String actBinTotalQty           = binTotalQty.getAttribute("value");
		String actBinBaseUom            = binBaseUOM.getAttribute("value");
		String actBinSearchTxt          = Boolean.toString(binSearchTxt.isDisplayed());
		String actBinSearchBtn          = Boolean.toString(binSearchBtn.isDisplayed());
		String actBinAutoAllocateBtn    = Boolean.toString(binAutoAllocateBtn.isDisplayed());
		String actBinPickBtn            = Boolean.toString(binPickBtn.isDisplayed());
		String actBinCancelBtn          = Boolean.toString(binCancelBtn.isDisplayed());
		//String actBinBalanceQty         = binBalanceTxt.getAttribute("value");

		String expBinTotalQty           = excelReader.getCellData(xlSheetName, 102, 6);
		excelReader.setCellData(xlfile, xlSheetName, 102, 7, actBinTotalQty);
		
		String expBinBaseUom            = excelReader.getCellData(xlSheetName, 103, 6);
		excelReader.setCellData(xlfile, xlSheetName, 103, 7, actBinBaseUom);
		
		String expBinSearchTxt          = excelReader.getCellData(xlSheetName, 104, 6);
		excelReader.setCellData(xlfile, xlSheetName, 104, 7, actBinSearchTxt);
		
		String expBinSearchBtn          = excelReader.getCellData(xlSheetName, 105, 6);
		excelReader.setCellData(xlfile, xlSheetName, 105, 7, actBinSearchBtn);
		
		String expBinAutoAllocateBtn    = excelReader.getCellData(xlSheetName, 106, 6);
		excelReader.setCellData(xlfile, xlSheetName, 106, 7, actBinAutoAllocateBtn);
		
		String expBinPickBtn            = excelReader.getCellData(xlSheetName, 107, 6);
		excelReader.setCellData(xlfile, xlSheetName, 107, 7, actBinPickBtn);
		
		String expBinCancelBtn          = excelReader.getCellData(xlSheetName, 108, 6);
		excelReader.setCellData(xlfile, xlSheetName, 108, 7, actBinCancelBtn);
		
		//String expBinBalanceQty         = excelReader.getCellData(xlSheetName, 109, 6);
		//excelReader.setCellData(xlfile, xlSheetName, 109, 7, actBinBalanceQty);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();

		String actTotalToBeAllocated 	= binRowTotalToBeAllocated.getText();
		String expTotalToBeAllocated 	= excelReader.getCellData(xlSheetName, 110, 6);
		excelReader.setCellData(xlfile, xlSheetName, 110, 7, actTotalToBeAllocated);
		
		System.out.println("***********************************BeforeAllocation**************************************");

		System.out.println("BinTotalQty         : "+actBinTotalQty			+"  Value Expected  "+expBinTotalQty);
		System.out.println("BinBaseUom          : "+actBinBaseUom			+"  Value Expected  "+expBinBaseUom);
		System.out.println("BinSearchTxt        : "+actBinSearchTxt			+"  Value Expected  "+expBinSearchTxt);
		System.out.println("BinSearchBtn        : "+actBinSearchBtn			+"  Value Expected  "+expBinSearchBtn);
		System.out.println("BinAutoAllocateBtn  : "+actBinAutoAllocateBtn	+"  Value Expected  "+expBinAutoAllocateBtn);
		System.out.println("BinPickBtn          : "+actBinPickBtn			+"  Value Expected  "+expBinPickBtn);
		System.out.println("BinCancelBtn        : "+actBinCancelBtn			+"  Value Expected  "+expBinCancelBtn);
		//System.out.println("BinBalanceQty       : "+actBinBalanceQty		  +"  Value Expected  "+expBinBalanceQty);

		System.out.println("TotalToBeAllocated    : "+actTotalToBeAllocated		+"  Value Expected  "+expTotalToBeAllocated);

		int count = binPopupBinsList.size();

		for (int i = 0; i < count; i++) 
		{
			String data = binPopupBinsList.get(i).getText();

			if (data.equalsIgnoreCase("Bin5")) 
			{
				getAction().doubleClick(binPopupIndexList.get(i)).build().perform();
			}
		}

		Thread.sleep(2000);

		String actTotalToBeAllocated2 	= binRowTotalToBeAllocated.getText();
		String actBinBalanceQty2        = binBalanceTxt.getAttribute("value");

		String expTotalToBeAllocated2 	= excelReader.getCellData(xlSheetName, 111, 6);
		excelReader.setCellData(xlfile, xlSheetName, 111, 7, actTotalToBeAllocated2);
		
		String expBinBalanceQty2        = excelReader.getCellData(xlSheetName, 112, 6);
		excelReader.setCellData(xlfile, xlSheetName, 112, 7, actBinBalanceQty2);

		System.out.println("***********************************AfterAllocation**************************************");

		System.out.println("TotalToBeAllocated    : "+actTotalToBeAllocated2	+"  Value Expected  "+expTotalToBeAllocated2);
		System.out.println("BinBalanceQty         : "+actBinBalanceQty2			+"  Value Expected  "+expBinBalanceQty2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		getAction().moveToElement(binOkBtn).click().build().perform();
		//binOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_7thColumn));
		String actBinColumn=select2ndRow_7thColumn.getText();

		String expBinColumn=excelReader.getCellData(xlSheetName, 113, 6);
		excelReader.setCellData(xlfile, xlSheetName, 114, 7, actBinColumn);
		
		System.out.println("Bin Column: "+actBinColumn+"  value expected  "+expBinColumn);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		getAction().moveToElement(voucherSaveBtn).click().build().perform();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 3");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 114, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 115, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 114, 7, actMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		
String expMessage1 = excelReader.getCellData(xlSheetName, 114, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 115, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 114, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))


		if(actBinTotalQty.equalsIgnoreCase(expBinTotalQty) && actBinBaseUom.equalsIgnoreCase(expBinBaseUom) && actBinSearchTxt.equalsIgnoreCase(expBinSearchTxt) 
				&& actBinSearchBtn.equalsIgnoreCase(expBinSearchBtn) && actBinAutoAllocateBtn.equalsIgnoreCase(expBinAutoAllocateBtn) && actBinPickBtn.equalsIgnoreCase(expBinPickBtn) 

				 && actTotalToBeAllocated2.equalsIgnoreCase(expTotalToBeAllocated2) && actBinBalanceQty2.equalsIgnoreCase(expBinBalanceQty2)

				 && actBinColumn.equalsIgnoreCase(expBinColumn)

				 && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 102, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 102, 8, resFail);
			return false;
		}
	}




	



	public boolean checkEditingAndVerifingDataOfSavedOpeningStocksVoucher3AddingRow3WithBinItemByAutoAllocation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo= excelReader.getCellData(xlSheetName, 116, 6);
		excelReader.setCellData(xlfile, xlSheetName, 116, 7, actDocNo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		String actWarehouse=warehouseTxt.getAttribute("value");
		String expWarehouse= excelReader.getCellData(xlSheetName, 117, 6);
		excelReader.setCellData(xlfile, xlSheetName, 117, 7, actWarehouse);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Item=select1stRow_1stColumn.getText();
		String expR1Item= excelReader.getCellData(xlSheetName, 118, 6);
		excelReader.setCellData(xlfile, xlSheetName, 118, 7, actR1Item);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
		String actR1Units=select1stRow_2ndColumn.getText();
		String expR1Units= excelReader.getCellData(xlSheetName, 119, 6);
		excelReader.setCellData(xlfile, xlSheetName, 119, 7, actR1Units);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		String actR1Quantity=select1stRow_3rdColumn.getText();
		String expR1Quantity= excelReader.getCellData(xlSheetName, 120, 6);
		excelReader.setCellData(xlfile, xlSheetName, 120, 7, actR1Quantity);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		String actR1Rate=select1stRow_4thColumn.getText();
		String expR1Rate= excelReader.getCellData(xlSheetName, 121, 6);
		excelReader.setCellData(xlfile, xlSheetName, 121, 7, actR1Rate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		String actR1Gross=select1stRow_5thColumn.getText();
		String expR1Gross= excelReader.getCellData(xlSheetName, 122, 6);
		excelReader.setCellData(xlfile, xlSheetName, 122, 7, actR1Gross);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
		String actR1Bin=select1stRow_7thColumn.getText();
		String expR1Bin= excelReader.getCellData(xlSheetName, 123, 6);
		excelReader.setCellData(xlfile, xlSheetName, 123, 7, actR1Bin);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Item=select2ndRow_1stColumn.getText();
		String expR2Item=excelReader.getCellData(xlSheetName, 124, 6);
		excelReader.setCellData(xlfile, xlSheetName, 124, 7, actR2Item);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_2ndColumn));
		String actR2Units=select2ndRow_2ndColumn.getText();
		String expR2Units= excelReader.getCellData(xlSheetName, 125, 6);
		excelReader.setCellData(xlfile, xlSheetName, 125, 7, actR2Units);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
		String actR2Quantity=select2ndRow_3rdColumn.getText();
		String expR2Quantity= excelReader.getCellData(xlSheetName, 126, 6);
		excelReader.setCellData(xlfile, xlSheetName, 126, 7, actR2Quantity);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
		String actR2Rate=select2ndRow_4thColumn.getText();
		String expR2Rate= excelReader.getCellData(xlSheetName, 127, 6);
		excelReader.setCellData(xlfile, xlSheetName, 127, 7, actR2Rate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
		String actR2Gross=select2ndRow_5thColumn.getText();
		String expR2Gross= excelReader.getCellData(xlSheetName, 128, 6);
		excelReader.setCellData(xlfile, xlSheetName, 128, 7, actR2Gross);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_7thColumn));
		String actR2Bin=select2ndRow_7thColumn.getText();
		String expR2Bin= excelReader.getCellData(xlSheetName, 129, 6);
		excelReader.setCellData(xlfile, xlSheetName, 129, 7, actR2Bin);

		System.out.println("**********************checkEditingAndVerifingDataOfSavedOpeningStocksVoucher3AddingRow3WithBinItem*********************");

		System.out.println("*********** Voucher No : "+actDocNo        +"  value expected  "+expDocNo);
		System.out.println("*********** Warehouse  : "+actWarehouse    +"  value expected  "+expWarehouse);
		System.out.println("*********************************Row1******************************************");
		System.out.println("*********** Item       : "+actR1Item       +"  value expected  "+expR1Item);
		System.out.println("*********** Units      : "+actR1Units      +"  value expected  "+expR1Units);
		System.out.println("*********** Quantity   : "+actR1Quantity   +"  value expected  "+expR1Quantity);
		System.out.println("*********** Rate       : "+actR1Rate       +"  value expected  "+expR1Rate);
		System.out.println("*********** Gross      : "+actR1Gross      +"  value expected  "+expR1Gross);
		System.out.println("*********** Bin        : "+actR1Bin        +"  value expected  "+expR1Bin);
		System.out.println("*********************************Row2******************************************");
		System.out.println("*********** Item       : "+actR2Item       +"  value expected  "+expR2Item);
		System.out.println("*********** Units      : "+actR2Units      +"  value expected  "+expR2Units);
		System.out.println("*********** Quantity   : "+actR2Quantity   +"  value expected  "+expR2Quantity);
		System.out.println("*********** Rate       : "+actR2Rate       +"  value expected  "+expR2Rate);
		System.out.println("*********** Gross      : "+actR2Gross      +"  value expected  "+expR2Gross);
		System.out.println("*********** Bin        : "+actR2Bin        +"  value expected  "+expR2Bin);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
		select3rdRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 116, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 117, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 118, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		String actBinTotalQty           = binTotalQty.getAttribute("value");
		String actBinBaseUom            = binBaseUOM.getAttribute("value");
		String actBinSearchTxt          = Boolean.toString(binSearchTxt.isDisplayed());
		String actBinSearchBtn          = Boolean.toString(binSearchBtn.isDisplayed());
		String actBinAutoAllocateBtn    = Boolean.toString(binAutoAllocateBtn.isDisplayed());
		String actBinPickBtn            = Boolean.toString(binPickBtn.isDisplayed());
		String actBinCancelBtn          = Boolean.toString(binCancelBtn.isDisplayed());
		//String actBinBalanceQty         = binBalanceTxt.getAttribute("value");

		String expBinTotalQty           = excelReader.getCellData(xlSheetName, 130, 6);
		excelReader.setCellData(xlfile, xlSheetName, 130, 7, actBinTotalQty);
		
		String expBinBaseUom            = excelReader.getCellData(xlSheetName, 131, 6);
		excelReader.setCellData(xlfile, xlSheetName, 131, 7, actBinBaseUom);
		
		String expBinSearchTxt          = excelReader.getCellData(xlSheetName, 132, 6);
		excelReader.setCellData(xlfile, xlSheetName, 132, 7, actBinSearchTxt);
		
		String expBinSearchBtn          = excelReader.getCellData(xlSheetName, 133, 6);
		excelReader.setCellData(xlfile, xlSheetName, 133, 7, actBinSearchBtn);
		
		String expBinAutoAllocateBtn    = excelReader.getCellData(xlSheetName, 134, 6);
		excelReader.setCellData(xlfile, xlSheetName, 134, 7, actBinAutoAllocateBtn);
		
		String expBinPickBtn            = excelReader.getCellData(xlSheetName, 135, 6);
		excelReader.setCellData(xlfile, xlSheetName, 135, 7, actBinPickBtn);
		
		String expBinCancelBtn          = excelReader.getCellData(xlSheetName, 136, 6);
		excelReader.setCellData(xlfile, xlSheetName, 136, 7, actBinCancelBtn);
		
		//String expBinBalanceQty         = excelReader.getCellData(xlSheetName, 137, 6);
		//excelReader.setCellData(xlfile, xlSheetName, 137, 7, actBinBalanceQty);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();

		String actTotalToBeAllocated 	= binRowTotalToBeAllocated.getText();
		String expTotalToBeAllocated 	= excelReader.getCellData(xlSheetName, 138, 6);
		excelReader.setCellData(xlfile, xlSheetName, 138, 7, actTotalToBeAllocated);
		
		System.out.println("***********************************BeforeAllocation**************************************");

		System.out.println("BinTotalQty        : "+actBinTotalQty			+"  Value Expected  "+expBinTotalQty);
		System.out.println("BinBaseUom         : "+actBinBaseUom			+"  Value Expected  "+expBinBaseUom);
		System.out.println("BinSearchTxt       : "+actBinSearchTxt			+"  Value Expected  "+expBinSearchTxt);
		System.out.println("BinSearchBtn       : "+actBinSearchBtn			+"  Value Expected  "+expBinSearchBtn);
		System.out.println("BinAutoAllocateBtn : "+actBinAutoAllocateBtn	+"  Value Expected  "+expBinAutoAllocateBtn);
		System.out.println("BinPickBtn         : "+actBinPickBtn			+"  Value Expected  "+expBinPickBtn);
		System.out.println("BinCancelBtn       : "+actBinCancelBtn			+"  Value Expected  "+expBinCancelBtn);
		//System.out.println("BinBalanceQty      : "+actBinBalanceQty		+"  Value Expected  "+expBinBalanceQty);

		System.out.println("TotalToBeAllocated    : "+actTotalToBeAllocated		+"  Value Expected  "+expTotalToBeAllocated);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
		binAutoAllocateBtn.click();

		Thread.sleep(2000);

		String actTotalToBeAllocated2 	= binRowTotalToBeAllocated.getText();
		String actBinBalanceQty2        = binBalanceTxt.getAttribute("value");

		String expTotalToBeAllocated2 	= excelReader.getCellData(xlSheetName, 139, 6);
		excelReader.setCellData(xlfile, xlSheetName, 139, 7, actTotalToBeAllocated2);
		
		String expBinBalanceQty2        = excelReader.getCellData(xlSheetName, 140, 6);
		excelReader.setCellData(xlfile, xlSheetName, 140, 7, actBinBalanceQty2);
		
		System.out.println("***********************************AfterAllocation**************************************");

		System.out.println("TotalToBeAllocated    : "+actTotalToBeAllocated2	+"  Value Expected  "+expTotalToBeAllocated2);
		System.out.println("BinBalanceQty         : "+actBinBalanceQty2			+"  Value Expected  "+expBinBalanceQty2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		binOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_7thColumn));
		String actBinColumn=select3rdRow_7thColumn.getText();

		String expBinColumn=excelReader.getCellData(xlSheetName, 141, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 141, 7, actBinColumn);

		System.out.println("Bin Column: "+actBinColumn+"  value expected  "+expBinColumn);

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 3");
		expMessage.add("Saving in background.");*/
		
	/*	String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 142, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 143, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 142, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
String expMessage1 = excelReader.getCellData(xlSheetName, 142, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 143, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 142, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))


		if (actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) && actR1Item.equalsIgnoreCase(expR1Item) 
				&& actR1Units.equalsIgnoreCase(expR1Units) && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) 
				&& actR1Gross.equalsIgnoreCase(expR1Gross) && actR1Bin.equalsIgnoreCase(expR1Bin) && actR2Item.equalsIgnoreCase(expR2Item) 
				&& actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) 
				&& actR2Gross.equalsIgnoreCase(expR2Gross) && actR2Bin.equalsIgnoreCase(expR2Bin) 

				&& actBinTotalQty.equalsIgnoreCase(expBinTotalQty) && actBinBaseUom.equalsIgnoreCase(expBinBaseUom) && actBinSearchTxt==expBinSearchTxt 
				&& actBinSearchBtn==expBinSearchBtn && actBinAutoAllocateBtn==expBinAutoAllocateBtn && actBinPickBtn==expBinPickBtn 
				&& actBinCancelBtn==expBinCancelBtn && actTotalToBeAllocated.equalsIgnoreCase(expTotalToBeAllocated)

				&& actTotalToBeAllocated2.equalsIgnoreCase(expTotalToBeAllocated2) && actBinBalanceQty2.equalsIgnoreCase(expBinBalanceQty2)

				&& actBinColumn.equalsIgnoreCase(expBinColumn)

				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 116, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 116, 8, resFail);
			return false;
		}
	}






	public boolean checkVerifingDataOfSavedOpeningStocksVoucher3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 144, 6);
		excelReader.setCellData(xlfile, xlSheetName, 144, 7, actDocNo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		String actWarehouse=warehouseTxt.getAttribute("value");
		String expWarehouse=excelReader.getCellData(xlSheetName, 145, 6);
		excelReader.setCellData(xlfile, xlSheetName, 145, 7, actWarehouse);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Item=select1stRow_1stColumn.getText();
		String expR1Item=excelReader.getCellData(xlSheetName, 146, 6);
		excelReader.setCellData(xlfile, xlSheetName, 146, 7, actR1Item);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
		String actR1Units=select1stRow_2ndColumn.getText();
		String expR1Units=excelReader.getCellData(xlSheetName, 147, 6);
		excelReader.setCellData(xlfile, xlSheetName, 147, 7, actR1Units);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		String actR1Quantity=select1stRow_3rdColumn.getText();
		String expR1Quantity=excelReader.getCellData(xlSheetName, 148, 6);
		excelReader.setCellData(xlfile, xlSheetName, 148, 7, actR1Quantity);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		String actR1Rate=select1stRow_4thColumn.getText();
		String expR1Rate=excelReader.getCellData(xlSheetName, 149, 6);
		excelReader.setCellData(xlfile, xlSheetName, 149, 7, actR1Rate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		String actR1Gross=select1stRow_5thColumn.getText();
		String expR1Gross=excelReader.getCellData(xlSheetName, 150, 6);
		excelReader.setCellData(xlfile, xlSheetName, 150, 7, actR1Gross);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
		String actR1Bin=select1stRow_7thColumn.getText();
		String expR1Bin=excelReader.getCellData(xlSheetName, 151, 6);
		excelReader.setCellData(xlfile, xlSheetName, 151, 7, actR1Bin);



		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Item=select2ndRow_1stColumn.getText();
		String expR2Item=excelReader.getCellData(xlSheetName, 152, 6);
		excelReader.setCellData(xlfile, xlSheetName, 152, 7, actR2Item);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_2ndColumn));
		String actR2Units=select2ndRow_2ndColumn.getText();
		String expR2Units=excelReader.getCellData(xlSheetName, 153, 6);
		excelReader.setCellData(xlfile, xlSheetName, 153, 7, actR2Units);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
		String actR2Quantity=select2ndRow_3rdColumn.getText();
		String expR2Quantity=excelReader.getCellData(xlSheetName, 154, 6);
		excelReader.setCellData(xlfile, xlSheetName, 154, 7, actR2Quantity);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
		String actR2Rate=select2ndRow_4thColumn.getText();
		String expR2Rate=excelReader.getCellData(xlSheetName, 155, 6);
		excelReader.setCellData(xlfile, xlSheetName, 155, 7, actR2Rate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
		String actR2Gross=select2ndRow_5thColumn.getText();
		String expR2Gross=excelReader.getCellData(xlSheetName, 156, 6);
		excelReader.setCellData(xlfile, xlSheetName, 156, 7, actR2Gross);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_7thColumn));
		String actR2Bin=select2ndRow_7thColumn.getText();
		String expR2Bin=excelReader.getCellData(xlSheetName, 157, 6);
		excelReader.setCellData(xlfile, xlSheetName, 157, 7, actR2Bin);



		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
		String actR3Item=select3rdRow_1stColumn.getText();
		String expR3Item=excelReader.getCellData(xlSheetName, 158, 6);
		excelReader.setCellData(xlfile, xlSheetName, 158, 7, actR3Item);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_2ndColumn));
		String actR3Units=select3rdRow_2ndColumn.getText();
		String expR3Units=excelReader.getCellData(xlSheetName, 159, 6);
		excelReader.setCellData(xlfile, xlSheetName, 159, 7, actR3Units);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_3rdColumn));
		String actR3Quantity=select3rdRow_3rdColumn.getText();
		String expR3Quantity=excelReader.getCellData(xlSheetName, 160, 6);
		excelReader.setCellData(xlfile, xlSheetName, 160, 7, actR3Quantity);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_4thColumn));
		String actR3Rate=select3rdRow_4thColumn.getText();
		String expR3Rate=excelReader.getCellData(xlSheetName, 161, 6);
		excelReader.setCellData(xlfile, xlSheetName, 161, 7, actR3Rate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_5thColumn));
		String actR3Gross=select3rdRow_5thColumn.getText();
		String expR3Gross=excelReader.getCellData(xlSheetName, 162, 6);
		excelReader.setCellData(xlfile, xlSheetName, 162, 7, actR3Gross);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_7thColumn));
		String actR3Bin=select3rdRow_7thColumn.getText();
		String expR3Bin=excelReader.getCellData(xlSheetName, 163, 6);
		excelReader.setCellData(xlfile, xlSheetName, 163, 7, actR3Bin);
		
		System.out.println("**********************checkVerifingDataOfSavedOpeningStocksVoucher3*********************");

		System.out.println("*********** Voucher No : "+actDocNo        +"  value expected  "+expDocNo);
		System.out.println("*********** Warehouse  : "+actWarehouse    +"  value expected  "+expWarehouse);
		System.out.println("*********************************Row1******************************************");
		System.out.println("*********** Item       : "+actR1Item       +"  value expected  "+expR1Item);
		System.out.println("*********** Units      : "+actR1Units      +"  value expected  "+expR1Units);
		System.out.println("*********** Quantity   : "+actR1Quantity   +"  value expected  "+expR1Quantity);
		System.out.println("*********** Rate       : "+actR1Rate       +"  value expected  "+expR1Rate);
		System.out.println("*********** Gross      : "+actR1Gross      +"  value expected  "+expR1Gross);
		System.out.println("*********** Bin        : "+actR1Bin        +"  value expected  "+expR1Bin);
		System.out.println("*********************************Row2******************************************");
		System.out.println("*********** Item       : "+actR2Item       +"  value expected  "+expR2Item);
		System.out.println("*********** Units      : "+actR2Units      +"  value expected  "+expR2Units);
		System.out.println("*********** Quantity   : "+actR2Quantity   +"  value expected  "+expR2Quantity);
		System.out.println("*********** Rate       : "+actR2Rate       +"  value expected  "+expR2Rate);
		System.out.println("*********** Gross      : "+actR2Gross      +"  value expected  "+expR2Gross);
		System.out.println("*********** Bin        : "+actR2Bin        +"  value expected  "+expR2Bin);
		System.out.println("*********************************Row3******************************************");
		System.out.println("*********** Item       : "+actR3Item       +"  value expected  "+expR3Item);
		System.out.println("*********** Units      : "+actR3Units      +"  value expected  "+expR3Units);
		System.out.println("*********** Quantity   : "+actR3Quantity   +"  value expected  "+expR3Quantity);
		System.out.println("*********** Rate       : "+actR3Rate       +"  value expected  "+expR3Rate);
		System.out.println("*********** Gross      : "+actR3Gross      +"  value expected  "+expR3Gross);
		System.out.println("*********** Bin        : "+actR3Bin        +"  value expected  "+expR3Bin);

		if (actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) 
				&& actR1Item.equalsIgnoreCase(expR1Item) && actR1Units.equalsIgnoreCase(expR1Units) && actR1Quantity.equalsIgnoreCase(expR1Quantity) 
				&& actR1Rate.equalsIgnoreCase(expR1Rate) && actR1Gross.equalsIgnoreCase(expR1Gross) && actR1Bin.equalsIgnoreCase(expR1Bin) 

				&& actR2Item.equalsIgnoreCase(expR2Item) && actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) 
				&& actR2Rate.equalsIgnoreCase(expR2Rate) && actR2Gross.equalsIgnoreCase(expR2Gross) && actR2Bin.equalsIgnoreCase(expR2Bin) 

				&& actR3Item.equalsIgnoreCase(expR3Item) && actR3Units.equalsIgnoreCase(expR3Units) && actR3Quantity.equalsIgnoreCase(expR3Quantity) 
				&& actR3Rate.equalsIgnoreCase(expR3Rate) && actR3Gross.equalsIgnoreCase(expR3Gross) && actR3Bin.equalsIgnoreCase(expR3Bin)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 144, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 144, 8, resFail);
			return false;
		}
	}










	public boolean checkSavingOpeningStockVoucher4WithBinItemByManuallyEnteringQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
		new_newBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 164, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 165, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 166, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 167, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		String actBinTotalQty           = binTotalQty.getAttribute("value");
		String actBinBaseUom            = binBaseUOM.getAttribute("value");
		String actBinSearchTxt          = Boolean.toString(binSearchTxt.isDisplayed());
		String actBinSearchBtn          = Boolean.toString(binSearchBtn.isDisplayed());
		String actBinAutoAllocateBtn    = Boolean.toString(binAutoAllocateBtn.isDisplayed());
		String actBinPickBtn            = Boolean.toString(binAutoAllocateBtn.isDisplayed());
		String actBinCancelBtn          = Boolean.toString(binCancelBtn.isDisplayed());
		//String actBinBalanceQty         = binBalanceTxt.getAttribute("value");

		String expBinTotalQty           = excelReader.getCellData(xlSheetName, 164, 6);
		excelReader.setCellData(xlfile, xlSheetName, 164, 7, actBinTotalQty);
		
		String expBinBaseUom            = excelReader.getCellData(xlSheetName, 165, 6);
		excelReader.setCellData(xlfile, xlSheetName, 165, 7, actBinBaseUom);
		
		String expBinSearchTxt          = excelReader.getCellData(xlSheetName, 166, 6);
		excelReader.setCellData(xlfile, xlSheetName, 166, 7, actBinSearchTxt);
		
		String expBinSearchBtn          = excelReader.getCellData(xlSheetName, 167, 6);
		excelReader.setCellData(xlfile, xlSheetName, 167, 7, actBinSearchBtn);
		
		String expBinAutoAllocateBtn    = excelReader.getCellData(xlSheetName, 168, 6);
		excelReader.setCellData(xlfile, xlSheetName, 168, 7, actBinAutoAllocateBtn);
		
		String expBinPickBtn            = excelReader.getCellData(xlSheetName, 169, 6);
		excelReader.setCellData(xlfile, xlSheetName, 169, 7, actBinPickBtn);
		
		String expBinCancelBtn          = excelReader.getCellData(xlSheetName, 170, 6);
		excelReader.setCellData(xlfile, xlSheetName, 170, 7, actBinCancelBtn);
		
		//String expBinBalanceQty         = excelReader.getCellData(xlSheetName, 171, 6);
		//excelReader.setCellData(xlfile, xlSheetName, 171, 7, actBinBalanceQty);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();

		String actTotalToBeAllocated 	= binRowTotalToBeAllocated.getText();
		String expTotalToBeAllocated 	= excelReader.getCellData(xlSheetName, 172, 6);
		excelReader.setCellData(xlfile, xlSheetName, 172, 7, actTotalToBeAllocated);
		
		System.out.println("***********************************BeforeAllocation**************************************");

		System.out.println("BinTotalQty         : "+actBinTotalQty			+"  Value Expected  "+expBinTotalQty);
		System.out.println("BinBaseUom          : "+actBinBaseUom			+"  Value Expected  "+expBinBaseUom);
		System.out.println("BinSearchTxt        : "+actBinSearchTxt			+"  Value Expected  "+expBinSearchTxt);
		System.out.println("BinSearchBtn        : "+actBinSearchBtn			+"  Value Expected  "+expBinSearchBtn);
		System.out.println("BinAutoAllocateBtn  : "+actBinAutoAllocateBtn	+"  Value Expected  "+expBinAutoAllocateBtn);
		System.out.println("BinPickBtn          : "+actBinPickBtn			+"  Value Expected  "+expBinPickBtn);
		System.out.println("BinCancelBtn        : "+actBinCancelBtn			+"  Value Expected  "+expBinCancelBtn);
		//System.out.println("BinBalanceQty     : "+actBinBalanceQty		+"  Value Expected  "+expBinBalanceQty);

		System.out.println("TotalToBeAllocated    : "+actTotalToBeAllocated		+"  Value Expected  "+expTotalToBeAllocated);

		int count = binPopupBinsList.size();

		for (int i = 0; i < count; i++) 
		{
			String data = binPopupBinsList.get(i).getText();
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 168, 5))) 
			{
				binPopupToBeAllocatedList.get(i).click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ToBeAllocateTxt));
				enter_ToBeAllocateTxt.sendKeys(excelReader.getCellData(xlSheetName, 169, 5));
				Thread.sleep(1000);
			}
		}

		for (int i = 0; i < count; i++) 
		{
			String data = binPopupBinsList.get(i).getText();
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 170, 5))) 
			{
				binPopupBinsList.get(i).click();
				binPopupToBeAllocatedList.get(i).click();
				Thread.sleep(1000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ToBeAllocateTxt));
				enter_ToBeAllocateTxt.sendKeys(excelReader.getCellData(xlSheetName, 171, 5));
				enter_ToBeAllocateTxt.sendKeys(Keys.TAB);
			}
		}

		Thread.sleep(2000);

		String actTotalToBeAllocated2 	= binRowTotalToBeAllocated.getText();
		String actBinBalanceQty2        = binBalanceTxt.getAttribute("value");

		String expTotalToBeAllocated2 	= excelReader.getCellData(xlSheetName, 173, 6);
		excelReader.setCellData(xlfile, xlSheetName, 173, 7, actTotalToBeAllocated2);
		
		String expBinBalanceQty2        = excelReader.getCellData(xlSheetName, 174, 6);
		excelReader.setCellData(xlfile, xlSheetName, 174, 7, actBinBalanceQty2);
		
		System.out.println("***********************************AfterAllocation**************************************");

		System.out.println("TotalToBeAllocated    : "+actTotalToBeAllocated2	+"  Value Expected  "+expTotalToBeAllocated2);
		System.out.println("BinBalanceQty         : "+actBinBalanceQty2			+"  Value Expected  "+expBinBalanceQty2);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		binOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
		String actR1BinColumn=select1stRow_7thColumn.getText();
		String expR1BinColumn=excelReader.getCellData(xlSheetName, 175, 6);
		excelReader.setCellData(xlfile, xlSheetName, 175, 7, actR1BinColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_7thColumn));
		String actR2BinColumn=select2ndRow_7thColumn.getText();
		String expR2BinColumn=excelReader.getCellData(xlSheetName, 176, 6);
		excelReader.setCellData(xlfile, xlSheetName, 177, 7, actR2BinColumn);

		System.out.println("Bin Column: "+actR1BinColumn+"  value expected  "+expR1BinColumn);
		System.out.println("Bin Column: "+actR2BinColumn+"  value expected  "+expR2BinColumn);

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 4");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 177, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 178, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 177, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		
String expMessage1 = excelReader.getCellData(xlSheetName, 177, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 178, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 177, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))


		if(actBinTotalQty.equalsIgnoreCase(expBinTotalQty) && actBinBaseUom.equalsIgnoreCase(expBinBaseUom) && actBinSearchTxt.equalsIgnoreCase(expBinSearchTxt) 
			&& actBinSearchBtn.equalsIgnoreCase(expBinSearchBtn) && actBinAutoAllocateBtn.equalsIgnoreCase(expBinAutoAllocateBtn) && actBinPickBtn.equalsIgnoreCase(expBinPickBtn) 
			&& actBinCancelBtn.equalsIgnoreCase(expBinCancelBtn) && actTotalToBeAllocated.equalsIgnoreCase(expTotalToBeAllocated)
			&& actTotalToBeAllocated2.equalsIgnoreCase(expTotalToBeAllocated2) && actBinBalanceQty2.equalsIgnoreCase(expBinBalanceQty2)
			&& actR1BinColumn.equalsIgnoreCase(expR1BinColumn) && actR2BinColumn.equalsIgnoreCase(expR2BinColumn)
			&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 164, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 164, 8, resFail);
			return false;
		}
	}




@FindBy(xpath="//*[@id='id_transactionentry_close']/a/span")
public static WebElement closeBtn;


	


	public boolean checkVerifingDataOfSavedOpeningStocksVoucher4AndDeleting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 179, 6);
		excelReader.setCellData(xlfile, xlSheetName, 179, 7, actDocNo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		String actWarehouse=warehouseTxt.getAttribute("value");
		String expWarehouse=excelReader.getCellData(xlSheetName, 180, 6);
		excelReader.setCellData(xlfile, xlSheetName, 180, 7, actWarehouse);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Item=select1stRow_1stColumn.getText();
		String expR1Item=excelReader.getCellData(xlSheetName, 181, 6);
		excelReader.setCellData(xlfile, xlSheetName, 181, 7, actR1Item);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
		String actR1Units=select1stRow_2ndColumn.getText();
		String expR1Units=excelReader.getCellData(xlSheetName, 182, 6);
		excelReader.setCellData(xlfile, xlSheetName, 182, 7, actR1Units);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		String actR1Quantity=select1stRow_3rdColumn.getText();
		String expR1Quantity=excelReader.getCellData(xlSheetName, 183, 6);
		excelReader.setCellData(xlfile, xlSheetName, 183, 7, actR1Quantity);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		String actR1Rate=select1stRow_4thColumn.getText();
		String expR1Rate=excelReader.getCellData(xlSheetName, 184, 6);
		excelReader.setCellData(xlfile, xlSheetName, 184, 7, actR1Rate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		String actR1Gross=select1stRow_5thColumn.getText();
		String expR1Gross=excelReader.getCellData(xlSheetName, 185, 6);
		excelReader.setCellData(xlfile, xlSheetName, 185, 7, actR1Gross);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
		String actR1Bin=select1stRow_7thColumn.getText();
		String expR1Bin=excelReader.getCellData(xlSheetName, 186, 6);
		excelReader.setCellData(xlfile, xlSheetName, 186, 7, actR1Bin);



		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Item=select2ndRow_1stColumn.getText();
		String expR2Item=excelReader.getCellData(xlSheetName, 187, 6);
		excelReader.setCellData(xlfile, xlSheetName, 187, 7, actR2Item);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_2ndColumn));
		String actR2Units=select2ndRow_2ndColumn.getText();
		String expR2Units=excelReader.getCellData(xlSheetName, 188, 6);
		excelReader.setCellData(xlfile, xlSheetName, 188, 7, actR2Units);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
		String actR2Quantity=select2ndRow_3rdColumn.getText();
		String expR2Quantity=excelReader.getCellData(xlSheetName, 189, 6);
		excelReader.setCellData(xlfile, xlSheetName, 189, 7, actR2Quantity);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
		String actR2Rate=select2ndRow_4thColumn.getText();
		String expR2Rate=excelReader.getCellData(xlSheetName, 190, 6);
		excelReader.setCellData(xlfile, xlSheetName, 190, 7, actR2Rate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
		String actR2Gross=select2ndRow_5thColumn.getText();
		String expR2Gross=excelReader.getCellData(xlSheetName, 191, 6);
		excelReader.setCellData(xlfile, xlSheetName, 191, 7, actR2Gross);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_7thColumn));
		String actR2Bin=select2ndRow_7thColumn.getText();
		String expR2Bin=excelReader.getCellData(xlSheetName, 192, 6);
		excelReader.setCellData(xlfile, xlSheetName, 192, 7, actR2Bin);

		System.out.println("**********************checkVerifingDataOfSavedOpeningStocksVoucher3*********************");

		System.out.println("*********** Voucher No : "+actDocNo        +"  value expected  "+expDocNo);
		System.out.println("*********** Warehouse  : "+actWarehouse    +"  value expected  "+expWarehouse);
		System.out.println("*********************************Row1******************************************");
		System.out.println("*********** Item       : "+actR1Item       +"  value expected  "+expR1Item);
		System.out.println("*********** Units      : "+actR1Units      +"  value expected  "+expR1Units);
		System.out.println("*********** Quantity   : "+actR1Quantity   +"  value expected  "+expR1Quantity);
		System.out.println("*********** Rate       : "+actR1Rate       +"  value expected  "+expR1Rate);
		System.out.println("*********** Gross      : "+actR1Gross      +"  value expected  "+expR1Gross);
		System.out.println("*********** Bin        : "+actR1Bin        +"  value expected  "+expR1Bin);
		System.out.println("*********************************Row2******************************************");
		System.out.println("*********** Item       : "+actR2Item       +"  value expected  "+expR2Item);
		System.out.println("*********** Units      : "+actR2Units      +"  value expected  "+expR2Units);
		System.out.println("*********** Quantity   : "+actR2Quantity   +"  value expected  "+expR2Quantity);
		System.out.println("*********** Rate       : "+actR2Rate       +"  value expected  "+expR2Rate);
		System.out.println("*********** Gross      : "+actR2Gross      +"  value expected  "+expR2Gross);
		System.out.println("*********** Bin        : "+actR2Bin        +"  value expected  "+expR2Bin);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		String expMessage=excelReader.getCellData(xlSheetName, 193, 6);
		String actMessage=checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 193, 7, actMessage);

		if (actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) 
				&& actR1Item.equalsIgnoreCase(expR1Item) && actR1Units.equalsIgnoreCase(expR1Units) && actR1Quantity.equalsIgnoreCase(expR1Quantity) 
				&& actR1Rate.equalsIgnoreCase(expR1Rate) && actR1Gross.equalsIgnoreCase(expR1Gross) && actR1Bin.equalsIgnoreCase(expR1Bin) 

				&& actR2Item.equalsIgnoreCase(expR2Item) && actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) 
				&& actR2Rate.equalsIgnoreCase(expR2Rate) && actR2Gross.equalsIgnoreCase(expR2Gross) && actR2Bin.equalsIgnoreCase(expR2Bin) 

				&& actMessage.equalsIgnoreCase(expMessage)) 
		{
			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
			
			closeBtn.click();
			*/
			excelReader.setCellData(xlfile, xlSheetName, 179, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 179, 8, resFail);
			return false;
		}
	}







	public boolean checkSavingOpeningStockVoucher4WithRMAItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 194, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 195, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 196, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 197, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
		rmaSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 198, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
		rmaAddBtn.click();

		Thread.sleep(2000);

		//actual
		String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
		String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
		String actRMA3=RMAPopupTableRow3Column1.getAttribute("data-value");
		String actRMA4=RMAPopupTableRow4Column1.getAttribute("data-value");
		String actRMA5=RMAPopupTableRow5Column1.getAttribute("data-value");
		String actRMA6=RMAPopupTableRow6Column1.getAttribute("data-value");
		String actRMA7=RMAPopupTableRow7Column1.getAttribute("data-value");
		String actRMA8=RMAPopupTableRow8Column1.getAttribute("data-value");
		String actRMA9=RMAPopupTableRow9Column1.getAttribute("data-value");
		String actRMA10=RMAPopupTableRow10Column1.getAttribute("data-value");
		String actRMA11=RMAPopupTableRow11Column1.getAttribute("data-value");
		String actRMA12=RMAPopupTableRow12Column1.getAttribute("data-value");

		//expected
		String expRMA1="RMA#R1OSV1";
		String expRMA2="RMA#R1OSV2";
		String expRMA3="RMA#R1OSV3";
		String expRMA4="RMA#R1OSV4";
		String expRMA5="RMA#R1OSV5";
		String expRMA6="RMA#R1OSV6";
		String expRMA7="RMA#R1OSV7";
		String expRMA8="RMA#R1OSV8";
		String expRMA9="RMA#R1OSV9";
		String expRMA10="RMA#R1OSV10";
		String expRMA11="RMA#R1OSV11";
		String expRMA12="RMA#R1OSV12";


		boolean Result = actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) &&
						 actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8) &&
						 actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10) && actRMA11.equalsIgnoreCase(expRMA11) && actRMA12.equalsIgnoreCase(expRMA12);
		
		String actRMAPopupDisplayingAsExpected = Boolean.toString(Result);
		String expRMAPopupDisplayingAsExpected = excelReader.getCellData(xlSheetName, 194, 6);
		excelReader.setCellData(xlfile, xlSheetName, 194, 7, actRMAPopupDisplayingAsExpected);
		
		
		System.out.println("********* Row1        : "+actRMA1  +"  value expected  "+expRMA1);
		System.out.println("********* Row2        : "+actRMA2  +"  value expected  "+expRMA2);
		System.out.println("********* Row3        : "+actRMA3  +"  value expected  "+expRMA3);
		System.out.println("********* Row4        : "+actRMA4  +"  value expected  "+expRMA4);
		System.out.println("********* Row5        : "+actRMA5  +"  value expected  "+expRMA5);
		System.out.println("********* Row6        : "+actRMA6  +"  value expected  "+expRMA6);
		System.out.println("********* Row7        : "+actRMA7  +"  value expected  "+expRMA7);
		System.out.println("********* Row8        : "+actRMA8  +"  value expected  "+expRMA8);
		System.out.println("********* Row9        : "+actRMA9  +"  value expected  "+expRMA9);
		System.out.println("********* Row10       : "+actRMA10 +"  value expected  "+expRMA10);
		System.out.println("********* Row11       : "+actRMA11 +"  value expected  "+expRMA11);
		System.out.println("********* Row12       : "+actRMA12 +"  value expected  "+expRMA12);
		
		System.out.println("RMAPopupDisplayingAsExpected  : "+actRMAPopupDisplayingAsExpected +"  Value Expected  "+expRMAPopupDisplayingAsExpected);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
		RMAPopupOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));

		String actRMAColumn=enter_RMA.getAttribute("value");

		String expRMAColumn=excelReader.getCellData(xlSheetName, 195, 6);
		excelReader.setCellData(xlfile, xlSheetName, 195, 7, actRMAColumn);

		System.out.println("******** RMA Column: "+actRMAColumn+"  value expected  "+expRMAColumn);

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 4");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 197, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 198, 6));
		
		String expMessage = expMsg.toString();

		excelReader.setCellData(xlfile, xlSheetName, 197, 7, actMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		
String expMessage1 = excelReader.getCellData(xlSheetName, 197, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 198, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 197, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))


		if(actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) 
				&& actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)	
				&& actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10) && actRMA11.equalsIgnoreCase(expRMA11) && actRMA12.equalsIgnoreCase(expRMA12)
				&& actRMAColumn.equalsIgnoreCase(expRMAColumn) && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 194, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 194, 8, resFail);
			return false;
		}
	}










	public boolean checkEditingAndVerifingDataOfSavedOpeningStocksVoucher4AddingRow2WithRMAItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 199, 6);
		excelReader.setCellData(xlfile, xlSheetName, 199, 7, actDocNo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		String actWarehouse=warehouseTxt.getAttribute("value");
		String expWarehouse=excelReader.getCellData(xlSheetName, 200, 6);
		excelReader.setCellData(xlfile, xlSheetName, 200, 7, actWarehouse);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actItem=select1stRow_1stColumn.getText();
		String expItem=excelReader.getCellData(xlSheetName, 201, 6);
		excelReader.setCellData(xlfile, xlSheetName, 201, 7, actItem);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
		String actUnits=select1stRow_2ndColumn.getText();
		String expUnits=excelReader.getCellData(xlSheetName, 202, 6);
		excelReader.setCellData(xlfile, xlSheetName, 202, 7, actUnits);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		String actQuantity=select1stRow_3rdColumn.getText();
		String expQuantity=excelReader.getCellData(xlSheetName, 203, 6);
		excelReader.setCellData(xlfile, xlSheetName, 203, 7, actQuantity);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		String actRate=select1stRow_4thColumn.getText();
		String expRate=excelReader.getCellData(xlSheetName, 204, 6);
		excelReader.setCellData(xlfile, xlSheetName, 204, 7, actRate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		String actGross=select1stRow_5thColumn.getText();
		String expGross=excelReader.getCellData(xlSheetName, 205, 6);
		excelReader.setCellData(xlfile, xlSheetName, 205, 7, actGross);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		String actRMA=select1stRow_9thColumn.getText();
		String expRMA=excelReader.getCellData(xlSheetName, 206, 6);
		excelReader.setCellData(xlfile, xlSheetName, 206, 7, actRMA);
		
		select1stRow_9thColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);

		Thread.sleep(2000);

		//actual
		String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
		String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
		String actRMA3=RMAPopupTableRow3Column1.getAttribute("data-value");
		String actRMA4=RMAPopupTableRow4Column1.getAttribute("data-value");
		String actRMA5=RMAPopupTableRow5Column1.getAttribute("data-value");
		String actRMA6=RMAPopupTableRow6Column1.getAttribute("data-value");
		String actRMA7=RMAPopupTableRow7Column1.getAttribute("data-value");
		String actRMA8=RMAPopupTableRow8Column1.getAttribute("data-value");
		String actRMA9=RMAPopupTableRow9Column1.getAttribute("data-value");
		String actRMA10=RMAPopupTableRow10Column1.getAttribute("data-value");
		String actRMA11=RMAPopupTableRow11Column1.getAttribute("data-value");
		String actRMA12=RMAPopupTableRow12Column1.getAttribute("data-value");

		//expected
		String expRMA1="RMA#R1OSV1";
		String expRMA2="RMA#R1OSV2";
		String expRMA3="RMA#R1OSV3";
		String expRMA4="RMA#R1OSV4";
		String expRMA5="RMA#R1OSV5";
		String expRMA6="RMA#R1OSV6";
		String expRMA7="RMA#R1OSV7";
		String expRMA8="RMA#R1OSV8";
		String expRMA9="RMA#R1OSV9";
		String expRMA10="RMA#R1OSV10";
		String expRMA11="RMA#R1OSV11";
		String expRMA12="RMA#R1OSV12";
		
		boolean Result = actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) &&
						 actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8) &&
						 actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10) && actRMA11.equalsIgnoreCase(expRMA11) && actRMA12.equalsIgnoreCase(expRMA12);
		
		String actRMAPopupDisplayingAsExpected = Boolean.toString(Result);
		String expRMAPopupDisplayingAsExpected = excelReader.getCellData(xlSheetName, 207, 6);
		excelReader.setCellData(xlfile, xlSheetName, 207, 7, actRMAPopupDisplayingAsExpected);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupCancelBtn));
		RMAPopupCancelBtn.click();

		System.out.println("**********************checkEditingAndVerifingDataOfSavedOpeningStocksVoucher4AddingRow2WithRMAItem*********************");

		System.out.println("*********** Voucher No : "+actDocNo      +"  value expected  "+expDocNo);
		System.out.println("*********** Warehouse  : "+actWarehouse  +"  value expected  "+expWarehouse);
		System.out.println("*********** Item       : "+actItem       +"  value expected  "+expItem);
		System.out.println("*********** Units      : "+actUnits      +"  value expected  "+expUnits);
		System.out.println("*********** Quantity   : "+actQuantity   +"  value expected  "+expQuantity);
		System.out.println("*********** Rate       : "+actRate       +"  value expected  "+expRate);
		System.out.println("*********** Gross      : "+actGross      +"  value expected  "+expGross);
		System.out.println(actRMA);
		System.out.println(expRMA);
		System.out.println("***********************************ROW1****************************************");
		System.out.println("********* Row1        : "+actRMA1  +"  value expected  "+expRMA1);
		System.out.println("********* Row2        : "+actRMA2  +"  value expected  "+expRMA2);
		System.out.println("********* Row3        : "+actRMA3  +"  value expected  "+expRMA3);
		System.out.println("********* Row4        : "+actRMA4  +"  value expected  "+expRMA4);
		System.out.println("********* Row5        : "+actRMA5  +"  value expected  "+expRMA5);
		System.out.println("********* Row6        : "+actRMA6  +"  value expected  "+expRMA6);
		System.out.println("********* Row7        : "+actRMA7  +"  value expected  "+expRMA7);
		System.out.println("********* Row8        : "+actRMA8  +"  value expected  "+expRMA8);
		System.out.println("********* Row9        : "+actRMA9  +"  value expected  "+expRMA9);
		System.out.println("********* Row10       : "+actRMA10 +"  value expected  "+expRMA10);
		System.out.println("********* Row11       : "+actRMA11 +"  value expected  "+expRMA11);
		System.out.println("********* Row12       : "+actRMA12 +"  value expected  "+expRMA12);


		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
			select2ndRow_1stColumn.click();*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 199, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 200, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 201, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
		rmaSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 202, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
		rmaAddBtn.click();

		Thread.sleep(2000);

		//actual
		String actR2RMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
		String actR2RMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
		String actR2RMA3=RMAPopupTableRow3Column1.getAttribute("data-value");
		String actR2RMA4=RMAPopupTableRow4Column1.getAttribute("data-value");
		String actR2RMA5=RMAPopupTableRow5Column1.getAttribute("data-value");
		String actR2RMA6=RMAPopupTableRow6Column1.getAttribute("data-value");
		String actR2RMA7=RMAPopupTableRow7Column1.getAttribute("data-value");
		String actR2RMA8=RMAPopupTableRow8Column1.getAttribute("data-value");
		String actR2RMA9=RMAPopupTableRow9Column1.getAttribute("data-value");
		String actR2RMA10=RMAPopupTableRow10Column1.getAttribute("data-value");
		String actR2RMA11=RMAPopupTableRow11Column1.getAttribute("data-value");
		String actR2RMA12=RMAPopupTableRow12Column1.getAttribute("data-value");

		//expected
		String expR2RMA1="RMA#R2OSV1";
		String expR2RMA2="RMA#R2OSV2";
		String expR2RMA3="RMA#R2OSV3";
		String expR2RMA4="RMA#R2OSV4";
		String expR2RMA5="RMA#R2OSV5";
		String expR2RMA6="RMA#R2OSV6";
		String expR2RMA7="RMA#R2OSV7";
		String expR2RMA8="RMA#R2OSV8";
		String expR2RMA9="RMA#R2OSV9";
		String expR2RMA10="RMA#R2OSV10";
		String expR2RMA11="RMA#R2OSV11";
		String expR2RMA12="RMA#R2OSV12";
		
		
		boolean R2Result = actR2RMA1.equalsIgnoreCase(expR2RMA1) && actR2RMA2.equalsIgnoreCase(expR2RMA2) && actR2RMA3.equalsIgnoreCase(expR2RMA3) && actR2RMA4.equalsIgnoreCase(expR2RMA4) &&
				 actR2RMA5.equalsIgnoreCase(expR2RMA5) && actR2RMA6.equalsIgnoreCase(expR2RMA6) && actR2RMA7.equalsIgnoreCase(expR2RMA7) && actR2RMA8.equalsIgnoreCase(expR2RMA8) &&
				 actR2RMA9.equalsIgnoreCase(expR2RMA9) && actR2RMA10.equalsIgnoreCase(expR2RMA10) && actR2RMA11.equalsIgnoreCase(expR2RMA11) && actR2RMA12.equalsIgnoreCase(expR2RMA12);

		String actR2RMAPopupDisplayingAsExpected = Boolean.toString(R2Result);
		String expR2RMAPopupDisplayingAsExpected = excelReader.getCellData(xlSheetName, 208, 6);
		excelReader.setCellData(xlfile, xlSheetName, 208, 7, actR2RMAPopupDisplayingAsExpected);


		System.out.println("***********************************ROW2****************************************");
		System.out.println("********* Row1        : "+actR2RMA1  +"  value expected  "+expR2RMA1);
		System.out.println("********* Row2        : "+actR2RMA2  +"  value expected  "+expR2RMA2);
		System.out.println("********* Row3        : "+actR2RMA3  +"  value expected  "+expR2RMA3);
		System.out.println("********* Row4        : "+actR2RMA4  +"  value expected  "+expR2RMA4);
		System.out.println("********* Row5        : "+actR2RMA5  +"  value expected  "+expR2RMA5);
		System.out.println("********* Row6        : "+actR2RMA6  +"  value expected  "+expR2RMA6);
		System.out.println("********* Row7        : "+actR2RMA7  +"  value expected  "+expR2RMA7);
		System.out.println("********* Row8        : "+actR2RMA8  +"  value expected  "+expR2RMA8);
		System.out.println("********* Row9        : "+actR2RMA9  +"  value expected  "+expR2RMA9);
		System.out.println("********* Row10       : "+actR2RMA10 +"  value expected  "+expR2RMA10);
		System.out.println("********* Row11       : "+actR2RMA11 +"  value expected  "+expR2RMA11);
		System.out.println("********* Row12       : "+actR2RMA12 +"  value expected  "+expR2RMA12);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
		RMAPopupOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));

		String actRMAColumn=enter_RMA.getAttribute("value");
		String expRMAColumn=excelReader.getCellData(xlSheetName, 209, 6);
		excelReader.setCellData(xlfile, xlSheetName, 209, 7, actRMAColumn);
		
		System.out.println(actRMAColumn);
		System.out.println(expRMAColumn);

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 4");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 210, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 211, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 210, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 210, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 211, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 210, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))


		if (actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) && actItem.equalsIgnoreCase(expItem) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate)
				&& actGross.equalsIgnoreCase(expGross)  && actRMA.equalsIgnoreCase(expRMA)

				&& actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) 
				&& actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)	
				&& actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10) && actRMA11.equalsIgnoreCase(expRMA11) && actRMA12.equalsIgnoreCase(expRMA12)

				&& actR2RMA1.equalsIgnoreCase(expR2RMA1) && actR2RMA2.equalsIgnoreCase(expR2RMA2) && actR2RMA3.equalsIgnoreCase(expR2RMA3) && actR2RMA4.equalsIgnoreCase(expR2RMA4) 
				&& actR2RMA5.equalsIgnoreCase(expR2RMA5) && actR2RMA6.equalsIgnoreCase(expR2RMA6) && actR2RMA7.equalsIgnoreCase(expR2RMA7) && actR2RMA8.equalsIgnoreCase(expR2RMA8)	
				&& actR2RMA9.equalsIgnoreCase(expR2RMA9) && actR2RMA10.equalsIgnoreCase(expR2RMA10) && actR2RMA11.equalsIgnoreCase(expR2RMA11) && actR2RMA12.equalsIgnoreCase(expR2RMA12)
				&& actRMAColumn.equalsIgnoreCase(expRMAColumn) 

				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 199, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 199, 8, resFail);
			return false;
		}
	}







	public boolean checkEditingAndVerifingDataOfSavedOpeningStocksVoucher4() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 212, 6);
		excelReader.setCellData(xlfile, xlSheetName, 212, 7, actDocNo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		String actWarehouse=warehouseTxt.getAttribute("value");
		String expWarehouse=excelReader.getCellData(xlSheetName, 213, 6);
		excelReader.setCellData(xlfile, xlSheetName, 213, 7, actWarehouse);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actItem=select1stRow_1stColumn.getText();
		String expItem=excelReader.getCellData(xlSheetName, 214, 6);
		excelReader.setCellData(xlfile, xlSheetName, 214, 7, actItem);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
		String actUnits=select1stRow_2ndColumn.getText();
		String expUnits=excelReader.getCellData(xlSheetName, 215, 6);
		excelReader.setCellData(xlfile, xlSheetName, 215, 7, actUnits);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		String actQuantity=select1stRow_3rdColumn.getText();
		String expQuantity=excelReader.getCellData(xlSheetName, 216, 6);
		excelReader.setCellData(xlfile, xlSheetName, 216, 7, actQuantity);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		String actRate=select1stRow_4thColumn.getText();
		String expRate=excelReader.getCellData(xlSheetName, 217, 6);
		excelReader.setCellData(xlfile, xlSheetName, 217, 7, actRate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		String actGross=select1stRow_5thColumn.getText();
		String expGross=excelReader.getCellData(xlSheetName, 218, 6);
		excelReader.setCellData(xlfile, xlSheetName, 218, 7, actGross);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		String actRMA=select1stRow_9thColumn.getText();
		String expRMA=excelReader.getCellData(xlSheetName, 219, 6);
		excelReader.setCellData(xlfile, xlSheetName, 219, 7, actRMA);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actItem2=select2ndRow_1stColumn.getText();
		String expItem2=excelReader.getCellData(xlSheetName, 220, 6);
		excelReader.setCellData(xlfile, xlSheetName, 220, 7, actItem2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_2ndColumn));
		String actUnits2=select2ndRow_2ndColumn.getText();
		String expUnits2=excelReader.getCellData(xlSheetName, 221, 6);
		excelReader.setCellData(xlfile, xlSheetName, 221, 7, actUnits2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
		String actQuantity2=select2ndRow_3rdColumn.getText();
		String expQuantity2=excelReader.getCellData(xlSheetName, 222, 6);
		excelReader.setCellData(xlfile, xlSheetName, 222, 7, actQuantity2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
		String actRate2=select2ndRow_4thColumn.getText();
		String expRate2=excelReader.getCellData(xlSheetName, 223, 6);
		excelReader.setCellData(xlfile, xlSheetName, 223, 7, actRate2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
		String actGross2=select2ndRow_5thColumn.getText();
		String expGross2=excelReader.getCellData(xlSheetName, 224, 6);
		excelReader.setCellData(xlfile, xlSheetName, 224, 7, actGross2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
		String actRMA2=select2ndRow_9thColumn.getText();
		String expRMA2=excelReader.getCellData(xlSheetName, 225, 6);
		excelReader.setCellData(xlfile, xlSheetName, 225, 7, actRMA2);

		System.out.println("**********************checkEditingAndVerifingDataOfSavedOpeningStocksVoucher4*********************");

		System.out.println("*********** Voucher No : "+actDocNo      +"  value expected  "+expDocNo);
		System.out.println("*********** Warehouse  : "+actWarehouse  +"  value expected  "+expWarehouse);
		System.out.println("*********** Item       : "+actItem       +"  value expected  "+expItem);
		System.out.println("*********** Units      : "+actUnits      +"  value expected  "+expUnits);
		System.out.println("*********** Quantity   : "+actQuantity   +"  value expected  "+expQuantity);
		System.out.println("*********** Rate       : "+actRate       +"  value expected  "+expRate);
		System.out.println("*********** Gross      : "+actGross      +"  value expected  "+expGross);
		System.out.println(actRMA);
		System.out.println(expRMA);
		System.out.println("*********** Item2       : "+actItem2       +"  value expected  "+expItem2);
		System.out.println("*********** Units2      : "+actUnits2      +"  value expected  "+expUnits2);
		System.out.println("*********** Quantity2   : "+actQuantity2   +"  value expected  "+expQuantity2);
		System.out.println("*********** Rate2       : "+actRate2       +"  value expected  "+expRate2);
		System.out.println("*********** Gross2      : "+actGross2      +"  value expected  "+expGross2);
		System.out.println(actRMA2);
		System.out.println(expRMA2);

		if (actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) && actItem.equalsIgnoreCase(expItem) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate)
				&& actGross.equalsIgnoreCase(expGross) && actRMA.equalsIgnoreCase(expRMA)
				&& actItem2.equalsIgnoreCase(expItem2) && actUnits2.equalsIgnoreCase(expUnits2) && actQuantity2.equalsIgnoreCase(expQuantity2) 
				&& actRate2.equalsIgnoreCase(expRate2) && actGross2.equalsIgnoreCase(expGross2) && actRMA2.equalsIgnoreCase(expRMA2)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 212, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 212, 8, resFail);
			return false;
		}
	}







	public boolean checkSavingOpeningStockVoucher5WithRMAItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
		new_newBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 226, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 227, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 228, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 229, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
		rmaSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 230, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
		rmaAddBtn.click();

		Thread.sleep(2000);

		//actual
		String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
		String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
		String actRMA3=RMAPopupTableRow3Column1.getAttribute("data-value");
		String actRMA4=RMAPopupTableRow4Column1.getAttribute("data-value");
		String actRMA5=RMAPopupTableRow5Column1.getAttribute("data-value");
		String actRMA6=RMAPopupTableRow6Column1.getAttribute("data-value");
		String actRMA7=RMAPopupTableRow7Column1.getAttribute("data-value");
		String actRMA8=RMAPopupTableRow8Column1.getAttribute("data-value");
		String actRMA9=RMAPopupTableRow9Column1.getAttribute("data-value");
		String actRMA10=RMAPopupTableRow10Column1.getAttribute("data-value");

		//expected
		String expRMA1="RMA#V2R1OSV1";
		String expRMA2="RMA#V2R1OSV2";
		String expRMA3="RMA#V2R1OSV3";
		String expRMA4="RMA#V2R1OSV4";
		String expRMA5="RMA#V2R1OSV5";
		String expRMA6="RMA#V2R1OSV6";
		String expRMA7="RMA#V2R1OSV7";
		String expRMA8="RMA#V2R1OSV8";
		String expRMA9="RMA#V2R1OSV9";
		String expRMA10="RMA#V2R1OSV10";

		
		boolean Result = actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) &&
				 actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8) &&
				 actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10);

		String actRMAPopupDisplayingAsExpected = Boolean.toString(Result);
		String expRMAPopupDisplayingAsExpected = excelReader.getCellData(xlSheetName, 226, 6);
		excelReader.setCellData(xlfile, xlSheetName, 226, 7, actRMAPopupDisplayingAsExpected);
		

		System.out.println("********* Row1        : "+actRMA1  +"  value expected  "+expRMA1);
		System.out.println("********* Row2        : "+actRMA2  +"  value expected  "+expRMA2);
		System.out.println("********* Row3        : "+actRMA3  +"  value expected  "+expRMA3);
		System.out.println("********* Row4        : "+actRMA4  +"  value expected  "+expRMA4);
		System.out.println("********* Row5        : "+actRMA5  +"  value expected  "+expRMA5);
		System.out.println("********* Row6        : "+actRMA6  +"  value expected  "+expRMA6);
		System.out.println("********* Row7        : "+actRMA7  +"  value expected  "+expRMA7);
		System.out.println("********* Row8        : "+actRMA8  +"  value expected  "+expRMA8);
		System.out.println("********* Row9        : "+actRMA9  +"  value expected  "+expRMA9);
		System.out.println("********* Row10       : "+actRMA10 +"  value expected  "+expRMA10);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
		RMAPopupOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));

		String actRMAColumn=enter_RMA.getAttribute("value");

		String expRMAColumn=excelReader.getCellData(xlSheetName, 227, 6);
		excelReader.setCellData(xlfile, xlSheetName, 227, 7, actRMAColumn);

		System.out.println(actRMAColumn);
		System.out.println(expRMAColumn);

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 5");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 229, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 230, 6));
		
		String expMessage = expMsg.toString();

		excelReader.setCellData(xlfile, xlSheetName, 229, 7, actMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 229, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 230, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 229, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		if(actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) 
				&& actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)	
				&& actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10) && actRMAColumn.equalsIgnoreCase(expRMAColumn)

				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 226, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 226, 8, resFail);
			return false;
		}
	}








	public boolean checkVerifingDataOfSavedOpeningStocksVoucher5AndDeleting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 231, 6);
		excelReader.setCellData(xlfile, xlSheetName, 231, 7, actDocNo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		String actWarehouse=warehouseTxt.getAttribute("value");
		String expWarehouse=excelReader.getCellData(xlSheetName, 232, 6);
		excelReader.setCellData(xlfile, xlSheetName, 232, 7, actWarehouse);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actItem=select1stRow_1stColumn.getText();
		String expItem=excelReader.getCellData(xlSheetName, 233, 6);
		excelReader.setCellData(xlfile, xlSheetName, 233, 7, actItem);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
		String actUnits=select1stRow_2ndColumn.getText();
		String expUnits=excelReader.getCellData(xlSheetName, 234, 6);
		excelReader.setCellData(xlfile, xlSheetName, 234, 7, actUnits);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		String actQuantity=select1stRow_3rdColumn.getText();
		String expQuantity=excelReader.getCellData(xlSheetName, 235, 6);
		excelReader.setCellData(xlfile, xlSheetName, 235, 7, actQuantity);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		String actRate=select1stRow_4thColumn.getText();
		String expRate=excelReader.getCellData(xlSheetName, 236, 6);
		excelReader.setCellData(xlfile, xlSheetName, 236, 7, actRate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		String actGross=select1stRow_5thColumn.getText();
		String expGross=excelReader.getCellData(xlSheetName, 237, 6);
		excelReader.setCellData(xlfile, xlSheetName, 237, 7, actGross);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		String actRMA=select1stRow_9thColumn.getText();
		String expRMA=excelReader.getCellData(xlSheetName, 238, 6);
		excelReader.setCellData(xlfile, xlSheetName, 238, 7, actRMA);

		System.out.println("**********************checkVerifingDataOfSavedOpeningStocksVoucher3AndDeleting*********************");

		System.out.println("*********** Voucher No : "+actDocNo      +"  value expected  "+expDocNo);
		System.out.println("*********** Warehouse  : "+actWarehouse  +"  value expected  "+expWarehouse);
		System.out.println("*********** Item       : "+actItem       +"  value expected  "+expItem);
		System.out.println("*********** Units      : "+actUnits      +"  value expected  "+expUnits);
		System.out.println("*********** Quantity   : "+actQuantity   +"  value expected  "+expQuantity);
		System.out.println("*********** Rate       : "+actRate       +"  value expected  "+expRate);
		System.out.println("*********** Gross      : "+actGross      +"  value expected  "+expGross);
		System.out.println(actRMA);
		System.out.println(expRMA);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		String expMessage=excelReader.getCellData(xlSheetName, 239, 6);
		String actMessage=checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 239, 7, actMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();

		Thread.sleep(1000);

		if (actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) && actItem.equalsIgnoreCase(expItem) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate)
				&& actGross.equalsIgnoreCase(expGross) && actRMA.equalsIgnoreCase(expRMA)
				&& actMessage.equalsIgnoreCase(expMessage)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 231, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 231, 8, resFail);
			return false;
		}
	}









	public boolean checkVerifingDetailsOfSavedOpeningStocksVouchersInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat todayDate = new SimpleDateFormat("dd/MM/yyyy");
		String currentDate = todayDate.format(cal.getTime());
		System.out.println("--------------Date Before Adding:: " + currentDate);

		int count = voucherHomeBodyList.size();

		ArrayList<String>  actVouchersList= new ArrayList<String>(); 

		for (int i = 0; i < count; i++) 
		{
			String data = voucherHomeBodyList.get(i).getText();

			if (data.isEmpty()==false && i!=9 && i!=10 && i!=24 && i!=25 && i!=39 && i!=40 && i!=54 && i!=55) 
			{
				System.out.println(i+" "+data);

				actVouchersList.add(data);
			}
		}


		ArrayList<String>  expVouchersList= new ArrayList<String>(); 

		expVouchersList.add(excelReader.getCellData(xlSheetName, 240, 6));
		expVouchersList.add("01/01/2020");
		expVouchersList.add(excelReader.getCellData(xlSheetName, 241, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 242, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 243, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 244, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 245, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 246, 6));
		expVouchersList.add("01/01/2020");
		expVouchersList.add(excelReader.getCellData(xlSheetName, 247, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 248, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 249, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 250, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 251, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 252, 6));
		expVouchersList.add("01/01/2020");
		expVouchersList.add(excelReader.getCellData(xlSheetName, 253, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 254, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 255, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 256, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 257, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 258, 6));
		expVouchersList.add("01/01/2020");
		expVouchersList.add(excelReader.getCellData(xlSheetName, 259, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 260, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 261, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 262, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 263, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 264, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 265, 6));

		System.out.println("**********************checkVerifingDetailsOfSavedOpeningStocksVouchersInHomepage*********************");

		System.out.println(actVouchersList);
		System.out.println(expVouchersList);
		
		String actList = actVouchersList.toString();
		String expList = expVouchersList.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 240, 7, actList);

		if (actList.equalsIgnoreCase(expList)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 240, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 240, 8, resFail);
			return false;
		}
	}





	

	public boolean checkSaveInCreateViewOfHomePageOfOpeningBalance() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createViewBtn));
		createViewBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewNameTxtInCreateView));
		viewNameTxtInCreateView.click();
		viewNameTxtInCreateView.sendKeys(excelReader.getCellData(xlSheetName, 266, 5));

		Select Authorization = new Select(authorizationInCreateView);
		Authorization.selectByVisibleText(excelReader.getCellData(xlSheetName, 267, 5));

		Select CheckStatus = new Select(checkSatusInCreateView);
		CheckStatus.selectByVisibleText(excelReader.getCellData(xlSheetName, 268, 5));

		Select QC = new Select(qcInCreateView);
		QC.selectByVisibleText(excelReader.getCellData(xlSheetName, 269, 5));

		Select SuspendStatus = new Select(suspendStatusInCreateView);
		SuspendStatus.selectByVisibleText(excelReader.getCellData(xlSheetName, 270, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveOptionInCreateView));
		saveOptionInCreateView.click();

		String expSaveViewMessage = excelReader.getCellData(xlSheetName, 266, 6);
		String actSaveViewMessage = checkValidationMessage(expSaveViewMessage);
		excelReader.setCellData(xlfile, xlSheetName, 266, 7, actSaveViewMessage);
		
		/*LogoutandLoginwithSU();
		Thread.sleep(1000);
		*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();
		
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksVoucher));
		openingStocksVoucher.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		
		Thread.sleep(2000);

		String actcreateViewNameDisplayInHomePage = createViewNameDisplayInHomePage.getText();
		String expcreateViewNameDisplayInHomePage = excelReader.getCellData(xlSheetName, 267, 6);
		excelReader.setCellData(xlfile, xlSheetName, 267, 7, actcreateViewNameDisplayInHomePage);
		
		System.out.println("View Display In Home page Value Actual    : " + actcreateViewNameDisplayInHomePage   +     "	 Value Expected : " + expcreateViewNameDisplayInHomePage);

		if(actSaveViewMessage.equalsIgnoreCase(expSaveViewMessage) && actcreateViewNameDisplayInHomePage.equalsIgnoreCase(expcreateViewNameDisplayInHomePage) )
		{
			
			excelReader.setCellData(xlfile, xlSheetName, 266, 8, resPass);
			return true;
			}
			
				
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 266, 8, resFail);
			return false;
		} 
	}


	






	public boolean checkSavingAndSuspendingOpeningStockVoucher5WithStockItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen Opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 271, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 272, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 273, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 274, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Saving in background.");
		expMessage.add("Voucher saved successfully : 5");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 271, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 272, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 271, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 271, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 272, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 271, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
		new_SuspendBtn.click();

		String expSuspendMessage1=excelReader.getCellData(xlSheetName, 273, 6);
		String expSuspendMessage2=excelReader.getCellData(xlSheetName, 274, 6);

		String actSuspendMessage=checkValidationMessage(expSuspendMessage1);
		
		excelReader.setCellData(xlfile, xlSheetName, 273, 7, actSuspendMessage);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)&& actSuspendMessage.startsWith(expSuspendMessage1) && actSuspendMessage.endsWith(expSuspendMessage2))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 271, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 271, 8, resFail);
			return false;
		}
	}








	public boolean checkVerifingDetailsOfSuspendedOpeningStocksVouchersInCreatedView() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createViewNameDisplayInHomePage));
		createViewNameDisplayInHomePage.click();

		Thread.sleep(2000);

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat todayDate = new SimpleDateFormat("dd/MM/yyyy");
		String currentDate = todayDate.format(cal.getTime());
		System.out.println("--------------Date Before Adding:: " + currentDate);

		int count = voucherHomeBodyList.size();

		ArrayList<String>  actVouchersList= new ArrayList<String>(); 

		for (int i = 0; i < count; i++) 
		{
			String data = voucherHomeBodyList.get(i).getText();

			if (data.isEmpty()==false && i!=9 && i!=10) 
			{
				System.out.println(i+" "+data);

				actVouchersList.add(data);
			}
		}


		ArrayList<String>  expVouchersList= new ArrayList<String>(); 

		expVouchersList.add(excelReader.getCellData(xlSheetName, 275, 6));
		expVouchersList.add("01/01/2020");
		expVouchersList.add(excelReader.getCellData(xlSheetName, 276, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 277, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 278, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 279, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 280, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 281, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 282, 6));

		System.out.println("**********************checkVerifingDetailsOfSuspendedOpeningStocksVouchersInCreatedView*********************");

		System.out.println(actVouchersList);
		System.out.println(expVouchersList);
		
		String actList = actVouchersList.toString();
		String expList = expVouchersList.toString();

		excelReader.setCellData(xlfile, xlSheetName, 275, 7, actList);
		
		if (actVouchersList.equals(expVouchersList)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 275, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 275, 8, resFail);
			return false;
		}
	}





	


	public boolean checkResavingOpeningStockVoucher5AndSuspendingFromHomepage() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 283, 5);

			/*String data="4";*/
			if(voucherGridDocNo.get(i).getText().equalsIgnoreCase(data))
			{
				getAction().doubleClick(voucherGridDocNo.get(i)).build().perform();
				break;
			}
		}

		Thread.sleep(2000);

		checkValidationMessage("Voucher loaded successfully");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Saving in background.");
		expMessage.add("Voucher saved successfully : 5");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 283, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 284, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 283, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 283, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 284, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 283, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(allVouchersLink));
		allVouchersLink.click();

		Thread.sleep(2000);

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 284, 5);

			if(voucherGridDocNo.get(i).getText().equalsIgnoreCase(data))
			{
				if (voucherGridIndexChkBox.get(i).isSelected()==false) 
				{
					voucherGridIndexChkBox.get(i).click();
					break;
				}
			}
		}

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(suspendBtn));
		suspendBtn.click();

		String expSuspendMessage = excelReader.getCellData(xlSheetName, 285, 6);
		String actSuspendMessage = checkValidationMessage(expSuspendMessage);

		Thread.sleep(2000);

		String actSuspendStatus = voucherHomeRow1SuspendedStatus.getText();
		String expSuspendStatus = excelReader.getCellData(xlSheetName, 286, 6);
		excelReader.setCellData(xlfile, xlSheetName, 286, 7, actSuspendStatus);

		System.out.println("Suspended Status : "+actSuspendStatus+"  Value Expected  "+expSuspendStatus);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2) && actSuspendMessage.equalsIgnoreCase(expSuspendMessage)
				&& actSuspendStatus.equalsIgnoreCase(expSuspendStatus))
		{
			excelReader.setCellData(xlfile, xlSheetName, 283, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 283, 8, resFail);
			return false;
		}
	}



	



	public boolean checkDeleteOptionInOpeningStocksHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 287, 5);

		/*	String data="4";*/
			if(voucherGridDocNo.get(i).getText().equalsIgnoreCase(data))
			{
				if (voucherGridIndexChkBox.get(i).isSelected()==false) 
				{
					voucherGridIndexChkBox.get(i).click();
					break;
				}
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHomeDeleteBtn));
		voucherHomeDeleteBtn.click();

		getWaitForAlert();

		String actAlertMessage=getAlert().getText();

		String expAlertMessage=excelReader.getCellData(xlSheetName, 287, 6);
				
		excelReader.setCellData(xlfile, xlSheetName, 287, 7, actAlertMessage);


		System.out.println("Alert Message on Deleteing the Voucher :  "+actAlertMessage +"  "+expAlertMessage);

		Thread.sleep(3000);

		getAlert().accept();

		String expMessage = excelReader.getCellData(xlSheetName, 288, 6);
		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 288, 7, actMessage);

		System.out.println("Message : "+actMessage+" Value Expected : "+expMessage);

		if(actAlertMessage.contains(expAlertMessage) && actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 287, 8, resPass);
			return true;
		}

		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 287, 8, resFail);
			return false;
		}
	}




/////
	public boolean checkSaveCreateViewOfHomePageOfOpeningBalance() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();
		
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksVoucher));
		openingStocksVoucher.click();

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createViewBtn));
		createViewBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewNameTxtInCreateView));
		viewNameTxtInCreateView.click();
		viewNameTxtInCreateView.sendKeys(excelReader.getCellData(xlSheetName, 266, 5));

		Select Authorization = new Select(authorizationInCreateView);
		Authorization.selectByVisibleText(excelReader.getCellData(xlSheetName, 267, 5));

		Select CheckStatus = new Select(checkSatusInCreateView);
		CheckStatus.selectByVisibleText(excelReader.getCellData(xlSheetName, 268, 5));

		Select QC = new Select(qcInCreateView);
		QC.selectByVisibleText(excelReader.getCellData(xlSheetName, 269, 5));

		Select SuspendStatus = new Select(suspendStatusInCreateView);
		SuspendStatus.selectByVisibleText(excelReader.getCellData(xlSheetName, 270, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveOptionInCreateView));
		saveOptionInCreateView.click();

		String expSaveViewMessage = "view saved successfully";
		String actSaveViewMessage = checkValidationMessage(expSaveViewMessage);
		
		
		

		String actcreateViewNameDisplayInHomePage = createViewNameDisplayInHomePage.getText();
		String expcreateViewNameDisplayInHomePage = "";
		
		System.out.println("View Display In Home page Value Actual    : " + actcreateViewNameDisplayInHomePage   +     "	 Value Expected : " + expcreateViewNameDisplayInHomePage);

		if(actSaveViewMessage.equalsIgnoreCase(expSaveViewMessage) && actcreateViewNameDisplayInHomePage.equalsIgnoreCase(expcreateViewNameDisplayInHomePage) )
		{
					
			return true;
			
		}
			
				
		else
		{
			return false;
		} 
	}
	
	@FindBy(xpath="//a[@id='btnEditView']")
	public static WebElement editViewBtn;
	
	@FindBy(xpath="//input[@id='txtViewName']")
	public static WebElement editViewNameTxt;
	
	@FindBy(xpath="//select[@id='cmbDataSet']")
	public static List<WebElement> editDataSetTxt;
	
	@FindBy(xpath="//select[@id='cmbCheckStatus']")
	public static WebElement editCheckStatusTxt;

	
	@FindBy(xpath="//select[@id='cmbQC']")
	public static List<WebElement> editQCTxt;

	
	@FindBy(xpath="//select[@id='Suspendstatus']")
	public static List<WebElement> editSuspendStatusTxt;
	
	@FindBy(xpath="(//*[@id='dvCreateEditView']//span/i)[1]")
	public static WebElement editViewCloseBtn;

	
public boolean checkEditViewofHomePageOpeningStocks() throws InterruptedException
{
	
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createViewNameDisplayInHomePage));
	getAction().moveToElement(createViewNameDisplayInHomePage).click().build().perform();
	
	//createViewNameDisplayInHomePage.click();
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editViewBtn));
	editViewBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editViewNameTxt));
	String actViewNameTxt=editViewNameTxt.getText();
	String expViewNameTxt="";
	
	//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable((By) editCheckStatusTxt));
	Select s1=new Select(editCheckStatusTxt);
	WebElement x=s1.getFirstSelectedOption();
	String txt=s1.getFirstSelectedOption().toString();
	System.out.println(txt);
	
	String actCheckStatusTxt=s1.getFirstSelectedOption().toString();
	String expCheckStatusTxt="";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable((By) editQCTxt));
	String actQCStatusTxt=editQCTxt.get(2).toString();
	String exQCStatusTxt="";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable((By) editSuspendStatusTxt));
	String actSuspendStatusTxt=editSuspendStatusTxt.get(2).toString();
	String exSuspendStatusTxt="";
	
	
	if(actViewNameTxt.equalsIgnoreCase(expViewNameTxt) && actCheckStatusTxt.equalsIgnoreCase(expCheckStatusTxt)
			&&actQCStatusTxt.equalsIgnoreCase(exQCStatusTxt)&&actSuspendStatusTxt.equalsIgnoreCase(exSuspendStatusTxt))
	{
		return true;
	}
	else
	{
		return false;
	}
}


public boolean checkValidationofSavedViewAfterRefreshtheBrowser() throws InterruptedException
{
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
	inventoryMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
	inventoryTransactionsMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
	inventoryTransactionsStocksMenu.click();
	
	Thread.sleep(1000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksVoucher));
	openingStocksVoucher.click();

	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
	homePannelOpenBtn.click();
	
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createViewNameDisplayInHomePage));
	
	boolean actCreateViewDisplay=createViewNameDisplayInHomePage.isSelected();
	boolean expCreateViewDisplay=true;
	
	if(actCreateViewDisplay==expCreateViewDisplay)
	{
		return true;
	}
	else
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();
		Thread.sleep(2000);
		
		 getDriver().navigate().refresh();
		 Thread.sleep(1999);
		

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt="su";
				      
		String pawslt="su";
				      
		LoginPage.enterUserName(unamelt);

		Thread.sleep(1000);
				
		LoginPage.enterPassword(pawslt);

		 Select oSelect = new Select(companyDropDownList);
		 oSelect.selectByValue("36");
		
		 LoginPage.clickOnSignInBtn();
		
		LoginPage.reLogin("su", "su", "Automation Company");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();
		
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksVoucher));
		openingStocksVoucher.click();

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createViewNameDisplayInHomePage));
		boolean actCreatedViewDisplay=createViewNameDisplayInHomePage.isDisplayed();
		boolean expCreatedViewDisplay=true;
		if(actCreatedViewDisplay==expCreatedViewDisplay)
		{
			return true;
		}
		else
		{
			return false;
		}
		
		
		
	}
	
	
}


///
@FindBy(xpath="//*[@id='btnDeleteview']/a")
public static WebElement deleteViewBtn;



public boolean checkValidationofDeletingViewinHomePageofOpeningStocks() throws InterruptedException, InvalidFormatException, IOException
{
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
	inventoryMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
	inventoryTransactionsMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
	inventoryTransactionsStocksMenu.click();
	
	Thread.sleep(1000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksVoucher));
	openingStocksVoucher.click();

	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
	homePannelOpenBtn.click();
	
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createViewNameDisplayInHomePage));
	
	createViewNameDisplayInHomePage.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteViewBtn));
	deleteViewBtn.click();
	
	getWaitForAlert();
	getAlert().accept();
	Thread.sleep(2000);
	
	String expDeleteMsg="View deleted successfully.";
	String actDeleteMsg=checkValidationMessage(expDeleteMsg);
	
	if(actDeleteMsg.equalsIgnoreCase(expDeleteMsg))
	{
		return true;
	}
	else
	{
		return false;
	
		
	}
	
	
}

public boolean checkDeletedViewinHomePageofOpeningStocks() throws InterruptedException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
	inventoryMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
	inventoryTransactionsMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
	inventoryTransactionsStocksMenu.click();
	
	Thread.sleep(1000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksVoucher));
	openingStocksVoucher.click();

	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
	homePannelOpenBtn.click();
	
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createViewNameDisplayInHomePage));
	boolean actCreateViewDisplay=createViewNameDisplayInHomePage.isDisplayed();
	boolean expCreateViewDisplay=false;
	if(actCreateViewDisplay==expCreateViewDisplay)
	{
		return true;
	}
	else
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
	userNameDisplayLogo.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
	logoutOption.click();
	Thread.sleep(2000);
	
	 getDriver().navigate().refresh();
	 Thread.sleep(1999);
	

	LoginPage lp=new LoginPage(getDriver()); 

	String unamelt="su";
			      
	String pawslt="su";
			      
	LoginPage.enterUserName(unamelt);

	Thread.sleep(1000);
			
	LoginPage.enterPassword(pawslt);

	 Select oSelect = new Select(companyDropDownList);
	 oSelect.selectByValue("36");
	
	 LoginPage.clickOnSignInBtn();
	
	LoginPage.reLogin("su", "su", "Automation Company");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
	inventoryMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
	inventoryTransactionsMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
	inventoryTransactionsStocksMenu.click();
	
	Thread.sleep(1000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksVoucher));
	openingStocksVoucher.click();

	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
	homePannelOpenBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createViewNameDisplayInHomePage));
	boolean actCreatedViewDisplay=createViewNameDisplayInHomePage.isDisplayed();
	boolean expCreatedViewDisplay=true;
	
	if(actCreatedViewDisplay==expCreatedViewDisplay)
	{
		return true;
	}
	else
	{
		return false;
	}
	}
}


public void LogoutandLoginwithSU() throws InterruptedException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
	userNameDisplayLogo.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
	logoutOption.click();
	
	Thread.sleep(3000);
	
	 getDriver().navigate().refresh();
	 Thread.sleep(1999);
	 
	// getDriver().navigate().refresh();
	 //Thread.sleep(1999);

	LoginPage lp=new LoginPage(getDriver()); 

	String unamelt="su";
			      
	String pawslt="su";
			      
	LoginPage.enterUserName(unamelt);

	Thread.sleep(1000);
			
	LoginPage.enterPassword(pawslt);

	 Select oSelect = new Select(companyDropDownList);
	 //oSelect.selectByVisibleText(Compname); 
	 oSelect.selectByValue("36");
	
	 LoginPage.clickOnSignInBtn();
	 
	 Thread.sleep(2000);
	 LoginPage.reLogin("su", "su", "Automation Company");
	 
	
}


public static boolean checkRestoreAutomationCompany() throws InterruptedException, IOException, AWTException
{
	String actMessage=BaseEngine.restoreCompany("SanityUptoMastersMounted","Automation Company");
	String expMessage="Restore company code : 010";
	
	System.err.println("Actual Text :"	+actMessage);
	System.err.println("Expected Text :"	+expMessage);
	
	if(actMessage.equalsIgnoreCase(expMessage))
	{
		return true;
	}
	else
	{
		return false;
	}
	
}	



	public OpeningStocksPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
