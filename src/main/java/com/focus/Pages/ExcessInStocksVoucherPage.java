package com.focus.Pages;

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

public class ExcessInStocksVoucherPage extends BaseEngine 
{
	private static String xlfile = getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlSheetName = "ExcessInStocksVoucherPage";
	
	public boolean checkSavingExcessInStocksVoucherWithAllItems() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(excessesInStocksVoucher));
		excessesInStocksVoucher.click();

		Thread.sleep(2000);

		checkDeleteLinkStatus();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 8, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);



		//Enter Secound  Line for BR COgs 	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(Keys.TAB);

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

		// Thrid  Row  FIFo Cogs Item 

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
		select3rdRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 17, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 18, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
		binAutoAllocateBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		binOkBtn.click();



		// fourth Row Rma ritem

	
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_1stColumn));
		select4thRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 19, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 20, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 21, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
		rmaSerialNumberTxtField.click();
		rmaSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 22, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
		rmaAddBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
		RMAPopupOkBtn.click();


		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		Thread.sleep(4000);

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

		expMsg.add(excelReader.getCellData(xlSheetName, 8, 6));
		//expMsg.add(excelReader.getCellData(xlSheetName, 9, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actMessage);*/
		
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 8, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 9, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))


		System.out.println("************************************* checkSavingexcessInStocksVoucherWithAllItems  ********************************************");

		

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
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





	public boolean checkCopyDocumentOptionAndSavingInExcessInStocks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));
		toggleBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(copyDocumentOption));
		copyDocumentOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cd_FirstChkBox));
		cd_FirstChkBox.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cd_OkBtn));
		cd_OkBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();

		// Actual

		int voucherGridBodyListCount = voucherGridBodyList.size();

		ArrayList<String> voucherGridBodyListArray = new ArrayList<String>();

		for(int i=0;i<voucherGridBodyListCount;i++)
		{
			String data = voucherGridBodyList.get(i).getText();
			voucherGridBodyListArray.add(data);
		}

		String actvoucherGridBodyList=voucherGridBodyListArray.toString();

		String expvoucherGridBodyList=excelReader.getCellData(xlSheetName, 23, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 23, 7, actvoucherGridBodyList);

		System.out.println("voucherGridBodyList Actual   : " + actvoucherGridBodyList);
		System.out.println("voucherGridBodyList Expected : " + expvoucherGridBodyList);



		//Enter Secound  Line for BR COgs 	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(Keys.TAB);

		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 23, 5));
		enter_Batch.sendKeys(Keys.TAB);




		// Thrid  Row  FIFO Cogs Item 

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
		select3rdRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
		binAutoAllocateBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		binOkBtn.click();

		// fourth Row Rma ritem

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_1stColumn));
		select4thRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);
*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
		rmaSerialNumberTxtField.click();
		rmaSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 24, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
		rmaAddBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
		RMAPopupOkBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		Thread.sleep(4000);

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

		//expMsg.add(excelReader.getCellData(xlSheetName, 24, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 25, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);*/
		
		
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 24, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 25, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)

		if(actvoucherGridBodyList.equalsIgnoreCase(expvoucherGridBodyList) && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 23, 8, resPass);
			return true;
		}

		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 23, 8, resFail);
			return false;
		}
	}

	
	
	
	

	public boolean checkUpdatingSavedVoucherInExcessInStocks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("");

		// Actual

		int voucherGridBodyListCount = voucherGridBodyList.size();

		ArrayList<String> voucherGridBodyListArray = new ArrayList<String>();

		for(int i=0;i<voucherGridBodyListCount;i++)
		{

			String data=null;

			if(i==22)
			{
				data =excelReader.getCellData(xlSheetName, 26, 5);
			}
			else
			{
				data = voucherGridBodyList.get(i).getText();
			}

			voucherGridBodyListArray.add(data);
		}

		String actvoucherGridBodyList=voucherGridBodyListArray.toString();

		String expvoucherGridBodyList=excelReader.getCellData(xlSheetName, 26, 6);

		excelReader.setCellData(xlfile, xlSheetName, 26, 7, actvoucherGridBodyList);
		
		System.out.println("voucherGridBodyList Actual   : " + actvoucherGridBodyList);
		System.out.println("voucherGridBodyList Expected : " + expvoucherGridBodyList);

		// Std Cogs

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
		select1stRow_6thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 27, 5));
		enter_Rate.sendKeys(Keys.TAB);

		//Enter Secound  Line for BR COgs 	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_6thColumn));
		select2ndRow_6thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 28, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.sendKeys(Keys.END);
		enter_Batch.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 29, 5));
		enter_Batch.sendKeys(Keys.TAB);

		// Thrid  Row  FIFo Cogs Item 

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_6thColumn));
		select3rdRow_6thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 30, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		// fourth Row Rma ritem

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_6thColumn));
		select4thRow_6thColumn.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 31, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);
		Thread.sleep(2000);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaClearBtn));
		rmaClearBtn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
		rmaSerialNumberTxtField.click();
		rmaSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 32, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
		rmaAddBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
		RMAPopupOkBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		Thread.sleep(4000);

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

		//expMsg.add(excelReader.getCellData(xlSheetName, 31, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 32, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);*/
		
	String expMessage1 = excelReader.getCellData(xlSheetName, 31, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 32, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
		//actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		if(actvoucherGridBodyList.equalsIgnoreCase(expvoucherGridBodyList) && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 26, 8, resPass);
			return true;
		}

		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 26, 8, resFail);
			return false;
		}
	}






	public boolean checkDeletingVoucherInExcessInStocks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		int bodyVoucherNolistCount = bodyVoucherNolist.size();

		for(int i=0;i<bodyVoucherNolistCount;i++)
		{
			String data = bodyVoucherNolist.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 33, 5)))
			{
				bodyChkBoxlist.get(i).click();

				break;
			}

		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHomeDeleteBtn));
		voucherHomeDeleteBtn.click();

		if(getIsAlertPresent())
		{
			getWaitForAlert();

			getAlert().accept();
		}


		String expMessage = excelReader.getCellData(xlSheetName, 33, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actMessage);
		
		System.out.println("Message : "+actMessage+" Value Expected : "+expMessage);

		Thread.sleep(1000);

		if (validationConfirmationMessage.getText().isEmpty()==false) 
		{
			checkValidationMessage("");
		}

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 33, 8, resPass);
			return true;
		}

		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 33, 8, resFail);
			return false;
		}
	}




	
	public boolean checkSavingAndSuspendingExcessInStocksVoucherInEntryPage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 34, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 35, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 36, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 37, 5));
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
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 34, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 35, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 34, 7, actMessage);
*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 34, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 35, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 34, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
		//actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
		new_SuspendBtn.click();

		String expSuspendMessage1=excelReader.getCellData(xlSheetName, 36, 6);
		String expSuspendMessage2=excelReader.getCellData(xlSheetName, 37, 6);

		String actSuspendMessage=checkValidationMessage(expSuspendMessage1);
		
		excelReader.setCellData(xlfile, xlSheetName, 36, 7, actSuspendMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
		Thread.sleep(2000);
		
		String actSuspendStatus = voucherHomeRow1SuspendedStatus.getText();
		String expSuspendStatus = excelReader.getCellData(xlSheetName, 38, 6);
		excelReader.setCellData(xlfile, xlSheetName, 38, 7, actMessage);
		
		System.out.println("Suspended Status : "+actSuspendStatus+"  Value Expected  "+expSuspendStatus);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2) && actSuspendMessage.startsWith(expSuspendMessage1) && actSuspendMessage.endsWith(expSuspendMessage2)
				&& actSuspendStatus.equalsIgnoreCase(expSuspendStatus))
		{
			excelReader.setCellData(xlfile, xlSheetName, 34, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 34, 8, resFail);
			return false;
		}
	}


	
	
	
	
	
	public boolean checkResavingExcessesInStocksVoucher3AndSuspendingFromHomepage() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 39, 5);

			if(voucherGridDocNo.get(i).getText().equalsIgnoreCase(data))
			{
				getAction().doubleClick(voucherGridDocNo.get(i)).build().perform();
				break;
			}
		}

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
		expMessage.add("Voucher saved successfully : 3");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 39, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 40, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 39, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 40, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
		//actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		
		Thread.sleep(2000);

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 40, 5);

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
		
		String expSuspendMessage = excelReader.getCellData(xlSheetName, 41, 6);
		String actSuspendMessage = checkValidationMessage(expSuspendMessage);
		
		Thread.sleep(2000);
		
		excelReader.setCellData(xlfile, xlSheetName, 41, 7, actSuspendMessage);
		
		String actSuspendStatus = voucherHomeRow1SuspendedStatus.getText();
		String expSuspendStatus = excelReader.getCellData(xlSheetName, 42, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 42, 7, actSuspendStatus);
		
		System.out.println("Suspended Status : "+actSuspendStatus+"  Value Expected  "+expSuspendStatus);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2) && actSuspendMessage.equalsIgnoreCase(expSuspendMessage)
				&& actSuspendStatus.equalsIgnoreCase(expSuspendStatus))
		{
			excelReader.setCellData(xlfile, xlSheetName, 39, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 39, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	public boolean checkDeletingExcessesInStockVoucher3FromHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 43, 5);

			if(voucherGridDocNo.get(i).getText().equalsIgnoreCase(data))
			{
				getAction().doubleClick(voucherGridDocNo.get(i)).build().perform();
				break;
			}
		}

		checkValidationMessage("Voucher loaded successfully");
		
		Thread.sleep(2000);			

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		String expMessage = excelReader.getCellData(xlSheetName, 43, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 43, 7, actMessage);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
		Thread.sleep(2000);

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 43, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 43, 8, resFail);
			return false;
		}	
	}
	
	
	
	
	

	public boolean checkVerifingDetailsOfSavedExcessInStocksVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(excessesInStocksVoucher));
		excessesInStocksVoucher.click();

		Thread.sleep(2000);*/

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

			if (data.isEmpty()==false && i!=9 && i!=10) 
			{
				System.out.println(data);

				actVouchersList.add(data);
			}
		}


		ArrayList<String>  expVouchersList= new ArrayList<String>(); 

		expVouchersList.add(excelReader.getCellData(xlSheetName, 44, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 45, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 46, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 47, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 48, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 49, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 50, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 51, 6));

		excelReader.setCellData(xlfile, xlSheetName, 44, 7, actVouchersList.toString());
		
		System.out.println("**********************checkVerifingDetailsOfSavedExcessInStocksVoucherInHomepage*********************");

		System.out.println(actVouchersList);
		System.out.println(expVouchersList);

		if (actVouchersList.equals(expVouchersList)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 44, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 44, 8, resFail);
			return false;
		}
	}







	public ExcessInStocksVoucherPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
