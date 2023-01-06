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

public class StockTransferVoucherPage extends BaseEngine 
{
	
	private static String xlfile = getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlSheetName = "StockTransferVoucherPage";
	
	@FindBy(xpath="//*[@id='id_body_85']")
	public static WebElement binIn;
	
	public boolean checkSavingStockTransferVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransfersVoucher));
		stockTransfersVoucher.click();

		Thread.sleep(2000);

		checkDeleteLinkStatus();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferheaderWarehouse1Txt));
		stockTransferheaderWarehouse1Txt.sendKeys(excelReader.getCellData(xlSheetName, 8, 5));
		Thread.sleep(2000);
		stockTransferheaderWarehouse1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_BodyWarehouse2Txt));
		enter_BodyWarehouse2Txt.sendKeys(Keys.END);
		enter_BodyWarehouse2Txt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_BodyWarehouse2Txt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		Thread.sleep(2000);
		enter_BodyWarehouse2Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Units));

		enter_Units.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		//select1stRow_4thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowCloseBtn));
		workFlowCloseBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_BodyWarehouse2Txt));
		enter_BodyWarehouse2Txt.sendKeys(Keys.END);
		enter_BodyWarehouse2Txt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_BodyWarehouse2Txt.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
		Thread.sleep(2000);
		enter_BodyWarehouse2Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 14, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 15, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowCloseBtn));
		workFlowCloseBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		batchPickOnFIFOIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
		batchOkIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
		select3rdRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_BodyWarehouse2Txt));
		enter_BodyWarehouse2Txt.sendKeys(Keys.END);
		enter_BodyWarehouse2Txt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_BodyWarehouse2Txt.sendKeys(excelReader.getCellData(xlSheetName, 17, 5));
		Thread.sleep(2000);
		enter_BodyWarehouse2Txt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 18, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 19, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowCloseBtn));
		workFlowCloseBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 20, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
		binAutoAllocateBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		binOkBtn.click();

		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binIn));
		binIn.sendKeys(Keys.TAB);
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(bin1Txt));

		bin1Txt.click();*/
	//	Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
		binAutoAllocateBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		binOkBtn.click();

		//fourth Row RMa
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_1stColumn));
		select4thRow_1stColumn.click();

		enter_BodyWarehouse2Txt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 21, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 22, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowCloseBtn));
		workFlowCloseBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 23, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaoutwardExpandBtn));
		rmaoutwardExpandBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAHeaderChkBox));
		searchRMAHeaderChkBox.click();	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
		searchRMAOkBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaoutwardOkBtn));
		rmaoutwardOkBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.TAB);

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
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 8, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 9, 6));
		
		String expMessage = expMsg.toString();
	
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actMessage);
	
		System.out.println(actMessage);
		System.out.println(expMessage);*/

		
		String expMessage1 = excelReader.getCellData(xlSheetName, 8, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 9, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
		//actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		
		
		if (actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
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








	public boolean checkSuspendOptionInSalesOrdersVoucherWithTransferedStock() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransfersVoucher));
		stockTransfersVoucher.click();

		Thread.sleep(2000);

		//checkDeleteLinkStatus();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
		new_SuspendBtn.click();

		String expMessage1=excelReader.getCellData(xlSheetName, 24, 6);
		String expMessage2=excelReader.getCellData(xlSheetName, 25, 6);
		String actMessage=checkValidationMessage(expMessage1);
		
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actMessage);

		if (actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 24, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 24, 8, resFail);
			return false;
		}
	}


	


	

	public boolean checkSavingVoucherUsingCopyDocumentInStockTransfer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		try 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));
			toggleBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(copyDocumentOption));
			copyDocumentOption.click();

			Thread.sleep(4000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cd_FirstChkBox));
			cd_FirstChkBox.click();

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cd_OkBtn));
			getAction().moveToElement(cd_OkBtn).click().build().perform();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferheaderWarehouse1Txt));
			stockTransferheaderWarehouse1Txt.click();
			stockTransferheaderWarehouse1Txt.sendKeys(Keys.END);
			stockTransferheaderWarehouse1Txt.sendKeys(Keys.SHIFT,Keys.HOME);
			stockTransferheaderWarehouse1Txt.sendKeys(excelReader.getCellData(xlSheetName, 26, 5));
			Thread.sleep(2000);
			stockTransferheaderWarehouse1Txt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_BodyWarehouse2Txt));
			enter_BodyWarehouse2Txt.sendKeys(Keys.END);
			enter_BodyWarehouse2Txt.sendKeys(Keys.SHIFT,Keys.HOME);
			enter_BodyWarehouse2Txt.sendKeys(excelReader.getCellData(xlSheetName, 27, 5));
			Thread.sleep(2000);
			enter_BodyWarehouse2Txt.sendKeys(Keys.TAB);


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
			select2ndRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_BodyWarehouse2Txt));
			enter_BodyWarehouse2Txt.sendKeys(Keys.END);
			enter_BodyWarehouse2Txt.sendKeys(Keys.SHIFT,Keys.HOME);
			enter_BodyWarehouse2Txt.sendKeys(excelReader.getCellData(xlSheetName, 28, 5));
			Thread.sleep(2000);
			enter_BodyWarehouse2Txt.sendKeys(Keys.TAB);


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
			select3rdRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_BodyWarehouse2Txt));
			enter_BodyWarehouse2Txt.sendKeys(Keys.END);
			enter_BodyWarehouse2Txt.sendKeys(Keys.SHIFT,Keys.HOME);
			enter_BodyWarehouse2Txt.sendKeys(excelReader.getCellData(xlSheetName, 29, 5));
			Thread.sleep(2000);
			enter_BodyWarehouse2Txt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_1stColumn));
			select4thRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_BodyWarehouse2Txt));
			enter_BodyWarehouse2Txt.sendKeys(Keys.END);
			enter_BodyWarehouse2Txt.sendKeys(Keys.SHIFT,Keys.HOME);
			enter_BodyWarehouse2Txt.sendKeys(excelReader.getCellData(xlSheetName, 30, 5));
			Thread.sleep(2000);
			enter_BodyWarehouse2Txt.sendKeys(Keys.TAB);


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_8thColumn));
			select2ndRow_8thColumn.click();

			Thread.sleep(1000);

			int count = batchPopupNameList.size();

			for (int i = 0; i < count; i++) 
			{
				String  data = batchPopupNameList.get(i).getText();

				if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 31, 5)))
				{
					getAction().doubleClick(batchPopupIndexList.get(i)).build().perform();
					break;
				}
			}
			
			Thread.sleep(2000);


			/*if (getIsAlertPresent()==true) 
				{
					System.out.println(" Alert Is DIsplayed ");
				   getAlert().accept();

	              Thread.sleep(1000);

				int count1 = batchPopupNameList.size();

				for (int i = 0; i < count1; i++) 
				{
					String  data = batchPopupNameList.get(i).getText();

					if (data.equalsIgnoreCase("BATCH#OSVR2"))
					{
						getAction().doubleClick(batchPopupIndexList.get(i)).build().perform();
						break;
					}
				}



				}*/

			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
				batchPickOnFIFOIcon.click();*/

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
			batchOkIcon.click();

			// Third Row BIn Finished Goods Item

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_9thColumn));
			select3rdRow_9thColumn.click();

			// Bin Outward

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
			binSearchBtn.click();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
			binAutoAllocateBtn.click();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
			binOkBtn.click();	


			// Bin Inward
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binIn));
			binIn.click();
			binIn.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
			binSearchBtn.click();

			Thread.sleep(1000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
			binAutoAllocateBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
			binOkBtn.click();	

			// Fourth Row Rma R Item

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_11thColumn));
			select4thRow_11thColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.SPACE);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaoutwardExpandBtn));
			rmaoutwardExpandBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAHeaderChkBox));
			searchRMAHeaderChkBox.click();	

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaoutwardOkBtn));
			rmaoutwardOkBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.TAB);

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
			
			/*String actMessage = SavingInBackground(2);

			HashSet<String> expMsg = new HashSet();

			expMsg.add(excelReader.getCellData(xlSheetName, 26, 6));
			expMsg.add(excelReader.getCellData(xlSheetName, 27, 6));
			
			String expMessage = expMsg.toString();
			
			excelReader.setCellData(xlfile, xlSheetName, 26, 7, actMessage);

			System.out.println(actMessage);
			System.out.println(expMessage);*/
			
			String expMessage1 = excelReader.getCellData(xlSheetName, 26, 6);
			
			String actMessage = checkValidationMessage(expMessage1);
			String expMessage2 = excelReader.getCellData(xlSheetName, 27, 6);
			
			excelReader.setCellData(xlfile, xlSheetName, 26, 7, actMessage);
		
			System.out.println("Actual Message    : "+actMessage);
			System.out.println("Expected Message  : "+expMessage1);
			//actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

			if (actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
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
		catch (Exception e) 
		{
			String exception =e.getMessage();
			System.err.println("CATCH BLOCK : "+ exception);

			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cd_CloseBtn));
			cd_CloseBtn.click();*/

			Thread.sleep(1000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
			previousBtn.click();

			checkValidationMessage("Voucher Loaded Successfully");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));
			toggleBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(copyToClipBoardOption));
			copyToClipBoardOption.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
			new_newBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));
			toggleBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pasteFromClipBoardOption));
			pasteFromClipBoardOption.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferheaderWarehouse1Txt));
			stockTransferheaderWarehouse1Txt.click();
			stockTransferheaderWarehouse1Txt.sendKeys(Keys.END);
			stockTransferheaderWarehouse1Txt.sendKeys(Keys.SHIFT,Keys.HOME);
			stockTransferheaderWarehouse1Txt.sendKeys(excelReader.getCellData(xlSheetName, 26, 5));
			Thread.sleep(2000);
			stockTransferheaderWarehouse1Txt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_BodyWarehouse2Txt));
			enter_BodyWarehouse2Txt.sendKeys(Keys.END);
			enter_BodyWarehouse2Txt.sendKeys(Keys.SHIFT,Keys.HOME);
			enter_BodyWarehouse2Txt.sendKeys(excelReader.getCellData(xlSheetName, 27, 5));
			Thread.sleep(2000);
			enter_BodyWarehouse2Txt.sendKeys(Keys.TAB);


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
			select2ndRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_BodyWarehouse2Txt));
			enter_BodyWarehouse2Txt.sendKeys(Keys.END);
			enter_BodyWarehouse2Txt.sendKeys(Keys.SHIFT,Keys.HOME);
			enter_BodyWarehouse2Txt.sendKeys(excelReader.getCellData(xlSheetName, 28, 5));
			Thread.sleep(2000);
			enter_BodyWarehouse2Txt.sendKeys(Keys.TAB);


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
			select3rdRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_BodyWarehouse2Txt));
			enter_BodyWarehouse2Txt.sendKeys(Keys.END);
			enter_BodyWarehouse2Txt.sendKeys(Keys.SHIFT,Keys.HOME);
			enter_BodyWarehouse2Txt.sendKeys(excelReader.getCellData(xlSheetName, 29, 5));
			Thread.sleep(2000);
			enter_BodyWarehouse2Txt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_1stColumn));
			select4thRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_BodyWarehouse2Txt));
			enter_BodyWarehouse2Txt.sendKeys(Keys.END);
			enter_BodyWarehouse2Txt.sendKeys(Keys.SHIFT,Keys.HOME);
			enter_BodyWarehouse2Txt.sendKeys(excelReader.getCellData(xlSheetName, 30, 5));
			Thread.sleep(2000);
			enter_BodyWarehouse2Txt.sendKeys(Keys.TAB);


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_8thColumn));
			select2ndRow_8thColumn.click();

			Thread.sleep(1000);

			int count = batchPopupNameList.size();

			for (int i = 0; i < count; i++) 
			{
				String  data = batchPopupNameList.get(i).getText();

				if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 31, 5)))
				{
					getAction().doubleClick(batchPopupIndexList.get(i)).build().perform();
					break;
				}
			}




			if (getIsAlertPresent()==true) 
			{
				System.out.println(" Alert Is DIsplayed ");
				getAlert().accept();

				Thread.sleep(1000);

				int count1 = batchPopupNameList.size();

				for (int i = 0; i < count1; i++) 
				{
					String  data = batchPopupNameList.get(i).getText();

					if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 31, 5)))
					{
						getAction().doubleClick(batchPopupIndexList.get(i)).build().perform();
						break;
					}
				}



			}

			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
				batchPickOnFIFOIcon.click();*/

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
			batchOkIcon.click();

			// Third Row BIn Finished Goods Item


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_9thColumn));
			select3rdRow_9thColumn.click();

			// Bin Outward

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
			binSearchBtn.click();

			Thread.sleep(4000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
			binAutoAllocateBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
			binOkBtn.click();	


			// Bin Inward
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binIn));
			binIn.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
			binSearchBtn.click();

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
			binAutoAllocateBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
			binOkBtn.click();	

			// Fourth Row Rma R Item

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_11thColumn));
			select4thRow_11thColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.SPACE);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaoutwardExpandBtn));
			rmaoutwardExpandBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAHeaderChkBox));
			searchRMAHeaderChkBox.click();	

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaoutwardOkBtn));
			rmaoutwardOkBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.TAB);

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
			
			/*String actMessage = SavingInBackground(2);

			HashSet<String> expMsg = new HashSet();

			expMsg.add(excelReader.getCellData(xlSheetName, 26, 6));
			expMsg.add(excelReader.getCellData(xlSheetName, 27, 6));
			
			String expMessage = expMsg.toString();
			
			excelReader.setCellData(xlfile, xlSheetName, 26, 7, actMessage);

			System.out.println(actMessage);
			System.out.println(expMessage);*/
			
			String expMessage1 = excelReader.getCellData(xlSheetName, 26, 6);
			
			String actMessage = checkValidationMessage(expMessage1);
			String expMessage2 = excelReader.getCellData(xlSheetName, 27, 6);
			
			excelReader.setCellData(xlfile, xlSheetName, 26, 7, actMessage);
		
			System.out.println("Actual Message    : "+actMessage);
			System.out.println("Expected Message  : "+expMessage1);
			//actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

			excelReader.setCellData(xlfile, xlSheetName, 26, 8, resFail);
			return false;
		}
	}






	




	public boolean checkUpdatingStockTransferVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
		select1stRow_6thColumn.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 32, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_6thColumn));
		select2ndRow_6thColumn.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 33, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_6thColumn));
		select3rdRow_6thColumn.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 34, 5));
		enter_Rate.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_6thColumn));
		select4thRow_6thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 35, 5));
		enter_Rate.sendKeys(Keys.TAB);


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

		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 32, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 33, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 32, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 33, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
		//actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		if (actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
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







	public boolean checkDeletingVoucherInStockTransfer() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 36, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded successfully");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		if(getIsAlertPresent())
		{
			getWaitForAlert();

			getAlert().accept();
		}

		String expMessage = excelReader.getCellData(xlSheetName, 36, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 36, 7, actMessage);

		System.out.println("Message : "+actMessage+" Value Expected : "+expMessage);

		Thread.sleep(2000);

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 36, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 36, 8, resFail);
			return false;
		}
	}




	
	
	
	public boolean checkSavingAndSuspendingStockTransferVoucherFromHomepage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferheaderWarehouse1Txt));
		stockTransferheaderWarehouse1Txt.sendKeys(excelReader.getCellData(xlSheetName, 37, 5));
		Thread.sleep(2000);
		stockTransferheaderWarehouse1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_BodyWarehouse2Txt));
		enter_BodyWarehouse2Txt.sendKeys(Keys.END);
		enter_BodyWarehouse2Txt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_BodyWarehouse2Txt.sendKeys(excelReader.getCellData(xlSheetName, 38, 5));
		Thread.sleep(2000);
		enter_BodyWarehouse2Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 39, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Units));
		
		enter_Units.sendKeys(Keys.TAB);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		//select1stRow_4thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 40, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowCloseBtn));
		workFlowCloseBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 41, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		getAction().moveToElement(voucherSaveBtn).build().perform();
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

		expMsg.add(excelReader.getCellData(xlSheetName, 37, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 38, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 37, 7, actMessage);
	
		System.out.println(actMessage);
		System.out.println(expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 37, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 38, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 37, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
		//actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		
		Thread.sleep(2000);

		int voucherGridDocNoCount = voucherGridDocNo.size();
		
		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 42, 5);

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
		
		String expSuspendMessage = excelReader.getCellData(xlSheetName, 39, 6);
		String actSuspendMessage = checkValidationMessage(expSuspendMessage);
		
		Thread.sleep(2000);
		
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actSuspendMessage);
		
		String actSuspendStatus = voucherHomeRow1SuspendedStatus.getText();
		String expSuspendStatus = excelReader.getCellData(xlSheetName, 40, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, actSuspendStatus);
		
		System.out.println("Suspended Status : "+actSuspendStatus+"  Value Expected  "+expSuspendStatus);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2) && actSuspendMessage.equalsIgnoreCase(expSuspendMessage)
				&& actSuspendStatus.equalsIgnoreCase(expSuspendStatus))
		{
			excelReader.setCellData(xlfile, xlSheetName, 37, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 37, 8, resFail);
			return false;
		}
	}

	
	
	
	
	
	
	public boolean checkDeleteOptionInStockTransferFromHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 43, 5);

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

		String expAlertMessage=excelReader.getCellData(xlSheetName, 43, 6);

		excelReader.setCellData(xlfile, xlSheetName, 43, 7, actAlertMessage);

		System.out.println("Alert Message on Deleteing the Voucher :  "+actAlertMessage +"  "+expAlertMessage);

		Thread.sleep(2000);

		getAlert().accept();

		String expMessage = excelReader.getCellData(xlSheetName, 44, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 44, 7, actMessage);

		System.out.println("Message : "+actMessage+" Value Expected : "+expMessage);

		if(actAlertMessage.contains(expAlertMessage) && actMessage.equalsIgnoreCase(expMessage))
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
	
	
	
	



	public boolean checkVerifingDetailsOfSavedStockTransferVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
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

		expVouchersList.add(excelReader.getCellData(xlSheetName, 45, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 46, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 47, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 48, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 49, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 50, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 51, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 52, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 53, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 54, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 55, 6));
		
		excelReader.setCellData(xlfile, xlSheetName, 45, 7, actVouchersList.toString());

		System.out.println("**********************checkVerifingDetailsOfSavedStockTransferVoucherInHomepage*********************");

		System.out.println(actVouchersList);
		System.out.println(expVouchersList);

		if (actVouchersList.equals(expVouchersList)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 45, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 45, 8, resFail);
			return false;
		}
	}














	public StockTransferVoucherPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
