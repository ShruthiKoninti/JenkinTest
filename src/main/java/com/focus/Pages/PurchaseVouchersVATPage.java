package com.focus.Pages;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashSet;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.omg.PortableServer.THREAD_POLICY_ID;
import org.openqa.selenium.By;
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
import com.testautomationguru.utility.PDFUtil;

public class PurchaseVouchersVATPage extends BaseEngine 
{
	private static String xlfile = getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlSheetName = "PurchaseVouchersVATPage";
	
	public static boolean checkRulesInPurchaseVoucherVat() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
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
		
		
		Thread.sleep(2000);
		
		 /*try {
				
				if(ExceptionerrorMessage.isDisplayed()) 
				{
					String actMessage=errorMessage.getText();
					System.out.println("Message Displayed : "+actMessage);
					
					errorMessageCloseBtn.click();
				}
		 }
		 catch(Exception e)
		 {
			 System.out.println(e.getMessage());
		 }

*/
		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 8, 5));
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		/* try {
				
				if(ExceptionerrorMessage1.isDisplayed()) 
				{
					String actMessage=errorMessage.getText();
					System.out.println("Meeage Displayed : "+actMessage);
					
					ExceptionerrorMessageCloseBtn.click();
				}
		 }
		 catch(Exception e)
		 {
			 System.out.println(e.getMessage());
		 }
*/

		String expRuleMessage = excelReader.getCellData(xlSheetName, 8, 6);
		String actRuleMessage = checkValidationMessage(expRuleMessage);
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actRuleMessage);

		String actNarration = purchaseVATNarrationTxt.getAttribute("value");
		String expNarration = excelReader.getCellData(xlSheetName, 9, 6);
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, actNarration);
		
		String actCaption = purchaseVATTableHeaderColumn2.getText();
		String expCaption = excelReader.getCellData(xlSheetName, 10, 6);
		excelReader.setCellData(xlfile, xlSheetName, 10, 7, actCaption);
		
		
		System.out.println(actRuleMessage+"   "+actNarration+"       "+ actCaption);

		if(actRuleMessage.equalsIgnoreCase(expRuleMessage) && actNarration.equalsIgnoreCase(expNarration) && actCaption.equalsIgnoreCase(expCaption))
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

	
	




	public static boolean checkSavingPurchaseVoucherVATVoucher1WithPurchaseOrderLink() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
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
		Thread.sleep(2000);
		
		
		
		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		documentNumberTxt.sendKeys(Keys.CONTROL, "l");
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowChkBox1));
		workFlowChkBox1.click();

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowOkBtnCL));
		workFlowOkBtnCL.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.TAB);	

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		Thread.sleep(2000);
		

		

		// Thread.sleep(2000);
		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT1");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);*/

		//HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 11, 6));
		//expMsg.add(excelReader.getCellData(xlSheetName, 12, 6));
		
		//String expMessage = expMsg.toString();
		
		//excelReader.setCellData(xlfile, xlSheetName, 11, 7, actMessage);

		//System.out.println(actMessage);
		//System.out.println(expMessage);

		//if(actMessage.equals(expMessage))*/
		
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








	public static boolean checkSavingPurchaseVoucherVATVoucher2WithBatchItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 14, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 15, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 17, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 18, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 19, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 20, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 21, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 22, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();

		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 23, 5));
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 2); 

		String next2DaysDate=df.format(cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 2Days>>>>>>>> : "+next2DaysDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next2DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++)
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT2");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 14, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 15, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 14, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);

		if(actMessage.equals(expMessage))*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 14, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 15, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 14, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 14, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 14, 8, resFail);
			return false;
		}
	}






	public static boolean checkEditingPurchaseVoucherVATVoucher2AddRow2AndUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 24, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 24, 6);
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actDocNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 25, 6);
		excelReader.setCellData(xlfile, xlSheetName, 25, 7, actVendor);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 26, 6);
		excelReader.setCellData(xlfile, xlSheetName, 26, 7, actCurrency);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 27, 6);
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actDepartment);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 28, 6);
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actPlaceOfSupply);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 29, 6);
		excelReader.setCellData(xlfile, xlSheetName, 29, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actWarehouse = select1stRow_1stColumn.getText();
		String expWarehouse = excelReader.getCellData(xlSheetName, 30, 6);
		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actWarehouse);
		
		String actItem = select1stRow_2ndColumn.getText();
		String expItem = excelReader.getCellData(xlSheetName, 31, 6);
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actItem);
		
		String actTaxCode = select1stRow_3rdColumn.getText();
		String expTaxCode = excelReader.getCellData(xlSheetName, 32, 6);
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actTaxCode);

		String actPurchaseAccount = select1stRow_4thColumn.getText();
		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 33, 6);
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actPurchaseAccount);
		
		String actUnits = select1stRow_5thColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 34, 6);
		excelReader.setCellData(xlfile, xlSheetName, 34, 7, actUnits);

		String actQuantity = select1stRow_9thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 35, 6);
		excelReader.setCellData(xlfile, xlSheetName, 35, 7, actQuantity);

		String actRate = select1stRow_11thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 36, 6);
		excelReader.setCellData(xlfile, xlSheetName, 36, 7, actRate);

		String actGross = select1stRow_12thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 37, 6);
		excelReader.setCellData(xlfile, xlSheetName, 37, 7, actGross);

		String actBatch = select1stRow_16thColumn.getText();
		String expBatch = excelReader.getCellData(xlSheetName, 38, 6);
		excelReader.setCellData(xlfile, xlSheetName, 38, 7, actBatch);
		
		String actExpiry = select1stRow_18thColumn.getText();
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		Calendar cal=Calendar.getInstance();
		cal.add(Calendar.DATE, 2); 
		String expExpiry=df.format(cal.getTime());


		System.out.println("*******************************checkEditingPurchaseVoucherVATVoucher1AddRow2AndUpdate**********************************");

		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		System.out.println("PlaceOfSupply   : "+actPlaceOfSupply	+"  Value Expected  "+expPlaceOfSupply);
		System.out.println("Jurisdiction    : "+actJurisdiction		+"  Value Expected  "+expJurisdiction);
		System.out.println("Warehouse       : "+actWarehouse		+"  Value Expected  "+expWarehouse);
		System.out.println("Item            : "+actItem				+"  Value Expected  "+expItem);
		System.out.println("TaxCode         : "+actTaxCode			+"  Value Expected  "+expTaxCode);
		System.out.println("PurchaseAccount : "+actPurchaseAccount	+"  Value Expected  "+expPurchaseAccount);
		System.out.println("Units           : "+actUnits			+"  Value Expected  "+expUnits);
		System.out.println("Quantity        : "+actQuantity			+"  Value Expected  "+expQuantity);
		System.out.println("Rate            : "+actRate				+"  Value Expected  "+expRate);
		System.out.println("Gross           : "+actGross			+"  Value Expected  "+expGross);
		System.out.println("Batch           : "+actBatch			+"  Value Expected  "+expBatch);
		System.out.println("Expiry          : "+actExpiry			+"  Value Expected  "+expExpiry);



		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 25, 5));

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 26, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 27, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
		select2ndRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 28, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_11thColumn));
		select2ndRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 29, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_14thColumn));
		select2ndRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();

		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 30, 5));
		enter_Batch.sendKeys(Keys.TAB);

		Calendar r2cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		r2cal.add(Calendar.DATE, 5); 

		String next5DaysDate=df.format(r2cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 5Days>>>>>>>> : "+next5DaysDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next5DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		
		Thread.sleep(3000);

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++)
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT2");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 39, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 40, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 39, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 40, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actWarehouse.equalsIgnoreCase(expWarehouse) 
				&& actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate) 
				&& actGross.equalsIgnoreCase(expGross) && actBatch.equalsIgnoreCase(expBatch) && actExpiry.equalsIgnoreCase(expExpiry) 
				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
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







	public static boolean checkVerifyingPurchaseVoucherVATVoucher2SavingVoucher3() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 41, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded Successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 41, 6);
		excelReader.setCellData(xlfile, xlSheetName, 41, 7, actDocNo);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 42, 6);
		excelReader.setCellData(xlfile, xlSheetName, 42, 7, actVendor);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 43, 6);
		excelReader.setCellData(xlfile, xlSheetName, 43, 7, actCurrency);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 44, 6);
		excelReader.setCellData(xlfile, xlSheetName, 44, 7, actDepartment);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 45, 6);
		excelReader.setCellData(xlfile, xlSheetName, 45, 7, actPlaceOfSupply);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 46, 6);
		excelReader.setCellData(xlfile, xlSheetName, 46, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 47, 6);
		excelReader.setCellData(xlfile, xlSheetName, 47, 7, actR1Warehouse);

		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 48, 6);
		excelReader.setCellData(xlfile, xlSheetName, 48, 7, actR1Item);

		String actR1TaxCode = select1stRow_3rdColumn.getText();
		String expR1TaxCode = excelReader.getCellData(xlSheetName, 49, 6);
		excelReader.setCellData(xlfile, xlSheetName, 49, 7, actR1TaxCode);

		String actR1PurchaseAccount = select1stRow_4thColumn.getText();
		String expR1PurchaseAccount = excelReader.getCellData(xlSheetName, 50, 6);
		excelReader.setCellData(xlfile, xlSheetName, 50, 7, actR1PurchaseAccount);

		String actR1Units = select1stRow_5thColumn.getText();
		String expR1Units = excelReader.getCellData(xlSheetName, 51, 6);
		excelReader.setCellData(xlfile, xlSheetName, 51, 7, actR1Units);

		String actR1Quantity = select1stRow_9thColumn.getText();
		String expR1Quantity = excelReader.getCellData(xlSheetName, 52, 6);
		excelReader.setCellData(xlfile, xlSheetName, 52, 7, actR1Quantity);

		String actR1Rate = select1stRow_11thColumn.getText();
		String expR1Rate = excelReader.getCellData(xlSheetName, 53, 6);
		excelReader.setCellData(xlfile, xlSheetName, 53, 7, actR1Rate);

		String actR1Gross = select1stRow_12thColumn.getText();
		String expR1Gross = excelReader.getCellData(xlSheetName, 54, 6);
		excelReader.setCellData(xlfile, xlSheetName, 54, 7, actR1Gross);

		String actR1Batch = select1stRow_16thColumn.getText();
		String expR1Batch = excelReader.getCellData(xlSheetName, 55, 6);
		excelReader.setCellData(xlfile, xlSheetName, 55, 7, actR1Batch);

		String actR1Expiry = select1stRow_18thColumn.getText();
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		Calendar R1cal=Calendar.getInstance();
		R1cal.add(Calendar.DATE, 2); 
		String expR1Expiry=df.format(R1cal.getTime());



		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 56, 6);
		excelReader.setCellData(xlfile, xlSheetName, 56, 7, actR2Warehouse);

		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 57, 6);
		excelReader.setCellData(xlfile, xlSheetName, 57, 7, actR2Item);

		String actR2TaxCode = select2ndRow_3rdColumn.getText();
		String expR2TaxCode = excelReader.getCellData(xlSheetName, 58, 6);
		excelReader.setCellData(xlfile, xlSheetName, 58, 7, actR2TaxCode);

		String actR2PurchaseAccount = select2ndRow_4thColumn.getText();
		String expR2PurchaseAccount = excelReader.getCellData(xlSheetName, 59, 6);
		excelReader.setCellData(xlfile, xlSheetName, 59, 7, actR2PurchaseAccount);

		String actR2Units = select2ndRow_5thColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 60, 6);
		excelReader.setCellData(xlfile, xlSheetName, 60, 7, actR2Units);

		String actR2Quantity = select2ndRow_9thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 61, 6);
		excelReader.setCellData(xlfile, xlSheetName, 61, 7, actR2Quantity);

		String actR2Rate = select2ndRow_11thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 62, 6);
		excelReader.setCellData(xlfile, xlSheetName, 62, 7, actR2Rate);

		String actR2Gross = select2ndRow_12thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 63, 6);
		excelReader.setCellData(xlfile, xlSheetName, 63, 7, actR2Gross);

		String actR2Batch = select2ndRow_16thColumn.getText();
		String expR2Batch = excelReader.getCellData(xlSheetName, 64, 6);
		excelReader.setCellData(xlfile, xlSheetName, 64, 7, actR2Batch);

		String actR2Expiry = select2ndRow_18thColumn.getText();
		Calendar R2cal=Calendar.getInstance();
		R2cal.add(Calendar.DATE, 5); 
		String expR2Expiry=df.format(R2cal.getTime());

		System.out.println("*******************************checkVerifyingPurchaseVoucherVATVoucher1SavingVoucher2**********************************");

		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		System.out.println("PlaceOfSupply   : "+actPlaceOfSupply	+"  Value Expected  "+expPlaceOfSupply);
		System.out.println("Jurisdiction    : "+actJurisdiction		+"  Value Expected  "+expJurisdiction);

		System.out.println("*******************************ROW1**********************************");

		System.out.println("Warehouse       : "+actR1Warehouse			+"  Value Expected  "+expR1Warehouse);
		System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
		System.out.println("TaxCode         : "+actR1TaxCode			+"  Value Expected  "+expR1TaxCode);
		System.out.println("PurchaseAccount : "+actR1PurchaseAccount	+"  Value Expected  "+expR1PurchaseAccount);
		System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
		System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
		System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
		System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);
		System.out.println("Batch           : "+actR1Batch				+"  Value Expected  "+expR1Batch);
		System.out.println("Expiry          : "+actR1Expiry				+"  Value Expected  "+expR1Expiry);

		System.out.println("*******************************ROW2**********************************");

		System.out.println("Warehouse       : "+actR2Warehouse			+"  Value Expected  "+expR2Warehouse);
		System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
		System.out.println("TaxCode         : "+actR2TaxCode			+"  Value Expected  "+expR2TaxCode);
		System.out.println("PurchaseAccount : "+actR2PurchaseAccount	+"  Value Expected  "+expR2PurchaseAccount);
		System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
		System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
		System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
		System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);
		System.out.println("Batch           : "+actR2Batch				+"  Value Expected  "+expR2Batch);
		System.out.println("Expiry          : "+actR2Expiry				+"  Value Expected  "+expR2Expiry);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
		new_newBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		documentNumberTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 42, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 43, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 44, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 45, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 46, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 47, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 48, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 49, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 50, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();

		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 51, 5));
		enter_Batch.sendKeys(Keys.TAB);

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 2); 

		String next2DaysDate=df.format(cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 2Days>>>>>>>> : "+next2DaysDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next2DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++)
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT3");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 65, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 66, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 65, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);
*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 65, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 66, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 65, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		Thread.sleep(2000);

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) 
				&& actR1Item.equalsIgnoreCase(expR1Item) && actR1TaxCode.equalsIgnoreCase(expR1TaxCode) && actR1PurchaseAccount.equalsIgnoreCase(expR1PurchaseAccount) 
				&& actR1Units.equalsIgnoreCase(expR1Units) && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) 
				&& actR1Gross.equalsIgnoreCase(expR1Gross) && actR1Batch.equalsIgnoreCase(expR1Batch) && actR1Expiry.equalsIgnoreCase(expR1Expiry)
				&& actR2Warehouse.equalsIgnoreCase(expR2Warehouse) 
				&& actR2Item.equalsIgnoreCase(expR2Item) && actR2TaxCode.equalsIgnoreCase(expR2TaxCode) && actR2PurchaseAccount.equalsIgnoreCase(expR2PurchaseAccount) 
				&& actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) 
				&& actR2Gross.equalsIgnoreCase(expR2Gross) && actR2Batch.equalsIgnoreCase(expR2Batch) && actR2Expiry.equalsIgnoreCase(expR2Expiry)
				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 41, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 41, 8, resFail);
			return false;
		}
	}









	public static boolean checkVerifyingAndDeletingPurchaseVoucherVATVoucher3() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 67, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 67, 6);
		excelReader.setCellData(xlfile, xlSheetName, 67, 7, actDocNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 68, 6);
		excelReader.setCellData(xlfile, xlSheetName, 68, 7, actVendor);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 69, 6);
		excelReader.setCellData(xlfile, xlSheetName, 69, 7, actCurrency);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 70, 6);
		excelReader.setCellData(xlfile, xlSheetName, 70, 7, actDepartment);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 71, 6);
		excelReader.setCellData(xlfile, xlSheetName, 71, 7, actPlaceOfSupply);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 72, 6);
		excelReader.setCellData(xlfile, xlSheetName, 72, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actWarehouse = select1stRow_1stColumn.getText();
		String expWarehouse = excelReader.getCellData(xlSheetName, 73, 6);
		excelReader.setCellData(xlfile, xlSheetName, 73, 7, actWarehouse);

		String actItem = select1stRow_2ndColumn.getText();
		String expItem = excelReader.getCellData(xlSheetName, 74, 6);
		excelReader.setCellData(xlfile, xlSheetName, 74, 7, actItem);

		String actTaxCode = select1stRow_3rdColumn.getText();
		String expTaxCode = excelReader.getCellData(xlSheetName, 75, 6);
		excelReader.setCellData(xlfile, xlSheetName, 75, 7, actTaxCode);

		String actPurchaseAccount = select1stRow_4thColumn.getText();
		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 76, 6);
		excelReader.setCellData(xlfile, xlSheetName, 76, 7, actPurchaseAccount);

		String actUnits = select1stRow_5thColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 77, 6);
		excelReader.setCellData(xlfile, xlSheetName, 77, 7, actUnits);

		String actQuantity = select1stRow_9thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 78, 6);
		excelReader.setCellData(xlfile, xlSheetName, 78, 7, actQuantity);

		String actRate = select1stRow_11thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 79, 6);
		excelReader.setCellData(xlfile, xlSheetName, 79, 7, actRate);

		String actGross = select1stRow_12thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 80, 6);
		excelReader.setCellData(xlfile, xlSheetName, 80, 7, actGross);

		String actBatch = select1stRow_16thColumn.getText();
		String expBatch = excelReader.getCellData(xlSheetName, 81, 6);
		excelReader.setCellData(xlfile, xlSheetName, 81, 7, actBatch);

		String actExpiry = select1stRow_18thColumn.getText();
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		Calendar cal=Calendar.getInstance();
		cal.add(Calendar.DATE, 2); 
		String expExpiry=df.format(cal.getTime());

		System.out.println("*******************************checkVerifyingAndDeletingPurchaseVoucherVATVoucher2**********************************");

		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		System.out.println("PlaceOfSupply   : "+actPlaceOfSupply	+"  Value Expected  "+expPlaceOfSupply);
		System.out.println("Jurisdiction    : "+actJurisdiction		+"  Value Expected  "+expJurisdiction);
		System.out.println("Warehouse       : "+actWarehouse		+"  Value Expected  "+expWarehouse);
		System.out.println("Item            : "+actItem				+"  Value Expected  "+expItem);
		System.out.println("TaxCode         : "+actTaxCode			+"  Value Expected  "+expTaxCode);
		System.out.println("PurchaseAccount : "+actPurchaseAccount	+"  Value Expected  "+expPurchaseAccount);
		System.out.println("Units           : "+actUnits			+"  Value Expected  "+expUnits);
		System.out.println("Quantity        : "+actQuantity			+"  Value Expected  "+expQuantity);
		System.out.println("Rate            : "+actRate				+"  Value Expected  "+expRate);
		System.out.println("Gross           : "+actGross			+"  Value Expected  "+expGross);
		System.out.println("Batch           : "+actBatch			+"  Value Expected  "+expBatch);
		System.out.println("Expiry          : "+actExpiry			+"  Value Expected  "+expExpiry);

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		String expMessage=excelReader.getCellData(xlSheetName, 82, 6);
		String actMessage=checkValidationMessage(expMessage);

		Thread.sleep(2000);

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actWarehouse.equalsIgnoreCase(expWarehouse) 
				&& actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate) 
				&& actGross.equalsIgnoreCase(expGross) && actBatch.equalsIgnoreCase(expBatch) && actExpiry.equalsIgnoreCase(expExpiry)
				&& actMessage.startsWith(expMessage))
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










	public static boolean checkSavingPurchaseVoucherVATVoucher3WithStockItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 83, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 84, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 85, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 86, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 87, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 88, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 89, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 90, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 91, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();
		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++)
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT3");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 83, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 84, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 83, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);

		Thread.sleep(2000);

		if(actMessage.equals(expMessage))*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 83, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 84, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 83, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

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









	public static boolean checkEditingPurchaseVoucherVATVoucher3AddRow2AndUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 92, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 92, 6);
		excelReader.setCellData(xlfile, xlSheetName, 92, 7, actDocNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 93, 6);
		excelReader.setCellData(xlfile, xlSheetName, 93, 7, actVendor);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 94, 6);
		excelReader.setCellData(xlfile, xlSheetName, 94, 7, actCurrency);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 95, 6);
		excelReader.setCellData(xlfile, xlSheetName, 95, 7, actDepartment);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 96, 6);
		excelReader.setCellData(xlfile, xlSheetName, 96, 7, actPlaceOfSupply);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 97, 6);
		excelReader.setCellData(xlfile, xlSheetName, 97, 7, actJurisdiction);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actWarehouse = select1stRow_1stColumn.getText();
		String expWarehouse = excelReader.getCellData(xlSheetName, 98, 6);
		excelReader.setCellData(xlfile, xlSheetName, 98, 7, actWarehouse);

		String actItem = select1stRow_2ndColumn.getText();
		String expItem = excelReader.getCellData(xlSheetName, 99, 6);
		excelReader.setCellData(xlfile, xlSheetName, 99, 7, actItem);

		String actTaxCode = select1stRow_3rdColumn.getText();
		String expTaxCode = excelReader.getCellData(xlSheetName, 100, 6);
		excelReader.setCellData(xlfile, xlSheetName, 100, 7, actTaxCode);

		String actPurchaseAccount = select1stRow_4thColumn.getText();
		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 101, 6);
		excelReader.setCellData(xlfile, xlSheetName, 101, 7, actPurchaseAccount);

		String actUnits = select1stRow_5thColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 102, 6);
		excelReader.setCellData(xlfile, xlSheetName, 102, 7, actUnits);

		String actQuantity = select1stRow_9thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 103, 6);
		excelReader.setCellData(xlfile, xlSheetName, 103, 7, actQuantity);

		String actRate = select1stRow_11thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 104, 6);
		excelReader.setCellData(xlfile, xlSheetName, 104, 7, actRate);

		String actGross = select1stRow_12thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 105, 6);
		excelReader.setCellData(xlfile, xlSheetName, 105, 7, actGross);

		System.out.println("*******************************checkEditingPurchaseVoucherVATVoucher2AddRow2AndUpdate**********************************");

		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		System.out.println("PlaceOfSupply   : "+actPlaceOfSupply	+"  Value Expected  "+expPlaceOfSupply);
		System.out.println("Jurisdiction    : "+actJurisdiction		+"  Value Expected  "+expJurisdiction);
		System.out.println("Warehouse       : "+actWarehouse		+"  Value Expected  "+expWarehouse);
		System.out.println("Item            : "+actItem				+"  Value Expected  "+expItem);
		System.out.println("TaxCode         : "+actTaxCode			+"  Value Expected  "+expTaxCode);
		System.out.println("PurchaseAccount : "+actPurchaseAccount	+"  Value Expected  "+expPurchaseAccount);
		System.out.println("Units           : "+actUnits			+"  Value Expected  "+expUnits);
		System.out.println("Quantity        : "+actQuantity			+"  Value Expected  "+expQuantity);
		System.out.println("Rate            : "+actRate				+"  Value Expected  "+expRate);
		System.out.println("Gross           : "+actGross			+"  Value Expected  "+expGross);


		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 93, 5));

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();
		enter_WarehouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 94, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 95, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
		select2ndRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 96, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_11thColumn));
		select2ndRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 97, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_14thColumn));
		select2ndRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();

		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++)
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT3");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 106, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 107, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 106, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);
*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 106, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 107, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 106, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		
		Thread.sleep(2000);

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actWarehouse.equalsIgnoreCase(expWarehouse) 
				&& actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate) 
				&& actGross.equalsIgnoreCase(expGross) 
				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 92, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 92, 8, resFail);
			return false;
		}
	}








	public static boolean checkVerifyingPurchaseVoucherVATVoucher3SavingVoucher4() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 108, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded Successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 108, 6);
		excelReader.setCellData(xlfile, xlSheetName, 108, 7, actDocNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 109, 6);
		excelReader.setCellData(xlfile, xlSheetName, 109, 7, actVendor);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 110, 6);
		excelReader.setCellData(xlfile, xlSheetName, 110, 7, actCurrency);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 111, 6);
		excelReader.setCellData(xlfile, xlSheetName, 111, 7, actDepartment);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 112, 6);
		excelReader.setCellData(xlfile, xlSheetName, 112, 7, actPlaceOfSupply);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 113, 6);
		excelReader.setCellData(xlfile, xlSheetName, 113, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 114, 6);
		excelReader.setCellData(xlfile, xlSheetName, 114, 7, actR1Warehouse);

		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 115, 6);
		excelReader.setCellData(xlfile, xlSheetName, 115, 7, actR1Item);

		String actR1TaxCode = select1stRow_3rdColumn.getText();
		String expR1TaxCode = excelReader.getCellData(xlSheetName, 116, 6);
		excelReader.setCellData(xlfile, xlSheetName, 116, 7, actR1TaxCode);

		String actR1PurchaseAccount = select1stRow_4thColumn.getText();
		String expR1PurchaseAccount = excelReader.getCellData(xlSheetName, 117, 6);
		excelReader.setCellData(xlfile, xlSheetName, 117, 7, actR1PurchaseAccount);

		String actR1Units = select1stRow_5thColumn.getText();
		String expR1Units = excelReader.getCellData(xlSheetName, 118, 6);
		excelReader.setCellData(xlfile, xlSheetName, 118, 7, actR1Units);

		String actR1Quantity = select1stRow_9thColumn.getText();
		String expR1Quantity = excelReader.getCellData(xlSheetName, 119, 6);
		excelReader.setCellData(xlfile, xlSheetName, 119, 7, actR1Quantity);

		String actR1Rate = select1stRow_11thColumn.getText();
		String expR1Rate = excelReader.getCellData(xlSheetName, 120, 6);
		excelReader.setCellData(xlfile, xlSheetName, 120, 7, actR1Rate);

		String actR1Gross = select1stRow_12thColumn.getText();
		String expR1Gross = excelReader.getCellData(xlSheetName, 121, 6);
		excelReader.setCellData(xlfile, xlSheetName, 121, 7, actR1Gross);


		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 122, 6);
		excelReader.setCellData(xlfile, xlSheetName, 122, 7, actR2Warehouse);

		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 123, 6);
		excelReader.setCellData(xlfile, xlSheetName, 123, 7, actR2Item);

		String actR2TaxCode = select2ndRow_3rdColumn.getText();
		String expR2TaxCode = excelReader.getCellData(xlSheetName, 124, 6);
		excelReader.setCellData(xlfile, xlSheetName, 124, 7, actR2TaxCode);

		String actR2PurchaseAccount = select2ndRow_4thColumn.getText();
		String expR2PurchaseAccount = excelReader.getCellData(xlSheetName, 125, 6);
		excelReader.setCellData(xlfile, xlSheetName, 125, 7, actR2PurchaseAccount);

		String actR2Units = select2ndRow_5thColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 126, 6);
		excelReader.setCellData(xlfile, xlSheetName, 126, 7, actR2Units);

		String actR2Quantity = select2ndRow_9thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 127, 6);
		excelReader.setCellData(xlfile, xlSheetName, 127, 7, actR2Quantity);

		String actR2Rate = select2ndRow_11thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 128, 6);
		excelReader.setCellData(xlfile, xlSheetName, 128, 7, actR2Rate);

		String actR2Gross = select2ndRow_12thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 129, 6);
		excelReader.setCellData(xlfile, xlSheetName, 129, 7, actR2Gross);

		System.out.println("*******************************checkVerifyingPurchaseVoucherVATVoucher1SavingVoucher2**********************************");

		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		System.out.println("PlaceOfSupply   : "+actPlaceOfSupply	+"  Value Expected  "+expPlaceOfSupply);
		System.out.println("Jurisdiction    : "+actJurisdiction		+"  Value Expected  "+expJurisdiction);

		System.out.println("*******************************ROW1**********************************");

		System.out.println("Warehouse       : "+actR1Warehouse			+"  Value Expected  "+expR1Warehouse);
		System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
		System.out.println("TaxCode         : "+actR1TaxCode			+"  Value Expected  "+expR1TaxCode);
		System.out.println("PurchaseAccount : "+actR1PurchaseAccount	+"  Value Expected  "+expR1PurchaseAccount);
		System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
		System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
		System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
		System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);

		System.out.println("*******************************ROW2**********************************");

		System.out.println("Warehouse       : "+actR2Warehouse			+"  Value Expected  "+expR2Warehouse);
		System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
		System.out.println("TaxCode         : "+actR2TaxCode			+"  Value Expected  "+expR2TaxCode);
		System.out.println("PurchaseAccount : "+actR2PurchaseAccount	+"  Value Expected  "+expR2PurchaseAccount);
		System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
		System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
		System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
		System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
		new_newBtn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 109, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 110, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 111, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 112, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 113, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 114, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 115, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 116, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 117, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();

		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++)
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT4");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 130, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 131, 6));
		
		String expMessage = expMsg.toString();

		excelReader.setCellData(xlfile, xlSheetName, 130, 7, actMessage);
		
		System.out.println(actMessage);
		System.out.println(expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 130, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 131, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 130, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))


		Thread.sleep(2000);

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) 
				&& actR1Item.equalsIgnoreCase(expR1Item) && actR1TaxCode.equalsIgnoreCase(expR1TaxCode) && actR1PurchaseAccount.equalsIgnoreCase(expR1PurchaseAccount) 
				&& actR1Units.equalsIgnoreCase(expR1Units) && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) 
				&& actR1Gross.equalsIgnoreCase(expR1Gross) && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) 
				&& actR2Item.equalsIgnoreCase(expR2Item) && actR2TaxCode.equalsIgnoreCase(expR2TaxCode) && actR2PurchaseAccount.equalsIgnoreCase(expR2PurchaseAccount) 
				&& actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) 
				&& actR2Gross.equalsIgnoreCase(expR2Gross) 
				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 108, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 108, 8, resFail);
			return false;
		}
	}








	public static boolean checkVerifyingAndDeletingPurchaseVoucherVATVoucher4() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 132, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 132, 6);
		excelReader.setCellData(xlfile, xlSheetName, 132, 7, actDocNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 133, 6);
		excelReader.setCellData(xlfile, xlSheetName, 133, 7, actVendor);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 134, 6);
		excelReader.setCellData(xlfile, xlSheetName, 134, 7, actCurrency);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 135, 6);
		excelReader.setCellData(xlfile, xlSheetName, 135, 7, actDepartment);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 136, 6);
		excelReader.setCellData(xlfile, xlSheetName, 136, 7, actPlaceOfSupply);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 137, 6);
		excelReader.setCellData(xlfile, xlSheetName, 137, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actWarehouse = select1stRow_1stColumn.getText();
		String expWarehouse = excelReader.getCellData(xlSheetName, 138, 6);
		excelReader.setCellData(xlfile, xlSheetName, 138, 7, actWarehouse);
		
		String actItem = select1stRow_2ndColumn.getText();
		String expItem = excelReader.getCellData(xlSheetName, 139, 6);
		excelReader.setCellData(xlfile, xlSheetName, 139, 7, actItem);

		String actTaxCode = select1stRow_3rdColumn.getText();
		String expTaxCode = excelReader.getCellData(xlSheetName, 140, 6);
		excelReader.setCellData(xlfile, xlSheetName, 140, 7, actTaxCode);

		String actPurchaseAccount = select1stRow_4thColumn.getText();
		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 141, 6);
		excelReader.setCellData(xlfile, xlSheetName, 141, 7, actPurchaseAccount);

		String actUnits = select1stRow_5thColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 142, 6);
		excelReader.setCellData(xlfile, xlSheetName, 142, 7, actUnits);

		String actQuantity = select1stRow_9thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 143, 6);
		excelReader.setCellData(xlfile, xlSheetName, 143, 7, actQuantity);

		String actRate = select1stRow_11thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 144, 6);
		excelReader.setCellData(xlfile, xlSheetName, 144, 7, actRate);

		String actGross = select1stRow_12thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 145, 6);
		excelReader.setCellData(xlfile, xlSheetName, 145, 7, actGross);

		System.out.println("*******************************checkVerifyingAndDeletingPurchaseVoucherVATVoucher2**********************************");

		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		System.out.println("PlaceOfSupply   : "+actPlaceOfSupply	+"  Value Expected  "+expPlaceOfSupply);
		System.out.println("Jurisdiction    : "+actJurisdiction		+"  Value Expected  "+expJurisdiction);
		System.out.println("Warehouse       : "+actWarehouse		+"  Value Expected  "+expWarehouse);
		System.out.println("Item            : "+actItem				+"  Value Expected  "+expItem);
		System.out.println("TaxCode         : "+actTaxCode			+"  Value Expected  "+expTaxCode);
		System.out.println("PurchaseAccount : "+actPurchaseAccount	+"  Value Expected  "+expPurchaseAccount);
		System.out.println("Units           : "+actUnits			+"  Value Expected  "+expUnits);
		System.out.println("Quantity        : "+actQuantity			+"  Value Expected  "+expQuantity);
		System.out.println("Rate            : "+actRate				+"  Value Expected  "+expRate);
		System.out.println("Gross           : "+actGross			+"  Value Expected  "+expGross);

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		String expMessage=excelReader.getCellData(xlSheetName, 146, 6);
		String actMessage=checkValidationMessage(expMessage);

		Thread.sleep(2000);
		
		excelReader.setCellData(xlfile, xlSheetName, 146, 7, actMessage);

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actWarehouse.equalsIgnoreCase(expWarehouse) 
				&& actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate) 
				&& actGross.equalsIgnoreCase(expGross) 
				&& actMessage.startsWith(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 132, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 132, 8, resFail);
			return false;
		}
	}











	public static boolean checkSavingPurchaseVoucherVATVoucher4WithBinItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 147, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 148, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 149, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 150, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 151, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 152, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 153, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 154, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 155, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();

		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		String actBinTotalQty           = binTotalQty.getAttribute("value");
		String actBinBaseUom            = binBaseUOM.getText();
		boolean actBinSearchTxt         = binSearchTxt.isDisplayed();
		boolean actBinSearchBtn         = binSearchBtn.isDisplayed();
		boolean actBinAutoAllocateBtn   = binAutoAllocateBtn.isDisplayed();
		boolean actBinPickBtn           = binPickBtn.isDisplayed();
		boolean actBinCancelBtn         = binCancelBtn.isDisplayed();
		//String actBinBalanceQty         = binBalanceTxt.getAttribute("value");

		String expBinTotalQty           = excelReader.getCellData(xlSheetName, 147, 6);
		excelReader.setCellData(xlfile, xlSheetName, 147, 7, actBinTotalQty);
		
		String expBinBaseUom            = excelReader.getCellData(xlSheetName, 148, 6);
		excelReader.setCellData(xlfile, xlSheetName, 148, 7, actBinBaseUom);
		
		boolean expBinSearchTxt         = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 149, 6)); //true;
		excelReader.setCellData(xlfile, xlSheetName, 149, 7, Boolean.toString(actBinSearchTxt));
		
		boolean expBinSearchBtn         = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 150, 6));
		excelReader.setCellData(xlfile, xlSheetName, 150, 7, Boolean.toString(actBinSearchBtn));
		
		boolean expBinAutoAllocateBtn   = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 151, 6));
		excelReader.setCellData(xlfile, xlSheetName, 151, 7, Boolean.toString(actBinAutoAllocateBtn));
		
		boolean expBinPickBtn           = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 152, 6));
		excelReader.setCellData(xlfile, xlSheetName, 152, 7, Boolean.toString(actBinPickBtn));
		
		boolean expBinCancelBtn         = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 153, 6));
		excelReader.setCellData(xlfile, xlSheetName, 153, 7, Boolean.toString(actBinCancelBtn));
		
		//String expBinBalanceQty         = excelReader.getCellData(xlSheetName, 154, 6);
		//excelReader.setCellData(xlfile, xlSheetName, 154, 7, actBinBalanceQty);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();

		String actTotalToBeAllocated 	= binRowTotalToBeAllocated.getText();
		String expTotalToBeAllocated 	= excelReader.getCellData(xlSheetName, 155, 6);
		excelReader.setCellData(xlfile, xlSheetName, 155, 7, actTotalToBeAllocated);

		System.out.println("***********************************BeforeAllocation**************************************");

		System.out.println("BinTotalQty        : "+actBinTotalQty			+"  Value Expected  "+expBinTotalQty);
		System.out.println("BinBaseUom         : "+actBinBaseUom			+"  Value Expected  "+expBinBaseUom);
		System.out.println("BinSearchTxt       : "+actBinSearchTxt			+"  Value Expected  "+expBinSearchTxt);
		System.out.println("BinSearchBtn       : "+actBinSearchBtn			+"  Value Expected  "+expBinSearchBtn);
		System.out.println("BinAutoAllocateBtn : "+actBinAutoAllocateBtn	+"  Value Expected  "+expBinAutoAllocateBtn);
		System.out.println("BinPickBtn         : "+actBinPickBtn			+"  Value Expected  "+expBinPickBtn);
		System.out.println("BinCancelBtn       : "+actBinCancelBtn			+"  Value Expected  "+expBinCancelBtn);
		//System.out.println("BinBalanceQty  : "+actBinBalanceQty			+"  Value Expected  "+expBinBalanceQty);

		System.out.println("TotalToBeAllocated    : "+actTotalToBeAllocated		+"  Value Expected  "+expTotalToBeAllocated);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
		binAutoAllocateBtn.click();

		Thread.sleep(2000);

		String actTotalToBeAllocated2 	= binRowTotalToBeAllocated.getText();
		String actBinBalanceQty2        = binBalanceTxt.getAttribute("value");

		String expTotalToBeAllocated2 	= excelReader.getCellData(xlSheetName, 156, 6);
		excelReader.setCellData(xlfile, xlSheetName, 156, 7, actTotalToBeAllocated2);
		
		String expBinBalanceQty2        = excelReader.getCellData(xlSheetName, 157, 6);
		excelReader.setCellData(xlfile, xlSheetName, 157, 7, actBinBalanceQty2);
		
		System.out.println("***********************************AfterAllocation**************************************");

		System.out.println("TotalToBeAllocated    : "+actTotalToBeAllocated2	+"  Value Expected  "+expTotalToBeAllocated2);
		System.out.println("BinBalanceQty         : "+actBinBalanceQty2			+"  Value Expected  "+expBinBalanceQty2);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		binOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_17thColumn));
		String actBinColumn=select1stRow_19thColumn.getText();

		String expBinColumn=excelReader.getCellData(xlSheetName, 158, 6);
		excelReader.setCellData(xlfile, xlSheetName, 158, 7, actBinColumn);
		
		System.out.println("Bin Column: "+actBinColumn+"  value expected  "+expBinColumn);

		if(actBinTotalQty.equalsIgnoreCase(expBinTotalQty) && actBinBaseUom.equalsIgnoreCase(expBinBaseUom) && actBinSearchTxt==expBinSearchTxt 
				&& actBinSearchBtn==expBinSearchBtn && actBinAutoAllocateBtn==expBinAutoAllocateBtn && actBinPickBtn==expBinPickBtn 
				&& actBinCancelBtn==expBinCancelBtn && actTotalToBeAllocated.equalsIgnoreCase(expTotalToBeAllocated)

				&& actTotalToBeAllocated2.equalsIgnoreCase(expTotalToBeAllocated2) && actBinBalanceQty2.equalsIgnoreCase(expBinBalanceQty2)

				&& actBinColumn.equalsIgnoreCase(expBinColumn))
		{
			excelReader.setCellData(xlfile, xlSheetName, 147, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 147, 8, resFail);
			return false;
		}
	}





	public boolean checkSavingPurchaseVoucherVATVoucher4With2RowsAndBinItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();
		enter_WarehouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 159, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_2ndColumn));
		select2ndRow_2ndColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 160, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
		select2ndRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 161, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_11thColumn));
		select2ndRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 162, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_14thColumn));
		select2ndRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();

		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		String actBinTotalQty           = binTotalQty.getAttribute("value");
		String actBinBaseUom            = binBaseUOM.getText();
		boolean actBinSearchTxt         = binSearchTxt.isDisplayed();
		boolean actBinSearchBtn         = binSearchBtn.isDisplayed();
		boolean actBinAutoAllocateBtn   = binAutoAllocateBtn.isDisplayed();
		boolean actBinPickBtn           = binPickBtn.isDisplayed();
		boolean actBinCancelBtn         = binCancelBtn.isDisplayed();
		//String actBinBalanceQty         = binBalanceTxt.getAttribute("value");

		String expBinTotalQty           = excelReader.getCellData(xlSheetName, 159, 6);
		excelReader.setCellData(xlfile, xlSheetName, 159, 7, actBinTotalQty);
		
		String expBinBaseUom            = excelReader.getCellData(xlSheetName, 160, 6);
		excelReader.setCellData(xlfile, xlSheetName, 160, 7, actBinBaseUom);
		
		boolean expBinSearchTxt         = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 161, 6));
		excelReader.setCellData(xlfile, xlSheetName, 161, 7, Boolean.toString(actBinSearchTxt));
		
		boolean expBinSearchBtn         = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 162, 6));
		excelReader.setCellData(xlfile, xlSheetName, 162, 7, Boolean.toString(actBinSearchBtn));
		
		boolean expBinAutoAllocateBtn   = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 163, 6));
		excelReader.setCellData(xlfile, xlSheetName, 163, 7, Boolean.toString(actBinAutoAllocateBtn));
		
		boolean expBinPickBtn           = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 164, 6));
		excelReader.setCellData(xlfile, xlSheetName, 164, 7, Boolean.toString(actBinPickBtn));
		
		boolean expBinCancelBtn         = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 165, 6));
		excelReader.setCellData(xlfile, xlSheetName, 165, 7, Boolean.toString(actBinCancelBtn));
		
		//String expBinBalanceQty         = excelReader.getCellData(xlSheetName, 166, 6);
		//excelReader.setCellData(xlfile, xlSheetName, 165, 7, actBinBalanceQty);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();

		String actTotalToBeAllocated 	= binRowTotalToBeAllocated.getText();
		String expTotalToBeAllocated 	= excelReader.getCellData(xlSheetName, 167, 6);
		excelReader.setCellData(xlfile, xlSheetName, 167, 7, actTotalToBeAllocated);

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

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 163, 5))) 
			{
				getAction().doubleClick(binPopupIndexList.get(i)).build().perform();
			}
		}

		Thread.sleep(2000);

		String actTotalToBeAllocated2 	= binRowTotalToBeAllocated.getText();
		String actBinBalanceQty2        = binBalanceTxt.getAttribute("value");

		String expTotalToBeAllocated2 	= excelReader.getCellData(xlSheetName, 168, 6);
		excelReader.setCellData(xlfile, xlSheetName, 168, 7, actTotalToBeAllocated2);
		
		String expBinBalanceQty2        = excelReader.getCellData(xlSheetName, 169, 6);
		excelReader.setCellData(xlfile, xlSheetName, 169, 7, actBinBalanceQty2);
		
		System.out.println("***********************************AfterAllocation**************************************");

		System.out.println("TotalToBeAllocated    : "+actTotalToBeAllocated2	+"  Value Expected  "+expTotalToBeAllocated2);
		System.out.println("BinBalanceQty         : "+actBinBalanceQty2			+"  Value Expected  "+expBinBalanceQty2);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		binOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_17thColumn));
		String actBinColumn=select2ndRow_19thColumn.getText();

		String expBinColumn=excelReader.getCellData(xlSheetName, 170, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 170, 7, actBinColumn);

		System.out.println("Bin Column: "+actBinColumn+"  value expected  "+expBinColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT4");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 171, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 172, 6));
		
		String expMessage = expMsg.toString();

		excelReader.setCellData(xlfile, xlSheetName, 171, 7, actMessage);
		
		System.out.println(actMessage);
		System.out.println(expMessage);
*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 171, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 172, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 171, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		Thread.sleep(2000);

		if(actBinTotalQty.equalsIgnoreCase(expBinTotalQty) && actBinBaseUom.equalsIgnoreCase(expBinBaseUom) && actBinSearchTxt==expBinSearchTxt 
				&& actBinSearchBtn==expBinSearchBtn && actBinAutoAllocateBtn==expBinAutoAllocateBtn && actBinPickBtn==expBinPickBtn 
				&& actBinCancelBtn==expBinCancelBtn && actTotalToBeAllocated.equalsIgnoreCase(expTotalToBeAllocated)
				  
				&& actTotalToBeAllocated2.equalsIgnoreCase(expTotalToBeAllocated2) && actBinBalanceQty2.equalsIgnoreCase(expBinBalanceQty2)
	
				&& actBinColumn.equalsIgnoreCase(expBinColumn)
	
				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 159, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 159, 8, resFail);
			return false;
		}
	}









	public static boolean checkEditingPurchaseVoucherVATVoucher4AddRow3AndUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 173, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 173, 6);
		excelReader.setCellData(xlfile, xlSheetName, 173, 7, actDocNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 174, 6);
		excelReader.setCellData(xlfile, xlSheetName, 174, 7, actVendor);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 175, 6);
		excelReader.setCellData(xlfile, xlSheetName, 175, 7, actCurrency);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 176, 6);
		excelReader.setCellData(xlfile, xlSheetName, 176, 7, actDepartment);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 177, 6);
		excelReader.setCellData(xlfile, xlSheetName, 177, 7, actPlaceOfSupply);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 178, 6);
		excelReader.setCellData(xlfile, xlSheetName, 178, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 179, 6);
		excelReader.setCellData(xlfile, xlSheetName, 179, 7, actR1Warehouse);

		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 180, 6);
		excelReader.setCellData(xlfile, xlSheetName, 180, 7, actR1Item);

		String actR1TaxCode = select1stRow_3rdColumn.getText();
		String expR1TaxCode = excelReader.getCellData(xlSheetName, 181, 6);
		excelReader.setCellData(xlfile, xlSheetName, 181, 7, actR1TaxCode);

		String actR1PurchaseAccount = select1stRow_4thColumn.getText();
		String expR1PurchaseAccount = excelReader.getCellData(xlSheetName, 182, 6);
		excelReader.setCellData(xlfile, xlSheetName, 182, 7, actR1PurchaseAccount);

		String actR1Units = select1stRow_5thColumn.getText();
		String expR1Units = excelReader.getCellData(xlSheetName, 183, 6);
		excelReader.setCellData(xlfile, xlSheetName, 183, 7, actR1Units);

		String actR1Quantity = select1stRow_9thColumn.getText();
		String expR1Quantity = excelReader.getCellData(xlSheetName, 184, 6);
		excelReader.setCellData(xlfile, xlSheetName, 184, 7, actR1Quantity);

		String actR1Rate = select1stRow_11thColumn.getText();
		String expR1Rate = excelReader.getCellData(xlSheetName, 185, 6);
		excelReader.setCellData(xlfile, xlSheetName, 185, 7, actR1Rate);

		String actR1Gross = select1stRow_12thColumn.getText();
		String expR1Gross = excelReader.getCellData(xlSheetName, 186, 6);
		excelReader.setCellData(xlfile, xlSheetName, 186, 7, actR1Gross);

		String actR1Bin = select2ndRow_17thColumn.getText();
		String expR1Bin = excelReader.getCellData(xlSheetName, 187, 6);
		excelReader.setCellData(xlfile, xlSheetName, 187, 7, actR1Bin);



		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 188, 6);
		excelReader.setCellData(xlfile, xlSheetName, 188, 7, actR2Warehouse);

		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 189, 6);
		excelReader.setCellData(xlfile, xlSheetName, 189, 7, actR2Item);

		String actR2TaxCode = select2ndRow_3rdColumn.getText();
		String expR2TaxCode = excelReader.getCellData(xlSheetName, 190, 6);
		excelReader.setCellData(xlfile, xlSheetName, 190, 7, actR2TaxCode);

		String actR2PurchaseAccount = select2ndRow_4thColumn.getText();
		String expR2PurchaseAccount = excelReader.getCellData(xlSheetName, 191, 6);
		excelReader.setCellData(xlfile, xlSheetName, 191, 7, actR2PurchaseAccount);

		String actR2Units = select2ndRow_5thColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 192, 6);
		excelReader.setCellData(xlfile, xlSheetName, 192, 7, actR2Units);

		String actR2Quantity = select2ndRow_9thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 193, 6);
		excelReader.setCellData(xlfile, xlSheetName, 193, 7, actR2Quantity);

		String actR2Rate = select2ndRow_11thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 194, 6);
		excelReader.setCellData(xlfile, xlSheetName, 194, 7, actR2Rate);

		String actR2Gross = select2ndRow_12thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 195, 6);
		excelReader.setCellData(xlfile, xlSheetName, 195, 7, actR2Gross);

		String actR2Bin = select2ndRow_17thColumn.getText();
		String expR2Bin = excelReader.getCellData(xlSheetName, 196, 6);
		excelReader.setCellData(xlfile, xlSheetName, 196, 7, actR2Bin);

		System.out.println("*******************************checkVerifyingPurchaseVoucherVATVoucher1SavingVoucher2**********************************");

		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		System.out.println("PlaceOfSupply   : "+actPlaceOfSupply	+"  Value Expected  "+expPlaceOfSupply);
		System.out.println("Jurisdiction    : "+actJurisdiction		+"  Value Expected  "+expJurisdiction);

		System.out.println("*******************************ROW1**********************************");

		System.out.println("Warehouse       : "+actR1Warehouse			+"  Value Expected  "+expR1Warehouse);
		System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
		System.out.println("TaxCode         : "+actR1TaxCode			+"  Value Expected  "+expR1TaxCode);
		System.out.println("PurchaseAccount : "+actR1PurchaseAccount	+"  Value Expected  "+expR1PurchaseAccount);
		System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
		System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
		System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
		System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);
		System.out.println("Bin             : "+actR1Bin				+"  Value Expected  "+expR1Bin);

		System.out.println("*******************************ROW2**********************************");

		System.out.println("Warehouse       : "+actR2Warehouse			+"  Value Expected  "+expR2Warehouse);
		System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
		System.out.println("TaxCode         : "+actR2TaxCode			+"  Value Expected  "+expR2TaxCode);
		System.out.println("PurchaseAccount : "+actR2PurchaseAccount	+"  Value Expected  "+expR2PurchaseAccount);
		System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
		System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
		System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
		System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);
		System.out.println("Bin             : "+actR2Bin				+"  Value Expected  "+expR2Bin);


		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
		select3rdRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();
		enter_WarehouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 174, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 175, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_9thColumn));
		select3rdRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 176, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_11thColumn));
		select3rdRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 177, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_14thColumn));
		select3rdRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();

		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		String actBinTotalQty           = binTotalQty.getAttribute("value");
		String actBinBaseUom            = binBaseUOM.getText();
		boolean actBinSearchTxt         = binSearchTxt.isDisplayed();
		boolean actBinSearchBtn         = binSearchBtn.isDisplayed();
		boolean actBinAutoAllocateBtn   = binAutoAllocateBtn.isDisplayed();
		boolean actBinPickBtn           = binPickBtn.isDisplayed();
		boolean actBinCancelBtn         = binCancelBtn.isDisplayed();
		//String actBinBalanceQty         = binBalanceTxt.getAttribute("value");

		String expBinTotalQty           = excelReader.getCellData(xlSheetName, 197, 6);
		excelReader.setCellData(xlfile, xlSheetName, 197, 7, actBinTotalQty);
		
		String expBinBaseUom            = excelReader.getCellData(xlSheetName, 198, 6);
		excelReader.setCellData(xlfile, xlSheetName, 198, 7, actBinBaseUom);
		
		boolean expBinSearchTxt         = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 199, 6));
		excelReader.setCellData(xlfile, xlSheetName, 199, 7, Boolean.toString(actBinSearchTxt));
		
		boolean expBinSearchBtn         = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 200, 6));
		excelReader.setCellData(xlfile, xlSheetName, 200, 7, Boolean.toString(actBinSearchBtn));
		
		boolean expBinAutoAllocateBtn   = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 201, 6));
		excelReader.setCellData(xlfile, xlSheetName, 201, 7, Boolean.toString(actBinAutoAllocateBtn));
		
		boolean expBinPickBtn           = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 202, 6));
		excelReader.setCellData(xlfile, xlSheetName, 202, 7, Boolean.toString(actBinPickBtn));
		
		boolean expBinCancelBtn         = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 203, 6));
		excelReader.setCellData(xlfile, xlSheetName, 203, 7, Boolean.toString(actBinCancelBtn));
		
		//String expBinBalanceQty         = excelReader.getCellData(xlSheetName, 204, 6);
		//excelReader.setCellData(xlfile, xlSheetName, 204, 7, actBinBalanceQty);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();

		String actTotalToBeAllocated 	= binRowTotalToBeAllocated.getText();

		String expTotalToBeAllocated 	= excelReader.getCellData(xlSheetName, 205, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 205, 7, actTotalToBeAllocated);

		System.out.println("***********************************BeforeAllocation**************************************");

		System.out.println("BinTotalQty         : "+actBinTotalQty+"  Value Expected  "+expBinTotalQty);
		System.out.println("BinBaseUom          : "+actBinBaseUom+"  Value Expected  "+expBinBaseUom);
		System.out.println("BinSearchTxt        : "+actBinSearchTxt+"  Value Expected  "+expBinSearchTxt);
		System.out.println("BinSearchBtn        : "+actBinSearchBtn+"  Value Expected  "+expBinSearchBtn);
		System.out.println("BinAutoAllocateBtn  : "+actBinAutoAllocateBtn+"  Value Expected  "+expBinAutoAllocateBtn);
		System.out.println("BinPickBtn          : "+actBinPickBtn+"  Value Expected  "+expBinPickBtn);
		System.out.println("BinCancelBtn        : "+actBinCancelBtn+"  Value Expected  "+expBinCancelBtn);
		//System.out.println("BinBalanceQty     : "+actBinBalanceQty+"  Value Expected  "+expBinBalanceQty);
		
		System.out.println("TotalToBeAllocated    : "+actTotalToBeAllocated		+"  Value Expected  "+expTotalToBeAllocated);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
		binAutoAllocateBtn.click();

		Thread.sleep(2000);

		String actTotalToBeAllocated2 	= binRowTotalToBeAllocated.getText();
		String actBinBalanceQty2        = binBalanceTxt.getAttribute("value");

		String expTotalToBeAllocated2 	= excelReader.getCellData(xlSheetName, 206, 6);
		excelReader.setCellData(xlfile, xlSheetName, 206, 7, actTotalToBeAllocated2);
		
		String expBinBalanceQty2        = excelReader.getCellData(xlSheetName, 207, 6);
		excelReader.setCellData(xlfile, xlSheetName, 207, 7, actBinBalanceQty2);

		System.out.println("***********************************AfterAllocation**************************************");
		
		System.out.println("TotalToBeAllocated    : "+actTotalToBeAllocated2	+"  Value Expected  "+expTotalToBeAllocated2);
		System.out.println("BinBalanceQty         : "+actBinBalanceQty2			+"  Value Expected  "+expBinBalanceQty2);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		binOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_17thColumn));
		String actBinColumn=select3rdRow_17thColumn.getText();

		String expBinColumn=excelReader.getCellData(xlSheetName, 208, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 208, 7, actBinColumn);

		System.out.println("Bin Column: "+actBinColumn+"  value expected  "+expBinColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++)
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT4");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 209, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 210, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 209, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 209, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 210, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 209, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) 
				&& actR1Item.equalsIgnoreCase(expR1Item) && actR1TaxCode.equalsIgnoreCase(expR1TaxCode) && actR1PurchaseAccount.equalsIgnoreCase(expR1PurchaseAccount) 
				&& actR1Units.equalsIgnoreCase(expR1Units) && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) 
				&& actR1Gross.equalsIgnoreCase(expR1Gross) && actR1Bin.equalsIgnoreCase(expR1Bin)
				&& actR2Warehouse.equalsIgnoreCase(expR2Warehouse) 
				&& actR2Item.equalsIgnoreCase(expR2Item) && actR2TaxCode.equalsIgnoreCase(expR2TaxCode) && actR2PurchaseAccount.equalsIgnoreCase(expR2PurchaseAccount) 
				&& actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) 
				&& actR2Gross.equalsIgnoreCase(expR2Gross) && actR2Bin.equalsIgnoreCase(expR2Bin) 

				&& actBinTotalQty.equalsIgnoreCase(expBinTotalQty) && actBinBaseUom.equalsIgnoreCase(expBinBaseUom) && actBinSearchTxt==expBinSearchTxt 
				&& actBinSearchBtn==expBinSearchBtn && actBinAutoAllocateBtn==expBinAutoAllocateBtn && actBinPickBtn==expBinPickBtn 
				&& actBinCancelBtn==expBinCancelBtn && actTotalToBeAllocated.equalsIgnoreCase(expTotalToBeAllocated)

				&& actTotalToBeAllocated2.equalsIgnoreCase(expTotalToBeAllocated2) && actBinBalanceQty2.equalsIgnoreCase(expBinBalanceQty2)

				&& actBinColumn.equalsIgnoreCase(expBinColumn)

				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 173, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 173, 8, resFail);
			return false;
		}
	}









	public static boolean checkVerifyingPurchaseVoucherVATVoucher4SavingVoucher5() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 211, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 211, 6);
		excelReader.setCellData(xlfile, xlSheetName, 211, 7, actDocNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 212, 6);
		excelReader.setCellData(xlfile, xlSheetName, 212, 7, actVendor);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 213, 6);
		excelReader.setCellData(xlfile, xlSheetName, 213, 7, actCurrency);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 214, 6);
		excelReader.setCellData(xlfile, xlSheetName, 214, 7, actDepartment);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 215, 6);
		excelReader.setCellData(xlfile, xlSheetName, 215, 7, actPlaceOfSupply);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 216, 6);
		excelReader.setCellData(xlfile, xlSheetName, 216, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 217, 6);
		excelReader.setCellData(xlfile, xlSheetName, 217, 7, actR1Warehouse);

		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 218, 6);
		excelReader.setCellData(xlfile, xlSheetName, 218, 7, actR1Item);

		String actR1TaxCode = select1stRow_3rdColumn.getText();
		String expR1TaxCode = excelReader.getCellData(xlSheetName, 219, 6);
		excelReader.setCellData(xlfile, xlSheetName, 219, 7, actR1TaxCode);

		String actR1PurchaseAccount = select1stRow_4thColumn.getText();
		String expR1PurchaseAccount = excelReader.getCellData(xlSheetName, 220, 6);
		excelReader.setCellData(xlfile, xlSheetName, 220, 7, actR1PurchaseAccount);

		String actR1Units = select1stRow_5thColumn.getText();
		String expR1Units = excelReader.getCellData(xlSheetName, 221, 6);
		excelReader.setCellData(xlfile, xlSheetName, 221, 7, actR1Units);

		String actR1Quantity = select1stRow_9thColumn.getText();
		String expR1Quantity = excelReader.getCellData(xlSheetName, 222, 6);
		excelReader.setCellData(xlfile, xlSheetName, 222, 7, actR1Quantity);

		String actR1Rate = select1stRow_11thColumn.getText();
		String expR1Rate = excelReader.getCellData(xlSheetName, 223, 6);
		excelReader.setCellData(xlfile, xlSheetName, 223, 7, actR1Rate);

		String actR1Gross = select1stRow_12thColumn.getText();
		String expR1Gross = excelReader.getCellData(xlSheetName, 224, 6);
		excelReader.setCellData(xlfile, xlSheetName, 224, 7, actR1Gross);

		String actR1Bin = select1stRow_17thColumn.getText();
		String expR1Bin = excelReader.getCellData(xlSheetName, 225, 6);
		excelReader.setCellData(xlfile, xlSheetName, 225, 7, actR1Bin);


		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 226, 6);
		excelReader.setCellData(xlfile, xlSheetName, 226, 7, actR2Warehouse);

		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 227, 6);
		excelReader.setCellData(xlfile, xlSheetName, 227, 7, actR2Item);

		String actR2TaxCode = select2ndRow_3rdColumn.getText();
		String expR2TaxCode = excelReader.getCellData(xlSheetName, 228, 6);
		excelReader.setCellData(xlfile, xlSheetName, 228, 7, actR2TaxCode);

		String actR2PurchaseAccount = select2ndRow_4thColumn.getText();
		String expR2PurchaseAccount = excelReader.getCellData(xlSheetName, 229, 6);
		excelReader.setCellData(xlfile, xlSheetName, 229, 7, actR2PurchaseAccount);

		String actR2Units = select2ndRow_5thColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 230, 6);
		excelReader.setCellData(xlfile, xlSheetName, 230, 7, actR2Units);

		String actR2Quantity = select2ndRow_9thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 231, 6);
		excelReader.setCellData(xlfile, xlSheetName, 231, 7, actR2Quantity);

		String actR2Rate = select2ndRow_11thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 232, 6);
		excelReader.setCellData(xlfile, xlSheetName, 232, 7, actR2Rate);

		String actR2Gross = select2ndRow_12thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 233, 6);
		excelReader.setCellData(xlfile, xlSheetName, 233, 7, actR2Gross);

		String actR2Bin = select2ndRow_17thColumn.getText();
		String expR2Bin = excelReader.getCellData(xlSheetName, 234, 6);
		excelReader.setCellData(xlfile, xlSheetName, 234, 7, actR2Bin);


		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
		String actR3Warehouse = select3rdRow_1stColumn.getText();
		String expR3Warehouse = excelReader.getCellData(xlSheetName, 235, 6);
		excelReader.setCellData(xlfile, xlSheetName, 235, 7, actR3Warehouse);

		String actR3Item = select3rdRow_2ndColumn.getText();
		String expR3Item = excelReader.getCellData(xlSheetName, 236, 6);
		excelReader.setCellData(xlfile, xlSheetName, 236, 7, actR3Item);

		String actR3TaxCode = select3rdRow_3rdColumn.getText();
		String expR3TaxCode = excelReader.getCellData(xlSheetName, 237, 6);
		excelReader.setCellData(xlfile, xlSheetName, 237, 7, actR3TaxCode);

		String actR3PurchaseAccount = select3rdRow_4thColumn.getText();
		String expR3PurchaseAccount = excelReader.getCellData(xlSheetName, 238, 6);
		excelReader.setCellData(xlfile, xlSheetName, 238, 7, actR3PurchaseAccount);

		String actR3Units = select3rdRow_5thColumn.getText();
		String expR3Units = excelReader.getCellData(xlSheetName, 239, 6);
		excelReader.setCellData(xlfile, xlSheetName, 239, 7, actR3Units);

		String actR3Quantity = select3rdRow_9thColumn.getText();
		String expR3Quantity = excelReader.getCellData(xlSheetName, 240, 6);
		excelReader.setCellData(xlfile, xlSheetName, 240, 7, actR3Quantity);

		String actR3Rate = select3rdRow_11thColumn.getText();
		String expR3Rate = excelReader.getCellData(xlSheetName, 241, 6);
		excelReader.setCellData(xlfile, xlSheetName, 241, 7, actR3Rate);

		String actR3Gross = select3rdRow_12thColumn.getText();
		String expR3Gross = excelReader.getCellData(xlSheetName, 242, 6);
		excelReader.setCellData(xlfile, xlSheetName, 242, 7, actR3Gross);

		String actR3Bin = select3rdRow_17thColumn.getText();
		String expR3Bin = excelReader.getCellData(xlSheetName, 243, 6);
		excelReader.setCellData(xlfile, xlSheetName, 243, 7, actR3Bin);


		System.out.println("*******************************checkVerifyingPurchaseVoucherVATVoucher1SavingVoucher2**********************************");

		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		System.out.println("PlaceOfSupply   : "+actPlaceOfSupply	+"  Value Expected  "+expPlaceOfSupply);
		System.out.println("Jurisdiction    : "+actJurisdiction		+"  Value Expected  "+expJurisdiction);

		System.out.println("*******************************ROW1**********************************");

		System.out.println("Warehouse       : "+actR1Warehouse			+"  Value Expected  "+expR1Warehouse);
		System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
		System.out.println("TaxCode         : "+actR1TaxCode			+"  Value Expected  "+expR1TaxCode);
		System.out.println("PurchaseAccount : "+actR1PurchaseAccount	+"  Value Expected  "+expR1PurchaseAccount);
		System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
		System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
		System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
		System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);
		System.out.println("Bin             : "+actR1Bin 				+"  Value Expected  "+expR1Bin);

		System.out.println("*******************************ROW2**********************************");

		System.out.println("Warehouse       : "+actR2Warehouse			+"  Value Expected  "+expR2Warehouse);
		System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
		System.out.println("TaxCode         : "+actR2TaxCode			+"  Value Expected  "+expR2TaxCode);
		System.out.println("PurchaseAccount : "+actR2PurchaseAccount	+"  Value Expected  "+expR2PurchaseAccount);
		System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
		System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
		System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
		System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);
		System.out.println("Bin             : "+actR2Bin 				+"  Value Expected  "+expR2Bin);

		System.out.println("*******************************ROW3**********************************");

		System.out.println("Warehouse       : "+actR3Warehouse			+"  Value Expected  "+expR3Warehouse);
		System.out.println("Item            : "+actR3Item				+"  Value Expected  "+expR3Item);
		System.out.println("TaxCode         : "+actR3TaxCode			+"  Value Expected  "+expR3TaxCode);
		System.out.println("PurchaseAccount : "+actR3PurchaseAccount	+"  Value Expected  "+expR3PurchaseAccount);
		System.out.println("Units           : "+actR3Units				+"  Value Expected  "+expR3Units);
		System.out.println("Quantity        : "+actR3Quantity			+"  Value Expected  "+expR3Quantity);
		System.out.println("Rate            : "+actR3Rate				+"  Value Expected  "+expR3Rate);
		System.out.println("Gross           : "+actR3Gross				+"  Value Expected  "+expR3Gross);
		System.out.println("Bin             : "+actR3Bin 				+"  Value Expected  "+expR3Bin);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
		new_newBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 212, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 213, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 214, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 215, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 216, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 217, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 218, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 219, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 220, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();
		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		String actBinTotalQty           = binTotalQty.getAttribute("value");
		String actBinBaseUom            = binBaseUOM.getText();
		String actBinSearchTxt         = Boolean.toString(binSearchTxt.isDisplayed());
		boolean actBinSearchBtn         = binSearchBtn.isDisplayed();
		boolean actBinAutoAllocateBtn   = binAutoAllocateBtn.isDisplayed();
		boolean actBinPickBtn           = binPickBtn.isDisplayed();
		boolean actBinCancelBtn         = binCancelBtn.isDisplayed();
		//String actBinBalanceQty         = binBalanceTxt.getAttribute("value");

		String expBinTotalQty           = excelReader.getCellData(xlSheetName, 244, 6);
		excelReader.setCellData(xlfile, xlSheetName, 244, 7, actBinTotalQty);
		
		String expBinBaseUom            = excelReader.getCellData(xlSheetName, 245, 6);
		excelReader.setCellData(xlfile, xlSheetName, 245, 7, actBinBaseUom);
		
		String expBinSearchTxt         = excelReader.getCellData(xlSheetName, 246, 6);
		excelReader.setCellData(xlfile, xlSheetName, 246, 7, actBinSearchTxt);
		
		boolean expBinSearchBtn         = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 247, 6));
		excelReader.setCellData(xlfile, xlSheetName, 247, 7, Boolean.toString(actBinSearchBtn));
		
		boolean expBinAutoAllocateBtn   = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 248, 6));
		excelReader.setCellData(xlfile, xlSheetName, 248, 7, Boolean.toString(actBinAutoAllocateBtn));
		
		boolean expBinPickBtn           = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 249, 6));
		excelReader.setCellData(xlfile, xlSheetName, 249, 7, Boolean.toString(actBinPickBtn));
		
		boolean expBinCancelBtn         = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 250, 6));
		excelReader.setCellData(xlfile, xlSheetName, 250, 7, Boolean.toString(actBinCancelBtn));
		
		//String expBinBalanceQty         = excelReader.getCellData(xlSheetName, 251, 6);
		//excelReader.setCellData(xlfile, xlSheetName, 251, 7, actBinBalanceQty);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();

		String actTotalToBeAllocated 	= binRowTotalToBeAllocated.getText();

		String expTotalToBeAllocated 	= excelReader.getCellData(xlSheetName, 252, 6);
		excelReader.setCellData(xlfile, xlSheetName, 252, 7, actTotalToBeAllocated);

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
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 221, 5))) 
			{
				binPopupToBeAllocatedList.get(i).click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ToBeAllocateTxt));
				enter_ToBeAllocateTxt.sendKeys(excelReader.getCellData(xlSheetName, 222, 5));
				//enter_ToBeAllocateTxt.sendKeys(Keys.TAB);
			}
		}
		
		for (int i = 0; i < count; i++) 
		{
			String data = binPopupBinsList.get(i).getText();
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 223, 5))) 
			{
				binPopupBinsList.get(i).click();
				binPopupToBeAllocatedList.get(i).click();
				Thread.sleep(1000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ToBeAllocateTxt));
				enter_ToBeAllocateTxt.sendKeys(excelReader.getCellData(xlSheetName, 224, 5));
				enter_ToBeAllocateTxt.sendKeys(Keys.TAB);
			}
		}


	/*	for (int i = 0; i < count; i++) 
		{
			String data = binPopupBinsList.get(i).getText();
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 223, 5))) 
			{
				binPopupToBeAllocatedList.get(i).click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ToBeAllocateTxt));
				//enter_ToBeAllocateTxt.click();
				enter_ToBeAllocateTxt.sendKeys(excelReader.getCellData(xlSheetName, 224, 5));
				enter_ToBeAllocateTxt.sendKeys(Keys.TAB);
			}
		}
*/
		Thread.sleep(2000);

		String actTotalToBeAllocated2 	= binRowTotalToBeAllocated.getText();
		String actBinBalanceQty2        = binBalanceTxt.getAttribute("value");

		String expTotalToBeAllocated2 	= excelReader.getCellData(xlSheetName, 253, 6);
		excelReader.setCellData(xlfile, xlSheetName, 253, 7, actTotalToBeAllocated2);
		
		String expBinBalanceQty2        = excelReader.getCellData(xlSheetName, 254, 6);
		excelReader.setCellData(xlfile, xlSheetName, 254, 7, actBinBalanceQty2);

		System.out.println("***********************************AfterAllocation**************************************");
		
		System.out.println("TotalToBeAllocated    : "+actTotalToBeAllocated2	+"  Value Expected  "+expTotalToBeAllocated2);
		System.out.println("BinBalanceQty         : "+actBinBalanceQty2			+"  Value Expected  "+expBinBalanceQty2);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		binOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_17thColumn));
		String actR1BinColumn=select1stRow_17thColumn.getText();
		String expR1BinColumn=excelReader.getCellData(xlSheetName, 255, 6);
		excelReader.setCellData(xlfile, xlSheetName, 255, 7, actR1BinColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_17thColumn));
		String actR2BinColumn=select2ndRow_17thColumn.getText();
		String expR2BinColumn=excelReader.getCellData(xlSheetName, 256, 6);
		excelReader.setCellData(xlfile, xlSheetName, 256, 7, actR2BinColumn);

		System.out.println("Bin Column: "+actR1BinColumn+"  value expected  "+expR1BinColumn);
		System.out.println("Bin Column: "+actR2BinColumn+"  value expected  "+expR2BinColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++)
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT5");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 257, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 258, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 257, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);
*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 257, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 258, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 257, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		Thread.sleep(2000);

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) 
				&& actR1Item.equalsIgnoreCase(expR1Item) && actR1TaxCode.equalsIgnoreCase(expR1TaxCode) && actR1PurchaseAccount.equalsIgnoreCase(expR1PurchaseAccount) 
				&& actR1Units.equalsIgnoreCase(expR1Units) && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) 
				&& actR1Gross.equalsIgnoreCase(expR1Gross) && actR1Bin.equalsIgnoreCase(expR1Bin) 

				&& actR2Warehouse.equalsIgnoreCase(expR2Warehouse) && actR2Item.equalsIgnoreCase(expR2Item) && actR2TaxCode.equalsIgnoreCase(expR2TaxCode) 
				&& actR2PurchaseAccount.equalsIgnoreCase(expR2PurchaseAccount) && actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) 
				&& actR2Rate.equalsIgnoreCase(expR2Rate) && actR2Gross.equalsIgnoreCase(expR2Gross) && actR2Bin.equalsIgnoreCase(expR2Bin) 

				&& actR3Warehouse.equalsIgnoreCase(expR3Warehouse) && actR3Item.equalsIgnoreCase(expR3Item) && actR3TaxCode.equalsIgnoreCase(expR3TaxCode) 
				&& actR3PurchaseAccount.equalsIgnoreCase(expR3PurchaseAccount) && actR3Units.equalsIgnoreCase(expR3Units) && actR3Quantity.equalsIgnoreCase(expR3Quantity) 
				&& actR3Rate.equalsIgnoreCase(expR3Rate) && actR3Gross.equalsIgnoreCase(expR3Gross) && actR3Bin.equalsIgnoreCase(expR3Bin) 

				&& actBinTotalQty.equalsIgnoreCase(expBinTotalQty) && actBinBaseUom.equalsIgnoreCase(expBinBaseUom) && actBinSearchTxt.equalsIgnoreCase(expBinSearchTxt) 
				&& actBinSearchBtn==expBinSearchBtn && actBinAutoAllocateBtn==expBinAutoAllocateBtn && actBinPickBtn==expBinPickBtn 
				&& actBinCancelBtn==expBinCancelBtn && actTotalToBeAllocated.equalsIgnoreCase(expTotalToBeAllocated)

				&& actTotalToBeAllocated2.equalsIgnoreCase(expTotalToBeAllocated2) && actBinBalanceQty2.equalsIgnoreCase(expBinBalanceQty2)

				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 211, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 211, 8, resFail);
			return false;
		}
	}








	public static boolean checkVerifyingAndDeletingPurchaseVoucherVATVoucher5() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 259, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 259, 6);
		excelReader.setCellData(xlfile, xlSheetName, 259, 7, actDocNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 260, 6);
		excelReader.setCellData(xlfile, xlSheetName, 260, 7, actVendor);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 261, 6);
		excelReader.setCellData(xlfile, xlSheetName, 261, 7, actCurrency);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 262, 6);
		excelReader.setCellData(xlfile, xlSheetName, 262, 7, actDepartment);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 263, 6);
		excelReader.setCellData(xlfile, xlSheetName, 263, 7, actPlaceOfSupply);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 264, 6);
		excelReader.setCellData(xlfile, xlSheetName, 264, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 265, 6);
		excelReader.setCellData(xlfile, xlSheetName, 265, 7, actR1Warehouse);

		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 266, 6);
		excelReader.setCellData(xlfile, xlSheetName, 266, 7, actR1Item);

		String actR1TaxCode = select1stRow_3rdColumn.getText();
		String expR1TaxCode = excelReader.getCellData(xlSheetName, 267, 6);
		excelReader.setCellData(xlfile, xlSheetName, 267, 7, actR1TaxCode);

		String actR1PurchaseAccount = select1stRow_4thColumn.getText();
		String expR1PurchaseAccount = excelReader.getCellData(xlSheetName, 268, 6);
		excelReader.setCellData(xlfile, xlSheetName, 268, 7, actR1PurchaseAccount);

		String actR1Units = select1stRow_5thColumn.getText();
		String expR1Units = excelReader.getCellData(xlSheetName, 269, 6);
		excelReader.setCellData(xlfile, xlSheetName, 269, 7, actR1Units);

		String actR1Quantity = select1stRow_9thColumn.getText();
		String expR1Quantity = excelReader.getCellData(xlSheetName, 270, 6);
		excelReader.setCellData(xlfile, xlSheetName, 270, 7, actR1Quantity);

		String actR1Rate = select1stRow_11thColumn.getText();
		String expR1Rate = excelReader.getCellData(xlSheetName, 271, 6);
		excelReader.setCellData(xlfile, xlSheetName, 271, 7, actR1Rate);

		String actR1Gross = select1stRow_12thColumn.getText();
		String expR1Gross = excelReader.getCellData(xlSheetName, 272, 6);
		excelReader.setCellData(xlfile, xlSheetName, 272, 7, actR1Gross);
		
		String actR1Bin = select1stRow_17thColumn.getText();
		String expR1Bin = excelReader.getCellData(xlSheetName, 273, 6);
		excelReader.setCellData(xlfile, xlSheetName, 273, 7, actR1Bin);


		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 274, 6);
		excelReader.setCellData(xlfile, xlSheetName, 274, 7, actR2Warehouse);

		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 275, 6);
		excelReader.setCellData(xlfile, xlSheetName, 275, 7, actR2Item);

		String actR2TaxCode = select2ndRow_3rdColumn.getText();
		String expR2TaxCode = excelReader.getCellData(xlSheetName, 276, 6);
		excelReader.setCellData(xlfile, xlSheetName, 276, 7, actR2TaxCode);

		String actR2PurchaseAccount = select2ndRow_4thColumn.getText();
		String expR2PurchaseAccount = excelReader.getCellData(xlSheetName, 277, 6);
		excelReader.setCellData(xlfile, xlSheetName, 277, 7, actR2PurchaseAccount);

		String actR2Units = select2ndRow_5thColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 278, 6);
		excelReader.setCellData(xlfile, xlSheetName, 278, 7, actR2Units);

		String actR2Quantity = select2ndRow_9thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 279, 6);
		excelReader.setCellData(xlfile, xlSheetName, 279, 7, actR2Quantity);

		String actR2Rate = select2ndRow_11thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 280, 6);
		excelReader.setCellData(xlfile, xlSheetName, 280, 7, actR2Rate);

		String actR2Gross = select2ndRow_12thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 281, 6);
		excelReader.setCellData(xlfile, xlSheetName, 281, 7, actR2Gross);

		String actR2Bin = select2ndRow_17thColumn.getText();
		String expR2Bin = excelReader.getCellData(xlSheetName, 282, 6);
		excelReader.setCellData(xlfile, xlSheetName, 283, 7, actR2Bin);


		System.out.println("*******************************checkVerifyingAndDeletingPurchaseVoucherVATVoucher4**********************************");

		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		System.out.println("PlaceOfSupply   : "+actPlaceOfSupply	+"  Value Expected  "+expPlaceOfSupply);
		System.out.println("Jurisdiction    : "+actJurisdiction		+"  Value Expected  "+expJurisdiction);

		System.out.println("*******************************ROW1**********************************");

		System.out.println("Warehouse       : "+actR1Warehouse			+"  Value Expected  "+expR1Warehouse);
		System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
		System.out.println("TaxCode         : "+actR1TaxCode			+"  Value Expected  "+expR1TaxCode);
		System.out.println("PurchaseAccount : "+actR1PurchaseAccount	+"  Value Expected  "+expR1PurchaseAccount);
		System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
		System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
		System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
		System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);
		System.out.println("Bin             : "+actR1Bin 				+"  Value Expected  "+expR1Bin);

		System.out.println("*******************************ROW2**********************************");

		System.out.println("Warehouse       : "+actR2Warehouse			+"  Value Expected  "+expR2Warehouse);
		System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
		System.out.println("TaxCode         : "+actR2TaxCode			+"  Value Expected  "+expR2TaxCode);
		System.out.println("PurchaseAccount : "+actR2PurchaseAccount	+"  Value Expected  "+expR2PurchaseAccount);
		System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
		System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
		System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
		System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);
		System.out.println("Bin             : "+actR2Bin 				+"  Value Expected  "+expR2Bin);

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		String expMessage=excelReader.getCellData(xlSheetName, 283, 6);
		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 283, 7, actMessage);

		Thread.sleep(2000);

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) 
				&& actR1Item.equalsIgnoreCase(expR1Item) && actR1TaxCode.equalsIgnoreCase(expR1TaxCode) && actR1PurchaseAccount.equalsIgnoreCase(expR1PurchaseAccount) 
				&& actR1Units.equalsIgnoreCase(expR1Units) && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) 
				&& actR1Gross.equalsIgnoreCase(expR1Gross) && actR1Bin.equalsIgnoreCase(expR1Bin) 

				&& actR2Warehouse.equalsIgnoreCase(expR2Warehouse) && actR2Item.equalsIgnoreCase(expR2Item) && actR2TaxCode.equalsIgnoreCase(expR2TaxCode) 
				&& actR2PurchaseAccount.equalsIgnoreCase(expR2PurchaseAccount) && actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) 
				&& actR2Rate.equalsIgnoreCase(expR2Rate) && actR2Gross.equalsIgnoreCase(expR2Gross) && actR2Bin.equalsIgnoreCase(expR2Bin)
				&& actMessage.startsWith(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 259, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 259, 8, resFail);
			return false;
		}
	}




@FindBy(xpath="//*[@id='id_body_38_input_container']/i")
public static WebElement rmaScrollBtn;



	public static boolean checkSavingPurchaseVoucherVATVoucher5WithRMAItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
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
		Thread.sleep(2000);
		
		
		
		checkValidationMessage("Screen opened");

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 284, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 285, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 286, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 287, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 288, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 289, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 290, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 291, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 292, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();

		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
		select1stRow_21stColumn.click();
		Thread.sleep(2000);
		*/
		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaClearBtn));
		rmaClearBtn.click();
		Thread.sleep(1000);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
		rmaSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 293, 5));

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
		String expRMA1="RMA#R1PVVAT1";
		String expRMA2="RMA#R1PVVAT2";
		String expRMA3="RMA#R1PVVAT3";
		String expRMA4="RMA#R1PVVAT4";
		String expRMA5="RMA#R1PVVAT5";
		String expRMA6="RMA#R1PVVAT6";
		String expRMA7="RMA#R1PVVAT7";
		String expRMA8="RMA#R1PVVAT8";
		String expRMA9="RMA#R1PVVAT9";
		String expRMA10="RMA#R1PVVAT10";
		String expRMA11="RMA#R1PVVAT11";
		String expRMA12="RMA#R1PVVAT12";


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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
		RMAPopupOkBtn.click();

		Thread.sleep(2000);
		
		boolean RMATableData = actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4)
						       && actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)
						       && actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10) && actRMA11.equalsIgnoreCase(expRMA11) && actRMA12.equalsIgnoreCase(expRMA12);
		
		excelReader.setCellData(xlfile, xlSheetName, 284, 7, Boolean.toString(RMATableData));
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));

		String actRMAColumn=enter_RMA.getAttribute("value");

		String expRMAColumn=excelReader.getCellData(xlSheetName, 285, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 285, 7, actRMAColumn);

		System.out.println("******** RMA Column: "+actRMAColumn+"  value expected  "+expRMAColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++)
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT5");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 292, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 293, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 292, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);
*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 292, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 293, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 292, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		if(actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) 
				&& actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)	
				&& actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10) && actRMA11.equalsIgnoreCase(expRMA11) && actRMA12.equalsIgnoreCase(expRMA12)
				&& actRMAColumn.equalsIgnoreCase(expRMAColumn) 
				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 284, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 284, 8, resFail);
			return false;
		}
	}









	public static boolean checkEditingPurchaseVoucherVATVoucher5AddRow2AndUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 294, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 294, 6);
		excelReader.setCellData(xlfile, xlSheetName, 294, 7, actDocNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 295, 6);
		excelReader.setCellData(xlfile, xlSheetName, 295, 7, actVendor);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 296, 6);
		excelReader.setCellData(xlfile, xlSheetName, 296, 7, actCurrency);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 297, 6);
		excelReader.setCellData(xlfile, xlSheetName, 297, 7, actDepartment);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 298, 6);
		excelReader.setCellData(xlfile, xlSheetName, 298, 7, actPlaceOfSupply);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 299, 6);
		excelReader.setCellData(xlfile, xlSheetName, 299, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actWarehouse = select1stRow_1stColumn.getText();
		String expWarehouse = excelReader.getCellData(xlSheetName, 300, 6);
		excelReader.setCellData(xlfile, xlSheetName, 300, 7, actWarehouse);

		String actItem = select1stRow_2ndColumn.getText();
		String expItem = excelReader.getCellData(xlSheetName, 301, 6);
		excelReader.setCellData(xlfile, xlSheetName, 301, 7, actItem);

		String actTaxCode = select1stRow_3rdColumn.getText();
		String expTaxCode = excelReader.getCellData(xlSheetName, 302, 6);
		excelReader.setCellData(xlfile, xlSheetName, 302, 7, actTaxCode);

		String actPurchaseAccount = select1stRow_4thColumn.getText();
		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 303, 6);
		excelReader.setCellData(xlfile, xlSheetName, 303, 7, actPurchaseAccount);

		String actUnits = select1stRow_5thColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 304, 6);
		excelReader.setCellData(xlfile, xlSheetName, 304, 7, actUnits);

		String actQuantity = select1stRow_9thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 305, 6);
		excelReader.setCellData(xlfile, xlSheetName, 305, 7, actQuantity);

		String actRate = select1stRow_11thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 306, 6);
		excelReader.setCellData(xlfile, xlSheetName, 306, 7, actRate);

		String actGross = select1stRow_12thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 307, 6);
		excelReader.setCellData(xlfile, xlSheetName, 307, 7, actGross);

		String actRMA = select1stRow_19thColumn.getText();
		String expRMA = excelReader.getCellData(xlSheetName, 308, 6);
		excelReader.setCellData(xlfile, xlSheetName, 308, 7, actRMA);

		System.out.println("*******************************checkEditingPurchaseVoucherVATVoucher4AddRow2AndUpdate**********************************");

		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		System.out.println("PlaceOfSupply   : "+actPlaceOfSupply	+"  Value Expected  "+expPlaceOfSupply);
		System.out.println("Jurisdiction    : "+actJurisdiction		+"  Value Expected  "+expJurisdiction);
		System.out.println("Warehouse       : "+actWarehouse		+"  Value Expected  "+expWarehouse);
		System.out.println("Item            : "+actItem				+"  Value Expected  "+expItem);
		System.out.println("TaxCode         : "+actTaxCode			+"  Value Expected  "+expTaxCode);
		System.out.println("PurchaseAccount : "+actPurchaseAccount	+"  Value Expected  "+expPurchaseAccount);
		System.out.println("Units           : "+actUnits			+"  Value Expected  "+expUnits);
		System.out.println("Quantity        : "+actQuantity			+"  Value Expected  "+expQuantity);
		System.out.println("Rate            : "+actRate				+"  Value Expected  "+expRate);
		System.out.println("Gross           : "+actGross			+"  Value Expected  "+expGross);
		System.out.println("RMA             : "+actRMA  			+"  Value Expected  "+expRMA);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();
		enter_WarehouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 295, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 296, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
		select2ndRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 297, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_11thColumn));
		select2ndRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 298, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_14thColumn));
		select2ndRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();

		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScrollBtn));
		rmaScrollBtn.click();
		Thread.sleep(1000);*/
		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
		rmaSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 299, 5));

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
		String expRMA1="RMA#R2PVVAT1";
		String expRMA2="RMA#R2PVVAT2";
		String expRMA3="RMA#R2PVVAT3";
		String expRMA4="RMA#R2PVVAT4";
		String expRMA5="RMA#R2PVVAT5";
		String expRMA6="RMA#R2PVVAT6";
		String expRMA7="RMA#R2PVVAT7";
		String expRMA8="RMA#R2PVVAT8";
		String expRMA9="RMA#R2PVVAT9";
		String expRMA10="RMA#R2PVVAT10";


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

		boolean RMATableData = actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4)
			       && actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)
			       && actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10);

		excelReader.setCellData(xlfile, xlSheetName, 309, 7, Boolean.toString(RMATableData));
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
		RMAPopupOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));

		String actRMAColumn=enter_RMA.getAttribute("value");

		String expRMAColumn=excelReader.getCellData(xlSheetName, 310, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 310, 7, actRMAColumn);

		System.out.println("******** RMA Column: "+actRMAColumn+"  value expected  "+expRMAColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++)
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT5");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 311, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 312, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 311, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 311, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 312, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 311, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		Thread.sleep(2000);

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actWarehouse.equalsIgnoreCase(expWarehouse) 
				&& actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate) 
				&& actGross.equalsIgnoreCase(expGross) && actRMA.equalsIgnoreCase(expRMA) 
				&& actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) 
				&& actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)	
				&& actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10) && actRMAColumn.equalsIgnoreCase(expRMAColumn) 
				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 294, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 294, 8, resFail);
			return false;
		}
	}







	public static boolean checkVerifyingPurchaseVoucherVATVoucher5SavingVoucher6() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 313, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 313, 6);
		excelReader.setCellData(xlfile, xlSheetName, 313, 7, actDocNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 314, 6);
		excelReader.setCellData(xlfile, xlSheetName, 314, 7, actVendor);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 315, 6);
		excelReader.setCellData(xlfile, xlSheetName, 315, 7, actCurrency);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 316, 6);
		excelReader.setCellData(xlfile, xlSheetName, 316, 7, actDepartment);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 317, 6);
		excelReader.setCellData(xlfile, xlSheetName, 317, 7, actPlaceOfSupply);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 318, 6);
		excelReader.setCellData(xlfile, xlSheetName, 318, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 319, 6);
		excelReader.setCellData(xlfile, xlSheetName, 319, 7, actR1Warehouse);

		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 320, 6);
		excelReader.setCellData(xlfile, xlSheetName, 320, 7, actR1Item);

		String actR1TaxCode = select1stRow_3rdColumn.getText();
		String expR1TaxCode = excelReader.getCellData(xlSheetName, 321, 6);
		excelReader.setCellData(xlfile, xlSheetName, 321, 7, actR1TaxCode);

		String actR1PurchaseAccount = select1stRow_4thColumn.getText();
		String expR1PurchaseAccount = excelReader.getCellData(xlSheetName, 322, 6);
		excelReader.setCellData(xlfile, xlSheetName, 322, 7, actR1PurchaseAccount);

		String actR1Units = select1stRow_5thColumn.getText();
		String expR1Units = excelReader.getCellData(xlSheetName, 323, 6);
		excelReader.setCellData(xlfile, xlSheetName, 323, 7, actR1Units);

		String actR1Quantity = select1stRow_9thColumn.getText();
		String expR1Quantity = excelReader.getCellData(xlSheetName, 324, 6);
		excelReader.setCellData(xlfile, xlSheetName, 324, 7, actR1Quantity);

		String actR1Rate = select1stRow_11thColumn.getText();
		String expR1Rate = excelReader.getCellData(xlSheetName, 325, 6);
		excelReader.setCellData(xlfile, xlSheetName, 325, 7, actR1Rate);

		String actR1Gross = select1stRow_12thColumn.getText();
		String expR1Gross = excelReader.getCellData(xlSheetName, 326, 6);
		excelReader.setCellData(xlfile, xlSheetName, 326, 7, actR1Gross);

		String actR1RMA = select1stRow_19thColumn.getText();
		String expR1RMA = excelReader.getCellData(xlSheetName, 327, 6);
		excelReader.setCellData(xlfile, xlSheetName, 327, 7, actR1RMA);




		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 328, 6);
		excelReader.setCellData(xlfile, xlSheetName, 328, 7, actR2Warehouse);

		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 329, 6);
		excelReader.setCellData(xlfile, xlSheetName, 329, 7, actR2Item);

		String actR2TaxCode = select2ndRow_3rdColumn.getText();
		String expR2TaxCode = excelReader.getCellData(xlSheetName, 330, 6);
		excelReader.setCellData(xlfile, xlSheetName, 330, 7, actR2TaxCode);

		String actR2PurchaseAccount = select2ndRow_4thColumn.getText();
		String expR2PurchaseAccount = excelReader.getCellData(xlSheetName, 331, 6);
		excelReader.setCellData(xlfile, xlSheetName, 331, 7, actR2PurchaseAccount);

		String actR2Units = select2ndRow_5thColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 332, 6);
		excelReader.setCellData(xlfile, xlSheetName, 332, 7, actR2Units);

		String actR2Quantity = select2ndRow_9thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 333, 6);
		excelReader.setCellData(xlfile, xlSheetName, 333, 7, actR2Quantity);

		String actR2Rate = select2ndRow_11thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 334, 6);
		excelReader.setCellData(xlfile, xlSheetName, 334, 7, actR2Rate);

		String actR2Gross = select2ndRow_12thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 335, 6);
		excelReader.setCellData(xlfile, xlSheetName, 335, 7, actR2Gross);
		
		String actR2RMA = select2ndRow_19thColumn.getText();
		String expR2RMA = excelReader.getCellData(xlSheetName, 336, 6);
		excelReader.setCellData(xlfile, xlSheetName, 336, 7, actR2RMA);

		System.out.println("*******************************checkVerifyingPurchaseVoucherVATVoucher4SavingVoucher5*********************************");

		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		System.out.println("PlaceOfSupply   : "+actPlaceOfSupply	+"  Value Expected  "+expPlaceOfSupply);
		System.out.println("Jurisdiction    : "+actJurisdiction		+"  Value Expected  "+expJurisdiction);

		System.out.println("*******************************ROW1**********************************");

		System.out.println("Warehouse       : "+actR1Warehouse			+"  Value Expected  "+expR1Warehouse);
		System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
		System.out.println("TaxCode         : "+actR1TaxCode			+"  Value Expected  "+expR1TaxCode);
		System.out.println("PurchaseAccount : "+actR1PurchaseAccount	+"  Value Expected  "+expR1PurchaseAccount);
		System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
		System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
		System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
		System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);
		System.out.println("RMA             : "+actR1RMA				+"  Value Expected  "+expR1RMA);

		System.out.println("*******************************ROW2**********************************");

		System.out.println("Warehouse       : "+actR2Warehouse			+"  Value Expected  "+expR2Warehouse);
		System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
		System.out.println("TaxCode         : "+actR2TaxCode			+"  Value Expected  "+expR2TaxCode);
		System.out.println("PurchaseAccount : "+actR2PurchaseAccount	+"  Value Expected  "+expR2PurchaseAccount);
		System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
		System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
		System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
		System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);
		System.out.println("RMA             : "+actR2RMA				+"  Value Expected  "+expR2RMA);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
		new_newBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 314, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
				voucherHeaderDueDate.click();
				voucherHeaderDueDate.sendKeys(Keys.TAB);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 315, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 316, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 317, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 318, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 319, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 320, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 321, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 322, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();

		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		Thread.sleep(2000);

	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScrollBtn));
		rmaScrollBtn.click();
		Thread.sleep(2000);
*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
		rmaSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 323, 5));

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
		String expRMA1="RMA#PVVAT1";
		String expRMA2="RMA#PVVAT2";
		String expRMA3="RMA#PVVAT3";
		String expRMA4="RMA#PVVAT4";
		String expRMA5="RMA#PVVAT5";
		String expRMA6="RMA#PVVAT6";
		String expRMA7="RMA#PVVAT7";
		String expRMA8="RMA#PVVAT8";
		String expRMA9="RMA#PVVAT9";
		String expRMA10="RMA#PVVAT10";
		String expRMA11="RMA#PVVAT11";
		String expRMA12="RMA#PVVAT12";


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

		
		boolean RMATableData = actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4)
			       && actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)
			       && actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10) && actRMA11.equalsIgnoreCase(expRMA11) && actRMA12.equalsIgnoreCase(expRMA12);

		excelReader.setCellData(xlfile, xlSheetName, 337, 7, Boolean.toString(RMATableData));
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
		RMAPopupOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));

		String actRMAColumn=enter_RMA.getAttribute("value");

		String expRMAColumn=excelReader.getCellData(xlSheetName, 338, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 338, 7, actRMAColumn);

		System.out.println("******** RMA Column: "+actRMAColumn+"  value expected  "+expRMAColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));

		newReferenceTxt.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++)
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT6");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 339, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 340, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 339, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);
*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 339, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 340, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 339, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		Thread.sleep(2000);

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) 
				&& actR1Item.equalsIgnoreCase(expR1Item) && actR1TaxCode.equalsIgnoreCase(expR1TaxCode) && actR1PurchaseAccount.equalsIgnoreCase(expR1PurchaseAccount) 
				&& actR1Units.equalsIgnoreCase(expR1Units) && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) 
				&& actR1Gross.equalsIgnoreCase(expR1Gross) && actR1RMA.equalsIgnoreCase(expR1RMA) 
				&& actR2Warehouse.equalsIgnoreCase(expR2Warehouse) 
				&& actR2Item.equalsIgnoreCase(expR2Item) && actR2TaxCode.equalsIgnoreCase(expR2TaxCode) && actR2PurchaseAccount.equalsIgnoreCase(expR2PurchaseAccount) 
				&& actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) 
				&& actR2Gross.equalsIgnoreCase(expR2Gross) && actR2RMA.equalsIgnoreCase(expR2RMA) 

				&& actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) 
				&& actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)	
				&& actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10) && actRMAColumn.equalsIgnoreCase(expRMAColumn) 

				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 313, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 313, 8, resFail);
			return false;
		}
	}










	public static boolean checkVerifyingAndDeletingPurchaseVoucherVATVoucher6() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 341, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 341, 6);
		excelReader.setCellData(xlfile, xlSheetName, 341, 7, actDocNo);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 342, 6);
		excelReader.setCellData(xlfile, xlSheetName, 342, 7, actVendor);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 343, 6);
		excelReader.setCellData(xlfile, xlSheetName, 343, 7, actCurrency);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 344, 6);
		excelReader.setCellData(xlfile, xlSheetName, 344, 7, actDepartment);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 345, 6);
		excelReader.setCellData(xlfile, xlSheetName, 345, 7, actPlaceOfSupply);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 346, 6);
		excelReader.setCellData(xlfile, xlSheetName, 346, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actWarehouse = select1stRow_1stColumn.getText();
		String expWarehouse = excelReader.getCellData(xlSheetName, 347, 6);
		excelReader.setCellData(xlfile, xlSheetName, 347, 7, actWarehouse);

		String actItem = select1stRow_2ndColumn.getText();
		String expItem = excelReader.getCellData(xlSheetName, 348, 6);
		excelReader.setCellData(xlfile, xlSheetName, 348, 7, actItem);

		String actTaxCode = select1stRow_3rdColumn.getText();
		String expTaxCode = excelReader.getCellData(xlSheetName, 349, 6);
		excelReader.setCellData(xlfile, xlSheetName, 349, 7, actTaxCode);

		String actPurchaseAccount = select1stRow_4thColumn.getText();
		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 350, 6);
		excelReader.setCellData(xlfile, xlSheetName, 350, 7, actPurchaseAccount);

		String actUnits = select1stRow_5thColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 351, 6);
		excelReader.setCellData(xlfile, xlSheetName, 351, 7, actUnits);

		String actQuantity = select1stRow_9thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 352, 6);
		excelReader.setCellData(xlfile, xlSheetName, 352, 7, actQuantity);

		String actRate = select1stRow_11thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 353, 6);
		excelReader.setCellData(xlfile, xlSheetName, 353, 7, actRate);

		String actGross = select1stRow_12thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 354, 6);
		excelReader.setCellData(xlfile, xlSheetName, 354, 7, actGross);

		String actRMA = select1stRow_19thColumn.getText();
		String expRMA = excelReader.getCellData(xlSheetName, 355, 6);
		excelReader.setCellData(xlfile, xlSheetName, 355, 7, actRMA);

		System.out.println("*******************************checkVerifyingAndDeletingPurchaseVoucherVATVoucher5**********************************");

		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		System.out.println("PlaceOfSupply   : "+actPlaceOfSupply	+"  Value Expected  "+expPlaceOfSupply);
		System.out.println("Jurisdiction    : "+actJurisdiction		+"  Value Expected  "+expJurisdiction);
		System.out.println("Warehouse       : "+actWarehouse		+"  Value Expected  "+expWarehouse);
		System.out.println("Item            : "+actItem				+"  Value Expected  "+expItem);
		System.out.println("TaxCode         : "+actTaxCode			+"  Value Expected  "+expTaxCode);
		System.out.println("PurchaseAccount : "+actPurchaseAccount	+"  Value Expected  "+expPurchaseAccount);
		System.out.println("Units           : "+actUnits			+"  Value Expected  "+expUnits);
		System.out.println("Quantity        : "+actQuantity			+"  Value Expected  "+expQuantity);
		System.out.println("Rate            : "+actRate				+"  Value Expected  "+expRate);
		System.out.println("Gross           : "+actGross			+"  Value Expected  "+expGross);
		System.out.println("RMA             : "+actRMA  			+"  Value Expected  "+expRMA);

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		String expMessage=excelReader.getCellData(xlSheetName, 356, 6);
		String actMessage=checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 356, 7, actMessage);
		
		Thread.sleep(2000);

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actWarehouse.equalsIgnoreCase(expWarehouse) 
				&& actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate) 
				&& actGross.equalsIgnoreCase(expGross) && actRMA.equalsIgnoreCase(expRMA)
				&& actMessage.startsWith(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 341, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 341, 8, resFail);
			return false;
		}
	}

	
	
	
	
	
	
	public static boolean checkSavingAndSuspendingPurchaseVoucherVATVoucher6WithStockItemInEntryPage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 357, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 358, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 359, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 360, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 361, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 362, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 363, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 364, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 365, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();

		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++)
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT6");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 357, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 358, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 357, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);
*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 357, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 358, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 357, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		Thread.sleep(2000);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 366, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded successfully");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
		new_SuspendBtn.click();

		String expSuspendMessage1=excelReader.getCellData(xlSheetName, 360, 6);
		String expSuspendMessage2=excelReader.getCellData(xlSheetName, 361, 6);

		String actSuspendMessage=checkValidationMessage(expSuspendMessage1);
		
		excelReader.setCellData(xlfile, xlSheetName, 360, 7, actSuspendMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
		Thread.sleep(2000);
		
		String actSuspendStatus = voucherHomeRow1SuspendedStatus.getText();
		String expSuspendStatus = excelReader.getCellData(xlSheetName, 362, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 362, 7, actSuspendStatus);
		
		System.out.println("Suspended Status : "+actSuspendStatus+"  Value Expected  "+expSuspendStatus);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)/* && actSuspendMessage.startsWith(expSuspendMessage1) 
				&& actSuspendMessage.endsWith(expSuspendMessage2)*/
				&& actSuspendStatus.equalsIgnoreCase(expSuspendStatus))
		{
			excelReader.setCellData(xlfile, xlSheetName, 357, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 357, 8, resFail);
			return false;
		}
	}





	
	public boolean checkResavingPurchaseVoucherVAT6AndSuspendingFromHomepage() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 367, 5);

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
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		
		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Saving in background.");
		expMessage.add("Voucher saved successfully : SU/IND/TEXT6");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 367, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 368, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 367, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 367, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 368, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 367, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		
		Thread.sleep(2000);

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 368, 5);

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
		
		String expSuspendMessage = excelReader.getCellData(xlSheetName, 369, 6);
		String actSuspendMessage = checkValidationMessage(expSuspendMessage);
		
		Thread.sleep(2000);
		
		excelReader.setCellData(xlfile, xlSheetName, 369, 7, actSuspendMessage);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherAll));
		voucherAll.click();
		*/
		String actSuspendStatus = voucherHomeRow1SuspendedStatus.getText();
		String expSuspendStatus = excelReader.getCellData(xlSheetName, 370, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 370, 7, actSuspendStatus);
		
		System.out.println("Suspended Status : "+actSuspendStatus+"  Value Expected  "+expSuspendStatus);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2) && actSuspendMessage.equalsIgnoreCase(expSuspendMessage)
				&& actSuspendStatus.equalsIgnoreCase(expSuspendStatus))
		{
			excelReader.setCellData(xlfile, xlSheetName, 367, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 367, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	public boolean checkDeleteOptionInPurchaseVoucherVATHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 371, 5);

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

		String expAlertMessage=excelReader.getCellData(xlSheetName, 371, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 371, 7, actAlertMessage);


		System.out.println("Alert Message on Deleteing the Voucher :  "+actAlertMessage +"  "+expAlertMessage);

		Thread.sleep(3000);

		getAlert().accept();

		String expMessage = excelReader.getCellData(xlSheetName, 372, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 372, 7, actMessage);

		System.out.println("Message : "+actMessage+" Value Expected : "+expMessage);

		if(actAlertMessage.contains(expAlertMessage) && actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 371, 8, resPass);
			return true;
		}

		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 371, 8, resFail);
			return false;
		}
	}
	
	
	
	


	public boolean checkVerifingDetailsOfSavedPurchaseVoucherVATInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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

			System.out.println(i+"."+data);

			if (data.isEmpty()==false && i!=9 && i!=10 && i!=23 && i!=24 && i!=37 && i!=38 && i!=51 && i!=52 && i!=65 && i!=66) 
			{
				actVouchersList.add(data);
			}
		}


		ArrayList<String>  expVouchersList= new ArrayList<String>(); 

		expVouchersList.add(excelReader.getCellData(xlSheetName, 373, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 374, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 375, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 376, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 377, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 378, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 379, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 380, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 381, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 382, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 383, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 384, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 385, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 386, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 387, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 388, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 389, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 390, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 391, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 392, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 393, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 394, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 395, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 396, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 397, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 398, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 399, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 400, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 401, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 402, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 403, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 404, 6));

		excelReader.setCellData(xlfile, xlSheetName, 373, 7, actVouchersList.toString());
		
		System.out.println("**********************checkVerifingDetailsOfSavedOpeningStocksVouchersInHomepage*********************");

		System.out.println(actVouchersList);
		System.out.println(expVouchersList);

		if (actVouchersList.equals(expVouchersList)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 373, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 373, 8, resFail);
			return false;
		}
	}





	public boolean checkPrintPdfOfPurchaseVoucherVAT() throws Exception, EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
		grid_ChkBox1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LayoutNameTxt));
		LayoutNameTxt.click();
		LayoutNameTxt.sendKeys(Keys.SPACE);
		Thread.sleep(2000);
		LayoutNameTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);


		Select s = new Select(LayoutOptionsdropdown);
		s.selectByIndex(3);

		Thread.sleep(2000);

		footerPrintBtn.click();

		Thread.sleep(5000);

		Robot robot = new Robot();
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_J);
		robot.keyRelease(KeyEvent.VK_J);
		robot.keyRelease(KeyEvent.VK_CONTROL);

		Thread.sleep(2000);

		robot.keyPress(KeyEvent.VK_TAB);
		robot.keyRelease(KeyEvent.VK_TAB);
		robot.keyPress(KeyEvent.VK_TAB);
		robot.keyRelease(KeyEvent.VK_TAB);

		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);

		Thread.sleep(2000);

		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_S);
		robot.keyRelease(KeyEvent.VK_S);
		robot.keyRelease(KeyEvent.VK_CONTROL);

		Thread.sleep(2000);

		Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\SavingPrintPDFInPurchaseVoucher.exe");

		Thread.sleep(4000);

		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_J);
		robot.keyRelease(KeyEvent.VK_J);
		robot.keyRelease(KeyEvent.VK_CONTROL);

		Thread.sleep(2000);

		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

		int actOpenWindowsCount = getDriver().getWindowHandles().size();
		int expOpenWindowsCount = 3;

		System.out.println("Number of Windows  : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);

		Thread.sleep(1000);

		getDriver().switchTo().window(openTabs.get(2)).close();
		Thread.sleep(1000);
		getDriver().switchTo().window(openTabs.get(1)).close();
		Thread.sleep(1000);
		getDriver().switchTo().window(openTabs.get(0));


		String actPDF = "C:\\Users\\D.n\\Downloads\\PurchaseVoucher.pdf";
		String expPDF = "D:\\FocusSanity\\FocusAI\\autoIt\\Prints\\ExpectedPurchaseVoucher.pdf";

		PDFUtil pdfutil = new PDFUtil();

		boolean result = pdfutil.compare(actPDF, expPDF);

		System.out.println("Compared Result  : "+result);

		if (result) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			return false;
		}
	}








	public boolean checkVerifingTriggersOptionInPurchaseVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys(excelReader.getCellData(xlSheetName, 405, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();

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

			System.out.println(i+"."+data);

			if (data.isEmpty()==false && i!=9 && i!=10 && i!=24 && i!=25 && i!=39 && i!=40 && i!=54 && i!=55 && i!=69 && i!=70) 
			{
				System.out.println(data);

				actVouchersList.add(data);
			}
		}


		ArrayList<String>  expVouchersList= new ArrayList<String>(); 

		expVouchersList.add(excelReader.getCellData(xlSheetName, 405, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 406, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 407, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 408, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 409, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 410, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 411, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 412, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 413, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 414, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 415, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 416, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 417, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 418, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 419, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 420, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 421, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 422, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 423, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 424, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 425, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 426, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 427, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 428, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 429, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 430, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 431, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 432, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 433, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 434, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 435, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 436, 6));
		
		excelReader.setCellData(xlfile, xlSheetName, 405, 7, actVouchersList.toString());

		System.out.println("**********************checkVerifingDetailsOfSavedOpeningStocksVouchersInHomepage*********************");

		System.out.println(actVouchersList);
		System.out.println(expVouchersList);

		if (actVouchersList.equals(expVouchersList)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 405, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 405, 8, resFail);
			return false;
		}
	}


///////
	
	@FindBy(xpath="//*[@class='icon-info-2 hiconright2']")
	public static WebElement voucherDetailsBtn;
	
	@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[14]/ul//li/a")
	public static List<WebElement> voucherDetailsList;
	
	@FindBy(xpath="//*[contains(text(),'Audit log details')]")
	public static WebElement auditLogDetails;
	
	@FindBy(xpath="//*[contains(text(),'Authorize history')]")
	public static WebElement authorizeHistory;
	
	@FindBy(xpath="//*[contains(text(),'Email sent details)]")
	public static WebElement emailSentDetails;
	
	@FindBy(xpath="//*[contains(text(),'Settings details')]")
	public static WebElement settingsDetails;
	
	@FindBy(xpath="//*[contains(text(),'Voucher details')]")
	public static WebElement voucherDetails;
	
	@FindBy(xpath="//*[@id='id_transaction_AuthorizeDetails']//table/tbody/tr[1]//td")
	public static List<WebElement> auditLogVoucherDetails1stRowList;
	
	@FindBy(xpath="//*[@id='id_transaction_AuthorizeDetails']//input")
	public static WebElement auditLogCloseBtn;
	
	
	public boolean checkAuditLogInformationOfSavedPurchaseVoucherVATInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		
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

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat todayDate = new SimpleDateFormat("dd/MM/yyyy");
		String currentDate = todayDate.format(cal.getTime());
		System.out.println("--------------Date Before Adding:: " + currentDate);

		int count = voucherHomeBodyList.size();

		

		for (int i = 0; i < count; i++) 
		{
			String data = voucherHomeBodyList.get(i).getText();

			System.out.println(i+"."+data);

			if (data.equals("SU/IND/TEXT4")) 
			{
				voucherHomeBodyList.get(i).click();
				voucherHomeBody2rdRow1stCol.click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherDetailsBtn));
		voucherDetailsBtn.click();
		Thread.sleep(2000);
		
		
		int count1 = voucherDetailsList.size();

		ArrayList<String>  actVouchersList= new ArrayList<String>(); 
		String actRow1List = null;
		String expRow1List = "[NDT57 : SU/IND/TEXT4, "+currentDate+", , , Unchecked, Modified, Test, SU, ]";

		for (int i = 0; i < count1; i++) 
		{
			String data = voucherDetailsList.get(i).getText();

			System.out.println(i+"."+data);

			if (data.equals("Audit log details")) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(auditLogDetails));
				auditLogDetails.click();
				Thread.sleep(3000);
				
				int reportsRow1ListCount = auditLogVoucherDetails1stRowList.size();
				ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
				for(int j=0;j<reportsRow1ListCount;j++)
				{
					if(j!=2)
					{
					String data1 = auditLogVoucherDetails1stRowList.get(j).getText();
					reportsRow1ListArray.add(data1);
					}
				}
				 actRow1List = reportsRow1ListArray.toString();
				
				
				System.err.println(actRow1List);
				System.err.println(expRow1List);
				
				
			}
			
			
		}
		
		if(actRow1List.equalsIgnoreCase(expRow1List))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(auditLogCloseBtn));
			auditLogCloseBtn.click();
				return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(auditLogCloseBtn));
			auditLogCloseBtn.click();
			return false;
		}
	}

	
	@FindBy(xpath="//*[@id='txtUsername']")
	public static WebElement createdUserTxt;
	
	@FindBy(xpath="//*[@id='txtDate']")
	public static WebElement createdDateTxt;
	
	
	public boolean checkAuthorizeHistoryInformationOfSavedPurchaseVoucherVATInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		Calendar cal=Calendar.getInstance();
		SimpleDateFormat todayDate = new SimpleDateFormat("dd/MM/yyyy");
		String currentDate = todayDate.format(cal.getTime());
		System.out.println("--------------Date Before Adding:: " + currentDate);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherDetailsBtn));
		voucherDetailsBtn.click();
		Thread.sleep(2000);
		
		
		int count1 = voucherDetailsList.size();

		ArrayList<String>  actVouchersList= new ArrayList<String>(); 
		String actRow1List = null;
		String expRow1List = null;
		
		String actCreatedUser=null;
		String expCreatedUser="Created by : SU";
		
		String actCreatedDate=null;
		String expCreatedDate="Created date : " + currentDate;

		for (int i = 0; i < count1; i++) 
		{
			String data = voucherDetailsList.get(i).getText();

			System.out.println(i+"."+data);

			if (data.equals("Authorize history")) 
				
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(authorizeHistory));
				authorizeHistory.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdUserTxt));
				actCreatedUser=createdUserTxt.getText();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdDateTxt));
				actCreatedDate=createdDateTxt.getText();
				
				
				System.err.println(actCreatedUser);
				
				System.err.println(expCreatedUser);
				
				System.err.println(actCreatedDate);
				
				System.err.println(expCreatedDate);
				
				
			}
			
			}
		
		if(actCreatedUser.equalsIgnoreCase(expCreatedUser) && actCreatedDate.contains(expCreatedDate))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(auditLogCloseBtn));
			auditLogCloseBtn.click();
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(auditLogCloseBtn));
			auditLogCloseBtn.click();
			return false;
		}
	}
	
	@FindBy(xpath="//*[@id='id_transaction_AuthorizeDetails']/div[2]/div/div[2]//label")
	public static List<WebElement> settingDetailsList;
	
	
	
	@FindBy(xpath="//*[@id='id_transaction_AuthorizeDetails']/div[2]/div/div[2]/div//label")
	public static List<WebElement> voucherDetailssubList;

	public boolean checkSettingDetailsInformationOfSavedPurchaseVoucherVATInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherDetailsBtn));
		voucherDetailsBtn.click();
		Thread.sleep(2000);
		
		
		int count1 = voucherDetailsList.size();

		ArrayList<String>  actVouchersList= new ArrayList<String>(); 
		String actRow1List = null;
		String expRow1List = null;
		
		

		for (int i = 0; i < count1; i++) 
		{
			String data = voucherDetailsList.get(i).getText();

			System.out.println(i+"."+data);

			if (data.equals("Settings details")) 
				
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsDetails));
				settingsDetails.click();
				Thread.sleep(2000);
				
			
			int reportsRow1ListCount = settingDetailsList.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int j=0;j<reportsRow1ListCount;j++)
			{
				String data2 = settingDetailsList.get(j).getText().trim();
				reportsRow1ListArray.add(data2);
			}
			 actRow1List = reportsRow1ListArray.toString().trim();
			 expRow1List = "[Detailed Posting, Yes, Document Numbering, Yes, Rule Count, 1, Trigger Count, 1, Update FA, Yes, Update Stock, Yes, View Count, 1]";
			}
			}
		
		System.err.println(actRow1List);
		
		System.err.println(expRow1List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) )
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(auditLogCloseBtn));
			auditLogCloseBtn.click();
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(auditLogCloseBtn));
			auditLogCloseBtn.click();
			return false;
		}
	}

	public boolean checkVoucherDetailsInformationOfSavedPurchaseVoucherVATInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		Calendar cal=Calendar.getInstance();
		SimpleDateFormat todayDate = new SimpleDateFormat("dd-MM-yyyy");
		String currentDate = todayDate.format(cal.getTime());
		System.out.println("--------------Date Before Adding:: " + currentDate);
		
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherDetailsBtn));
		voucherDetailsBtn.click();
		Thread.sleep(2000);
		
		
		int count1 = voucherDetailsList.size();

		ArrayList<String>  actVouchersList= new ArrayList<String>(); 
		String actRow1List = null;
		String expRow1List = null;
		
		

		for (int i = 0; i < count1; i++) 
		{
			String data = voucherDetailsList.get(i).getText();

			System.out.println(i+"."+data);

			if (data.equals("Voucher details")) 
				
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherDetails));
				voucherDetails.click();
				Thread.sleep(2000);
								
				
			
			
			int reportsRow1ListCount = voucherDetailssubList.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int j=0;j<reportsRow1ListCount;j++)
			{
				if(j!=7)
				{
				String data2 = voucherDetailssubList.get(j).getText().trim();
				reportsRow1ListArray.add(data2);
				}
			}
			 actRow1List = reportsRow1ListArray.toString().trim();
			 expRow1List = "[Status, Authorized, Date, "+currentDate+", Created by, SU, Created on, Net, ? 31.50, Line Count, 3]";
			}
			}
		
		System.err.println(actRow1List);
		
		System.err.println(expRow1List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) )
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(auditLogCloseBtn));
			auditLogCloseBtn.click();
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(auditLogCloseBtn));
			auditLogCloseBtn.click();
			return false;
		}
	}

	
	
	@FindBy(xpath="//*[@id='tblFooterReportRender']/div[2]")
	public static WebElement footerDetailsTxt;
	
	@FindBy(xpath="//*[@id='id_footer']/strong")
	public static WebElement footerCompanyTxt;
	
public boolean checkVoucherFooterDetailsatHomePage() throws InterruptedException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	financialsMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
	purchaseVouchersVat.click();

	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(footerDetailsTxt));
	String actText=footerDetailsTxt.getText();
	
	String s1=actText.substring(0, 19);
		
	Calendar cal=Calendar.getInstance();
	DateFormat df = new SimpleDateFormat("dd-MM-yyyy");
	String currentDate = df.format(cal.getTime());
	
	String s2= s1.concat(currentDate);
	String s3=actText.substring(37, actText.length());
	String expText=s2.concat(s3);
	
	System.err.println("Actual Text :"	+	actText);
	
	System.err.println("Expected Text :"	+	expText);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(footerCompanyTxt));
	String expCompNameatFooter=footerCompanyTxt.getText();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	companyLogo.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
	String actGetLoginCompanyNameInformation = companyName.getText();
	
	System.out.println("Company Name Display Value Actual                : " + actGetLoginCompanyNameInformation +  " Value Expected : " + expCompNameatFooter);

	if(actText.startsWith(expText) && actGetLoginCompanyNameInformation.equalsIgnoreCase(expCompNameatFooter))
	{
	
		return true;
	}
	else
	{
		return false;
	}
}



///////////////////////


public boolean checkSearchingVendorandItemsListwithoutAstrickSigninPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

	getDriver().navigate().refresh();
	Thread.sleep(1000);
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

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
	vendorAccountTxt.click();
	Thread.sleep(2000);
	vendorAccountTxt.sendKeys("en");
	Thread.sleep(3000);
	
	ArrayList<String>  actVendorAccountList= new ArrayList<String>(); 

	int vendorcount=vendorAccountListCount.size();

	System.err.println(vendorcount);

	for(int i=0 ; i < vendorcount ;i++)
	{
		String data=vendorAccountListCount.get(i).getText();
		actVendorAccountList.add(data);
	}

	boolean actVendorList= actVendorAccountList.isEmpty();
	boolean expVendorList= true;
		
	System.out.println(actVendorList);
	System.out.println(expVendorList);
	
	//vendorAccountTxt.sendKeys(Keys.TAB);
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
	enter_WarehouseTxt.click();
	enter_WarehouseTxt.sendKeys(Keys.TAB);



	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.click();
	enter_ItemTxt.sendKeys("fo");	
	Thread.sleep(2000);
	
	ArrayList<String>  actItemList= new ArrayList<String>(); 
	int itemscount=pvvGridItemList.size();

	for(int i=0 ; i < itemscount ;i++)
	{
		String data=pvvGridItemList.get(i).getText();
		actItemList.add(data);
	
	}

	boolean actItemsList=actItemList.isEmpty();
	boolean expItemsList=true;
	
	Thread.sleep(2000);
	
	System.out.println(actItemsList);
	System.out.println(expItemsList);
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
	new_CloseBtn.click();
	
	*/
	if(actVendorList==expVendorList  && actItemsList==expItemsList)
	{
		return true;
	} 
	else 
	{
		return false;
	}

}


public boolean checkSearchingofVendorandItemsListusingAstricksigninPurchaseVoucherVAT() throws InterruptedException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
	vendorAccountTxt.click();
	Thread.sleep(2000);
	vendorAccountTxt.sendKeys("*en");
	Thread.sleep(3000);
	
	ArrayList<String>  actVendorAccountList= new ArrayList<String>(); 

	int vendorcount=vendorAccountListCount.size();

	System.err.println(vendorcount);

	for(int i=0 ; i < vendorcount ;i++)
	{
		String data=vendorAccountListCount.get(i).getText();
		actVendorAccountList.add(data);
	}

	boolean actVendorList= actVendorAccountList.isEmpty();
	boolean expVendorList= true;
	
	String actVendorsList=null;
	
	if(actVendorList==false)
	{
		
		actVendorsList=actVendorAccountList.toString();
	}
	
	String expVendorsList="[Pricebook Vendor Account One, Pricebook Vendor Account Two, Vendor A, Vendor B]";
	
	vendorAccountTxt.sendKeys(Keys.TAB);

		
	System.out.println(actVendorsList);
	System.out.println(expVendorsList);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
	enter_WarehouseTxt.click();
	enter_WarehouseTxt.sendKeys(Keys.TAB);



	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.click();
	enter_ItemTxt.sendKeys("*fo");	
	Thread.sleep(2000);
	
	ArrayList<String>  actItemList= new ArrayList<String>(); 
	int itemscount=pvvGridItemList.size();

	for(int i=0 ; i < itemscount ;i++)
	{
		String data=pvvGridItemList.get(i).getText();
		actItemList.add(data);
	
	}

	boolean actItemsList=actItemList.isEmpty();
	boolean expItemsList=true;
	
	Thread.sleep(2000);
	String actitemsArry=null;
	if(actItemsList==false)
	{
		actitemsArry=actItemList.toString();
	}
	
	String expItemsArry="[BATCH FIFO ITEM, BATCH IGNORE EXP LIFO ITEM, FIFO COGS ITEM]";
	
	
	System.out.println(actitemsArry);
	System.out.println(expItemsArry);
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
	new_CloseBtn.click();
	
	*/
	if(actVendorsList.equalsIgnoreCase(expVendorsList)  && actitemsArry.equalsIgnoreCase(expItemsArry))
	{
		return true;
	} 
	else 
	{
		return false;
	}
}

////Elements at customization at voucher level


@FindBy(xpath="(//a[@title='Add Field'])[1]")
public static WebElement PV_EditLayoutHeaderAddFieldBtn;

@FindBy(xpath="//*[@id='editScreen_FieldsCustomization_Caption']")
public static WebElement editScreenCaptionTxt;


@FindBy(xpath="(//a[@title='Load Field'])[2]")
public static WebElement PV_EditLayoutHeaderLoadFieldBtn;

@FindBy(xpath="//table[@id='editLayout_headerTable']//tbody//tr[1]//td[1]/span[1]")
public static WebElement PV_EditLayoutHeader1stRowEditBtn;

@FindBy(xpath="//table[@id='editLayout_headerTable']//tbody//tr[1]//td[1]/span[2]")
public static WebElement PV_EditLayoutHeader1stRowDelBtn;

@FindBy(xpath="//table[@id='editLayout_headerTable']//tbody//tr[2]//td[1]/span[1]")
public static WebElement PV_EditLayoutHeader2ndRowEditBtn;

@FindBy(xpath="//table[@id='editLayout_headerTable']//tbody//tr[2]//td[1]/span[2]")
public static WebElement PV_EditLayoutHeader2ndRowDelBtn;


@FindBy(xpath="//table[@id='editLayout_headerTable']//tbody//tr[3]//td[1]/span[1]")
public static WebElement PV_EditLayoutHeader3rdRowEditBtn;

@FindBy(xpath="//table[@id='editLayout_headerTable']//tbody//tr[3]//td[1]/span[2]")
public static WebElement PV_EditLayoutHeader3rdRowDelBtn;


@FindBy(xpath="//table[@id='editLayout_bodyTable']//tr[2]//td//span[1]")
public static WebElement PV_EditLayoutBody2ndRowEditBtn;

@FindBy(xpath="//table[@id='editLayout_bodyTable']//tr[2]//td//span[2]")
public static WebElement PV_EditLayoutBody2ndRowDelBtn;



@FindBy(xpath="//table[@id='editScreen_bodyTable']//tr[6]//td//span[1]")
public static WebElement PV_EditScreenBody6thRowEditBtn;

@FindBy(xpath="//table[@id='editScreen_bodyTable']//tr[6]//td//span[2]")
public static WebElement PV_EditScreenBody6thRowDelBtn;

@FindBy(xpath="//table[@id='editScreen_bodyTable']//tr[7]//td//span[1]")
public static WebElement PV_EditScreenBody7thRowEditBtn;

@FindBy(xpath="//table[@id='editScreen_bodyTable']//tr[7]//td//span[2]")
public static WebElement PV_EditScreenBody7thRowDelBtn;


@FindBy(xpath="//table[@id='editScreen_footerTable']//tr[3]//td//span[1]")
public static WebElement PV_EditScreenFooter3rdRowEditBtn;

@FindBy(xpath="//table[@id='editScreen_footerTable']//tr[3]//td//span[2]")
public static WebElement PV_EditScreenFooter3rdRowDelBtn;



@FindBy(xpath="//*[@id='dvCustomizeEditLayoutField']/div[2]/div/div[3]/input[2]")
public static WebElement PV_editLayoutAddFieldCloseBtn;  


@FindBy(xpath="//*[@id='id_transactionentry_field_popup']//button/i")
public static WebElement editLayoutCloseBtn;  

@FindBy(xpath="(//a[@title='Add Field'])[2]")
public static WebElement PV_EditScreenLayoutBodyAddFieldBtn;

@FindBy(xpath="//*[@id='id_transactionentry_header1_section']//label")
public static List<WebElement> PV_HeaderList;






/*
 * editLayoutCaptionTxt
 * editLayoutApplyBtn
 * 
 * 
 * editScreenBodyRow3Edit
 */

@FindBy(xpath="(//*[@class='icon-custamize btn-img '])[1]")
public static WebElement PV_HeaderCustBtn;

@FindBy(xpath="(//*[@class='icon-custamize btn-img '])[2]")
public static WebElement PV_BodyCustBtn;

@FindBy(xpath="(//*[@class='icon-custamize btn-img '])[3]")
public static WebElement PV_FooterCustBtn;

@FindBy(xpath="//*[@id='id_transactionentry_field_popup']//div[1]/h5")
public static WebElement PV_CustHeaderLabel;


@FindBy(xpath="//*[@id='txtEditLayout_FieldsCustomization_FieldDetails_regularExp']")
public static WebElement PV_EditLayoutRegularExpTxt;


@FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_errorMsg']")
public static WebElement PV_EditLayoutErrorMsgTxt;

@FindBy(xpath="//*[@id='id_transactionentry_header1_section']//div[16]//input")
public static WebElement PV_AddedHeaderField;

@FindBy(xpath="//*[@id='id_transactionentry_header1_section']//div[16]//label")
public static WebElement PV_AddedHeaderFieldLabel;




public boolean checkCreatingEditLayoutHeaderatPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
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
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_HeaderCustBtn));
	PV_HeaderCustBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_CustHeaderLabel));
	String actHeaderLbl=PV_CustHeaderLabel.getText();
	String expHeaderLbl="Customize Header";
	
	System.out.println("Actual Header Label :	" +	actHeaderLbl);
	System.out.println("Expect Header Label :	" +	expHeaderLbl);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_EditLayoutHeaderAddFieldBtn));
	PV_EditLayoutHeaderAddFieldBtn.click();
	Thread.sleep(1000);
	
	if(getIsAlertPresent()){
		getWaitForAlert();
		Thread.sleep(1000);
		getAlert().accept();
		Thread.sleep(1000);
	}
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCaptionTxt));
	editLayoutCaptionTxt.click();
	editLayoutCaptionTxt.sendKeys("CustHeader");
	Thread.sleep(1000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_EditLayoutRegularExpTxt));
	PV_EditLayoutRegularExpTxt.click();
	PV_EditLayoutRegularExpTxt.sendKeys("Focus");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_EditLayoutErrorMsgTxt));
	PV_EditLayoutErrorMsgTxt.click();
	PV_EditLayoutErrorMsgTxt.sendKeys("Enter Wrong Data in the CustHeader Field");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutApplyBtn));
	editLayoutApplyBtn.click();
	Thread.sleep(1000);
	String expMessage="Field saved successfully";
	String actMessage=	checkValidationMessage(expMessage);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_editLayoutAddFieldCloseBtn));
	PV_editLayoutAddFieldCloseBtn.click();
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCloseBtn));
	editLayoutCloseBtn.click();
	Thread.sleep(4000);
	
	
	int actHeaderListAfterAdd=PV_HeaderList.size();
	int expHeaderListAfterAdd=16;
	
	System.out.println("Actual Added Field Count "	+ actHeaderListAfterAdd);
	System.out.println("Expected Added Field Count "	+ expHeaderListAfterAdd);

	ArrayList<String> headerArray=new ArrayList<String>();
	
	for(int i=0;i<actHeaderListAfterAdd;i++)
	{
		String data=PV_HeaderList.get(i).getText();
		headerArray.add(data);
	}
	
	String actAddedFieldDisplayafterAdd=headerArray.toString();
	String expAddedFieldDisplayafterAdd="[Document No., Date, , Vendor Account, , , Due Date, Currency, Exchange Rate, Department, LocExchangeRate, Place of supply, Jurisdiction, Narration, Permit No, CustHeader]";
	
	System.out.println("Actual Added Field after Delete "	+ actAddedFieldDisplayafterAdd);
	System.out.println("Expected Added Field after Delete "	+ expAddedFieldDisplayafterAdd);
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_AddedHeaderField));
	
	boolean actAddedFieldDisplay=PV_AddedHeaderField.isDisplayed();
	boolean expAddedFieldDisplay=true;
	
	System.out.println("Actual Added Field Displayed "	+ actAddedFieldDisplay);
	System.out.println("Expected Added Field Displayed "	+ expAddedFieldDisplay);
	
	//checking Regular Expression other than Focus
	PV_AddedHeaderField.click();
	Thread.sleep(1000);
	PV_AddedHeaderField.sendKeys("focus");
	Thread.sleep(1000);
	PV_AddedHeaderField.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	checkValidationMessage("Enter Wrong Data in the CustHeader Field");
	
	
	if(actHeaderLbl.equalsIgnoreCase(expHeaderLbl) && actMessage.equalsIgnoreCase(expMessage)
			&& actHeaderListAfterAdd==expHeaderListAfterAdd && actAddedFieldDisplayafterAdd.equalsIgnoreCase(expAddedFieldDisplayafterAdd)
			&&actAddedFieldDisplay==expAddedFieldDisplay)
	{
	
	return true;
	}
	else
	{
		return false;
	}
	
}

public boolean checkEdittheCreatedFieldinPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_HeaderCustBtn));
	PV_HeaderCustBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_EditLayoutHeader3rdRowEditBtn));
	PV_EditLayoutHeader3rdRowEditBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCaptionTxt));
	editLayoutCaptionTxt.click();
	editLayoutCaptionTxt.clear();
	editLayoutCaptionTxt.sendKeys("CustHeaderUpdated");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutApplyBtn));
	editLayoutApplyBtn.click();
	Thread.sleep(1000);
	String expMessage="Field saved successfully";
	String actMessage=	checkValidationMessage(expMessage);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_editLayoutAddFieldCloseBtn));
	PV_editLayoutAddFieldCloseBtn.click();
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCloseBtn));
	editLayoutCloseBtn.click();
	Thread.sleep(4000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_AddedHeaderFieldLabel));
	String actUpdatedFieldName=PV_AddedHeaderFieldLabel.getText();
	String expUpdatedFieldName="CustHeaderUpdated";
	
	System.out.println("Actual Added Field Displayed "	+ actUpdatedFieldName);
	System.out.println("Expected Added Field Displayed "	+ expUpdatedFieldName);
	
	
	if(actUpdatedFieldName.equalsIgnoreCase(expUpdatedFieldName) &&actMessage.equalsIgnoreCase(expMessage))
	
		{
		return true;
		}
	else
	{
		return false;
	}
	
}


public boolean checkDeletetheCreatedFieldatPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_HeaderCustBtn));
	PV_HeaderCustBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_EditLayoutHeader3rdRowDelBtn));
	PV_EditLayoutHeader3rdRowDelBtn.click();
	Thread.sleep(2000);
	
	getWaitForAlert();
	Thread.sleep(2000);
	
	getAlert().accept();
	
	
	String expMessage="Field deleted successfully";
	String actMessage=checkValidationMessage(expMessage);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCloseBtn));
	editLayoutCloseBtn.click();
	Thread.sleep(4000);
	
	
	int actHeaderListAfterDel=PV_HeaderList.size();
	int expHeaderListAfterDel=15;

	ArrayList<String> headerArray=new ArrayList<String>();
	
	for(int i=0;i<actHeaderListAfterDel;i++)
	{
		String data=PV_HeaderList.get(i).getText();
		headerArray.add(data);
	}
	
	String actAddedFieldDisplayafterDel=headerArray.toString();
	String expAddedFieldDisplayafterDel="[Document No., Date, , Vendor Account, , , Due Date, Currency, Exchange Rate, Department, LocExchangeRate, Place of supply, Jurisdiction, Narration, Permit No]";
	
	System.out.println("Actual Added Field after Delete "	+ actAddedFieldDisplayafterDel);
	System.out.println("Expected Added Field after Delete "	+ expAddedFieldDisplayafterDel);
	
	if(actMessage.equalsIgnoreCase(expMessage) && actAddedFieldDisplayafterDel.equalsIgnoreCase(expAddedFieldDisplayafterDel))
	
		{return true;
		
		}
	else
	{
		return false;
	}
}


//@FindBy(xpath="//table[@id='formulaGrideditScreen_FieldsCustomization_PreLoad']//tr//td[3]")
@FindBy(xpath="//table[@id='formulaGrideditScreen_FieldsCustomization_PreLoad']//tr//td[3][@class='FGridColumn  ']")
public static List<WebElement> PV_formulaList;

@FindBy(xpath="//select[@id='editScreen_FieldsCustomization_behaviour']")
public static WebElement PV_BehaviourTxt;

@FindBy(xpath="//*[@id='editScreen_FieldsCustomization_Formula_textbox']")
public static WebElement  PV_FormulaTxt;

@FindBy(xpath="//*[@id='editScreen_FieldsCustomization_Formula_formulaText']")
public static WebElement  PV_FormulaformulaTxt;


@FindBy(xpath="//*[@id='editScreen_FieldsCustomization_Formula_Ok']")
public static WebElement  PV_FormulaOkBtn;


@FindBy(xpath="//table[@id='id_transaction_entry_detail_table']//tr[1]//th[10]")
public static WebElement  PV_AddedBodyField;


@FindBy(xpath="//*[@id='formulaGrideditScreen_FieldsCustomization_PreLoad_col_10-2']")
public static WebElement PV_formula;


@FindBy(xpath="//*[@id='id_global_externalmodule_container']")
public static WebElement PV_frame;



@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_head']//th//div[1][@class='no-select']")
public static List<WebElement>  PV_BodyFieldsList;


@FindBy(xpath="//*[@id='id_transactionentry_footer_panel_entry']//div//label")
public static List<WebElement>  PV_FooterFieldsList;


public boolean checkCustomisationScreenBodyatPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_BodyCustBtn));
	PV_BodyCustBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_CustHeaderLabel));
	String actHeaderLbl=PV_CustHeaderLabel.getText();
	String expHeaderLbl="Customize Body";
	
	System.out.println("Actual Body Label :	" +	actHeaderLbl);
	System.out.println("Expect Body Label :	" +	expHeaderLbl);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_EditLayoutHeaderAddFieldBtn));
	PV_EditLayoutHeaderAddFieldBtn.click();
	Thread.sleep(1000);
	
	if(getIsAlertPresent()){
		getWaitForAlert();
		Thread.sleep(1000);
		getAlert().accept();
		Thread.sleep(1000);
	}
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenCaptionTxt));
	editScreenCaptionTxt.click();
	editScreenCaptionTxt.sendKeys("CustBody");
	Thread.sleep(1000);
	
		
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_BehaviourTxt));
	Select s=new Select(PV_BehaviourTxt);
	s.selectByVisibleText("Formula");
	PV_BehaviourTxt.sendKeys(Keys.TAB);
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_FormulaTxt));
	PV_FormulaTxt.click();
	Thread.sleep(2000);
	
	ArrayList<String> formulaArray=new ArrayList<String>();

	int rows=getDriver().findElements(By.xpath("//*[@id='formulaGrideditScreen_FieldsCustomization_Formula']//tr")).size();
	

	String beforePath="//*[@id='formulaGrideditScreen_FieldsCustomization_Formula']//tr[";
	String afterPath="]//td[3][@class='FGridColumn  ']";
	
	
	for( int j=1;j<rows-2;j++)
	{
		String data=(getDriver().findElement(By.xpath(beforePath + j + afterPath))).getText();
		
			if(data.isEmpty()==false)
			{
			formulaArray.add(data);
			}
		
	}
		
	
	int actFormulaListCount=PV_formulaList.size();
	int expFormulaListCount=107;
	
		
	String actFormulaArrayList=formulaArray.toString();
	String expFormulaArrayList="[Transaction, Quantity, Alternate Qty, Gross Amt, Stock Value, Account, Account 2, Exchange Rate, Qty in base unit, Currency, Value of the tag(MasterId of the tag), Net Amount, Net Amount in original currency, Total Quantity, Total Gross, ScreenField, Discount Input, Discount Calculated Value, FD% Input, FD% Calculated Value, RD Input, RD Calculated Value, VAT advance Input, VAT advance Calculated Value, Avg Rate Input, Avg Rate Calculated Value, Avg Rate(O) Input, Avg Rate(O) Calculated Value, VAT Input, VAT Calculated Value, Taxable Input, Taxable Calculated Value, LayoutField, Narration Input, Tax Code Input, Permit No Input, MasterField, Name, Code, Group, Status, Do Not Restrict, CreatedBy, Modified By, Created Date, Modified Date, iLocationId, Edited From, Closing Date, Allow Other Companies To View Record, Sync Received Date, Editing Location, Do not maintain bin, Currency, Bins, Phys Inventory Counting Frequency, Warehouse Type, Name, Code, Group, Status, Do Not Restrict, CreatedBy, Modified By, Created Date, Modified Date, iLocationId, Edited From, Closing Date, Allow Other Companies To View Record, Sync Received Date, Editing Location, Local Currency, Jurisdiction, Name, Code, Group, Status, Do Not Restrict, CreatedBy, Modified By, Created Date, Modified Date, iLocationId, Edited From, Closing Date, Allow Other Companies To View Record, Sync Received Date, Editing Location, Taxable, Name, Code, Group, Status, Do Not Restrict, CreatedBy, Modified By, Created Date, Modified Date, iLocationId, Edited From, Closing Date, Allow Other Companies To View Record, Sync Received Date, Editing Location]";
	
	
		
	System.out.println("Actual Formula List :	" +	actFormulaArrayList);
	System.out.println("Expect Formula List :	" +	expFormulaArrayList);
	
	System.out.println("Actual Formula List Count :	" +	actFormulaListCount);
	System.out.println("Expect Formula List Count:	" +	expFormulaListCount);
	
	PV_FormulaTxt.sendKeys("qty");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_FormulaOkBtn));
	PV_FormulaOkBtn.click();
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
	editScreenApplyBtn.click();
	Thread.sleep(1000);
	String expMessage="Field saved successfully";
	
	String actMessage=checkValidationMessage(expMessage);
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_editLayoutAddFieldCloseBtn));
	PV_editLayoutAddFieldCloseBtn.click();
	Thread.sleep(1000);
	*/
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCloseBtn));
	editLayoutCloseBtn.click();
	Thread.sleep(4000);
	
	
	
	int actHeaderListAfterDel=PV_BodyFieldsList.size();
	int expHeaderListAfterDel=21;

	ArrayList<String> bodyArray=new ArrayList<String>();
	
	int rows1=getDriver().findElements(By.xpath("//tr[@id='id_transaction_entry_detail_table_row_heading']//th")).size();
		
	System.err.println(rows1);

		String beforePath1="//tr[@id='id_transaction_entry_detail_table_row_heading']//th[";
		String afterPath1="]";
		
		
		for( int k=2;k<=rows1;k++)
		{
				
			String data=(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).getAttribute("title");
			
				System.err.println(data);
				if(data.isEmpty()==false)
				{
					bodyArray.add(data);
				}
				
				getAction().moveToElement(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).build().perform();
				Thread.sleep(1500);
			
		}
			

	String actAddedFieldDisplayafterDel=bodyArray.toString();
	String expAddedFieldDisplayafterDel="[Warehouse, Item, Tax Code, Purchase Account, Units, RD, Avg Rate, Avg Rate(O), CustBody, Quantity, L-Purchases Orders, Rate, Gross, Discount, VAT, Taxable, Batch, Bins, Expiry Date, RMA]";
	
	System.out.println("Actual Added Field  "	+ actAddedFieldDisplayafterDel);
	System.out.println("Expected Added Field  "	+ expAddedFieldDisplayafterDel);
	
	
	
	
	if(actHeaderLbl.equalsIgnoreCase(expHeaderLbl)
			&& actFormulaArrayList.equalsIgnoreCase(expFormulaArrayList) 
			&& actAddedFieldDisplayafterDel.equalsIgnoreCase(expAddedFieldDisplayafterDel))
	{
	
	return true;
	}
	else
	{
		return false;
	}
	
		
}

public boolean checkEditCustomisationScreenBodyatPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_BodyCustBtn));
	PV_BodyCustBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_EditScreenBody7thRowEditBtn));
	PV_EditScreenBody7thRowEditBtn.click();
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenCaptionTxt));
	editScreenCaptionTxt.click();
	editScreenCaptionTxt.clear();
	editScreenCaptionTxt.sendKeys("CustBodyUpdate");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
	editScreenApplyBtn.click();
	Thread.sleep(1000);
	checkValidationMessage("Data saved successfully");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCloseBtn));
	editLayoutCloseBtn.click();
	Thread.sleep(4000);
	
	int actBodyListAfterEdit=PV_BodyFieldsList.size();
	int expBodyListAfterEdit=20;

	System.out.println("Actual List  after Edit Added Field "	+ actBodyListAfterEdit);
	System.out.println("Expected List after Edit Added Field "	+ expBodyListAfterEdit);
	
	
	ArrayList<String> bodyArray=new ArrayList<String>();
	int rows1=getDriver().findElements(By.xpath("//tr[@id='id_transaction_entry_detail_table_row_heading']//th")).size();
		
		String beforePath1="//tr[@id='id_transaction_entry_detail_table_row_heading']//th[";
		String afterPath1="]";
	
		
		for( int k=2;k<=rows1;k++)
		{
				
			String data=(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).getAttribute("title");
			
			System.err.println(data);
				if(data.isEmpty()==false)
				{
					bodyArray.add(data);
				}
				
		//			getAction().moveToElement(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).build().perform();
				Thread.sleep(1500);
			
		}
			
		
	
	
	
	String actAddedFieldDisplayafterDel=bodyArray.toString();
	String expAddedFieldDisplayafterDel="[Warehouse, Item, Tax Code, Purchase Account, Units, RD, Avg Rate, Avg Rate(O), CustBodyUpdate, Quantity, L-Purchases Orders, Rate, Gross, Discount, VAT, Taxable, Batch, Bins, Expiry Date, RMA]";
	
	System.out.println("Actual Added Field after Edit "	+ actAddedFieldDisplayafterDel);
	System.out.println("Expected Added Field after Edit "	+ expAddedFieldDisplayafterDel);
	
	
	
	if(actAddedFieldDisplayafterDel.equalsIgnoreCase(expAddedFieldDisplayafterDel))
	{
	return true;
	}
	else
	{
		return false;
	}
	
	
}


public boolean checkDeleteAddedFieldinBodyatPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_BodyCustBtn));
	PV_BodyCustBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_EditScreenBody7thRowDelBtn));
	PV_EditScreenBody7thRowDelBtn.click();
	Thread.sleep(2000);
	
	getWaitForAlert();
	Thread.sleep(2000);
	
	getAlert().accept();
	
	String expMessage="Data deleted successfully";
	String actMessage=checkValidationMessage(expMessage);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCloseBtn));
	editLayoutCloseBtn.click();
	Thread.sleep(4000);
	
	
	int actBodyListAfterDel=PV_BodyFieldsList.size();
	int expBodyListAfterDel=20;
	
	System.out.println("Actual Added Field List after Delete "	+ actBodyListAfterDel);
	System.out.println("Expected Added Field List after Delete "	+ actBodyListAfterDel);

	ArrayList<String> bodyArray=new ArrayList<String>();
	int rows1=getDriver().findElements(By.xpath("//tr[@id='id_transaction_entry_detail_table_row_heading']//th")).size();
		
		String beforePath1="//tr[@id='id_transaction_entry_detail_table_row_heading']//th[";
		String afterPath1="]";
	
		
		for( int k=2;k<=rows1;k++)
		{
				
			String data=(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).getAttribute("title");
			
			System.err.println(data);
				if(data.isEmpty()==false)
				{
					bodyArray.add(data);
				}
				
		//			getAction().moveToElement(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).build().perform();
				Thread.sleep(1500);
			
		}
		
		String actAddedFieldDisplayafterDel=bodyArray.toString();
		String expAddedFieldDisplayafterDel="[Warehouse, Item, Tax Code, Purchase Account, Units, RD, Avg Rate, Avg Rate(O), Quantity, L-Purchases Orders, Rate, Gross, Discount, VAT, Taxable, Batch, Bins, Expiry Date, RMA]";
		
	System.out.println("Actual Added Field after Delete "	+ actAddedFieldDisplayafterDel);
	System.out.println("Expected Added Field after Delete "	+ expAddedFieldDisplayafterDel);
	
	if(actMessage.equalsIgnoreCase(expMessage) && actAddedFieldDisplayafterDel.equalsIgnoreCase(expAddedFieldDisplayafterDel))
	
		{return true;
		
		}
	else
	{
		return false;
	}
	

}



public boolean checkCustomizationLayoutBodyatPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_BodyCustBtn));
	PV_BodyCustBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_CustHeaderLabel));
	String actHeaderLbl=PV_CustHeaderLabel.getText();
	String expHeaderLbl="Customize Body";
	
	System.out.println("Actual Body Label :	" +	actHeaderLbl);
	System.out.println("Expect Body Label :	" +	expHeaderLbl);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_EditScreenLayoutBodyAddFieldBtn));
	PV_EditScreenLayoutBodyAddFieldBtn.click();
	Thread.sleep(1000);
	
	if(getIsAlertPresent()){
		getWaitForAlert();
		Thread.sleep(1000);
		getAlert().accept();
		Thread.sleep(1000);
	}
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCaptionTxt));
	editLayoutCaptionTxt.click();
	editLayoutCaptionTxt.sendKeys("CustBodyLayout");
	Thread.sleep(1000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutApplyBtn));
	editLayoutApplyBtn.click();
	Thread.sleep(1000);
	String expMessage="Field saved successfully";
	String actMessage=	checkValidationMessage(expMessage);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_editLayoutAddFieldCloseBtn));
	PV_editLayoutAddFieldCloseBtn.click();
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCloseBtn));
	editLayoutCloseBtn.click();
	Thread.sleep(4000);
	
	
	
	int actHeaderListAfterDel=PV_BodyFieldsList.size();
	int expHeaderListAfterDel=20;

ArrayList<String> bodyArray=new ArrayList<String>();
	
	int rows1=getDriver().findElements(By.xpath("//tr[@id='id_transaction_entry_detail_table_row_heading']//th")).size();
		
	System.err.println(rows1);

		String beforePath1="//tr[@id='id_transaction_entry_detail_table_row_heading']//th[";
		String afterPath1="]";
		
		
		for( int k=2;k<=rows1;k++)
		{
				
			String data=(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).getAttribute("title");
			
				System.err.println(data);
				if(data.isEmpty()==false)
				{
					bodyArray.add(data);
				}
				
				getAction().moveToElement(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).build().perform();
				Thread.sleep(1500);
			
		}
			

	String actAddedFieldDisplayafterDel=bodyArray.toString();
	String expAddedFieldDisplayafterDel="[Warehouse, Item, Tax Code, Purchase Account, Units, RD, Avg Rate, Avg Rate(O), Quantity, L-Purchases Orders, Rate, Gross, Discount, VAT, Taxable, Batch, Bins, Expiry Date, RMA, CustBodyLayout]";
	
	
	System.out.println("Actual Added Field  "	+ actAddedFieldDisplayafterDel);
	System.out.println("Expected Added Field  "	+ expAddedFieldDisplayafterDel);
	
	
	
	if(actHeaderLbl.equalsIgnoreCase(expHeaderLbl) && actMessage.equalsIgnoreCase(expMessage)
			&& actAddedFieldDisplayafterDel.equalsIgnoreCase(expAddedFieldDisplayafterDel))
	{
	
	return true;
	}
	else
	{
		return false;
	}
	
	
	
	}



public boolean checkEditCreatedLayoutBodyatPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_BodyCustBtn));
	PV_BodyCustBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_EditLayoutBody2ndRowEditBtn));
	PV_EditLayoutBody2ndRowEditBtn.click();
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCaptionTxt));
	editLayoutCaptionTxt.click();
	editLayoutCaptionTxt.clear();
	editLayoutCaptionTxt.sendKeys("CustBodyLayoutUpdate");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutApplyBtn));
	editLayoutApplyBtn.click();
	Thread.sleep(1000);
	String expMessage="Field saved successfully";
	String actMessage=	checkValidationMessage(expMessage);
	
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_editLayoutAddFieldCloseBtn));
	PV_editLayoutAddFieldCloseBtn.click();
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCloseBtn));
	editLayoutCloseBtn.click();
	Thread.sleep(4000);
	
	int actHeaderListAfterDel=PV_BodyFieldsList.size();
	int expHeaderListAfterDel=21;

ArrayList<String> bodyArray=new ArrayList<String>();
	
	int rows1=getDriver().findElements(By.xpath("//tr[@id='id_transaction_entry_detail_table_row_heading']//th")).size();
		
	System.err.println(rows1);

		String beforePath1="//tr[@id='id_transaction_entry_detail_table_row_heading']//th[";
		String afterPath1="]";
		
		
		for( int k=2;k<=rows1;k++)
		{
				
			String data=(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).getAttribute("title");
			
				System.err.println(data);
				if(data.isEmpty()==false)
				{
					bodyArray.add(data);
				}
				
				getAction().moveToElement(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).build().perform();
				Thread.sleep(1500);
			
		}
			

	String actAddedFieldDisplayafterDel=bodyArray.toString();
	String expAddedFieldDisplayafterDel="[Warehouse, Item, Tax Code, Purchase Account, Units, RD, Avg Rate, Avg Rate(O), Quantity, L-Purchases Orders, Rate, Gross, Discount, VAT, Taxable, Batch, Bins, Expiry Date, RMA, CustBodyLayoutUpdate]";
	
	System.out.println("Actual Added Field  "	+ actAddedFieldDisplayafterDel);
	System.out.println("Expected Added Field  "	+ expAddedFieldDisplayafterDel);
	
	if(actAddedFieldDisplayafterDel.equalsIgnoreCase(expAddedFieldDisplayafterDel))
	{
	return true;
	}
	else
	{
		return false;
	}
	
}



public boolean checkDeleteAddedFieldinLayoutBodyatPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_BodyCustBtn));
	PV_BodyCustBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_EditLayoutBody2ndRowDelBtn));
	PV_EditLayoutBody2ndRowDelBtn.click();
	Thread.sleep(2000);
	
	getWaitForAlert();
	Thread.sleep(2000);
	
	getAlert().accept();
	
	
	String expMessage="Field deleted successfully";
	String actMessage=checkValidationMessage(expMessage);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCloseBtn));
	editLayoutCloseBtn.click();
	Thread.sleep(4000);
	
	
	int actHeaderListAfterDel=PV_BodyFieldsList.size();
	int expHeaderListAfterDel=15;

ArrayList<String> bodyArray=new ArrayList<String>();
	
	int rows1=getDriver().findElements(By.xpath("//tr[@id='id_transaction_entry_detail_table_row_heading']//th")).size();
		
	System.err.println(rows1);

		String beforePath1="//tr[@id='id_transaction_entry_detail_table_row_heading']//th[";
		String afterPath1="]";
		
		
		for( int k=2;k<=rows1;k++)
		{
				
			String data=(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).getAttribute("title");
			
				System.err.println(data);
				if(data.isEmpty()==false)
				{
					bodyArray.add(data);
				}
				
				getAction().moveToElement(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).build().perform();
				Thread.sleep(1500);
			
		}
			

	String actAddedFieldDisplayafterDel=bodyArray.toString();
	String expAddedFieldDisplayafterDel="[Warehouse, Item, Tax Code, Purchase Account, Units, RD, Avg Rate, Avg Rate(O), Quantity, L-Purchases Orders, Rate, Gross, Discount, VAT, Taxable, Batch, Bins, Expiry Date, RMA]";
	
	
	System.out.println("Actual Added Field  "	+ actAddedFieldDisplayafterDel);
	System.out.println("Expected Added Field  "	+ expAddedFieldDisplayafterDel);
	
	
	if(actMessage.equalsIgnoreCase(expMessage) && actAddedFieldDisplayafterDel.equalsIgnoreCase(expAddedFieldDisplayafterDel))
	
		{return true;
		
		}
	else
	{
		return false;
	}
}

public boolean checkCustomizationFooteratPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_FooterCustBtn));
	PV_FooterCustBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_CustHeaderLabel));
	String actHeaderLbl=PV_CustHeaderLabel.getText();
	String expHeaderLbl="Customize Footer";
	
	System.out.println("Actual Body Label :	" +	actHeaderLbl);
	System.out.println("Expect Body Label :	" +	expHeaderLbl);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_EditLayoutHeaderAddFieldBtn));
	PV_EditLayoutHeaderAddFieldBtn.click();
	Thread.sleep(1000);
	
	if(getIsAlertPresent()){
		getWaitForAlert();
		Thread.sleep(1000);
		getAlert().accept();
		Thread.sleep(1000);
	}
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenCaptionTxt));
	editScreenCaptionTxt.click();
	editScreenCaptionTxt.sendKeys("CustFooter");
	Thread.sleep(1000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
	editScreenApplyBtn.click();
	Thread.sleep(1000);
	String expMessage="Data saved successfully";
	String actMessage=	checkValidationMessage(expMessage);
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_editLayoutAddFieldCloseBtn));
	PV_editLayoutAddFieldCloseBtn.click();
	Thread.sleep(1000);*/
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCloseBtn));
	editLayoutCloseBtn.click();
	Thread.sleep(4000);
	
	
	
	int actHeaderListAfterDel=PV_FooterFieldsList.size();
	int expHeaderListAfterDel=20;

ArrayList<String> bodyArray=new ArrayList<String>();
	
	int rows1=getDriver().findElements(By.xpath("//*[@id='id_transactionentry_footer_panel_entry']//div//label")).size();
		
	System.err.println(rows1);

		String beforePath1="//*[@id='id_transactionentry_footer_panel_entry']//div[";
		String afterPath1="]/label";
		
		
		for( int k=1;k<=rows1;k++)
		{
				
			String data=(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).getText();
			
				System.err.println(data);
				if(data.isEmpty()==false)
				{
					bodyArray.add(data);
				}
				
				getAction().moveToElement(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).build().perform();
				Thread.sleep(1500);
			
		}
			

	String actAddedFieldDisplayafterDel=bodyArray.toString();
	String expAddedFieldDisplayafterDel="[FD%, VAT advance, CustFooter]";
	
	
	System.out.println("Actual Added Footer Field  "	+ actAddedFieldDisplayafterDel);
	System.out.println("Expected Added Footer Field  "	+ expAddedFieldDisplayafterDel);
	
	
	
	if(actHeaderLbl.equalsIgnoreCase(expHeaderLbl) && actMessage.equalsIgnoreCase(expMessage)
			&& actAddedFieldDisplayafterDel.equalsIgnoreCase(expAddedFieldDisplayafterDel))
	{
	
	return true;
	}
	else
	{
		return false;
	}
	

	}


public boolean checkEditCustomizedFooterFieldatPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_FooterCustBtn));
	PV_FooterCustBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_EditScreenFooter3rdRowEditBtn));
	PV_EditScreenFooter3rdRowEditBtn.click();
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenCaptionTxt));
	editScreenCaptionTxt.click();
	editScreenCaptionTxt.clear();
	editScreenCaptionTxt.sendKeys("CustFooterUpdate");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
	editScreenApplyBtn.click();
	Thread.sleep(1000);
	checkValidationMessage("Data saved successfully");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCloseBtn));
	editLayoutCloseBtn.click();
	Thread.sleep(4000);
	
	int actBodyListAfterEdit=PV_FooterFieldsList.size();
	int expBodyListAfterEdit=20;

	System.out.println("Actual List  after Edit Added Field "	+ actBodyListAfterEdit);
	System.out.println("Expected List after Edit Added Field "	+ expBodyListAfterEdit);
	
	
ArrayList<String> bodyArray=new ArrayList<String>();
	
	int rows1=getDriver().findElements(By.xpath("//*[@id='id_transactionentry_footer_panel_entry']//div//label")).size();
		
	System.err.println(rows1);

		String beforePath1="//*[@id='id_transactionentry_footer_panel_entry']//div[";
		String afterPath1="]/label";
		
		
		for( int k=1;k<=rows1;k++)
		{
				
			String data=(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).getText();
			
				System.err.println(data);
				if(data.isEmpty()==false)
				{
					bodyArray.add(data);
				}
				
				getAction().moveToElement(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).build().perform();
				Thread.sleep(1500);
			
		}
			

	String actAddedFieldDisplayafterDel=bodyArray.toString();
	String expAddedFieldDisplayafterDel="[FD%, VAT advance, CustFooterUpdate]";
	
	
	
	
	
	System.out.println("Actual Added Field after Edit "	+ actAddedFieldDisplayafterDel);
	System.out.println("Expected Added Field after Edit "	+ expAddedFieldDisplayafterDel);
	
	
	
	if(actAddedFieldDisplayafterDel.equalsIgnoreCase(expAddedFieldDisplayafterDel))
	{
	return true;
	}
	else
	{
		return false;
	}
	
	

}

public boolean checkDeleteAddedFieldinFooteratPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{


	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_FooterCustBtn));
	PV_FooterCustBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PV_EditScreenFooter3rdRowDelBtn));
	PV_EditScreenFooter3rdRowDelBtn.click();
	Thread.sleep(2000);
	
	getWaitForAlert();
	Thread.sleep(2000);
	
	getAlert().accept();
	
	String expMessage="Data deleted successfully";
	String actMessage=checkValidationMessage(expMessage);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCloseBtn));
	editLayoutCloseBtn.click();
	Thread.sleep(4000);
	
	
	int actBodyListAfterDel=PV_FooterFieldsList.size();
	int expBodyListAfterDel=20;
	
	System.out.println("Actual Added Field List after Delete "	+ actBodyListAfterDel);
	System.out.println("Expected Added Field List after Delete "	+ actBodyListAfterDel);

ArrayList<String> bodyArray=new ArrayList<String>();
	
	int rows1=getDriver().findElements(By.xpath("//*[@id='id_transactionentry_footer_panel_entry']//div//label")).size();
		
	System.err.println(rows1);

		String beforePath1="//*[@id='id_transactionentry_footer_panel_entry']//div[";
		String afterPath1="]/label";
		
		
		for( int k=1;k<=rows1;k++)
		{
				
			String data=(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).getText();
			
				System.err.println(data);
				if(data.isEmpty()==false)
				{
					bodyArray.add(data);
				}
				
				getAction().moveToElement(getDriver().findElement(By.xpath(beforePath1 + k + afterPath1))).build().perform();
				Thread.sleep(1500);
			
		}
			

	String actAddedFieldDisplayafterDel=bodyArray.toString();
	String expAddedFieldDisplayafterDel="[FD%, VAT advance]";
	
	
	
	
	System.out.println("Actual Added Field after Delete "	+ actAddedFieldDisplayafterDel);
	System.out.println("Expected Added Field after Delete "	+ expAddedFieldDisplayafterDel);
	
	if(actMessage.equalsIgnoreCase(expMessage) && actAddedFieldDisplayafterDel.equalsIgnoreCase(expAddedFieldDisplayafterDel))
	
		{return true;
		
		}
	else
	{
		return false;
	}
	


}


	public PurchaseVouchersVATPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
