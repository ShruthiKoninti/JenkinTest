package com.focus.Pages;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashSet;
import java.util.List;

import org.apache.pdfbox.contentstream.operator.graphics.MoveTo;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.Alert;
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
import com.sun.tools.xjc.reader.ExtensionBindingChecker;
import com.testautomationguru.utility.PDFUtil;

public class BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage extends BaseEngine
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "BuyerAndSellerValidations";
	
	
	@FindBy (xpath="//div[@id='id_focus_msgbox_main']")
	private static WebElement ServerErrorPopup;

	@FindBy (xpath="//div[@id='id_focus_msgbox_title']/div[2]/span")
	private static WebElement ServerErrorPopupCloseBtn;

	@FindBy (xpath="//*[@id='id_focus_msgbox_detail']/span/h2/i")
	private static WebElement ServerErrorPopupTxt;
	
	public static void CloseServerError() throws InterruptedException
	{
		try 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ServerErrorPopupCloseBtn));
			ServerErrorPopupCloseBtn.click();
		} 
		catch (Exception e) 
		{
			System.err.println("**************NO SERVER ERROR EXIST**************");
		}
	}
	
	public static boolean checkSavingPurchaseVoucherVATVoucher1WithoutRateFormula() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
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
		
		
		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		//documentNumberTxt.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		//vendorAccountTxt.click();
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 8, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
					voucherHeaderDueDate.click();
					voucherHeaderDueDate.sendKeys(Keys.TAB);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
					voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
					voucherHeaderCurrency.sendKeys(Keys.SPACE);

					int currencycount=currencyListCount.size();

					System.err.println(currencycount);

					for(int i=0 ; i < currencycount ;i++)
					{
						String data=currencyListCount.get(i).getText();

						if (data.equalsIgnoreCase("INR")) 
						{
							currencyListCount.get(i).click();
							break;
						}
					}
					voucherHeaderCurrency.sendKeys(Keys.TAB);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 14, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		String actRate1 = select1stRow_11thColumn.getText();
		String expRate1 = excelReader.getCellData(xlSheetName, 8, 6);
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actRate1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
		select1stRow_12thColumn.click();
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
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 15, 5));
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


		// Row 2 with Item Group item

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(Keys.END);
		enter_WarehouseTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 17, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
		select2ndRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 18, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_11thColumn));
		String actRate2 = select2ndRow_11thColumn.getText();
		String expRate2 = excelReader.getCellData(xlSheetName, 9, 6);
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, expRate2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_12thColumn));
		select2ndRow_12thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_13thColumn));
		select2ndRow_13thColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();

		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);
		
		Thread.sleep(2000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
		rmaSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 19, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
		rmaAddBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
		RMAPopupOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));

		String actRMAColumn=enter_RMA.getAttribute("value");

		String expRMAColumn=excelReader.getCellData(xlSheetName, 10, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 10, 7, actRMAColumn);

		System.out.println(actRMAColumn);
		System.out.println(expRMAColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		Thread.sleep(2000);
		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i--) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}*/
		
		String actMessage = SavingInBackground(1);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : SU/AME/TEXT1");
		expMsg.add("Saving in background.");*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 17, 6);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 17, 7, actMessage);

		System.out.println("************************************************************************");

		System.out.println("Rates1 : "+actRate1+"  Value Expected  "+expRate1);
		//System.out.println("Rates2 : "+actRate2+"  Value Expected  "+expRate2);

		if(actRate1.equalsIgnoreCase(expRate1) /*&& actRate2.equalsIgnoreCase(expRate2)*/ 
				&& actMessage.equals(expMessage))
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







	public static boolean checkVerifyingPurchaseVoucherVATVoucher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 20, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("");

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 20, 6);
		excelReader.setCellData(xlfile, xlSheetName, 20, 7, actDocNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 21, 6);
		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actVendor);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 22, 6);
		excelReader.setCellData(xlfile, xlSheetName, 22, 7, actCurrency);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 23, 6);
		excelReader.setCellData(xlfile, xlSheetName, 23, 7, actDepartment);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 24, 6);
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actPlaceOfSupply);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 25, 6);
		excelReader.setCellData(xlfile, xlSheetName, 25, 7, actJurisdiction);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 26, 6);
		excelReader.setCellData(xlfile, xlSheetName, 26, 7, actR1Warehouse);
		
		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 27, 6);
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actR1Item);

		String actR1TaxCode = select1stRow_3rdColumn.getText();
		String expR1TaxCode = excelReader.getCellData(xlSheetName, 28, 6);
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actR1TaxCode);
		
		String actPurchaseAccount = select1stRow_4thColumn.getText();
		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 29, 6);
		excelReader.setCellData(xlfile, xlSheetName, 29, 7, actPurchaseAccount);

		String actUnits = select1stRow_5thColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 30, 6);
		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actUnits);

		String actQuantity = select1stRow_9thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 31, 6);
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actQuantity);

		String actRate = select1stRow_11thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 32, 6);
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actRate);

		String actGross = select1stRow_12thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 33, 6);
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actGross);

		String actBatch = select1stRow_16thColumn.getText();
		String expBatch = excelReader.getCellData(xlSheetName, 34, 6);
		excelReader.setCellData(xlfile, xlSheetName, 34, 7, actBatch);

		String actExpiry = select1stRow_18thColumn.getText();
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		Calendar cal=Calendar.getInstance();
		cal.add(Calendar.DATE, 2); 
		String expExpiry=df.format(cal.getTime());



		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 35, 6);
		excelReader.setCellData(xlfile, xlSheetName, 35, 7, actR2Warehouse);

		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 36, 6);
		excelReader.setCellData(xlfile, xlSheetName, 36, 7, actR2Item);

		String actR2TaxCode = select2ndRow_3rdColumn.getText();
		String expR2TaxCode = excelReader.getCellData(xlSheetName, 37, 6);
		excelReader.setCellData(xlfile, xlSheetName, 37, 7, actR2TaxCode);

		String actR2PurchaseAccount = select2ndRow_4thColumn.getText();
		String expR2PurchaseAccount = excelReader.getCellData(xlSheetName, 38, 6);
		excelReader.setCellData(xlfile, xlSheetName, 38, 7, actR2PurchaseAccount);

		String actR2Units = select2ndRow_5thColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 39, 6);
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actR2Units);

		String actR2Quantity = select2ndRow_9thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 40, 6);
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, actR2Quantity);

		String actR2Rate = select2ndRow_11thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 41, 6);
		excelReader.setCellData(xlfile, xlSheetName, 41, 7, actR2Rate);

		String actR2Gross = select2ndRow_12thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 42, 6);
		excelReader.setCellData(xlfile, xlSheetName, 42, 7, actR2Gross);

		String actR2RMA = select2ndRow_19thColumn.getText();
		String expR2RMA = excelReader.getCellData(xlSheetName, 43, 6);
		excelReader.setCellData(xlfile, xlSheetName, 43, 7, actR2RMA);

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
		System.out.println("PurchaseAccount : "+actPurchaseAccount	    +"  Value Expected  "+expPurchaseAccount);
		System.out.println("Units           : "+actUnits			    +"  Value Expected  "+expUnits);
		System.out.println("Quantity        : "+actQuantity			    +"  Value Expected  "+expQuantity);
		System.out.println("Rate            : "+actRate				    +"  Value Expected  "+expRate);
		System.out.println("Gross           : "+actGross			    +"  Value Expected  "+expGross);
		System.out.println("Batch           : "+actBatch			    +"  Value Expected  "+expBatch);
		System.out.println("Expiry          : "+actExpiry			    +"  Value Expected  "+expExpiry);

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

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) 
				&& actR1Item.equalsIgnoreCase(expR1Item) && actR1TaxCode.equalsIgnoreCase(expR1TaxCode) && actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate) 
				&& actGross.equalsIgnoreCase(expGross) && actBatch.equalsIgnoreCase(expBatch) && actExpiry.equalsIgnoreCase(expExpiry)

				&& actR2Warehouse.equalsIgnoreCase(expR2Warehouse) 
				&& actR2Item.equalsIgnoreCase(expR2Item) && actR2TaxCode.equalsIgnoreCase(expR2TaxCode) && actR2PurchaseAccount.equalsIgnoreCase(expR2PurchaseAccount) 
				&& actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) 
				&& actR2Gross.equalsIgnoreCase(expR2Gross) && actR2RMA.equalsIgnoreCase(expR2RMA))
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 20, 8, resPass);
			return true;
		} 
		else 
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 20, 8, resFail);
			return false;
		}
	}



	



	public static boolean checkUpdatingBuyerPricebookFormulaForPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsBtn));
		settingsBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
		editScreenTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBehaviourDropdown));
		Select behaviour = new Select(editScreenBehaviourDropdown);
		behaviour.selectByVisibleText(excelReader.getCellData(xlSheetName, 44, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBehaviourFormulaTxt));
		editScreenBehaviourFormulaTxt.click();
		editScreenBehaviourFormulaTxt.clear();
		editScreenBehaviourFormulaTxt.sendKeys(excelReader.getCellData(xlSheetName, 45, 5));
		editScreenBehaviourFormulaTxt.sendKeys(Keys.TAB);
		//getAction().moveToElement(editScreenBehaviourFormulaTxt).pause(1000).click().build().perform();
		Thread.sleep(2000);

		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", editScreenBehaviourFormulaOkBtn);

		Thread.sleep(5000);
		
		//getAction().moveToElement(editScreenBehaviourFormulaOkBtn).build().perform();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBehaviourFormulaOkBtn));
		editScreenBehaviourFormulaOkBtn.click();
		
	((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", updateBtn);
		
		//getAction().moveToElement(updateBtn).build().perform();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		Thread.sleep(2000);

		System.out.println("******************checkUpdatingBuyerPricebookFormulaForPurchaseVoucherVAT*****************");

		String expMessage = excelReader.getCellData(xlSheetName, 44, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 44, 7, actMessage);

		Thread.sleep(2000);

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 44, 8, resPass);
			return true;
		}
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 44, 8, resFail);
			return false;
		}
	}







	public static boolean checkSavingPurchaseVoucherVATVoucher2WithRateFormula() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getDriver().navigate().refresh();
		
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

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 46, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
					voucherHeaderDueDate.click();
					voucherHeaderDueDate.sendKeys(Keys.TAB);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
					voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
					voucherHeaderCurrency.sendKeys(Keys.SPACE);

					int currencycount=currencyListCount.size();

					System.err.println(currencycount);

					for(int i=0 ; i < currencycount ;i++)
					{
						String data=currencyListCount.get(i).getText();

						if (data.equalsIgnoreCase("INR")) 
						{
							currencyListCount.get(i).click();
							break;
						}
					}
					voucherHeaderCurrency.sendKeys(Keys.TAB);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 47, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 48, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 49, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 50, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 51, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 52, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		String actRate1 = select1stRow_11thColumn.getText();
		String expRate1 = excelReader.getCellData(xlSheetName, 46, 6);
		excelReader.setCellData(xlfile, xlSheetName, 46, 7, actRate1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
		select1stRow_12thColumn.click();
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
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 53, 5));
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


		// Row 2 with Item Group item

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(Keys.END);
		enter_WarehouseTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 54, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 55, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_11thColumn));
		String actRate2 = select2ndRow_11thColumn.getText();
		String expRate2 = excelReader.getCellData(xlSheetName, 47, 6);
		excelReader.setCellData(xlfile, xlSheetName, 47, 7, actRate2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
		select2ndRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 56, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_12thColumn));
		select2ndRow_12thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_14thColumn));
		select2ndRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));

		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
		rmaSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 57, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
		rmaAddBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
		RMAPopupOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));

		String actRMAColumn=enter_RMA.getAttribute("value");

		String expRMAColumn=excelReader.getCellData(xlSheetName, 48, 6);
		excelReader.setCellData(xlfile, xlSheetName, 48, 7, actRMAColumn);

		System.out.println(actRMAColumn);
		System.out.println(expRMAColumn);
*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

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
		
		expMessage.add("Voucher saved successfully : SU/AME/TEXT2");
		expMessage.add("Saving in background.");*/
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : SU/AME/TEXT2");
		expMsg.add("Saving in background.");*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 55, 6);
		excelReader.setCellData(xlfile, xlSheetName, 55, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		System.out.println("************************************************************************");

		System.out.println("Rates1 : "+actRate1+"  Value Expected  "+expRate1);
		//System.out.println("Rates2 : "+actRate2+"  Value Expected  "+expRate2);

		if(actRate1.equalsIgnoreCase(expRate1) /*&& actRate2.equalsIgnoreCase(expRate2) /*&& actBackgroundSavingMessage.equalsIgnoreCase(expBackgroundSavingMessage)*/
				&& actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 46, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 46, 8, resFail);
			return false;
		}
	}







	public static boolean checkVerifyingPurchaseVoucherVATVoucher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 58, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher Loaded Successfully");

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 58, 6);
		excelReader.setCellData(xlfile, xlSheetName, 58, 7, actDocNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 59, 6);
		excelReader.setCellData(xlfile, xlSheetName, 59, 7, actVendor);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 60, 6);
		excelReader.setCellData(xlfile, xlSheetName, 60, 7, actCurrency);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 61, 6);
		excelReader.setCellData(xlfile, xlSheetName, 61, 7, actDepartment);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 62, 6);
		excelReader.setCellData(xlfile, xlSheetName, 62, 7, actPlaceOfSupply);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 63, 6);
		excelReader.setCellData(xlfile, xlSheetName, 63, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 64, 6);
		excelReader.setCellData(xlfile, xlSheetName, 64, 7, actR1Warehouse);

		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 65, 6);
		excelReader.setCellData(xlfile, xlSheetName, 65, 7, actR1Item);

		String actR1TaxCode = select1stRow_3rdColumn.getText();
		String expR1TaxCode = excelReader.getCellData(xlSheetName, 66, 6);
		excelReader.setCellData(xlfile, xlSheetName, 66, 7, actR1TaxCode);

		String actPurchaseAccount = select1stRow_4thColumn.getText();
		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 67, 6);
		excelReader.setCellData(xlfile, xlSheetName, 67, 7, actPurchaseAccount);

		String actUnits = select1stRow_5thColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 68, 6);
		excelReader.setCellData(xlfile, xlSheetName, 68, 7, actUnits);

		String actQuantity = select1stRow_9thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 69, 6);
		excelReader.setCellData(xlfile, xlSheetName, 69, 7, actQuantity);

		String actRate = select1stRow_11thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 70, 6);
		excelReader.setCellData(xlfile, xlSheetName, 70, 7, actRate);

		String actGross = select1stRow_12thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 71, 6);
		excelReader.setCellData(xlfile, xlSheetName, 71, 7, actGross);

		String actBatch = select1stRow_16thColumn.getText();
		String expBatch = excelReader.getCellData(xlSheetName, 72, 6);
		excelReader.setCellData(xlfile, xlSheetName, 72, 7, actBatch);

		String actExpiry = select1stRow_18thColumn.getText();
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		Calendar cal=Calendar.getInstance();
		cal.add(Calendar.DATE, 2); 
		String expExpiry=df.format(cal.getTime());



		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 73, 6);
		excelReader.setCellData(xlfile, xlSheetName, 73, 7, actR2Warehouse);

		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 74, 6);
		excelReader.setCellData(xlfile, xlSheetName, 74, 7, actR2Item);

		String actR2TaxCode = select2ndRow_3rdColumn.getText();
		String expR2TaxCode = excelReader.getCellData(xlSheetName, 75, 6);
		excelReader.setCellData(xlfile, xlSheetName, 75, 7, actR2TaxCode);

		String actR2PurchaseAccount = select2ndRow_4thColumn.getText();
		String expR2PurchaseAccount = excelReader.getCellData(xlSheetName, 76, 6);
		excelReader.setCellData(xlfile, xlSheetName, 76, 7, actR2PurchaseAccount);

		String actR2Units = select2ndRow_5thColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 77, 6);
		excelReader.setCellData(xlfile, xlSheetName, 77, 7, actR2Units);

		String actR2Quantity = select2ndRow_9thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 78, 6);
		excelReader.setCellData(xlfile, xlSheetName, 78, 7, actR2Quantity);

		String actR2Rate = select2ndRow_11thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 79, 6);
		excelReader.setCellData(xlfile, xlSheetName, 79, 7, actR2Rate);

		String actR2Gross = select2ndRow_12thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 80, 6);
		excelReader.setCellData(xlfile, xlSheetName, 80, 7, actR2Gross);

		String actR2RMA = select2ndRow_19thColumn.getText();
		String expR2RMA = excelReader.getCellData(xlSheetName, 81, 6);
		excelReader.setCellData(xlfile, xlSheetName, 81, 7, actR2RMA);

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
		System.out.println("PurchaseAccount : "+actPurchaseAccount	    +"  Value Expected  "+expPurchaseAccount);
		System.out.println("Units           : "+actUnits			    +"  Value Expected  "+expUnits);
		System.out.println("Quantity        : "+actQuantity			    +"  Value Expected  "+expQuantity);
		System.out.println("Rate            : "+actRate				    +"  Value Expected  "+expRate);
		System.out.println("Gross           : "+actGross			    +"  Value Expected  "+expGross);
		System.out.println("Batch           : "+actBatch			    +"  Value Expected  "+expBatch);
		System.out.println("Expiry          : "+actExpiry			    +"  Value Expected  "+expExpiry);

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

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) 
				&& actR1Item.equalsIgnoreCase(expR1Item) && actR1TaxCode.equalsIgnoreCase(expR1TaxCode) && actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate) 
				&& actGross.equalsIgnoreCase(expGross) /*&& actBatch.equalsIgnoreCase(expBatch)*/ && actExpiry.equalsIgnoreCase(expExpiry)

				&& actR2Warehouse.equalsIgnoreCase(expR2Warehouse) 
				&& actR2Item.equalsIgnoreCase(expR2Item) && actR2TaxCode.equalsIgnoreCase(expR2TaxCode) && actR2PurchaseAccount.equalsIgnoreCase(expR2PurchaseAccount) 
				&& actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) 
				&& actR2Gross.equalsIgnoreCase(expR2Gross) && actR2RMA.equalsIgnoreCase(expR2RMA))
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 58, 8, resPass);
			return true;
		} 
		else 
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 58, 8, resFail);
			return false;
		}
	}







	public static boolean checkUpdatingBuyerPricebookFormulaWithAbbrivationForPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsBtn));
		settingsBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
		editScreenTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBehaviourFormulaTxt));
		editScreenBehaviourFormulaTxt.click();
		editScreenBehaviourFormulaTxt.sendKeys(Keys.END);
		editScreenBehaviourFormulaTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(1000);
		editScreenBehaviourFormulaTxt.sendKeys(excelReader.getCellData(xlSheetName, 82, 5));
		editScreenBehaviourFormulaTxt.sendKeys(Keys.TAB);
		//getAction().moveToElement(editScreenBehaviourFormulaTxt).pause(1000).click().build().perform();
		Thread.sleep(2000);
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", editScreenBehaviourFormulaOkBtn);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBehaviourFormulaOkBtn));
		editScreenBehaviourFormulaOkBtn.click();
		
		Thread.sleep(2000);
		
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", updateBtn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		Thread.sleep(2000);

		System.out.println("******************checkUpdatingBuyerPricebookFormulaWithAbbrivationForPurchaseVoucherVAT*****************");

		String expMessage = excelReader.getCellData(xlSheetName, 82, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 82, 7, actMessage);

		Thread.sleep(2000);

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 82, 8, resPass);
			return true;
		}
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 82, 8, resFail);
			return false;
		}
	}







	public static boolean checkSavingPurchaseVoucherVATVoucher3WithRateFormulaWithAbbrevation() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
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

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 83, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
					voucherHeaderDueDate.click();
					voucherHeaderDueDate.sendKeys(Keys.TAB);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
					voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
					voucherHeaderCurrency.sendKeys(Keys.SPACE);

					int currencycount=currencyListCount.size();

					System.err.println(currencycount);

					for(int i=0 ; i < currencycount ;i++)
					{
						String data=currencyListCount.get(i).getText();

						if (data.equalsIgnoreCase("INR")) 
						{
							currencyListCount.get(i).click();
							break;
						}
					}
					voucherHeaderCurrency.sendKeys(Keys.TAB);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 84, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 85, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 86, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 87, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 88, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 89, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
		select1stRow_12thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		String actRate1 = select1stRow_11thColumn.getText();
		String expRate1 = excelReader.getCellData(xlSheetName, 83, 6);
		excelReader.setCellData(xlfile, xlSheetName, 83, 7, actRate1);

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
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 90, 5));
		enter_Batch.sendKeys(Keys.TAB);

		/*DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

					Calendar cal=Calendar.getInstance();
					SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
					cal.add(Calendar.DATE, 2); 

					String next2DaysDate=df.format(cal.getTime());

					System.out.println(">>>>>>>>>Expiry Date After Adding 2Days>>>>>>>> : "+next2DaysDate);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(excelReader.getCellData(xlSheetName, 91, 5));
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(excelReader.getCellData(xlSheetName, 91, 5));
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		// Row 2 with Item Group item

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(Keys.END);
		enter_WarehouseTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 92, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 93, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_11thColumn));
		String actRate2 = select2ndRow_11thColumn.getText();
		String expRate2 = excelReader.getCellData(xlSheetName, 84, 6);
		excelReader.setCellData(xlfile, xlSheetName, 84, 7, actRate1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
		select2ndRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 94, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_12thColumn));
		select2ndRow_12thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_14thColumn));
		select2ndRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.click();

		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.click();
		enter_PvTaxable.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
		rmaSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 95, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
		rmaAddBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
		RMAPopupOkBtn.click();

	/*	Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));

		String actRMAColumn=enter_RMA.getAttribute("value");

		String expRMAColumn=excelReader.getCellData(xlSheetName, 85, 6);
		excelReader.setCellData(xlfile, xlSheetName, 85, 7, actRate1);

		System.out.println(actRMAColumn);
		System.out.println(expRMAColumn);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
	Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

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

		expMessage.add("Voucher saved successfully : SU/AME/TEXT3");
		expMessage.add("Saving in background.");*/

		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : SU/AME/TEXT3");
		expMsg.add("Saving in background.");*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 92, 6);
		excelReader.setCellData(xlfile, xlSheetName, 92, 7, actMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		System.out.println("************************************************************************");

		System.out.println("Rates1 : "+actRate1+"  Value Expected  "+expRate1);
	//	System.out.println("Rates2 : "+actRate2+"  Value Expected  "+expRate2);

		if(actRate1.equalsIgnoreCase(expRate1)/* && actRate2.equalsIgnoreCase(expRate2) /*&&  actBackgroundSavingMessage.equalsIgnoreCase(expBackgroundSavingMessage)*/
				&& actMessage.equals(expMessage))
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







	public static boolean checkVerifyingPurchaseVoucherVATVoucher3() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 96, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher Loaded Successfully");

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 96, 6);
		excelReader.setCellData(xlfile, xlSheetName, 96, 7, actDocNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 97, 6);
		excelReader.setCellData(xlfile, xlSheetName, 97, 7, actVendor);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 98, 6);
		excelReader.setCellData(xlfile, xlSheetName, 98, 7, actCurrency);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 99, 6);
		excelReader.setCellData(xlfile, xlSheetName, 99, 7, actDepartment);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 100, 6);
		excelReader.setCellData(xlfile, xlSheetName, 100, 7, actPlaceOfSupply);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 101, 6);
		excelReader.setCellData(xlfile, xlSheetName, 101, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 102, 6);
		excelReader.setCellData(xlfile, xlSheetName, 102, 7, actR1Warehouse);

		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 103, 6);
		excelReader.setCellData(xlfile, xlSheetName, 103, 7, actR1Item);

		String actR1TaxCode = select1stRow_3rdColumn.getText();
		String expR1TaxCode = excelReader.getCellData(xlSheetName, 104, 6);
		excelReader.setCellData(xlfile, xlSheetName, 104, 7, actR1TaxCode);

		String actPurchaseAccount = select1stRow_4thColumn.getText();
		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 105, 6);
		excelReader.setCellData(xlfile, xlSheetName, 105, 7, actPurchaseAccount);

		String actUnits = select1stRow_5thColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 106, 6);
		excelReader.setCellData(xlfile, xlSheetName, 106, 7, actUnits);

		String actQuantity = select1stRow_9thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 107, 6);
		excelReader.setCellData(xlfile, xlSheetName, 107, 7, actQuantity);

		String actRate = select1stRow_11thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 108, 6);
		excelReader.setCellData(xlfile, xlSheetName, 108, 7, actRate);

		String actGross = select1stRow_12thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 109, 6);
		excelReader.setCellData(xlfile, xlSheetName, 109, 7, actGross);

		String actBatch = select1stRow_16thColumn.getText();
		String expBatch = excelReader.getCellData(xlSheetName, 110, 6);
		excelReader.setCellData(xlfile, xlSheetName, 110, 7, actBatch);

		String actExpiry = select1stRow_18thColumn.getText();
		String expExpiry=excelReader.getCellData(xlSheetName, 111, 6);
		excelReader.setCellData(xlfile, xlSheetName, 111, 7, actExpiry);



		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 112, 6);
		excelReader.setCellData(xlfile, xlSheetName, 112, 7, actR2Warehouse);

		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 113, 6);
		excelReader.setCellData(xlfile, xlSheetName, 113, 7, actR2Item);

		String actR2TaxCode = select2ndRow_3rdColumn.getText();
		String expR2TaxCode = excelReader.getCellData(xlSheetName, 114, 6);
		excelReader.setCellData(xlfile, xlSheetName, 114, 7, actR2TaxCode);

		String actR2PurchaseAccount = select2ndRow_4thColumn.getText();
		String expR2PurchaseAccount = excelReader.getCellData(xlSheetName, 115, 6);
		excelReader.setCellData(xlfile, xlSheetName, 115, 7, actR2PurchaseAccount);

		String actR2Units = select2ndRow_5thColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 116, 6);
		excelReader.setCellData(xlfile, xlSheetName, 116, 7, actR2Units);

		String actR2Quantity = select2ndRow_9thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 117, 6);
		excelReader.setCellData(xlfile, xlSheetName, 117, 7, actR2Quantity);

		String actR2Rate = select2ndRow_11thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 118, 6);
		excelReader.setCellData(xlfile, xlSheetName, 118, 7, actR2Rate);

		String actR2Gross = select2ndRow_12thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 119, 6);
		excelReader.setCellData(xlfile, xlSheetName, 119, 7, actR2Gross);

		String actR2RMA = select2ndRow_19thColumn.getText();
		String expR2RMA = excelReader.getCellData(xlSheetName, 120, 6);
		excelReader.setCellData(xlfile, xlSheetName, 120, 7, actR2RMA);

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
		System.out.println("PurchaseAccount : "+actPurchaseAccount	    +"  Value Expected  "+expPurchaseAccount);
		System.out.println("Units           : "+actUnits			    +"  Value Expected  "+expUnits);
		System.out.println("Quantity        : "+actQuantity			    +"  Value Expected  "+expQuantity);
		System.out.println("Rate            : "+actRate				    +"  Value Expected  "+expRate);
		System.out.println("Gross           : "+actGross			    +"  Value Expected  "+expGross);
		System.out.println("Batch           : "+actBatch			    +"  Value Expected  "+expBatch);
		System.out.println("Expiry          : "+actExpiry			    +"  Value Expected  "+expExpiry);

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

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) 
				&& actR1Item.equalsIgnoreCase(expR1Item) && actR1TaxCode.equalsIgnoreCase(expR1TaxCode) && actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate) 
				&& actGross.equalsIgnoreCase(expGross) /*&& actBatch.equalsIgnoreCase(expBatch)*/ && actExpiry.equalsIgnoreCase(expExpiry)

				&& actR2Warehouse.equalsIgnoreCase(expR2Warehouse) 
				&& actR2Item.equalsIgnoreCase(expR2Item) && actR2TaxCode.equalsIgnoreCase(expR2TaxCode) && actR2PurchaseAccount.equalsIgnoreCase(expR2PurchaseAccount) 
				&& actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) 
				&& actR2Gross.equalsIgnoreCase(expR2Gross) && actR2RMA.equalsIgnoreCase(expR2RMA))
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 96, 8, resPass);
			return true;
		} 
		else 
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 96, 8, resFail);
			return false;
		}
	}







	public static boolean checkChangingRatePreloadBehaviourToRatePurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsBtn));
		settingsBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
		editScreenTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBehaviourDropdown));
		ScrollIntoView(editScreenBehaviourDropdown);
		Select behaviour = new Select(editScreenBehaviourDropdown);
		behaviour.selectByVisibleText(excelReader.getCellData(xlSheetName, 121, 5));
		
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", updateBtn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		System.out.println("******************checkChangingRatePreloadBehaviourToRatePurchaseVoucherVAT*****************");

		String expMessage = excelReader.getCellData(xlSheetName, 121, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 121, 7, actMessage);

		Thread.sleep(2000);

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 121, 8, resPass);
			return true;
		}
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 121, 8, resFail);
			return false;
		}
	}







	public static boolean checkSavingPurchaseVoucherVATVoucher4WithStockItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
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

		checkValidationMessage("Screen Opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 122, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 123, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 124, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 125, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 126, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 127, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 128, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 129, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 130, 5));
		Thread.sleep(2000);
		enter_Rate.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

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

		expMessage.add("Voucher saved successfully : SU/DUB/TEXT1");
		expMessage.add("Saving in background.");*/
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : SU/DUB/TEXT1");
		expMsg.add("Saving in background.");*/
		
		String expMessage = /*expMsg.toString()*/ excelReader.getCellData(xlSheetName, 128, 6);
		excelReader.setCellData(xlfile, xlSheetName, 128, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		Thread.sleep(2000);

		if(actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 122, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 122, 8, resFail);
			return false;
		}
	}


	

	
	

	public static boolean checkPostingDetailsWithEqualsToAndUpdatingRateToGreaterThanStandardRate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 131, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 131, 6);
		excelReader.setCellData(xlfile, xlSheetName, 131, 7, actDocNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 132, 6);
		excelReader.setCellData(xlfile, xlSheetName, 132, 7, actVendor);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 133, 6);
		excelReader.setCellData(xlfile, xlSheetName, 133, 7, actCurrency);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 134, 6);
		excelReader.setCellData(xlfile, xlSheetName, 134, 7, actDepartment);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 135, 6);
		excelReader.setCellData(xlfile, xlSheetName, 135, 7, actPlaceOfSupply);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 136, 6);
		excelReader.setCellData(xlfile, xlSheetName, 136, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actWarehouse = select1stRow_1stColumn.getText();
		String expWarehouse = excelReader.getCellData(xlSheetName, 137, 6);
		excelReader.setCellData(xlfile, xlSheetName, 137, 7, actWarehouse);

		String actItem = select1stRow_2ndColumn.getText();
		String expItem = excelReader.getCellData(xlSheetName, 138, 6);
		excelReader.setCellData(xlfile, xlSheetName, 138, 7, actItem);

		String actTaxCode = select1stRow_3rdColumn.getText();
		String expTaxCode = excelReader.getCellData(xlSheetName, 139, 6);
		excelReader.setCellData(xlfile, xlSheetName, 139, 7, actTaxCode);

		String actPurchaseAccount = select1stRow_4thColumn.getText();
		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 140, 6);
		excelReader.setCellData(xlfile, xlSheetName, 140, 7, actPurchaseAccount);

		String actUnits = select1stRow_5thColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 141, 6);
		excelReader.setCellData(xlfile, xlSheetName, 141, 7, actUnits);

		String actQuantity = select1stRow_9thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 142, 6);
		excelReader.setCellData(xlfile, xlSheetName, 142, 7, actQuantity);

		String actRate = select1stRow_11thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 143, 6);
		excelReader.setCellData(xlfile, xlSheetName, 143, 7, actRate);

		String actGross = select1stRow_12thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 144, 6);
		excelReader.setCellData(xlfile, xlSheetName, 144, 7, actGross);

		System.out.println("**************checkPostingDetailsWithEqualsToAndUpdatingRateToGreaterThanStandardRate*****************");

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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));	
		toggleBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postingDetailsBtn));	
		postingDetailsBtn.click();

		Thread.sleep(4000);

		int postingDetailsDebitListCount = postingDetailsDebitList.size();

		ArrayList<String> postingDetailsDebitListArray = new ArrayList<String>();

		for (int i=0;i<postingDetailsDebitListCount;i++)
		{
			String data = postingDetailsDebitList.get(i).getText();
			postingDetailsDebitListArray.add(data);

		}

		String actpostingDetailsDebitList = postingDetailsDebitListArray.toString();
		String exppostingDetailsDebitList = excelReader.getCellData(xlSheetName, 145, 6);
		excelReader.setCellData(xlfile, xlSheetName, 145, 7, actpostingDetailsDebitList);

		System.out.println("actpostingDetailsDebitList : "+actpostingDetailsDebitList);
		System.out.println("exppostingDetailsDebitList : "+exppostingDetailsDebitList);



		int postingDetailsCreditListCount = postingDetailsCreditList.size();

		ArrayList<String> postingDetailsCreditListArray = new ArrayList<String>();

		for (int i=0;i<postingDetailsCreditListCount;i++)
		{
			String data = postingDetailsCreditList.get(i).getText();
			postingDetailsCreditListArray.add(data);

		}

		String actpostingDetailsCreditList = postingDetailsCreditListArray.toString();
		String exppostingDetailsCreditList = excelReader.getCellData(xlSheetName, 146, 6);
		excelReader.setCellData(xlfile, xlSheetName, 146, 7, actpostingDetailsCreditList);

		System.out.println("actpostingDetailsCreditList : "+actpostingDetailsCreditList);
		System.out.println("exppostingDetailsCreditList : "+exppostingDetailsCreditList);


		String actpostingDetailsDebitSum = postingDetailsDebitSum.getText();
		String exppostingDetailsDebitSum = excelReader.getCellData(xlSheetName, 147, 6);
		excelReader.setCellData(xlfile, xlSheetName, 147, 7, actpostingDetailsDebitSum);

		String actpostingDetailsCreditSum = postingDetailsCreditSum.getText();
		String exppostingDetailsCreditSum = excelReader.getCellData(xlSheetName, 148, 6);
		excelReader.setCellData(xlfile, xlSheetName, 148, 7, actpostingDetailsCreditSum);

		System.out.println("postingDetailsDebitSum : "+actpostingDetailsDebitSum  +" Value Expected : "+exppostingDetailsDebitSum);
		System.out.println("postingDetailsCreditSum : "+actpostingDetailsCreditSum+" Value Expected : "+exppostingDetailsCreditSum);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postingDetailsCloseBtn));
		postingDetailsCloseBtn.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.sendKeys(Keys.END);
		enter_Rate.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 132, 5));
		Thread.sleep(2000);
		enter_Rate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

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

		expMessage.add("Voucher saved successfully : SU/DUB/TEXT1");
		expMessage.add("Saving in background.");*/

		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : SU/DUB/TEXT1");
		expMsg.add("Saving in background.");*/
		
		String expMessage = /*expMsg.toString()*/excelReader.getCellData(xlSheetName, 149, 6);
		excelReader.setCellData(xlfile, xlSheetName, 149, 7, actMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		Thread.sleep(2000);

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actWarehouse.equalsIgnoreCase(expWarehouse) 
				&& actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate) 
				&& actGross.equalsIgnoreCase(expGross) 

				&& actpostingDetailsDebitList.equalsIgnoreCase(exppostingDetailsDebitList)
				&& actpostingDetailsCreditList.equalsIgnoreCase(exppostingDetailsCreditList)

				&& actpostingDetailsDebitSum.equalsIgnoreCase(exppostingDetailsDebitSum) 
				&& actpostingDetailsCreditSum.equalsIgnoreCase(exppostingDetailsCreditSum)

				&& actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 131, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 131, 8, resFail);
			return false;
		}
	}







	public static boolean checkPostingDetailsWithGreaterThanAndUpdatingRateToLessThanStandardRate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 151, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 151, 6);
		excelReader.setCellData(xlfile, xlSheetName, 151, 7, actDocNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 152, 6);
		excelReader.setCellData(xlfile, xlSheetName, 152, 7, actVendor);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 153, 6);
		excelReader.setCellData(xlfile, xlSheetName, 153, 7, actCurrency);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 154, 6);
		excelReader.setCellData(xlfile, xlSheetName, 154, 7, actDepartment);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 155, 6);
		excelReader.setCellData(xlfile, xlSheetName, 155, 7, actPlaceOfSupply);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 156, 6);
		excelReader.setCellData(xlfile, xlSheetName, 156, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actWarehouse = select1stRow_1stColumn.getText();
		String expWarehouse = excelReader.getCellData(xlSheetName, 157, 6);
		excelReader.setCellData(xlfile, xlSheetName, 157, 7, actWarehouse);

		String actItem = select1stRow_2ndColumn.getText();
		String expItem = excelReader.getCellData(xlSheetName, 158, 6);
		excelReader.setCellData(xlfile, xlSheetName, 158, 7, actItem);

		String actTaxCode = select1stRow_3rdColumn.getText();
		String expTaxCode = excelReader.getCellData(xlSheetName, 159, 6);
		excelReader.setCellData(xlfile, xlSheetName, 159, 7, actTaxCode);

		String actPurchaseAccount = select1stRow_4thColumn.getText();
		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 160, 6);
		excelReader.setCellData(xlfile, xlSheetName, 160, 7, actPurchaseAccount);

		String actUnits = select1stRow_5thColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 161, 6);
		excelReader.setCellData(xlfile, xlSheetName, 161, 7, actUnits);

		String actQuantity = select1stRow_9thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 162, 6);
		excelReader.setCellData(xlfile, xlSheetName, 162, 7, actQuantity);

		String actRate = select1stRow_11thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 163, 6);
		excelReader.setCellData(xlfile, xlSheetName, 163, 7, actRate);

		String actGross = select1stRow_12thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 164, 6);
		excelReader.setCellData(xlfile, xlSheetName, 164, 7, actGross);

		System.out.println("**************checkPostingDetailsWithGreaterThanAndUpdatingRateToLessThanStandardRate*****************");

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


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));	
		toggleBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postingDetailsBtn));	
		postingDetailsBtn.click();

		Thread.sleep(2000);

		int postingDetailsDebitListCount = postingDetailsDebitList.size();

		ArrayList<String> postingDetailsDebitListArray = new ArrayList<String>();

		for (int i=0;i<postingDetailsDebitListCount;i++)
		{
			String data = postingDetailsDebitList.get(i).getText();
			postingDetailsDebitListArray.add(data);
		}

		String actpostingDetailsDebitList = postingDetailsDebitListArray.toString();
		String exppostingDetailsDebitList = excelReader.getCellData(xlSheetName, 165, 6);
		excelReader.setCellData(xlfile, xlSheetName, 165, 7, actpostingDetailsDebitList);

		System.out.println("actpostingDetailsDebitList : "+actpostingDetailsDebitList);
		System.out.println("exppostingDetailsDebitList : "+exppostingDetailsDebitList);



		int postingDetailsCreditListCount = postingDetailsCreditList.size();

		ArrayList<String> postingDetailsCreditListArray = new ArrayList<String>();

		for (int i=0;i<postingDetailsCreditListCount;i++)
		{
			String data = postingDetailsCreditList.get(i).getText();
			postingDetailsCreditListArray.add(data);

		}

		String actpostingDetailsCreditList = postingDetailsCreditListArray.toString();
		String exppostingDetailsCreditList = excelReader.getCellData(xlSheetName, 166, 6);
		excelReader.setCellData(xlfile, xlSheetName, 166, 7, actpostingDetailsCreditList);

		System.out.println("actpostingDetailsCreditList : "+actpostingDetailsCreditList);
		System.out.println("exppostingDetailsCreditList : "+exppostingDetailsCreditList);


		String actpostingDetailsDebitSum = postingDetailsDebitSum.getText();
		String exppostingDetailsDebitSum = excelReader.getCellData(xlSheetName, 167, 6);
		excelReader.setCellData(xlfile, xlSheetName, 167, 7, actpostingDetailsDebitSum);

		String actpostingDetailsCreditSum = postingDetailsCreditSum.getText();
		String exppostingDetailsCreditSum = excelReader.getCellData(xlSheetName, 168, 6);
		excelReader.setCellData(xlfile, xlSheetName, 168, 7, actpostingDetailsCreditSum);

		System.out.println("postingDetailsDebitSum : "+actpostingDetailsDebitSum  +" Value Expected : "+exppostingDetailsDebitSum);
		System.out.println("postingDetailsCreditSum : "+actpostingDetailsCreditSum+" Value Expected : "+exppostingDetailsCreditSum);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postingDetailsCloseBtn));
		postingDetailsCloseBtn.click();


		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.sendKeys(Keys.END);
		enter_Rate.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 152, 5));
		Thread.sleep(2000);
		enter_Rate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		Thread.sleep(2000);
		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/DUB/TEXT1");
		expMessage.add("Saving in background.");*/

		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : SU/DUB/TEXT1");
		expMsg.add("Saving in background.");*/
		
		String expMessage = /*expMsg.toString()*/excelReader.getCellData(xlSheetName, 169, 6);
		excelReader.setCellData(xlfile, xlSheetName, 169, 7, actMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		Thread.sleep(2000);

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actWarehouse.equalsIgnoreCase(expWarehouse) 
				&& actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate) 
				&& actGross.equalsIgnoreCase(expGross) 

				&& actpostingDetailsDebitList.equalsIgnoreCase(exppostingDetailsDebitList)
				&& actpostingDetailsCreditList.equalsIgnoreCase(exppostingDetailsCreditList)

				&& actpostingDetailsDebitSum.equalsIgnoreCase(exppostingDetailsDebitSum) 
				&& actpostingDetailsCreditSum.equalsIgnoreCase(exppostingDetailsCreditSum)

				&& actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 151, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 151, 8, resFail);
			return false;
		}
	}








	public static boolean checkPostingDetailsWithRateLessThanStandardRate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(excelReader.getCellData(xlSheetName, 170, 5));
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 170, 6);
		excelReader.setCellData(xlfile, xlSheetName, 170, 7, actDocNo);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 171, 6);
		excelReader.setCellData(xlfile, xlSheetName, 171, 7, actVendor);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 172, 6);
		excelReader.setCellData(xlfile, xlSheetName, 172, 7, actCurrency);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 173, 6);
		excelReader.setCellData(xlfile, xlSheetName, 173, 7, actDepartment);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 174, 6);
		excelReader.setCellData(xlfile, xlSheetName, 174, 7, actPlaceOfSupply);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 175, 6);
		excelReader.setCellData(xlfile, xlSheetName, 175, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actWarehouse = select1stRow_1stColumn.getText();
		String expWarehouse = excelReader.getCellData(xlSheetName, 176, 6);
		excelReader.setCellData(xlfile, xlSheetName, 176, 7, actWarehouse);

		String actItem = select1stRow_2ndColumn.getText();
		String expItem = excelReader.getCellData(xlSheetName, 177, 6);
		excelReader.setCellData(xlfile, xlSheetName, 177, 7, actItem);

		String actTaxCode = select1stRow_3rdColumn.getText();
		String expTaxCode = excelReader.getCellData(xlSheetName, 178, 6);
		excelReader.setCellData(xlfile, xlSheetName, 178, 7, actTaxCode);

		String actPurchaseAccount = select1stRow_4thColumn.getText();
		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 179, 6);
		excelReader.setCellData(xlfile, xlSheetName, 179, 7, actPurchaseAccount);

		String actUnits = select1stRow_5thColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 180, 6);
		excelReader.setCellData(xlfile, xlSheetName, 180, 7, actUnits);

		String actQuantity = select1stRow_9thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 181, 6);
		excelReader.setCellData(xlfile, xlSheetName, 181, 7, actQuantity);

		String actRate = select1stRow_11thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 182, 6);
		excelReader.setCellData(xlfile, xlSheetName, 182, 7, actRate);

		String actGross = select1stRow_12thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 183, 6);
		excelReader.setCellData(xlfile, xlSheetName, 183, 7, actGross);

		System.out.println("**************checkPostingDetailsWithRateLessThanStandardRate*****************");

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


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));	
		toggleBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postingDetailsBtn));	
		postingDetailsBtn.click();

		Thread.sleep(2000);

		int postingDetailsDebitListCount = postingDetailsDebitList.size();

		ArrayList<String> postingDetailsDebitListArray = new ArrayList<String>();

		for (int i=0;i<postingDetailsDebitListCount;i++)
		{
			String data = postingDetailsDebitList.get(i).getText();
			postingDetailsDebitListArray.add(data);

		}

		String actpostingDetailsDebitList = postingDetailsDebitListArray.toString();
		String exppostingDetailsDebitList = excelReader.getCellData(xlSheetName, 184, 6);
		excelReader.setCellData(xlfile, xlSheetName, 184, 7, actpostingDetailsDebitList);

		System.out.println("actpostingDetailsDebitList : "+actpostingDetailsDebitList);
		System.out.println("exppostingDetailsDebitList : "+exppostingDetailsDebitList);



		int postingDetailsCreditListCount = postingDetailsCreditList.size();

		ArrayList<String> postingDetailsCreditListArray = new ArrayList<String>();

		for (int i=0;i<postingDetailsCreditListCount;i++)
		{
			String data = postingDetailsCreditList.get(i).getText();
			postingDetailsCreditListArray.add(data);

		}

		String actpostingDetailsCreditList = postingDetailsCreditListArray.toString();
		String exppostingDetailsCreditList = excelReader.getCellData(xlSheetName, 185, 6);
		excelReader.setCellData(xlfile, xlSheetName, 185, 7, actpostingDetailsCreditList);
		
		System.out.println("actpostingDetailsCreditList : "+actpostingDetailsCreditList);
		System.out.println("exppostingDetailsCreditList : "+exppostingDetailsCreditList);


		String actpostingDetailsDebitSum = postingDetailsDebitSum.getText();
		String exppostingDetailsDebitSum = excelReader.getCellData(xlSheetName, 186, 6);
		excelReader.setCellData(xlfile, xlSheetName, 186, 7, actpostingDetailsDebitSum);

		String actpostingDetailsCreditSum = postingDetailsCreditSum.getText();
		String exppostingDetailsCreditSum = excelReader.getCellData(xlSheetName, 187, 6);
		excelReader.setCellData(xlfile, xlSheetName, 187, 7, actpostingDetailsCreditSum);

		System.out.println("postingDetailsDebitSum : "+actpostingDetailsDebitSum  +" Value Expected : "+exppostingDetailsDebitSum);
		System.out.println("postingDetailsCreditSum : "+actpostingDetailsCreditSum+" Value Expected : "+exppostingDetailsCreditSum);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postingDetailsCloseBtn));
		postingDetailsCloseBtn.click();

		Thread.sleep(2000);

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
				&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actWarehouse.equalsIgnoreCase(expWarehouse) 
				&& actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) 
				&& actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate) 
				&& actGross.equalsIgnoreCase(expGross) 

				&& actpostingDetailsDebitList.equalsIgnoreCase(exppostingDetailsDebitList)
				&& actpostingDetailsCreditList.equalsIgnoreCase(exppostingDetailsCreditList)

				&& actpostingDetailsDebitSum.equalsIgnoreCase(exppostingDetailsDebitSum) 
				&& actpostingDetailsCreditSum.equalsIgnoreCase(exppostingDetailsCreditSum))
		{
			excelReader.setCellData(xlfile, xlSheetName, 170, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 170, 8, resFail);
			return false;
		}
	}








	public static boolean checkSavingPurchaseVoucherVATVoucher5WithStockItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
		new_newBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 188, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		if (validationConfirmationMessage.getText().isEmpty()==false) 
		{
			checkValidationMessage("Rules");
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 189, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 190, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 191, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 192, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 193, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 194, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 195, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_13thColumn));
		select1stRow_13thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 196, 5));
		Thread.sleep(2000);
		enter_Rate.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		if (validationConfirmationMessage.getText().isEmpty()==false) 
		{
			checkValidationMessage("Rules");
		}

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

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

		expMessage.add("Voucher saved successfully : SU/DUB/TEXT2");
		expMessage.add("Saving in background.");*/

		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : SU/DUB/TEXT2");
		expMsg.add("Saving in background.");*/
		
		String expMessage = /*expMsg.toString()*/ excelReader.getCellData(xlSheetName, 188, 6);
		excelReader.setCellData(xlfile, xlSheetName, 188, 7, actMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		Thread.sleep(2000);

		if(actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 188, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 188, 8, resFail);
			return false;
		}
	}








	public static boolean checkPurchaseVoucherVATWithBinItemAutoAllocateInBinPopup() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
		new_newBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 197, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 198, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 199, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 200, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 201, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 202, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 203, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 204, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_13thColumn));
		select1stRow_13thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 205, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_16thColumn));
		select1stRow_16thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
		binAutoAllocateBtn.click();

		Thread.sleep(2000);

		String actBinBalanceQty1        = binBalanceTxt.getAttribute("value");
		String expBinBalanceQty1        = excelReader.getCellData(xlSheetName, 197, 6);
		excelReader.setCellData(xlfile, xlSheetName, 197, 7, actBinBalanceQty1);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
		binAutoAllocateBtn.click();

		Thread.sleep(2000);

		String actBinBalanceQty2        = binBalanceTxt.getAttribute("value");
		String expBinBalanceQty2        = excelReader.getCellData(xlSheetName, 198, 6);
		excelReader.setCellData(xlfile, xlSheetName, 198, 7, actBinBalanceQty2);


		System.out.println("BinBalanceQty        : "+actBinBalanceQty1      +"  Value Expected  "+expBinBalanceQty1);
		System.out.println("BinBalanceQty1       : "+actBinBalanceQty2      +"  Value Expected  "+expBinBalanceQty2);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binCancelBtn));
		binCancelBtn.click();

		Thread.sleep(2000);

		if(actBinBalanceQty1.equalsIgnoreCase(expBinBalanceQty1) && actBinBalanceQty2.equalsIgnoreCase(expBinBalanceQty2))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();

			getWaitForAlert();
			getAlert().accept();

			Thread.sleep(2000);
			
			excelReader.setCellData(xlfile, xlSheetName, 197, 8, resPass);

			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();

			getWaitForAlert();
			getAlert().accept();

			Thread.sleep(2000);

			excelReader.setCellData(xlfile, xlSheetName, 197, 8, resFail);
			
			return false;
		}
	}









	public boolean checkPrintScreenOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		Thread.sleep(2000);
		
		String actLayoutNameTxt = Boolean.toString(LayoutNameTxt.isDisplayed());
		String expLayoutNameTxt = excelReader.getCellData(xlSheetName, 206, 6);
		excelReader.setCellData(xlfile, xlSheetName, 206, 7, actLayoutNameTxt);
		
		String actcreateLayoutBtn = Boolean.toString(createLayoutBtn.isDisplayed());
		String expcreateLayoutBtn = excelReader.getCellData(xlSheetName, 207, 6); 
		excelReader.setCellData(xlfile, xlSheetName, 207, 7, actcreateLayoutBtn);
		
		String actLayoutName_Editlayout = Boolean.toString(LayoutName_Editlayout.isDisplayed());
		String expLayoutName_Editlayout = excelReader.getCellData(xlSheetName, 208, 6);
		excelReader.setCellData(xlfile, xlSheetName, 208, 7, actLayoutName_Editlayout);
		
		String actDocumentSet = Boolean.toString(DocumentSet.isDisplayed());
		String expDocumentSet = excelReader.getCellData(xlSheetName, 209, 6);
		excelReader.setCellData(xlfile, xlSheetName, 209, 7, actDocumentSet);
		
		String actDocumentSet_CreatelayoutBtn = Boolean.toString(DocumentSet_CreatelayoutBtn.isDisplayed());
		String expDocumentSet_CreatelayoutBtn = excelReader.getCellData(xlSheetName, 210, 6);
		excelReader.setCellData(xlfile, xlSheetName, 210, 7, actDocumentSet_CreatelayoutBtn);
		
		String actDocumentSet_Editlayout = Boolean.toString(DocumentSet_Editlayout.isDisplayed());
		String expDocumentSet_Editlayout = excelReader.getCellData(xlSheetName, 211, 6);
		excelReader.setCellData(xlfile, xlSheetName, 211, 7, actDocumentSet_Editlayout);

		String actLayoutOptionsdropdown = Boolean.toString(LayoutOptionsdropdown.isDisplayed());
		String expLayoutOptionsdropdown = excelReader.getCellData(xlSheetName, 212, 6);
		excelReader.setCellData(xlfile, xlSheetName, 212, 7, actLayoutOptionsdropdown);

		String actPreviewBtn = Boolean.toString(PreviewBtn.isDisplayed());
		String expPreviewBtn = excelReader.getCellData(xlSheetName, 213, 6);
		excelReader.setCellData(xlfile, xlSheetName, 213, 7, actPreviewBtn);
		
		String actPrintBtn = Boolean.toString(PrintBtn.isDisplayed());
		String expPrintBtn = excelReader.getCellData(xlSheetName, 214, 6);
		excelReader.setCellData(xlfile, xlSheetName, 214, 7, actPrintBtn);
		
		String actCloseBtn = Boolean.toString(CloseBtn.isDisplayed());
		String expCloseBtn = excelReader.getCellData(xlSheetName, 215, 6);
		excelReader.setCellData(xlfile, xlSheetName, 215, 7, actCloseBtn);
		
		System.out.println("LayoutNameTxt               : "+actLayoutNameTxt               +"  Value Expected  "+expLayoutNameTxt);
		System.out.println("createLayoutBtn             : "+actcreateLayoutBtn             +"  Value Expected  "+expcreateLayoutBtn);
		System.out.println("LayoutName_Editlayout       : "+actLayoutName_Editlayout       +"  Value Expected  "+expLayoutName_Editlayout);
		System.out.println("DocumentSet                 : "+actDocumentSet                 +"  Value Expected  "+expDocumentSet);
		System.out.println("DocumentSet_CreatelayoutBtn : "+actDocumentSet_CreatelayoutBtn +"  Value Expected  "+expDocumentSet_CreatelayoutBtn);
		System.out.println("DocumentSet_Editlayout      : "+actDocumentSet_Editlayout      +"  Value Expected  "+expDocumentSet_Editlayout);
		System.out.println("LayoutOptionsdropdown       : "+actLayoutOptionsdropdown       +"  Value Expected  "+expLayoutOptionsdropdown);
		System.out.println("PreviewBtn                  : "+actPreviewBtn                  +"  Value Expected  "+expPreviewBtn);
		System.out.println("PrintBtn                    : "+actPrintBtn                    +"  Value Expected  "+expPrintBtn);
		System.out.println("CloseBtn                    : "+actCloseBtn                    +"  Value Expected  "+expCloseBtn);
		
		if(actLayoutNameTxt.equalsIgnoreCase(expLayoutNameTxt) && actcreateLayoutBtn.equalsIgnoreCase(expcreateLayoutBtn) 
				&& actLayoutName_Editlayout.equalsIgnoreCase(expLayoutName_Editlayout)
				&& actDocumentSet.equalsIgnoreCase(expDocumentSet) && actDocumentSet_CreatelayoutBtn.equalsIgnoreCase(expDocumentSet_CreatelayoutBtn) 
				&& actDocumentSet_Editlayout.equalsIgnoreCase(expDocumentSet_Editlayout) && actLayoutOptionsdropdown.equalsIgnoreCase(expLayoutOptionsdropdown)
				&& actPreviewBtn.equalsIgnoreCase(expPreviewBtn) && actPrintBtn.equalsIgnoreCase(expPrintBtn) && actCloseBtn.equalsIgnoreCase(expCloseBtn))
		{
			excelReader.setCellData(xlfile, xlSheetName, 206, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 206, 8, resFail);
			return false;
		}
	}




	public boolean checkLayoutOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(LayoutOptionsdropdown));
		LayoutOptionsdropdown.click();
		Select s=new Select(LayoutOptionsdropdown);
		int actval=s.getOptions().size();

		Thread.sleep(2000);
		
		String actOutputOptions = Integer.toString(actval);
		String expOutputOptions = excelReader.getCellData(xlSheetName, 216, 6);
		excelReader.setCellData(xlfile, xlSheetName, 216, 7, actOutputOptions);

		System.out.println("Output Options Count : "+actOutputOptions+"  Value Expected  "+expOutputOptions);
		
		if(actOutputOptions.equalsIgnoreCase(expOutputOptions))
		{
			excelReader.setCellData(xlfile, xlSheetName, 216, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 216, 8, resFail);
			return false;
		}
	}







	public boolean checkInvoiceDesignOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(createLayoutBtn));
		createLayoutBtn.click();

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(Newdropdown));
		
		String actNewdropdown = Boolean.toString(Newdropdown.isDisplayed());
		String expNewdropdown = excelReader.getCellData(xlSheetName, 217, 6);
		excelReader.setCellData(xlfile, xlSheetName, 217, 7, actNewdropdown);
		
		String actOpendropdown = Boolean.toString(Opendropdown.isDisplayed());
		String expOpendropdown = excelReader.getCellData(xlSheetName, 218, 6);
		excelReader.setCellData(xlfile, xlSheetName, 218, 7, actOpendropdown);
		
		String actSavedropdown = Boolean.toString(Savedropdown.isDisplayed());
		String expSavedropdown = excelReader.getCellData(xlSheetName, 219, 6);
		excelReader.setCellData(xlfile, xlSheetName, 219, 7, actSavedropdown);
		
		String actRemovedropdown = Boolean.toString(Removedropdown.isDisplayed());
		String expRemovedropdown = excelReader.getCellData(xlSheetName, 220, 6);
		excelReader.setCellData(xlfile, xlSheetName, 220, 7, actRemovedropdown);
		
		String actPageNo = Boolean.toString(PageNo.isDisplayed());
		String expPageNo = excelReader.getCellData(xlSheetName, 221, 6);
		excelReader.setCellData(xlfile, xlSheetName, 221, 7, actPageNo);
		
		String actControldrpdown = Boolean.toString(Controldrpdown.isDisplayed());
		String expControldrpdown = excelReader.getCellData(xlSheetName, 222, 6);
		excelReader.setCellData(xlfile, xlSheetName, 222, 7, actControldrpdown);
		
		String actAlignmentdropdown = Boolean.toString(Alignmentdropdown.isDisplayed());
		String expAlignmentdropdown = excelReader.getCellData(xlSheetName, 223, 6);
		excelReader.setCellData(xlfile, xlSheetName, 223, 7, actAlignmentdropdown);
		
		String actSizedropdown = Boolean.toString(Sizedropdown.isDisplayed());
		String expSizedropdown = excelReader.getCellData(xlSheetName, 224, 6);
		excelReader.setCellData(xlfile, xlSheetName, 224, 7, actSizedropdown);
		
		String actPrintdropdown = Boolean.toString(Printdropdown.isDisplayed());
		String expPrintdropdown = excelReader.getCellData(xlSheetName, 225, 6);
		excelReader.setCellData(xlfile, xlSheetName, 225, 7, actPrintdropdown);
		
		String actExitBtn = Boolean.toString(ExitBtn.isDisplayed());
		String expExitBtn = excelReader.getCellData(xlSheetName, 226, 6);
		excelReader.setCellData(xlfile, xlSheetName, 226, 7, actExitBtn);
		
		String actpagesetupTab = Boolean.toString(pagesetupTab.isDisplayed());
		String exppagesetupTab = excelReader.getCellData(xlSheetName, 227, 6);
		excelReader.setCellData(xlfile, xlSheetName, 227, 7, actpagesetupTab);
		
		String actdesignerTab = Boolean.toString(designerTab.isDisplayed());
		String expdesignerTab = excelReader.getCellData(xlSheetName, 228, 6);
		excelReader.setCellData(xlfile, xlSheetName, 228, 7, actdesignerTab);

		String actpreviewTab = Boolean.toString(previewTab.isDisplayed());
		String exppreviewTab = excelReader.getCellData(xlSheetName, 229, 6);
		excelReader.setCellData(xlfile, xlSheetName, 229, 7, actpreviewTab);
		
		String actFieldsTab = Boolean.toString(FieldsTab.isDisplayed());
		String expFieldsTab = excelReader.getCellData(xlSheetName, 230, 6);
		excelReader.setCellData(xlfile, xlSheetName, 230, 7, actFieldsTab);
		
		String actPropertiesTab = Boolean.toString(PropertiesTab.isDisplayed());
		String expPropertiesTab = excelReader.getCellData(xlSheetName, 231, 6);
		excelReader.setCellData(xlfile, xlSheetName, 231, 7, actPropertiesTab);
		
		if(actNewdropdown.equalsIgnoreCase(expNewdropdown) && actOpendropdown.equalsIgnoreCase(expOpendropdown) && actSavedropdown.equalsIgnoreCase(expSavedropdown)
				&& actRemovedropdown.equalsIgnoreCase(expRemovedropdown) && actPageNo.equalsIgnoreCase(expPageNo) && actControldrpdown.equalsIgnoreCase(expControldrpdown) 
				&& actAlignmentdropdown.equalsIgnoreCase(expAlignmentdropdown) && actSizedropdown.equalsIgnoreCase(expSizedropdown)
				&& actPrintdropdown.equalsIgnoreCase(expPrintdropdown) && actExitBtn.equalsIgnoreCase(expExitBtn) && actpagesetupTab.equalsIgnoreCase(exppagesetupTab)
				&& actdesignerTab.equalsIgnoreCase(expdesignerTab) && actpreviewTab.equalsIgnoreCase(exppreviewTab) && actFieldsTab.equalsIgnoreCase(expFieldsTab) 
				&& actPropertiesTab.equalsIgnoreCase(expPropertiesTab))
		{
			excelReader.setCellData(xlfile, xlSheetName, 217, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 217, 8, resFail);
			return false;
		}
	}




	



	public boolean checkFieldsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(FieldsTab));
		FieldsTab.click();
		
		String actFinance_Purchases_HeaderExpandBtn = Boolean.toString(Finance_Purchases_HeaderExpandBtn.isDisplayed());
		String expFinance_Purchases_HeaderExpandBtn = excelReader.getCellData(xlSheetName, 232, 6);
		excelReader.setCellData(xlfile, xlSheetName, 232, 7, actFinance_Purchases_HeaderExpandBtn);
		
		String actPSS_BodyExpandbtn = Boolean.toString(PSS_BodyExpandbtn.isDisplayed());
		String expPSS_BodyExpandbtn = excelReader.getCellData(xlSheetName, 233, 6);
		excelReader.setCellData(xlfile, xlSheetName, 233, 7, actPSS_BodyExpandbtn);
		
		String actUserDetailsExpandbtn = Boolean.toString(UserDetailsExpandbtn.isDisplayed());
		String expUserDetailsExpandbtn = excelReader.getCellData(xlSheetName, 234, 6);
		excelReader.setCellData(xlfile, xlSheetName, 234, 7, actUserDetailsExpandbtn);
		
		String actMiscellaneousExpandBtn = Boolean.toString(MiscellaneousExpandBtn.isDisplayed());
		String expMiscellaneousExpandBtn = excelReader.getCellData(xlSheetName, 235, 6);
		excelReader.setCellData(xlfile, xlSheetName, 235, 7, actMiscellaneousExpandBtn);
		
		String actCompanyExpandBtn = Boolean.toString(CompanyExpandBtn.isDisplayed());
		String expCompanyExpandBtn = excelReader.getCellData(xlSheetName, 236, 6);
		excelReader.setCellData(xlfile, xlSheetName, 236, 7, actCompanyExpandBtn);

		if(actFinance_Purchases_HeaderExpandBtn.equalsIgnoreCase(expFinance_Purchases_HeaderExpandBtn) && actPSS_BodyExpandbtn.equalsIgnoreCase(expPSS_BodyExpandbtn)
				&& actUserDetailsExpandbtn.equalsIgnoreCase(expUserDetailsExpandbtn) && actMiscellaneousExpandBtn.equalsIgnoreCase(expMiscellaneousExpandBtn)
				&& actCompanyExpandBtn.equalsIgnoreCase(expCompanyExpandBtn))
		{
			excelReader.setCellData(xlfile, xlSheetName, 232, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 232, 8, resFail);
			return false;
		}
	}








	


	public static boolean checkDragandDropHeaderFiledsAndBodyFiledsIntoLayout() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickonHeader));
		clickonHeader.click();

		Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\LayoutHeaderFields.exe");

		Thread.sleep(60000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickonHeader));
		clickonHeader.click();
		Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnBody));
		clickOnBody.click();
		Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(Finance_Purchases_Header_WareHouseExpandBtn));
		Finance_Purchases_Header_WareHouseExpandBtn.click();

		Thread.sleep(5000);

		Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\LayoutBodyFields.exe");

		Thread.sleep(111000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(Savedropdown));
		getAction().moveToElement(Savedropdown).click().build().perform();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBttn));
		SaveBttn.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(layoutNameTextField));
		layoutNameTextField.sendKeys(excelReader.getCellData(xlSheetName, 237, 5));

		Thread.sleep(1000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(Save_OkBtn));
		Save_OkBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 237, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 237, 7, actMessage);

		Thread.sleep(2000);

		String actTitle = SavedlayoutNameTitle.getText();
		String expTitle = excelReader.getCellData(xlSheetName, 238, 6);
		excelReader.setCellData(xlfile, xlSheetName, 238, 7, actTitle);

		if(actMessage.equalsIgnoreCase(expMessage) && actTitle.equalsIgnoreCase(expTitle))
		{
			excelReader.setCellData(xlfile, xlSheetName, 237, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 237, 8, resFail);
			return false;
		}
	}

	
	
	



	public static boolean checkChangeFieldPropertiesInLyout() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatTxtField));
		getAction().contextClick(vatTxtField).build().perform();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(PropertiesBtn));
		PropertiesBtn.click();

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(widthTxt2));
		getAction().doubleClick(widthTxt2).build().perform();
		widthTxt2.sendKeys(excelReader.getCellData(xlSheetName, 239, 5));

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(grossTxtField));
		getAction().contextClick(grossTxtField).build().perform();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(PropertiesBtn));
		PropertiesBtn.click();

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(widthTxt2));
		getAction().doubleClick(widthTxt2).build().perform();
		widthTxt2.sendKeys(excelReader.getCellData(xlSheetName, 240, 5));

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateTxtField));
		getAction().contextClick(rateTxtField).build().perform();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(PropertiesBtn));
		PropertiesBtn.click();

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(widthTxt2));
		getAction().doubleClick(widthTxt2).build().perform();
		widthTxt2.sendKeys(excelReader.getCellData(xlSheetName, 241, 5));

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(quantityTxtField));
		getAction().contextClick(quantityTxtField).build().perform();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(PropertiesBtn));
		PropertiesBtn.click();

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(widthTxt2));
		getAction().doubleClick(widthTxt2).build().perform();
		widthTxt2.sendKeys(excelReader.getCellData(xlSheetName, 242, 5));

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxtfiled));
		getAction().contextClick(purchaseAccountTxtfiled).build().perform();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(PropertiesBtn));
		PropertiesBtn.click();

		Thread.sleep(4000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(widthTxt));
		getAction().doubleClick(widthTxt).build().perform();
		widthTxt.sendKeys(excelReader.getCellData(xlSheetName, 243, 5));

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemTxtField));
		getAction().contextClick(itemTxtField).build().perform();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(PropertiesBtn));
		PropertiesBtn.click();

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(widthTxt));
		getAction().doubleClick(widthTxt).build().perform();
		widthTxt.sendKeys(excelReader.getCellData(xlSheetName, 244, 5));

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxtField));
		getAction().contextClick(warehouseTxtField).build().perform();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(PropertiesBtn));
		PropertiesBtn.click();

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(widthTxt));
		getAction().doubleClick(widthTxt).build().perform();
		widthTxt.sendKeys(excelReader.getCellData(xlSheetName, 245, 5));

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatTxtField));
		getAction().contextClick(vatTxtField).build().perform();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(PropertiesBtn));
		PropertiesBtn.click();

		Thread.sleep(2000);

		Select sc= new Select(Txtalignment);
		sc.selectByVisibleText(excelReader.getCellData(xlSheetName, 246, 5));

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(grossTxtField));
		getAction().contextClick(grossTxtField).build().perform();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(PropertiesBtn));
		PropertiesBtn.click();

		Thread.sleep(2000);

		sc.selectByVisibleText(excelReader.getCellData(xlSheetName, 247, 5));

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateTxtField));
		getAction().contextClick(rateTxtField).build().perform();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(PropertiesBtn));
		PropertiesBtn.click();

		Thread.sleep(2000);

		sc.selectByVisibleText(excelReader.getCellData(xlSheetName, 248, 5));

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(quantityTxtField));
		getAction().contextClick(quantityTxtField).build().perform();

		Thread.sleep(4000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(PropertiesBtn));
		PropertiesBtn.click();

		Thread.sleep(2000);

		sc.selectByVisibleText(excelReader.getCellData(xlSheetName, 249, 5));

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(Savedropdown));
		Savedropdown.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBttn));
		SaveBttn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 239, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 239, 7, actMessage);

		Thread.sleep(2000);

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExitBtn));
			ExitBtn.click();

			Thread.sleep(3000);
			
			excelReader.setCellData(xlfile, xlSheetName, 239, 8, resPass);
			
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExitBtn));
			ExitBtn.click();

			Thread.sleep(3000);

			excelReader.setCellData(xlfile, xlSheetName, 239, 8, resFail);
			
			return false;
		}
	}







	public static boolean checkSavingPurchasesVatDesignLayout() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(createLayoutBtn));
		createLayoutBtn.click();

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(FieldsTab));
		FieldsTab.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickonHeader));
		clickonHeader.click();

		Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\LayoutHeaderFields.exe");

		Thread.sleep(60000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickonHeader));
		clickonHeader.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnBody));
		clickOnBody.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(Finance_Purchases_Header_WareHouseExpandBtn));
		Finance_Purchases_Header_WareHouseExpandBtn.click();

		Thread.sleep(5000);

		Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\LayoutBodyFields.exe");

		Thread.sleep(111000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(bodyGrid));
		getAction().contextClick(bodyGrid).build().perform();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(PropertiesBtn));
		PropertiesBtn.click();

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(bodyGridTopTextFiled));
		getAction().doubleClick(bodyGridTopTextFiled).build().perform();
		bodyGridTopTextFiled.sendKeys(excelReader.getCellData(xlSheetName, 250, 5));

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(Savedropdown));
		getAction().moveToElement(Savedropdown).click().build().perform();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBttn));
		SaveBttn.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(layoutNameTextField));
		layoutNameTextField.sendKeys(excelReader.getCellData(xlSheetName, 251, 5));

		Thread.sleep(1000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(Save_OkBtn));
		Save_OkBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 250, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 250, 7, actMessage);

		Thread.sleep(2000);

		String actTitle = SavedlayoutNameTitle.getText();
		String expTitle = excelReader.getCellData(xlSheetName, 251, 6);
		excelReader.setCellData(xlfile, xlSheetName, 251, 7, actTitle);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExitBtn));
		ExitBtn.click();

		Thread.sleep(3000);

		if (actMessage.equalsIgnoreCase(expMessage) && actTitle.equalsIgnoreCase(expTitle))
		{
			excelReader.setCellData(xlfile, xlSheetName, 250, 8, resPass);
			return true;

		} 
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 250, 8, resFail);
			return false;
		}
	}







	public static boolean checkEditingLayoutByAddingFields() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LayoutNameTxt));
		LayoutNameTxt.click();
		LayoutNameTxt.sendKeys(Keys.END);
		LayoutNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		LayoutNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 252, 5));
		Thread.sleep(2000);
		LayoutNameTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(LayoutName_Editlayout));
		LayoutName_Editlayout.click();

		Thread.sleep(2000);

		Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\EditLayout.exe");

		Thread.sleep(35000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(Savedropdown));
		Savedropdown.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBttn));
		SaveBttn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 252, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 252, 7, actMessage);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExitBtn));
		ExitBtn.click();

		Thread.sleep(3000);

		if (actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 252, 8, resPass);
			return true;

		} else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 252, 8, resFail);
			return false;
		}
	}



	

	
	

	public static boolean checkUpdatedLayoutAndDeleteTheLayout() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LayoutNameTxt));
		LayoutNameTxt.click();
		LayoutNameTxt.sendKeys(Keys.END);
		LayoutNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		LayoutNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 253, 5));
		Thread.sleep(2000);
		LayoutNameTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(LayoutName_Editlayout));
		LayoutName_Editlayout.click();

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateFiled));
		String actExchangeField = Boolean.toString(exchangeRateFiled.isDisplayed());
		String expExchangeField = excelReader.getCellData(xlSheetName, 253, 6);
		excelReader.setCellData(xlfile, xlSheetName, 253, 7, actExchangeField);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(locexchangeRateFiled));
		String actlocExchangeField = Boolean.toString(locexchangeRateFiled.isDisplayed());
		String explocExchangeField = excelReader.getCellData(xlSheetName, 254, 6);
		excelReader.setCellData(xlfile, xlSheetName, 254, 7, actlocExchangeField);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(discountFiledInbody));
		String actDiscountField = Boolean.toString(discountFiledInbody.isDisplayed());
		String expDiscountField = excelReader.getCellData(xlSheetName, 255, 6);
		excelReader.setCellData(xlfile, xlSheetName, 255, 7, actDiscountField);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(Removedropdown));
		Removedropdown.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(RemoveBtn));
		RemoveBtn.click();

		getWaitForAlert();

		getAlert().accept();

		String expMessage = excelReader.getCellData(xlSheetName, 256, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 256, 7, actMessage);

		System.out.println("ExchangeField 	  : " + actExchangeField + " 	 Value Expected : " + expExchangeField);
		System.out.println("locExchangeField  : " + actlocExchangeField + "  Value Expected : " + explocExchangeField);
		System.out.println("DiscountField     : " + actDiscountField + "  Value Expected : " + expDiscountField);

		if (actExchangeField == expExchangeField && actlocExchangeField == explocExchangeField 
				&& actDiscountField == expDiscountField && actMessage.equalsIgnoreCase(expMessage))  
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExitBtn));
			ExitBtn.click();

			Thread.sleep(3000);
			
			excelReader.setCellData(xlfile, xlSheetName, 253, 8, resPass);

			return true;
		} 
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExitBtn));
			ExitBtn.click();

			Thread.sleep(3000);
			
			excelReader.setCellData(xlfile, xlSheetName, 253, 8, resFail);

			return false;
		}
	}






	String FileName1 = null;

	public boolean checkPrintPdfOfPurchaseVoucherVATWithCreatedLayout() throws Exception, EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int count = grid_VoucherNoList.size();

		for (int i = 0; i < count; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 257, 5))) 
			{
				if (grid_CheckBoxList.get(i).isSelected()==false) 
				{
					grid_CheckBoxList.get(i).click();
					break;
				}
			}
		}      

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LayoutNameTxt));
		LayoutNameTxt.click();
		LayoutNameTxt.sendKeys(Keys.END);
		LayoutNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		LayoutNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 258, 5));
		Thread.sleep(2000);
		LayoutNameTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);

		Select s = new Select(LayoutOptionsdropdown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 259, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(footerPrintBtn));
		footerPrintBtn.click();

		Thread.sleep(5000);

		FileName1 = checkDownloadedFileName(getDriver());

		System.err.println(FileName1);

		/*Robot robot = new Robot();
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

						Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\SavingPrintPDFInPurchaseVoucherWithCreatedLayout.exe");

						Thread.sleep(5000);

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
					 	getDriver().switchTo().window(openTabs.get(0));*/


		String actPDF = getBaseDir()+"\\autoIt\\ExportFiles\\"+FileName1;
		String expPDF = getBaseDir()+"\\autoIt\\ImportFiles\\PurchaseVoucherVATPrintWithCreatedLayout.pdf";

		PDFUtil pdfutil = new PDFUtil();

		boolean result = pdfutil.compare(actPDF, expPDF);

		String actData = pdfutil.getText(actPDF);
		String expData = pdfutil.getText(expPDF);
		
		excelReader.setCellData(xlfile, xlSheetName, 258, 6, expData);
		excelReader.setCellData(xlfile, xlSheetName, 258, 7, actData);

		System.err.println(actData);
		System.err.println(expData);

		System.out.println("Compared Result  : "+result);

		if (actData.equalsIgnoreCase(expData)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 257, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 257, 8, resFail);
			return false;
		}
	}









	String FileName2 = null;

	public boolean checkPrintPdfOfPurchaseVoucherVAT() throws Exception, EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
					financialsMenu.click();

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
					financialsTransactionMenu.click();

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
					financialsTransactionsPurchaseMenu.click();

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
					purchaseVouchersVat.click();

					Thread.sleep(2000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int count = grid_VoucherNoList.size();

		for (int i = 0; i < count; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 260, 5))) 
			{
				if (grid_CheckBoxList.get(i).isSelected()==false) 
				{
					grid_CheckBoxList.get(i).click();
					break;
				}
			}
		}      

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LayoutNameTxt));
		LayoutNameTxt.click();
		LayoutNameTxt.sendKeys(Keys.END);
		LayoutNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		LayoutNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 261, 5));
		Thread.sleep(2000);
		LayoutNameTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);


		Select s = new Select(LayoutOptionsdropdown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 262, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(footerPrintBtn));
		footerPrintBtn.click();

		Thread.sleep(5000);

		FileName2 = checkDownloadedFileName(getDriver());

		/*Robot robot = new Robot();
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

					Thread.sleep(5000);

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
				 	getDriver().switchTo().window(openTabs.get(0));*/



		String actPDF = getBaseDir()+"\\autoIt\\ExportFiles\\"+FileName2;
		String expPDF = getBaseDir()+"\\autoIt\\ImportFiles\\PurchaseVoucherVATPrint.pdf";

		PDFUtil pdfutil = new PDFUtil();

		boolean result = pdfutil.compare(actPDF, expPDF);

		String actData = pdfutil.getText(actPDF);
		String expData = pdfutil.getText(expPDF);
		excelReader.setCellData(xlfile, xlSheetName, 260, 6, expData);
		excelReader.setCellData(xlfile, xlSheetName, 260, 7, actData);
		
		System.err.println(actData);
		System.err.println(expData);

		System.out.println("Compared Result  : "+result);

		if (actData.equalsIgnoreCase(expData)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 260, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 260, 8, resFail);
			return false;
		}
	}





	



	String FileName3 = null;

	public boolean checkPrintPdfOfPurchaseVoucherVATWithImportedXML() throws Exception, EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int count = grid_VoucherNoList.size();

		for (int i = 0; i < count; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 263, 5))) 
			{
				if (grid_CheckBoxList.get(i).isSelected()==false) 
				{
					grid_CheckBoxList.get(i).click();

					Thread.sleep(2000);
				}
			}
		}   

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createLayoutBtn));
		createLayoutBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Opendropdown));
		Opendropdown.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fromXMLBtn));
		fromXMLBtn.click();

		Thread.sleep(2000);

		Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\ImportingPurchaseVoucherVATPrintLayout.exe");

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Savedropdown));
		Savedropdown.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PrintSaveBttn));
		PrintSaveBttn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveTxtField));
		SaveTxtField.click();
		SaveTxtField.sendKeys(excelReader.getCellData(xlSheetName, 264, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Save_OkBtn));
		Save_OkBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 263, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 263, 7, actMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExitBtn));
		ExitBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		for (int i = 0; i < count; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 265, 5))) 
			{
				if (grid_CheckBoxList.get(i).isSelected()==false) 
				{
					grid_CheckBoxList.get(i).click();
					Thread.sleep(2000);
				}
			}
		}   

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LayoutNameTxt));
		LayoutNameTxt.click();
		LayoutNameTxt.sendKeys(Keys.END);
		LayoutNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		LayoutNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 266, 5));
		Thread.sleep(1000);
		LayoutNameTxt.sendKeys(Keys.TAB);

		Select s = new Select(LayoutOptionsdropdown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 267, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(footerPrintBtn));
		footerPrintBtn.click();

		Thread.sleep(5000);

		FileName3 = checkDownloadedFileName(getDriver());

		/*Robot robot = new Robot();
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

					Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\SavingPrintPDFInPurchaseVoucherWithImportedLayout.exe");

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
				 	getDriver().switchTo().window(openTabs.get(0));*/


		String actPDF = getBaseDir()+"\\autoIt\\ExportFiles\\"+FileName3;
		String expPDF = getBaseDir()+"\\autoIt\\ImportFiles\\PurchaseVoucherVATPrintWithImportedXML.pdf";

		PDFUtil pdfutil = new PDFUtil();

		boolean result = pdfutil.compare(actPDF, expPDF);

		String actData = pdfutil.getText(actPDF);
		String expData = pdfutil.getText(expPDF);
		excelReader.setCellData(xlfile, xlSheetName, 264, 6, expData);
		excelReader.setCellData(xlfile, xlSheetName, 264, 7, actData);

		System.err.println(actData);
		System.err.println(expData);

		System.out.println("Compared Result  : "+result);

		if (actMessage.equalsIgnoreCase(expMessage) && actData.equalsIgnoreCase(expData)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 263, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 263, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	

	String FileName4 = null;

	public boolean checkEmailOptionInPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int vcount = grid_VoucherNoList.size();

		for (int i = 0; i < vcount; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 268, 5))) 
			{
				if (grid_CheckBoxList.get(i).isSelected()==false) 
				{
					grid_CheckBoxList.get(i).click();
				}
			}
		}   

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LayoutNameTxt));
		LayoutNameTxt.click();
		LayoutNameTxt.sendKeys(Keys.END);
		LayoutNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		LayoutNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 269, 5));
		Thread.sleep(1000);
		LayoutNameTxt.sendKeys(Keys.TAB);

		Select s = new Select(LayoutOptionsdropdown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 270, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printScreenPickEmailChkbox));
		if (printScreenPickEmailChkbox.isSelected()==true) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printScreenPickEmailChkbox));
			printScreenPickEmailChkbox.click();
		} 

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printScreenSentEmailTxt));
		printScreenSentEmailTxt.click();
		printScreenSentEmailTxt.sendKeys(excelReader.getCellData(xlSheetName, 271, 5));
		printScreenSentEmailTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printSCreenSubjectTxt));
		printSCreenSubjectTxt.click();
		printSCreenSubjectTxt.sendKeys(excelReader.getCellData(xlSheetName, 272, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSendBtn));
		emailSendBtn.click();

		String ExpMessage=excelReader.getCellData(xlSheetName, 268, 6);
		String actMessage=checkValidationMessage(ExpMessage);
		excelReader.setCellData(xlfile, xlSheetName, 268, 7, actMessage);

		Thread.sleep(2000);

		Robot robot = new Robot();   
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_T);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_T);

		Thread.sleep(2000);

		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

		System.out.println("openTabs"+openTabs);

		getDriver().switchTo().window(openTabs.get(0));

		getDriver().switchTo().window(openTabs.get(1));

		getDriver().get(excelReader.getCellData(xlSheetName, 273, 5));

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
		userNameTxt.click();

		userNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 274, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
		NextBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
		PasswordTxt.click();

		PasswordTxt.sendKeys(excelReader.getCellData(xlSheetName, 275, 5));
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
		NextBtn.click();

		Thread.sleep(8000);

		int count = row1mailFromList.size();

		for (int i = 0; i < count; i++) 
		{
			String mailFromTxt = row1mailFromList.get(i).getText();

			if (mailFromTxt.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 276, 5))) 
			{
				row1mailFromList.get(i).click();
				break;
			}
		}

		Thread.sleep(3000);

		getAction().moveToElement(mailAttachmentDownloadBtn).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn));
		mailAttachmentDownloadBtn.click();

		Thread.sleep(5000);

		FileName4 = checkDownloadedFileName(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
		mailDeleteButton.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gmailUserBtn));
		gmailUserBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gmailSignOutBtn));
		gmailSignOutBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		removeAccountBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		removeDeleteBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		yesRemoveBtn.click();

		Thread.sleep(1000);

		getDriver().switchTo().window(openTabs.get(0));

		Thread.sleep(1000);

		getDriver().switchTo().window(openTabs.get(1)).close();

		Thread.sleep(1000);

		getDriver().switchTo().window(openTabs.get(0));

		String actPDF = getBaseDir()+"\\autoIt\\ExportFiles\\"+FileName4;
		String expPDF = getBaseDir()+"\\autoIt\\ImportFiles\\PurchaseVoucherVATPrintWithImportedXML.pdf";

		PDFUtil pdfutil = new PDFUtil();

		boolean result = pdfutil.compare(actPDF, expPDF);

		String actData = pdfutil.getText(actPDF);
		String expData = pdfutil.getText(expPDF);
		excelReader.setCellData(xlfile, xlSheetName, 269, 6, expData);
		excelReader.setCellData(xlfile, xlSheetName, 269, 7, actData);

		System.err.println(actData);
		System.err.println(expData);

		System.out.println("Compared Result  : "+result);

		if (actMessage.equalsIgnoreCase(ExpMessage) && actData.equalsIgnoreCase(expData)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 268, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 268, 8, resFail);
			return false;
		}
	}









	public static boolean checkRestoreCompanyForDocumentSet() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataMangementMenu));
		dataMangementMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restore));
		restore.click();

		boolean actRestoreCompanyBtnOfLoginPage  =restoreCompanyBtnOfLoginPage.isDisplayed();
		boolean actRestoreInputTxt               =restoreInputTxt.isDisplayed();
		boolean actRestoreBrowseBtn              =restoreBrowseBtn.isDisplayed();

		boolean expRestoreCompanyBtnOfLoginPage  =true;
		boolean expRestoreInputTxt               =true;
		boolean expRestoreBrowseBtn              =true;

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreBrowseBtn));
		restoreBrowseBtn.click();

		Thread.sleep(3000);

		Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\DocumentSet.exe");

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreCompanyBtnOfLoginPage));
		restoreCompanyBtnOfLoginPage.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreOverwriteYesBtn));
		restoreOverwriteYesBtn.click();

		getWaitForAlert();
		String actAlertTxt = getAlert().getText();
		getAlert().accept();

		System.err.println(actAlertTxt);

		String expAlertTxt = "Restore company code : 010";

		Thread.sleep(3000);

		if(actAlertTxt.equalsIgnoreCase(expAlertTxt))
		{
			return true;
		}
		else
		{
			return false;
		}
	}










	public static boolean checkCreatingDocumentSet() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_CreatelayoutBtn));
		DocumentSet_CreatelayoutBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSetTxt));
		DocumentSetTxt.click();
		DocumentSetTxt.sendKeys(excelReader.getCellData(xlSheetName, 277, 5));
		DocumentSetTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Select1stRowFilter));
		DocumentSet_Select1stRowFilter.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_EnterFilter));
		DocumentSet_EnterFilter.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterBtn));
		DocumentSet_AdvanceFilterBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterConjunctionDrpdwn));
		Select conjuction = new Select(DocumentSet_AdvanceFilterConjunctionDrpdwn);
		conjuction.selectByVisibleText(excelReader.getCellData(xlSheetName, 278, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterSelectFieldTxt));
		DocumentSet_AdvanceFilterSelectFieldTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterSelectFieldVendorAC));
		DocumentSet_AdvanceFilterSelectFieldVendorAC.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterSelectOperatorDrpdwn));
		Select selectOperator = new Select(DocumentSet_AdvanceFilterSelectOperatorDrpdwn);
		selectOperator.selectByVisibleText(excelReader.getCellData(xlSheetName, 279, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterCompareWithDrpdwn));
		Select compareWith = new Select(DocumentSet_AdvanceFilterCompareWithDrpdwn);
		compareWith.selectByVisibleText(excelReader.getCellData(xlSheetName, 280, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterValueTxt));
		DocumentSet_AdvanceFilterValueTxt.click();
		DocumentSet_AdvanceFilterValueTxt.sendKeys(excelReader.getCellData(xlSheetName, 281, 5));
		Thread.sleep(2000);
		DocumentSet_AdvanceFilterValueTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterOkBtn));
		DocumentSet_AdvanceFilterOkBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Select1stRowLayout));
		DocumentSet_Select1stRowLayout.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_EnterLayout));
		DocumentSet_EnterLayout.sendKeys(excelReader.getCellData(xlSheetName, 282, 5));
		Thread.sleep(3000);
		DocumentSet_EnterLayout.sendKeys(Keys.TAB);

		Thread.sleep(2000);



		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Select2ndRowFilter));
		DocumentSet_Select2ndRowFilter.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_EnterFilter));
		DocumentSet_EnterFilter.click();

		Thread.sleep(3000);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterBtn));
					DocumentSet_AdvanceFilterBtn.click();

					Thread.sleep(2000);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterConjunctionDrpdwn));
					conjuction.selectByVisibleText("Where");

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterSelectFieldTxt));
					DocumentSet_AdvanceFilterSelectFieldTxt.click();

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterSelectFieldVendorAC));
					DocumentSet_AdvanceFilterSelectFieldVendorAC.click();

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterSelectOperatorDrpdwn));
					selectOperator.selectByVisibleText("Equal to");

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterCompareWithDrpdwn));
					compareWith.selectByVisibleText("value");*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterValueCol));
		DocumentSet_AdvanceFilterValueCol.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterValueTxt));
		DocumentSet_AdvanceFilterValueTxt.sendKeys(Keys.END);
		DocumentSet_AdvanceFilterValueTxt.sendKeys(Keys.SHIFT, Keys.HOME);
		DocumentSet_AdvanceFilterValueTxt.sendKeys(excelReader.getCellData(xlSheetName, 283, 5));
		Thread.sleep(2000);
		DocumentSet_AdvanceFilterValueTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterOkBtn));
		DocumentSet_AdvanceFilterOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Select2ndRowLayout));
		DocumentSet_Select2ndRowLayout.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_EnterLayout));
		DocumentSet_EnterLayout.sendKeys(excelReader.getCellData(xlSheetName, 284, 5));
		Thread.sleep(2000);
		DocumentSet_EnterLayout.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Createlayout_SaveBtn));
		DocumentSet_Createlayout_SaveBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 277, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 277, 7, actMessage);

		Thread.sleep(2000);

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 277, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 277, 8, resFail);
			return false;
		}
	}










	public static boolean checkCreatingDocumentSetForDelete() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSetTxt));
		DocumentSetTxt.click();
		DocumentSetTxt.sendKeys(excelReader.getCellData(xlSheetName, 285, 5));
		DocumentSetTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Select1stRowFilter));
		DocumentSet_Select1stRowFilter.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_EnterFilter));
		DocumentSet_EnterFilter.click();

		Thread.sleep(2000);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterBtn));
					DocumentSet_AdvanceFilterBtn.click();

					Thread.sleep(3000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterConjunctionDrpdwn));
		Select conjuction = new Select(DocumentSet_AdvanceFilterConjunctionDrpdwn);
		conjuction.selectByVisibleText(excelReader.getCellData(xlSheetName, 286, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterSelectFieldTxt));
		DocumentSet_AdvanceFilterSelectFieldTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterSelectFieldVendorAC));
		DocumentSet_AdvanceFilterSelectFieldVendorAC.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterSelectOperatorDrpdwn));
		Select selectOperator = new Select(DocumentSet_AdvanceFilterSelectOperatorDrpdwn);
		selectOperator.selectByVisibleText(excelReader.getCellData(xlSheetName, 287, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterCompareWithDrpdwn));
		Select compareWith = new Select(DocumentSet_AdvanceFilterCompareWithDrpdwn);
		compareWith.selectByVisibleText(excelReader.getCellData(xlSheetName, 288, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterValueTxt));
		DocumentSet_AdvanceFilterValueTxt.click();
		DocumentSet_AdvanceFilterValueTxt.sendKeys(excelReader.getCellData(xlSheetName, 289, 5));
		Thread.sleep(2000);
		DocumentSet_AdvanceFilterValueTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterOkBtn));
		DocumentSet_AdvanceFilterOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Select1stRowLayout));
		DocumentSet_Select1stRowLayout.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_EnterLayout));
		DocumentSet_EnterLayout.sendKeys(excelReader.getCellData(xlSheetName, 290, 5));
		Thread.sleep(2000);
		DocumentSet_EnterLayout.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Createlayout_SaveBtn));
		DocumentSet_Createlayout_SaveBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 285, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 285, 7, actMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Createlayout_CancelBtn));
			DocumentSet_Createlayout_CancelBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 285, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Createlayout_CancelBtn));
			DocumentSet_Createlayout_CancelBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 285, 8, resFail);
			return false;
		}
	}









	public static boolean checkEditingAndUpdatingDocumentSetForDelete() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet));
		DocumentSet.click();
		DocumentSet.sendKeys(Keys.END);
		DocumentSet.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		DocumentSet.sendKeys(excelReader.getCellData(xlSheetName, 291, 5));
		Thread.sleep(2000);
		DocumentSet.sendKeys(Keys.TAB);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Editlayout));
		DocumentSet_Editlayout.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Select2ndRowFilter));
		DocumentSet_Select2ndRowFilter.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_EnterFilter));
		DocumentSet_EnterFilter.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterBtn));
		DocumentSet_AdvanceFilterBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterConjunctionDrpdwn));
		Select conjuction = new Select(DocumentSet_AdvanceFilterConjunctionDrpdwn);
		conjuction.selectByVisibleText(excelReader.getCellData(xlSheetName, 292, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterSelectFieldTxt));
		DocumentSet_AdvanceFilterSelectFieldTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterSelectFieldVendorAC));
		DocumentSet_AdvanceFilterSelectFieldVendorAC.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterSelectOperatorDrpdwn));
		Select selectOperator = new Select(DocumentSet_AdvanceFilterSelectOperatorDrpdwn);
		selectOperator.selectByVisibleText(excelReader.getCellData(xlSheetName, 293, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterCompareWithDrpdwn));
		Select compareWith = new Select(DocumentSet_AdvanceFilterCompareWithDrpdwn);
		compareWith.selectByVisibleText(excelReader.getCellData(xlSheetName, 294, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterValueTxt));
		DocumentSet_AdvanceFilterValueTxt.click();
		DocumentSet_AdvanceFilterValueTxt.sendKeys(excelReader.getCellData(xlSheetName, 295, 5));
		Thread.sleep(2000);
		DocumentSet_AdvanceFilterValueTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_AdvanceFilterOkBtn));
		DocumentSet_AdvanceFilterOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Select2ndRowLayout));
		DocumentSet_Select2ndRowLayout.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_EnterLayout));
		DocumentSet_EnterLayout.sendKeys(excelReader.getCellData(xlSheetName, 296, 5));
		Thread.sleep(2000);
		DocumentSet_EnterLayout.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Createlayout_SaveBtn));
		DocumentSet_Createlayout_SaveBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 291, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 291, 7, actMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Createlayout_CancelBtn));
			DocumentSet_Createlayout_CancelBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 291, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Createlayout_CancelBtn));
			DocumentSet_Createlayout_CancelBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 291, 8, resFail);
			return false;
		}
	}








	public static boolean checkDeletingDocumentSetForDelete() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet));
		DocumentSet.click();
		DocumentSet.sendKeys(Keys.END);
		DocumentSet.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		DocumentSet.sendKeys(excelReader.getCellData(xlSheetName, 297, 5));
		Thread.sleep(2000);
		DocumentSet.sendKeys(Keys.TAB);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Editlayout));
		DocumentSet_Editlayout.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Createlayout_DeleteBtn));
		DocumentSet_Createlayout_DeleteBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 297, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 297, 7, actMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Createlayout_CancelBtn));
			DocumentSet_Createlayout_CancelBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 297, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet_Createlayout_CancelBtn));
			DocumentSet_Createlayout_CancelBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 297, 8, resFail);
			return false;
		}
	}






	String FileName5 = null;

	public boolean checkPrintPdfOfPurchaseVoucherVATWithDocumentSetVendorA() throws Exception, EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int count = grid_VoucherNoList.size();

		for (int i = 0; i < count; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 298, 5))) 
			{
				if (grid_CheckBoxList.get(i).isSelected()==false) 
				{
					grid_CheckBoxList.get(i).click();
					break;
				}
			}
		}      

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet));
		DocumentSet.click();
		DocumentSet.sendKeys(Keys.END);
		DocumentSet.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		DocumentSet.sendKeys(excelReader.getCellData(xlSheetName, 299, 5));
		Thread.sleep(2000);
		DocumentSet.sendKeys(Keys.TAB);
		Thread.sleep(2000);


		Select s = new Select(LayoutOptionsdropdown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 300, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(footerPrintBtn));
		footerPrintBtn.click();

		Thread.sleep(5000);

		FileName5 = checkDownloadedFileName(getDriver());

		/*Robot robot = new Robot();
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

					Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\SavingPrintPDFInPurchaseVoucherWithDocumentSetVendorA.exe");

					Thread.sleep(5000);

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
				 	getDriver().switchTo().window(openTabs.get(0));*/


		String actPDF = getBaseDir()+"\\autoIt\\ExportFiles\\"+FileName5;
		String expPDF = getBaseDir()+"\\autoIt\\ImportFiles\\PurchaseVoucherVATPrintWithDocumentSetVendorA.pdf";

		PDFUtil pdfutil = new PDFUtil();

		boolean result = pdfutil.compare(actPDF, expPDF);

		String actData = pdfutil.getText(actPDF);
		String expData = pdfutil.getText(expPDF);
		excelReader.setCellData(xlfile, xlSheetName, 298, 6, expData);
		excelReader.setCellData(xlfile, xlSheetName, 298, 7, actData);

		System.err.println(actData);
		System.err.println(expData);

		System.out.println("Compared Result  : "+result);

		if (actData.equalsIgnoreCase(expData)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 298, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 298, 8, resFail);
			return false;
		}
	}







	String FileName6 = null;

	public boolean checkPrintPdfOfPurchaseVoucherVATWithDocumentSetVendorB() throws Exception, EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int count = grid_VoucherNoList.size();

		for (int i = 0; i < count; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 301, 5)) || VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 302, 5))) 
			{
				if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 301, 5)) && grid_CheckBoxList.get(i).isSelected()==true) 
				{
					grid_CheckBoxList.get(i).click();
				}

				if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 302, 5)) && grid_CheckBoxList.get(i).isSelected()==false) 
				{
					grid_CheckBoxList.get(i).click();
				}
			}
		}      

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DocumentSet));
		DocumentSet.click();
		DocumentSet.sendKeys(Keys.END);
		DocumentSet.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		DocumentSet.sendKeys(excelReader.getCellData(xlSheetName, 303, 5));
		Thread.sleep(2000);
		DocumentSet.sendKeys(Keys.TAB);
		Thread.sleep(2000);


		Select s = new Select(LayoutOptionsdropdown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 304, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(footerPrintBtn));
		footerPrintBtn.click();

		Thread.sleep(5000);

		FileName6 = checkDownloadedFileName(getDriver());

		/*Robot robot = new Robot();
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

					Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\SavingPrintPDFInPurchaseVoucherWithDocumentSetVendorB.exe");

					Thread.sleep(5000);

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
				 	getDriver().switchTo().window(openTabs.get(0));*/


		String actPDF = getBaseDir()+"\\autoIt\\ExportFiles\\"+FileName6;
		String expPDF = getBaseDir()+"\\autoIt\\ImportFiles\\PurchaseVoucherVATPrintWithDocumentSetVendorB.pdf";

		PDFUtil pdfutil = new PDFUtil();

		boolean result = pdfutil.compare(actPDF, expPDF);

		String actData = pdfutil.getText(actPDF);
		String expData = pdfutil.getText(expPDF);
		excelReader.setCellData(xlfile, xlSheetName, 301, 6, expData);
		excelReader.setCellData(xlfile, xlSheetName, 301, 7, actData);

		System.err.println(actData);
		System.err.println(expData);

		System.out.println("Compared Result  : "+result);

		if (actData.equalsIgnoreCase(expData)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 301, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 301, 8, resFail);
			return false;
		}
	}








	public boolean checksalesInvoiceVatWithBuyerSellerPriceBooklist() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
		financialTransactionSalesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
		salesInvoiceVATVoucher.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 306, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 307, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.END);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SHIFT,Keys.HOME);
		salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 308, 5));
		Thread.sleep(2000);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.sendKeys(Keys.END);
		jurisdictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 309, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 310, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 311, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		String actRate1 = select1stRow_14thColumn.getText();
		String expRate1 = excelReader.getCellData(xlSheetName, 306, 6);
		excelReader.setCellData(xlfile, xlSheetName, 306, 7, actRate1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 312, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 313, 5));
		enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
		getAction().moveToElement(select1stRow_19thColumn).click().build().perform();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		batchPickOnFIFOIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
		batchOkIcon.click();




		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.click();

		pvWareHouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 315, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 316, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
		select2ndRow_5thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_14thColumn));
		String actRate2 = select2ndRow_14thColumn.getText();
		String expRate2 = excelReader.getCellData(xlSheetName, 307, 6);
		excelReader.setCellData(xlfile, xlSheetName, 307, 7, actRate2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_8thColumn));
		select2ndRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 317, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 318, 5));
		enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_11thColumn));
		select2ndRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_21thColumn));
		select2ndRow_21thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
		RMAPopupSerialNoExpansionBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
		searchRMASelectAllChkBox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
		searchRMAOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaoutwardOkBtn));
		rmaoutwardOkBtn.click();

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 1");
		expMessage.add("Saving in background.");*/

		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : 1");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 308, 6);
		excelReader.setCellData(xlfile, xlSheetName, 308, 7, actMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		System.out.println("************************************************************************");

		System.out.println("Rates1 : "+actRate1+"  Value Expected  "+expRate1);
		System.out.println("Rates2 : "+actRate2+"  Value Expected  "+expRate2);

		if(actMessage.equals(expMessage) && actRate1.equalsIgnoreCase(expRate1) && actRate2.equalsIgnoreCase(expRate2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 306, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 306, 8, resFail);
			return false;
		}
	}








	public boolean checkSavedVoucherInSalesINvoiceVAT() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		String expMessage = "Voucher Loaded Successfully";

		String actMessage = checkValidationMessage(expMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		String actDocNo = documentNumberTxt.getAttribute("value");
		String actCustomerAccount = customerAccountTxt.getAttribute("value");
		String actDepartment = departmentTxt.getAttribute("value");


		String actWarehouse = select1stRow_1stColumn.getText();
		String actItem = select1stRow_2ndColumn.getText();
		String actQuantity = select1stRow_11thColumn.getText();
		String actRate = select1stRow_14thColumn.getText();
		String actGross = select1stRow_15thColumn.getText();
		String actBatch = select1stRow_19thColumn.getText();


		String actWarehouse1 = select2ndRow_1stColumn.getText();
		String actItem1 = select2ndRow_2ndColumn.getText();
		String actQuantity1 = select2ndRow_11thColumn.getText();
		String actRate1 = select2ndRow_14thColumn.getText();
		String actGross1 = select2ndRow_15thColumn.getText();
		String actRMA = select2ndRow_21stColumn.getText();

		String expDocNo = excelReader.getCellData(xlSheetName, 318, 6);
		excelReader.setCellData(xlfile, xlSheetName, 318, 7, actDocNo);
		String expCustomerAccount = excelReader.getCellData(xlSheetName, 319, 6);
		excelReader.setCellData(xlfile, xlSheetName, 319, 7, actCustomerAccount);
		String expDepartment = excelReader.getCellData(xlSheetName, 320, 6);
		excelReader.setCellData(xlfile, xlSheetName, 320, 7, actDepartment);

		
		String expWarehouse = excelReader.getCellData(xlSheetName, 321, 6);
		excelReader.setCellData(xlfile, xlSheetName, 321, 7, actWarehouse);
		String expItem = excelReader.getCellData(xlSheetName, 322, 6);
		excelReader.setCellData(xlfile, xlSheetName, 322, 7, actItem);
		String expQuantity = excelReader.getCellData(xlSheetName, 323, 6);
		excelReader.setCellData(xlfile, xlSheetName, 323, 7, actQuantity);
		String expRate = excelReader.getCellData(xlSheetName, 324, 6);
		excelReader.setCellData(xlfile, xlSheetName, 324, 7, actRate);
		String expGross = excelReader.getCellData(xlSheetName, 325, 6);
		excelReader.setCellData(xlfile, xlSheetName, 325, 7, actGross);
		String expBatch = excelReader.getCellData(xlSheetName, 326, 6);
		excelReader.setCellData(xlfile, xlSheetName, 326, 7, actBatch);


		String expWarehouse1 = excelReader.getCellData(xlSheetName, 327, 6);
		excelReader.setCellData(xlfile, xlSheetName, 327, 7, actWarehouse1);
		String expItem1 = excelReader.getCellData(xlSheetName, 328, 6);
		excelReader.setCellData(xlfile, xlSheetName, 328, 7, actItem1);
		String expQuantity1 = excelReader.getCellData(xlSheetName, 329, 6);
		excelReader.setCellData(xlfile, xlSheetName, 329, 7, actQuantity1);
		String expRate1 = excelReader.getCellData(xlSheetName, 330, 6);
		excelReader.setCellData(xlfile, xlSheetName, 330, 7, actRate1);
		String expGross1 = excelReader.getCellData(xlSheetName, 331, 6);
		excelReader.setCellData(xlfile, xlSheetName, 331, 7, actGross1);
		String expRMA = excelReader.getCellData(xlSheetName, 332, 6);
		excelReader.setCellData(xlfile, xlSheetName, 332, 7, actRMA);

		System.out.println("********* Voucher No        : " + actDocNo + "  value expected  " + expDocNo);
		System.out.println("********* Customer Account  : " + actCustomerAccount + "  value expected  " + expCustomerAccount);
		System.out.println("********* Department        : " + actDepartment + "  value expected  " + expDepartment);
		System.out.println("********* Warehouse         : " + actWarehouse + "  value expected  " + expWarehouse);
		System.out.println("********* Item              : " + actItem + "  value expected  " + expItem);
		System.out.println("********* Quantity          : " + actQuantity + "  value expected  " + expQuantity);
		System.out.println("********* Rate              : " + actRate + "  value expected  " + expRate);
		System.out.println("********* Gross             : " + actGross + "  value expected  " + expGross);
		System.out.println("********* BATCH Col         : " + actBatch + "  value expected  " + expBatch);


		System.out.println("********* Warehouse11         : " + actWarehouse1 + "  value expected  " + expWarehouse1);
		System.out.println("********* Item1              : " + actItem1 + "  value expected  " + expItem1);
		System.out.println("********* Quantity1          : " + actQuantity1 + "  value expected  " + expQuantity1);
		System.out.println("********* Rate1              : " + actRate1 + "  value expected  " + expRate1);
		System.out.println("********* Gross1             : " + actGross1 + "  value expected  " + expGross1);
		System.out.println("********* RMA Col            : " + actRMA + "  value expected  " + expRMA);

		Thread.sleep(2000);

		if (actDocNo.equalsIgnoreCase(expDocNo) && actCustomerAccount.equalsIgnoreCase(expCustomerAccount)
				&& actDepartment.equalsIgnoreCase(expDepartment) && actWarehouse.equalsIgnoreCase(expWarehouse)
				&& actItem.equalsIgnoreCase(expItem) && actQuantity.equalsIgnoreCase(expQuantity)
				&& actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross)
				/*&& actBatch.equalsIgnoreCase(expBatch)*/

				&& actWarehouse1.equalsIgnoreCase(expWarehouse1)
				&& actItem1.equalsIgnoreCase(expItem1) && actQuantity1.equalsIgnoreCase(expQuantity1)
				&& actRate1.equalsIgnoreCase(expRate1) && actGross1.equalsIgnoreCase(expGross1)
				/*&& actRMA.equalsIgnoreCase(expRMA)*/)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 318, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 318, 8, resFail);
			return false;
		}
	}








	public static boolean checkUpdatingBuyerPricebookFormulaForSalesInvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceSettingsBtn));
		salesInvoiceSettingsBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
		editScreenTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBehaviourDropdown));
		ScrollIntoView(editScreenBehaviourDropdown);
		Select behaviour = new Select(editScreenBehaviourDropdown);
		behaviour.selectByVisibleText(excelReader.getCellData(xlSheetName, 333, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBehaviourFormulaTxt));
		editScreenBehaviourFormulaTxt.sendKeys(excelReader.getCellData(xlSheetName, 334, 5));
		editScreenBehaviourFormulaTxt.sendKeys(Keys.TAB);
		getAction().moveToElement(editScreenBehaviourFormulaTxt).pause(1000).click().build().perform();
		
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBehaviourFormulaOkBtn));
		editScreenBehaviourFormulaOkBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		System.out.println("******************checkUpdatingBuyerPricebookFormulaForSalesInvoiceVAT*****************");

		String expMessage = excelReader.getCellData(xlSheetName, 333, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 333, 7, actMessage);

		Thread.sleep(2000);

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 333, 8, resPass);
			return true;
		}
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 333, 8, resFail);
			return false;
		}
	}





	public static boolean checkSavingSalesInvoiceVATVoucher2WithRateFormula() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
		financialTransactionSalesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
		salesInvoiceVATVoucher.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(Keys.END);
		customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 335, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 336, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
		salesInvoiceVATPlaceOFSupply.click();
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.END);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SHIFT,Keys.HOME);
		salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 337, 5));
		Thread.sleep(2000);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(Keys.END);
		jurisdictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 338, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 339, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 340, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		String actRate1 = select1stRow_14thColumn.getText();
		String expRate1 = excelReader.getCellData(xlSheetName, 335, 6);
		excelReader.setCellData(xlfile, xlSheetName, 335, 7, actRate1);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 341, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 342, 5));
		enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
		select1stRow_19thColumn.click();

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));

					enter_Batch.sendKeys(Keys.SPACE);;
		 */
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		batchPickOnFIFOIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
		batchOkIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.click();
		pvWareHouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 343, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 344, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
		select2ndRow_5thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_14thColumn));
		String actRate2 = select2ndRow_14thColumn.getText();
		String expRate2 = excelReader.getCellData(xlSheetName, 337, 6);
		excelReader.setCellData(xlfile, xlSheetName, 337, 7, actRate2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_8thColumn));
		select2ndRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 345, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 346, 5));
		enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_11thColumn));
		select2ndRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_21thColumn));
		select2ndRow_21thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
		RMAPopupSerialNoExpansionBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
		searchRMASelectAllChkBox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
		searchRMAOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaoutwardOkBtn));
		rmaoutwardOkBtn.click();

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 2");
		expMessage.add("Saving in background.");*/
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : 2");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 337, 6);
		excelReader.setCellData(xlfile, xlSheetName, 337, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		System.out.println("************************************************************************");

		System.out.println("Rates1 : "+actRate1+"  Value Expected  "+expRate1);
		System.out.println("Rates2 : "+actRate2+"  Value Expected  "+expRate2);

		if(actMessage.equals(expMessage) && actRate1.equalsIgnoreCase(expRate1) && actRate2.equalsIgnoreCase(expRate2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 335, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 335, 8, resFail);
			return false;
		}
	}





	public boolean checkSavedVoucher2InSalesINvoiceVATWithFormulaControl() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher Loaded Successfully");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		String actDocNo = documentNumberTxt.getAttribute("value");
		String actCustomerAccount = customerAccountTxt.getAttribute("value");
		String actDepartment = departmentTxt.getAttribute("value");


		String actWarehouse = select1stRow_1stColumn.getText();
		String actItem = select1stRow_2ndColumn.getText();
		String actQuantity = select1stRow_11thColumn.getText();
		String actRate = select1stRow_14thColumn.getText();
		String actGross = select1stRow_15thColumn.getText();
		String actBatch = select1stRow_19thColumn.getText();


		String actWarehouse1 = select2ndRow_1stColumn.getText();
		String actItem1 = select2ndRow_2ndColumn.getText();
		String actQuantity1 = select2ndRow_11thColumn.getText();
		String actRate1 = select2ndRow_14thColumn.getText();
		String actGross1 = select2ndRow_15thColumn.getText();
		String actRMA = select2ndRow_21stColumn.getText();

		String expDocNo = excelReader.getCellData(xlSheetName, 347, 6);
		excelReader.setCellData(xlfile, xlSheetName, 347, 7, actDocNo);
		
		String expCustomerAccount = excelReader.getCellData(xlSheetName, 348, 6);
		excelReader.setCellData(xlfile, xlSheetName, 348, 7, actCustomerAccount);
		
		String expDepartment = excelReader.getCellData(xlSheetName, 349, 6);
		excelReader.setCellData(xlfile, xlSheetName, 349, 7, actDepartment);


		String expWarehouse = excelReader.getCellData(xlSheetName, 350, 6);
		excelReader.setCellData(xlfile, xlSheetName, 350, 7, actWarehouse);
		
		String expItem = excelReader.getCellData(xlSheetName, 351, 6);
		excelReader.setCellData(xlfile, xlSheetName, 351, 7, actItem);
		
		String expQuantity = excelReader.getCellData(xlSheetName, 352, 6);
		excelReader.setCellData(xlfile, xlSheetName, 352, 7, actQuantity);
		
		
		String expRate = excelReader.getCellData(xlSheetName, 353, 6);
		excelReader.setCellData(xlfile, xlSheetName, 353, 7, actRate);
		
		String expGross = excelReader.getCellData(xlSheetName, 354, 6);
		excelReader.setCellData(xlfile, xlSheetName, 354, 7, actGross);
		
		String expBatch = excelReader.getCellData(xlSheetName, 355, 6);
		excelReader.setCellData(xlfile, xlSheetName, 355, 7, actBatch);



		String expWarehouse1 = excelReader.getCellData(xlSheetName, 356, 6);
		excelReader.setCellData(xlfile, xlSheetName, 356, 7, actWarehouse1);
		
		String expItem1 = excelReader.getCellData(xlSheetName, 357, 6);
		excelReader.setCellData(xlfile, xlSheetName, 357, 7, actItem1);
		
		String expQuantity1 = excelReader.getCellData(xlSheetName, 358, 6);
		excelReader.setCellData(xlfile, xlSheetName, 358, 7, actQuantity1);
		
		String expRate1 = excelReader.getCellData(xlSheetName, 359, 6);
		excelReader.setCellData(xlfile, xlSheetName, 359, 7, actRate1);
		
		String expGross1 = excelReader.getCellData(xlSheetName, 360, 6);
		excelReader.setCellData(xlfile, xlSheetName, 360, 7, actGross1);
		
		String expRMA = excelReader.getCellData(xlSheetName, 361, 6);
		excelReader.setCellData(xlfile, xlSheetName, 361, 7, actRMA);

		System.out.println("********* Voucher No        : " + actDocNo + "  value expected  " + expDocNo);
		System.out.println("********* Customer Account  : " + actCustomerAccount + "  value expected  " + expCustomerAccount);
		System.out.println("********* Department        : " + actDepartment + "  value expected  " + expDepartment);
		System.out.println("********* Warehouse         : " + actWarehouse + "  value expected  " + expWarehouse);
		System.out.println("********* Item              : " + actItem + "  value expected  " + expItem);
		System.out.println("********* Quantity          : " + actQuantity + "  value expected  " + expQuantity);
		System.out.println("********* Rate              : " + actRate + "  value expected  " + expRate);
		System.out.println("********* Gross             : " + actGross + "  value expected  " + expGross);
		System.out.println("********* BATCH Col         : " + actBatch + "  value expected  " + expBatch);


		System.out.println("********* Warehouse11        : " + actWarehouse1 + "  value expected  " + expWarehouse1);
		System.out.println("********* Item1              : " + actItem1 + "  value expected  " + expItem1);
		System.out.println("********* Quantity1          : " + actQuantity1 + "  value expected  " + expQuantity1);
		System.out.println("********* Rate1              : " + actRate1 + "  value expected  " + expRate1);
		System.out.println("********* Gross1             : " + actGross1 + "  value expected  " + expGross1);
		System.out.println(actRMA);
		System.out.println(expRMA);

		if (actDocNo.equalsIgnoreCase(expDocNo) && actCustomerAccount.equalsIgnoreCase(expCustomerAccount)
				&& actDepartment.equalsIgnoreCase(expDepartment) && actWarehouse.equalsIgnoreCase(expWarehouse)
				&& actItem.equalsIgnoreCase(expItem) && actQuantity.equalsIgnoreCase(expQuantity)
				&& actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross)
				/*&& actBatch.equalsIgnoreCase(expBatch)*/


				&& actWarehouse1.equalsIgnoreCase(expWarehouse1)
				&& actItem1.equalsIgnoreCase(expItem1) && actQuantity1.equalsIgnoreCase(expQuantity1)
				&& actRate1.equalsIgnoreCase(expRate1) && actGross1.equalsIgnoreCase(expGross1)
				/*&& actRMA.equalsIgnoreCase(expRMA)*/)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 347, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 347, 8, resFail);
			return false;
		}
	}







	public static boolean checkUpdatingBuyerPricebookFormulaForSalesInvoiceVATVoucher3() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceSettingsBtn));
		salesInvoiceSettingsBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
		editScreenTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBehaviourFormulaTxt));
		editScreenBehaviourFormulaTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		editScreenBehaviourFormulaTxt.sendKeys(excelReader.getCellData(xlSheetName, 362, 5));
		editScreenBehaviourFormulaTxt.sendKeys(Keys.TAB);
		getAction().moveToElement(editScreenBehaviourFormulaTxt).pause(1000).click().build().perform();
		
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBehaviourFormulaOkBtn));
		editScreenBehaviourFormulaOkBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		System.out.println("******************checkUpdatingBuyerPricebookFormulaForSalesInvoiceVATVoucher3*****************");

		String expMessage = excelReader.getCellData(xlSheetName, 362, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 362, 7, actMessage);

		Thread.sleep(2000);

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 362, 8, resPass);
			return true;
		}
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 362, 8, resFail);
			return false;
		}
	}



	
	


	public static boolean checkSavingSalesInvoiceVATVoucher3WithRateFormula() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
		financialTransactionSalesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
		salesInvoiceVATVoucher.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.click();
		customerAccountTxt.sendKeys(Keys.END);
		customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 363, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 364, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
		salesInvoiceVATPlaceOFSupply.click();
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.END);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SHIFT,Keys.HOME);
		salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 365, 5));
		Thread.sleep(2000);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(Keys.END);
		jurisdictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 366, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.click();
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 367, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 368, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 369, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 370, 5));
		enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
		select1stRow_19thColumn.click();


		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		batchPickOnFIFOIcon.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
		batchOkIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		String actRate1 = select1stRow_14thColumn.getText();
		String expRate1 = excelReader.getCellData(xlSheetName, 363, 6);
		excelReader.setCellData(xlfile, xlSheetName, 363, 7, actRate1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.click();
		pvWareHouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 371, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 372, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
		select2ndRow_5thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_8thColumn));
		select2ndRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 373, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 374, 5));
		enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_11thColumn));
		select2ndRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_21thColumn));
		select2ndRow_21thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
		RMAPopupSerialNoExpansionBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
		searchRMASelectAllChkBox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
		searchRMAOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaoutwardOkBtn));
		rmaoutwardOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_14thColumn));
		String actRate2 = select2ndRow_14thColumn.getText();
		String expRate2 = excelReader.getCellData(xlSheetName, 364, 6);
		excelReader.setCellData(xlfile, xlSheetName, 364, 7, actRate2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 3");
		expMessage.add("Saving in background.");*/
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : 3");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 365, 6);
		excelReader.setCellData(xlfile, xlSheetName, 365, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		System.out.println("************************************************************************");

		System.out.println("Rates1 : "+actRate1+"  Value Expected  "+expRate1);
		System.out.println("Rates2 : "+actRate2+"  Value Expected  "+expRate2);

		if(actMessage.equals(expMessage) && actRate1.equalsIgnoreCase(expRate1) && actRate2.equalsIgnoreCase(expRate2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 363, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 363, 8, resFail);
			return false;
		}
	}




	

	
	
	
	public boolean checkSavedVoucher3InSalesINvoiceVATWithFormulaControl() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher Loaded Successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		String actDocNo = documentNumberTxt.getAttribute("value");
		String actCustomerAccount = customerAccountTxt.getAttribute("value");
		String actDepartment = departmentTxt.getAttribute("value");


		String actWarehouse = select1stRow_1stColumn.getText();
		String actItem = select1stRow_2ndColumn.getText();
		String actQuantity = select1stRow_11thColumn.getText();
		String actRate = select1stRow_14thColumn.getText();
		String actGross = select1stRow_15thColumn.getText();
		String actBatch = select1stRow_19thColumn.getText();


		String actWarehouse1 = select2ndRow_1stColumn.getText();
		String actItem1 = select2ndRow_2ndColumn.getText();
		String actQuantity1 = select2ndRow_11thColumn.getText();
		String actRate1 = select2ndRow_14thColumn.getText();
		String actGross1 = select2ndRow_15thColumn.getText();
		String actRMA = select2ndRow_21stColumn.getText();

		String expDocNo = excelReader.getCellData(xlSheetName, 375, 6);
		excelReader.setCellData(xlfile, xlSheetName, 375, 7, actDocNo);
		
		String expCustomerAccount = excelReader.getCellData(xlSheetName, 376, 6);
		excelReader.setCellData(xlfile, xlSheetName, 376, 7, actCustomerAccount);
		
		String expDepartment = excelReader.getCellData(xlSheetName, 377, 6);
		excelReader.setCellData(xlfile, xlSheetName, 377, 7, actDepartment);

		String expWarehouse = excelReader.getCellData(xlSheetName, 378, 6);
		excelReader.setCellData(xlfile, xlSheetName, 378, 7, actWarehouse);
		
		String expItem = excelReader.getCellData(xlSheetName, 379, 6);
		excelReader.setCellData(xlfile, xlSheetName, 379, 7, actItem);
		
		String expQuantity = excelReader.getCellData(xlSheetName, 380, 6);
		excelReader.setCellData(xlfile, xlSheetName, 380, 7, actQuantity);
		
		String expRate = excelReader.getCellData(xlSheetName, 381, 6);
		excelReader.setCellData(xlfile, xlSheetName, 381, 7, actRate);
		
		String expGross = excelReader.getCellData(xlSheetName, 382, 6);
		excelReader.setCellData(xlfile, xlSheetName, 382, 7, actGross);
		
		String expBatch = excelReader.getCellData(xlSheetName, 383, 6);
		excelReader.setCellData(xlfile, xlSheetName, 383, 7, actBatch);


		String expWarehouse1 = excelReader.getCellData(xlSheetName, 384, 6);
		excelReader.setCellData(xlfile, xlSheetName, 384, 7, actWarehouse1);
		
		String expItem1 = excelReader.getCellData(xlSheetName, 385, 6);
		excelReader.setCellData(xlfile, xlSheetName, 385, 7, actItem1);
		
		String expQuantity1 = excelReader.getCellData(xlSheetName, 386, 6);
		excelReader.setCellData(xlfile, xlSheetName, 386, 7, actQuantity1);
		
		String expRate1 = excelReader.getCellData(xlSheetName, 387, 6);
		excelReader.setCellData(xlfile, xlSheetName, 387, 7, actRate1);
		
		String expGross1 = excelReader.getCellData(xlSheetName, 388, 6);
		excelReader.setCellData(xlfile, xlSheetName, 388, 7, actGross1);
		
		String expRMA = excelReader.getCellData(xlSheetName, 389, 6);
		excelReader.setCellData(xlfile, xlSheetName, 389, 7, actRMA);

		System.out.println("********* Voucher No        : " + actDocNo + "  value expected  " + expDocNo);
		System.out.println("********* Customer Account  : " + actCustomerAccount + "  value expected  " + expCustomerAccount);
		System.out.println("********* Department        : " + actDepartment + "  value expected  " + expDepartment);
		System.out.println("********* Warehouse         : " + actWarehouse + "  value expected  " + expWarehouse);
		System.out.println("********* Item              : " + actItem + "  value expected  " + expItem);
		System.out.println("********* Quantity          : " + actQuantity + "  value expected  " + expQuantity);
		System.out.println("********* Rate              : " + actRate + "  value expected  " + expRate);
		System.out.println("********* Gross             : " + actGross + "  value expected  " + expGross);
		System.out.println("********* BATCH Col         : " + actBatch + "  value expected  " + expBatch);

		System.out.println("********* Warehouse11         : " + actWarehouse1 + "  value expected  " + expWarehouse1);
		System.out.println("********* Item1              : " + actItem1 + "  value expected  " + expItem1);
		System.out.println("********* Quantity1          : " + actQuantity1 + "  value expected  " + expQuantity1);
		System.out.println("********* Rate1              : " + actRate1 + "  value expected  " + expRate1);
		System.out.println("********* Gross1             : " + actGross1 + "  value expected  " + expGross1);
		System.out.println("********* RMA Col            : " + actRMA + "  value expected  " + expRMA);

		if (actDocNo.equalsIgnoreCase(expDocNo) && actCustomerAccount.equalsIgnoreCase(expCustomerAccount)
				&& actDepartment.equalsIgnoreCase(expDepartment) && actWarehouse.equalsIgnoreCase(expWarehouse)
				&& actItem.equalsIgnoreCase(expItem) && actQuantity.equalsIgnoreCase(expQuantity)
				&& actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross)
				/*&& actBatch.equalsIgnoreCase(expBatch)*/


				&& actWarehouse1.equalsIgnoreCase(expWarehouse1)
				&& actItem1.equalsIgnoreCase(expItem1) && actQuantity1.equalsIgnoreCase(expQuantity1)
				&& actRate1.equalsIgnoreCase(expRate1) && actGross1.equalsIgnoreCase(expGross1)
				/*&& actRMA.equalsIgnoreCase(expRMA)*/)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 375, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 375, 8, resFail);
			return false;
		}
	}







	public static boolean checkChangingRatePreloadBehaviourToRateSalesInvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceSettingsBtn));
		salesInvoiceSettingsBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
		editScreenTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBehaviourDropdown));
		ScrollIntoView(editScreenBehaviourDropdown);
		Select behaviour = new Select(editScreenBehaviourDropdown);
		behaviour.selectByVisibleText(excelReader.getCellData(xlSheetName, 390, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		System.out.println("******************checkChangingRatePreloadBehaviourToRateSalesInvoiceVAT*****************");

		String expMessage = excelReader.getCellData(xlSheetName, 390, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 390, 7, actMessage);

		Thread.sleep(2000);

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 390, 8, resPass);
			return true;
		}
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 390, 8, resFail);
			return false;
		}
	}








	public static boolean checkSavingSalesInvoiceVATVoucher4WithRateAsLessThanCreditLimit() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
		financialTransactionSalesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
		salesInvoiceVATVoucher.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.click();
		customerAccountTxt.sendKeys(Keys.END);
		customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 391, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 392, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(Keys.END);
		jurisdictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 393, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
		salesInvoiceVATPlaceOFSupply.click();
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.END);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SHIFT,Keys.HOME);
		salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 394, 5));
		Thread.sleep(2000);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 395, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 396, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 397, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 398, 5));
		enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 399, 5));
		enter_Rate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 4");
		expMessage.add("Saving in background.");*/
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : 4");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();*/

		String expMessage = excelReader.getCellData(xlSheetName, 391, 6);
		excelReader.setCellData(xlfile, xlSheetName, 391, 7, actMessage);
				
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		System.out.println("************************************************************************");

		if(actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 391, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 391, 8, resFail);
			return false;
		}
	}







	public boolean checkSavedVoucher4AndTryingToUpdateVoucherRateGreaterThanCreditLimit() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher Loaded Successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		String actDocNo = documentNumberTxt.getAttribute("value");
		String actCustomerAccount = customerAccountTxt.getAttribute("value");
		String actDepartment = departmentTxt.getAttribute("value");


		String actWarehouse = select1stRow_1stColumn.getText();
		String actItem = select1stRow_2ndColumn.getText();
		String actQuantity = select1stRow_11thColumn.getText();
		String actRate = select1stRow_14thColumn.getText();
		String actGross = select1stRow_15thColumn.getText();
		String actBatch = select1stRow_19thColumn.getText();

		String expDocNo = excelReader.getCellData(xlSheetName, 400, 6);
		excelReader.setCellData(xlfile, xlSheetName, 400, 7, actDocNo);
		
		String expCustomerAccount = excelReader.getCellData(xlSheetName, 401, 6);
		excelReader.setCellData(xlfile, xlSheetName, 401, 7, actCustomerAccount);
		
		String expDepartment = excelReader.getCellData(xlSheetName, 402, 6);
		excelReader.setCellData(xlfile, xlSheetName, 402, 7, actDepartment);

		String expWarehouse = excelReader.getCellData(xlSheetName, 403, 6);
		excelReader.setCellData(xlfile, xlSheetName, 403, 7, actWarehouse);
		
		String expItem = excelReader.getCellData(xlSheetName, 404, 6);
		excelReader.setCellData(xlfile, xlSheetName, 404, 7, actItem);
		
		String expQuantity = excelReader.getCellData(xlSheetName, 405, 6);
		excelReader.setCellData(xlfile, xlSheetName, 405, 7, actQuantity);
		
		String expRate = excelReader.getCellData(xlSheetName, 406, 6);
		excelReader.setCellData(xlfile, xlSheetName, 406, 7, actDocNo);
		
		String expGross = excelReader.getCellData(xlSheetName, 407, 6);
		excelReader.setCellData(xlfile, xlSheetName, 407, 7, actDocNo);

		System.out.println("*************checkSavedVoucher4AndTryingToUpdateVoucherRateGreaterThanCreditLimit***********");

		System.out.println("********* Voucher No        : " + actDocNo + "  value expected  " + expDocNo);
		System.out.println("********* Customer Account  : " + actCustomerAccount + "  value expected  " + expCustomerAccount);
		System.out.println("********* Department        : " + actDepartment + "  value expected  " + expDepartment);
		System.out.println("********* Warehouse         : " + actWarehouse + "  value expected  " + expWarehouse);
		System.out.println("********* Item              : " + actItem + "  value expected  " + expItem);
		System.out.println("********* Quantity          : " + actQuantity + "  value expected  " + expQuantity);
		System.out.println("********* Rate              : " + actRate + "  value expected  " + expRate);
		System.out.println("********* Gross             : " + actGross + "  value expected  " + expGross);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(Keys.END);
		enter_Rate.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 400, 5));
		enter_Rate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 408, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 408, 7, actMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefcancel));
		billRefcancel.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();

		getWaitForAlert();
		getAlert().accept();

		Thread.sleep(3000);

		if (actDocNo.equalsIgnoreCase(expDocNo) && actCustomerAccount.equalsIgnoreCase(expCustomerAccount)
				&& actDepartment.equalsIgnoreCase(expDepartment) && actWarehouse.equalsIgnoreCase(expWarehouse)
				&& actItem.equalsIgnoreCase(expItem) && actQuantity.equalsIgnoreCase(expQuantity)
				&& actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross)

				&& actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 400, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 400, 8, resFail);
			return false;
		}
	}






	String FileName7 = null;
	public boolean checkPrintPdfOfSalesInvoiceVAT() throws Exception, EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
		financialTransactionSalesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
		salesInvoiceVATVoucher.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int count = grid_VoucherNoList.size();

		for (int i = 0; i < count; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 409, 5))) 
			{
				if (grid_CheckBoxList.get(i).isSelected()==false) 
				{
					grid_CheckBoxList.get(i).click();
					break;
				}
			}
		}      

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LayoutNameTxt));
		LayoutNameTxt.click();
		LayoutNameTxt.sendKeys(Keys.END);
		LayoutNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		LayoutNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 410, 5));
		Thread.sleep(2000);
		LayoutNameTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);


		Select s = new Select(LayoutOptionsdropdown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 411, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(footerPrintBtn));
		footerPrintBtn.click();

		Thread.sleep(5000);

		FileName7 = checkDownloadedFileName(getDriver());

		/*Robot robot = new Robot();
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

					Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\SavingPrintPDFInSalesInvoice.exe");

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
				 	getDriver().switchTo().window(openTabs.get(0));*/


		String actPDF = getBaseDir()+"\\autoIt\\ImportFiles\\SalesInvoiceVATPrint.pdf";
		String expPDF = getBaseDir()+"\\autoIt\\ExportFiles\\"+FileName7;

		PDFUtil pdfutil = new PDFUtil();

		boolean result = pdfutil.compare(actPDF, expPDF);

		String actData = pdfutil.getText(actPDF);
		String expData = pdfutil.getText(expPDF);
		
		excelReader.setCellData(xlfile, xlSheetName, 409, 6, expData);
		excelReader.setCellData(xlfile, xlSheetName, 409, 7, actData);

		System.err.println(actData);
		System.err.println(expData);

		System.out.println("Compared Result  : "+result);

		if (actData.equalsIgnoreCase(expData)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 409, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 409, 8, resFail);
			return false;
		}
	}






	String FileName8 = null;

	public boolean checkPrintPdfOfSalesInvoiceVATWithImportedXML() throws Exception, EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int count = grid_VoucherNoList.size();

		for (int i = 0; i < count; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 412, 5))) 
			{
				if (grid_CheckBoxList.get(i).isSelected()==false) 
				{
					grid_CheckBoxList.get(i).click();
					break;
				}
			}
		}   

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createLayoutBtn));
		createLayoutBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Opendropdown));
		Opendropdown.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fromXMLBtn));
		fromXMLBtn.click();

		Thread.sleep(2000);

		Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\ImportingSalesInvoiceVATPrintLayout.exe");

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Savedropdown));
		Savedropdown.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PrintSaveBttn));
		PrintSaveBttn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveTxtField));
		SaveTxtField.click();
		SaveTxtField.sendKeys(excelReader.getCellData(xlSheetName, 413, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Save_OkBtn));
		Save_OkBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 412, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 412, 7, actMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExitBtn));
		ExitBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		for (int i = 0; i < count; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 414, 5))) 
			{
				if (grid_CheckBoxList.get(i).isSelected()==false) 
				{
					grid_CheckBoxList.get(i).click();
					break;
				}
			}
		}   

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LayoutNameTxt));
		LayoutNameTxt.click();
		LayoutNameTxt.sendKeys(Keys.END);
		LayoutNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		LayoutNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 415, 5));
		Thread.sleep(1000);
		LayoutNameTxt.sendKeys(Keys.TAB);

		Select s = new Select(LayoutOptionsdropdown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 416, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(footerPrintBtn));
		footerPrintBtn.click();

		Thread.sleep(5000);

		FileName8 = checkDownloadedFileName(getDriver());

		String actPDF = getBaseDir()+"\\autoIt\\ExportFiles\\"+FileName8;
		String expPDF = getBaseDir()+"\\autoIt\\ImportFiles\\SalesInvoiceVATPrint.pdf";

		PDFUtil pdfutil = new PDFUtil();

		boolean result = pdfutil.compare(actPDF, expPDF);

		String actData = pdfutil.getText(actPDF);
		String expData = pdfutil.getText(expPDF);
		excelReader.setCellData(xlfile, xlSheetName, 413, 6, expData);
		excelReader.setCellData(xlfile, xlSheetName, 413, 7, actData);

		System.err.println(actData);
		System.err.println(expData);

		System.out.println("Compared Result  : "+result);

		if (actMessage.equalsIgnoreCase(expMessage) && actData.equalsIgnoreCase(expData)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 412, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 412, 8, resFail);
			return false;
		}
	}










	// Mail Starts From Here 


	String FileName9 = null;

	public boolean checkEmailOptionInSalesInvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int vcount = grid_VoucherNoList.size();

		for (int i = 0; i < vcount; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 417, 5))) 
			{
				if (grid_CheckBoxList.get(i).isSelected()==false) 
				{
					grid_CheckBoxList.get(i).click();
					break;
				}
			}
		}   

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printBtn));
		printBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LayoutNameTxt));
		LayoutNameTxt.click();
		LayoutNameTxt.sendKeys(Keys.END);
		LayoutNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		LayoutNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 418, 5));
		Thread.sleep(1000);
		LayoutNameTxt.sendKeys(Keys.TAB);

		Select s = new Select(LayoutOptionsdropdown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 419, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printScreenPickEmailChkbox));
		if (printScreenPickEmailChkboxIsSelected.isSelected()==true) 
		{
			printScreenPickEmailChkbox.click();
		} 

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printScreenSentEmailTxt));
		printScreenSentEmailTxt.click();
		printScreenSentEmailTxt.sendKeys(excelReader.getCellData(xlSheetName, 420, 5));
		printScreenSentEmailTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(printSCreenSubjectTxt));
		printSCreenSubjectTxt.click();
		printSCreenSubjectTxt.sendKeys(excelReader.getCellData(xlSheetName, 421, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSendBtn));
		emailSendBtn.click();

		String ExpMessage=excelReader.getCellData(xlSheetName, 417, 6);
		String actMessage=checkValidationMessage(ExpMessage);
		excelReader.setCellData(xlfile, xlSheetName, 417, 7, actMessage);

		Thread.sleep(2000);

		Robot robot = new Robot();   
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_T);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_T);

		Thread.sleep(2000);

		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

		System.out.println("openTabs"+openTabs);

		getDriver().switchTo().window(openTabs.get(0));

		getDriver().switchTo().window(openTabs.get(1));


		getDriver().get("https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
		userNameTxt.click();

		userNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 422, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
		NextBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
		PasswordTxt.click();

		PasswordTxt.sendKeys(excelReader.getCellData(xlSheetName, 423, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
		NextBtn.click();

		Thread.sleep(8000);

		int count = row1mailFromList.size();

		for (int i = 0; i < count; i++) 
		{
			String mailFromTxt = row1mailFromList.get(i).getText();

			if (mailFromTxt.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 424, 5))) 
			{
				row1mailFromList.get(i).click();
				break;
			}
		}

		Thread.sleep(3000);

		getAction().moveToElement(mailAttachmentDownloadBtn).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn));
		mailAttachmentDownloadBtn.click();

		Thread.sleep(5000);

		FileName9 = checkDownloadedFileName(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
		mailDeleteButton.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gmailUserBtn));
		gmailUserBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gmailSignOutBtn));
		gmailSignOutBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		removeAccountBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		removeDeleteBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		yesRemoveBtn.click();

		Thread.sleep(1000);

		getDriver().switchTo().window(openTabs.get(0));

		Thread.sleep(1000);

		getDriver().switchTo().window(openTabs.get(1)).close();

		Thread.sleep(1000);

		getDriver().switchTo().window(openTabs.get(0));

		String actPDF = getBaseDir()+"\\autoIt\\ImportFiles\\SalesInvoiceVATPrint.pdf";
		String expPDF = getBaseDir()+"\\autoIt\\ExportFiles\\"+FileName9;

		PDFUtil pdfutil = new PDFUtil();

		String actData = pdfutil.getText(actPDF);
		String expData = pdfutil.getText(expPDF);
		excelReader.setCellData(xlfile, xlSheetName, 418, 6, expData);
		excelReader.setCellData(xlfile, xlSheetName, 418, 7, actData);

		System.err.println(actData);
		System.err.println(expData);

		boolean result = pdfutil.compare(actPDF, expPDF);

		System.out.println("Compared Result  : "+result);

		if (actMessage.equalsIgnoreCase(ExpMessage) && actData.equalsIgnoreCase(expData)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 417, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 417, 8, resFail);
			return false;
		}
	}









	public boolean checkEraseAllTransactionsAfterCompletingAllValidations() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataMangementMenu));
		dataMangementMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseAll));
		eraseAll.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseTranscationsRadio));
		eraseTranscationsRadio.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseAllOkBtn));
		eraseAllOkBtn.click();

		if(getIsAlertPresent())
		{
			getWaitForAlert();

			getAlert().accept();
		}

		String expMessage = excelReader.getCellData(xlSheetName, 509, 6);

		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 509, 7, actMessage);

		System.out.println("************************************* checkEraseAllTransactionsAfterCompletingAllValidations  *********************************");
		System.out.println("Message  :  "+actMessage +" Value Expected : "+expMessage);

		if (actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 509, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 509, 8, resFail);
			return false;
		}
	}









	public static boolean checkSavingPurchaseVoucherVATVoucherForDebitNoteSellerPricebookValidations() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
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

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();
		vendorAccountTxt.sendKeys("Vendor B");	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys("INDIA");	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys("Abu Dhabi");
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys("DUBAI");
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();
		enter_WarehouseTxt.sendKeys("HYDERABAD");
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys("BR COGS ITEM");
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("100");


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys("25");


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
		select1stRow_12thColumn.click();
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
		enter_Batch.sendKeys("BatchHyderabad");
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 7); 

		String next7DaysDate=df.format(cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 7Days>>>>>>>> : "+next7DaysDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next7DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);


		// Row 2 with Item Group item

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(Keys.END);
		enter_WarehouseTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_WarehouseTxt.sendKeys("SECUNDERABAD");
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys("BR COGS ITEM");
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
		select2ndRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("100");


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_11thColumn));
		select2ndRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys("25");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_12thColumn));
		select2ndRow_12thColumn.click();
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
		enter_Batch.sendKeys("BatchSecunderabad");
		enter_Batch.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next7DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		Thread.sleep(2000);
		

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		Thread.sleep(2000);

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT1");
		expMessage.add("Saving in background.");*/

		String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : SU/IND/TEXT1");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		System.out.println("************************************************************************");

		if(actMessage.equals(expMessage))
		{
			return true;
		} 
		else 
		{
			return false;
		}
	}





	


	public boolean checkSavingDeliveryNoteVoucher1WithNoTagsAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsSalesMenu));
		inventoryTransactionsSalesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deliveryNotesBtn));
		deliveryNotesBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 467, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 468, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 469, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 470, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 471, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 472, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt0Txt));
		enter_salrt0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt1Txt));
		enter_salrt1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp0Txt));
		enter_salrtGrp0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp1Txt));
		enter_salrtGrp1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		batchPickOnFIFOIcon.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
		batchOkIcon.click();

		Thread.sleep(2000);

		ArrayList<String> row1 = new ArrayList<String>();

		int count = voucherBodyRow1List.size();

		for (int j = 0; j < count; j++) 
		{
			String data = voucherBodyRow1List.get(j).getText();
			row1.add(data);
		}

		String actRow1List = row1.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 467, 6);
		excelReader.setCellData(xlfile, xlSheetName, 467, 7, actRow1List);

		Thread.sleep(1000);

		System.out.println(actRow1List);
		System.out.println(expRow1List);

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
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : 1");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 470, 6);
		excelReader.setCellData(xlfile, xlSheetName, 470, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actRow1List.equalsIgnoreCase(expRow1List) && actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 467, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 467, 8, resFail);
			return false;
		}
	}







	public boolean checkSavingDeliveryNoteVoucher2WithPricebookAccountOneAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 473, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 474, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 475, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 476, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 477, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 478, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt0Txt));
		enter_salrt0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt1Txt));
		enter_salrt1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp0Txt));
		enter_salrtGrp0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp1Txt));
		enter_salrtGrp1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		batchPickOnFIFOIcon.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
		batchOkIcon.click();

		Thread.sleep(2000);

		ArrayList<String> row1 = new ArrayList<String>();

		int count = voucherBodyRow1List.size();

		for (int j = 0; j < count; j++) 
		{
			String data = voucherBodyRow1List.get(j).getText();
			row1.add(data);
		}

		String actRow1List = row1.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 473, 6);
		excelReader.setCellData(xlfile, xlSheetName, 473, 7, actRow1List);

		Thread.sleep(1000);

		System.out.println(actRow1List);
		System.out.println(expRow1List);

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
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : 2");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 476, 6);
		excelReader.setCellData(xlfile, xlSheetName, 476, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actRow1List.equalsIgnoreCase(expRow1List) && actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 473, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 473, 8, resFail);
			return false;
		}
	}



	



	public boolean checkSavingDeliveryNoteVoucher3WithPricebookAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 479, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 480, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 481, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 482, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 483, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 484, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt0Txt));
		enter_salrt0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt1Txt));
		enter_salrt1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp0Txt));
		enter_salrtGrp0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp1Txt));
		enter_salrtGrp1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		batchPickOnFIFOIcon.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
		batchOkIcon.click();

		Thread.sleep(2000);

		ArrayList<String> row1 = new ArrayList<String>();

		int count = voucherBodyRow1List.size();

		for (int j = 0; j < count; j++) 
		{
			String data = voucherBodyRow1List.get(j).getText();
			row1.add(data);
		}

		String actRow1List = row1.toString();
		String expRow1List = "excelReader.getCellData(xlSheetName, 479, 6)";
		excelReader.setCellData(xlfile, xlSheetName, 479, 7, actRow1List);

		Thread.sleep(1000);

		System.out.println(actRow1List);
		System.out.println(expRow1List);

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
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : 3");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 482, 6);
		excelReader.setCellData(xlfile, xlSheetName, 482, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actRow1List.equalsIgnoreCase(expRow1List) && actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 479, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 479, 8, resFail);
			return false;
		}
	}







	public boolean checkSavingDeliveryNoteVoucher4WithGroupOneAccountOneAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 485, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 486, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 487, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 488, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 489, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 490, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt0Txt));
		enter_salrt0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt1Txt));
		enter_salrt1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp0Txt));
		enter_salrtGrp0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp1Txt));
		enter_salrtGrp1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		batchPickOnFIFOIcon.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
		batchOkIcon.click();

		Thread.sleep(2000);

		ArrayList<String> row1 = new ArrayList<String>();

		int count = voucherBodyRow1List.size();

		for (int j = 0; j < count; j++) 
		{
			String data = voucherBodyRow1List.get(j).getText();
			row1.add(data);
		}

		String actRow1List = row1.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 485, 6);
		excelReader.setCellData(xlfile, xlSheetName, 485, 7, actRow1List);
		
		Thread.sleep(1000);

		System.out.println(actRow1List);
		System.out.println(expRow1List);

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
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : 4");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 488, 6);
		excelReader.setCellData(xlfile, xlSheetName, 488, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actRow1List.equalsIgnoreCase(expRow1List) && actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 485, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 485, 8, resFail);
			return false;
		}
	}








	public boolean checkSavingDeliveryNoteVoucher5WithGroupOneAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 491, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 492, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 493, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 494, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 495, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 496, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt0Txt));
		enter_salrt0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt1Txt));
		enter_salrt1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp0Txt));
		enter_salrtGrp0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp1Txt));
		enter_salrtGrp1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		batchPickOnFIFOIcon.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
		batchOkIcon.click();

		Thread.sleep(2000);

		ArrayList<String> row1 = new ArrayList<String>();

		int count = voucherBodyRow1List.size();

		for (int j = 0; j < count; j++) 
		{
			String data = voucherBodyRow1List.get(j).getText();
			row1.add(data);
		}

		String actRow1List = row1.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 491, 6);
		excelReader.setCellData(xlfile, xlSheetName, 491, 7, actRow1List);

		Thread.sleep(1000);

		System.out.println(actRow1List);
		System.out.println(expRow1List);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMsg = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMsg.add(data);
		}*/
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMessage = new HashSet();

		expMessage.add("Saving in background.");
		expMessage.add("Voucher saved successfully : 5");*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 494, 6);
		excelReader.setCellData(xlfile, xlSheetName, 494, 7, actMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actRow1List.equalsIgnoreCase(expRow1List) && actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 491, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 491, 8, resFail);
			return false;
		}
	}






	public boolean checkSavingDeliveryNoteVoucher6WithGroupTwoAccountOneAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 497, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 498, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 499, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 500, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 501, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 502, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt0Txt));
		enter_salrt0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt1Txt));
		enter_salrt1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp0Txt));
		enter_salrtGrp0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp1Txt));
		enter_salrtGrp1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		batchPickOnFIFOIcon.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
		batchOkIcon.click();

		Thread.sleep(2000);

		ArrayList<String> row1 = new ArrayList<String>();

		int count = voucherBodyRow1List.size();

		for (int j = 0; j < count; j++) 
		{
			String data = voucherBodyRow1List.get(j).getText();
			row1.add(data);
		}

		String actRow1List = row1.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 497, 6);
		excelReader.setCellData(xlfile, xlSheetName, 497, 7, actRow1List);

		Thread.sleep(1000);

		System.out.println(actRow1List);
		System.out.println(expRow1List);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 6");
		expMessage.add("Saving in background.");*/
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : 6");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 500, 6);
		excelReader.setCellData(xlfile, xlSheetName, 500, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actRow1List.equalsIgnoreCase(expRow1List) && actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 497, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 497, 8, resFail);
			return false;
		}
	}








	public boolean checkSavingDeliveryNoteVoucher7WithGroupTwoAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 503, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 504, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 505, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 506, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 507, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 508, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt0Txt));
		enter_salrt0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt1Txt));
		enter_salrt1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp0Txt));
		enter_salrtGrp0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp1Txt));
		enter_salrtGrp1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		batchPickOnFIFOIcon.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
		batchOkIcon.click();

		Thread.sleep(2000);

		ArrayList<String> row1 = new ArrayList<String>();

		int count = voucherBodyRow1List.size();

		for (int j = 0; j < count; j++) 
		{
			String data = voucherBodyRow1List.get(j).getText();
			row1.add(data);
		}

		String actRow1List = row1.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 503, 6);
		excelReader.setCellData(xlfile, xlSheetName, 503, 7, actRow1List);

		Thread.sleep(1000);

		System.out.println(actRow1List);
		System.out.println(expRow1List);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 7");
		expMessage.add("Saving in background.");*/
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Saving in background.");
		expMsg.add("Voucher saved successfully : 7");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 506, 6);
		excelReader.setCellData(xlfile, xlSheetName, 506, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actRow1List.equalsIgnoreCase(expRow1List) && actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 503, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 503, 8, resFail);
			return false;
		}
	}


	





	/*public boolean checkSavingDeliveryNoteVoucher7WithGroupTwoAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
					inventoryMenu.click();

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
					inventoryTransactionsMenu.click();

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsSalesMenu));
					inventoryTransactionsSalesMenu.click();

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deliveryNotesBtn));
					deliveryNotesBtn.click();

					Thread.sleep(2000);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
					newBtn.click();

					checkValidationMessage("Screen opened");

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
					customerAccountTxt.sendKeys("Group Two Account Two");
					Thread.sleep(2000);
					customerAccountTxt.sendKeys(Keys.TAB);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
					warehouseTxt.sendKeys("HYDERABAD");
					Thread.sleep(2000);
					warehouseTxt.sendKeys(Keys.TAB);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
					departmentTxt.sendKeys("DUBAI");
					Thread.sleep(2000);
					departmentTxt.sendKeys(Keys.TAB);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
					select1stRow_1stColumn.click();
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
					enter_ItemTxt.sendKeys("BR COGS ITEM");

					int itemsCount=itemComboList.size();

					for (int i = 0; i < itemsCount; i++) 
					{
						String item=itemComboList.get(i).getText();

						if (item.equalsIgnoreCase("BR COGS ITEM")) 
						{
							itemComboList.get(i).click();
							enter_ItemTxt.sendKeys(Keys.TAB);
							break;
						}
					}

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
					enter_UnitTxt.sendKeys(Keys.TAB);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
					enter_Quantity.sendKeys("1");
					enter_Quantity.sendKeys(Keys.TAB);

					Thread.sleep(3000);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
					enter_Rate.sendKeys("5");
					enter_Rate.sendKeys(Keys.TAB);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
					enter_Gross.sendKeys(Keys.TAB);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt0Txt));
					enter_salrt0Txt.sendKeys(Keys.TAB);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt1Txt));
					enter_salrt1Txt.sendKeys(Keys.TAB);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp0Txt));
					enter_salrtGrp0Txt.sendKeys(Keys.TAB);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp1Txt));
					enter_salrtGrp1Txt.sendKeys(Keys.TAB);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
					batchPickOnFIFOIcon.click();

					Thread.sleep(2000);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
			    	batchOkIcon.click();

					Thread.sleep(2000);

					ArrayList<String> row1 = new ArrayList<String>();

					int count = voucherBodyRow1List.size();

					for (int j = 0; j < count; j++) 
					{
						String data = voucherBodyRow1List.get(j).getText();
						row1.add(data);
					}

					String actRow1List = row1.toString();
					String expRow1List = "[1, BR COGS ITEM, Pcs, 1.00, , 5.00, 5.00, 10.00, 10.10, 7000.00, 7000.10, BATCHHYDERABAD, , ]";

					Thread.sleep(1000);

					System.out.println(actRow1List);
					System.out.println(expRow1List);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
					voucherSaveBtn.click();

					HashSet<String> actMessage = new HashSet();

					for (int i = 0; i < 2; i++) 
					{
						String data = checkValidationMessage("");
						actMessage.add(data);
					}

					HashSet<String> expMessage = new HashSet();

					expMessage.add("Voucher saved successfully : 7");
					expMessage.add("Saving in background.");

					System.out.println("Actual Message    : "+actMessage);
					System.out.println("Expected Message  : "+expMessage);

					if(actRow1List.equalsIgnoreCase(expRow1List) && actMessage.equals(expMessage))
					{
						return true;
					}
					else
					{
						return false;
					}
				}
	 */






	

	public boolean checkSavingMaterialReceiptNoteVoucher1WithNoTagsAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		inventoryTransactionsPurchasesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(materialReceiptNotesVoucher));
		materialReceiptNotesVoucher.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 425, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 426, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 427, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		select1stRow_4thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 428, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 429, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrt0Txt));
		enter_purrt0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrt1Txt));
		enter_purrt1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrtGrp0Txt));
		enter_purrtGrp0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrtGrp1Txt));
		enter_purrtGrp1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 430, 5));
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 5); 

		String next5DaysDate=df.format(cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 5 Days>>>>>>>> : "+next5DaysDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next5DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		ArrayList<String> row1 = new ArrayList<String>();

		int count = voucherBodyRow1List.size();

		for (int j = 0; j < count; j++) 
		{
			String data = voucherBodyRow1List.get(j).getText();
			row1.add(data);
		}

		String actRow1List = row1.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 425, 6)+next5DaysDate+", ]";
		excelReader.setCellData(xlfile, xlSheetName, 425, 7, actRow1List);

		Thread.sleep(1000);

		System.out.println(actRow1List);
		System.out.println(expRow1List);

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

		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : 1");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 428, 6);
		excelReader.setCellData(xlfile, xlSheetName, 428, 7, actMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actRow1List.equalsIgnoreCase(expRow1List) && actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 425, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 425, 8, resFail);
			return false;
		}
	}







	public boolean checkSavingMaterialReceiptNoteVoucher2WithPricebookVendorAccountOneAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 431, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 432, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 433, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		select1stRow_4thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 434, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 435, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrt0Txt));
		enter_purrt0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrt1Txt));
		enter_purrt1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrtGrp0Txt));
		enter_purrtGrp0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrtGrp1Txt));
		enter_purrtGrp1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 436, 5));
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 5); 

		String next5DaysDate=df.format(cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 5 Days>>>>>>>> : "+next5DaysDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next5DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		ArrayList<String> row1 = new ArrayList<String>();

		int count = voucherBodyRow1List.size();

		for (int j = 0; j < count; j++) 
		{
			String data = voucherBodyRow1List.get(j).getText();
			row1.add(data);
		}

		String actRow1List = row1.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 431, 6)+next5DaysDate+", ]";
		excelReader.setCellData(xlfile, xlSheetName, 431, 7, actRow1List);

		Thread.sleep(1000);

		System.out.println(actRow1List);
		System.out.println(expRow1List);

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
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : 2");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 434, 6);
		excelReader.setCellData(xlfile, xlSheetName, 434, 7, actMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actRow1List.equalsIgnoreCase(expRow1List) && actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 431, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 431, 8, resFail);
			return false;
		}
	}







	public boolean checkSavingMaterialReceiptsNoteVoucher3WithPricebookVendorAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 437, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 438, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 439, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		select1stRow_4thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 440, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 441, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrt0Txt));
		enter_purrt0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrt1Txt));
		enter_purrt1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrtGrp0Txt));
		enter_purrtGrp0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrtGrp1Txt));
		enter_purrtGrp1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 442, 5));
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 5); 

		String next5DaysDate=df.format(cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 5 Days>>>>>>>> : "+next5DaysDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next5DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		ArrayList<String> row1 = new ArrayList<String>();

		int count = voucherBodyRow1List.size();

		for (int j = 0; j < count; j++) 
		{
			String data = voucherBodyRow1List.get(j).getText();
			row1.add(data);
		}

		String actRow1List = row1.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 437, 6)+next5DaysDate+", ]";
		excelReader.setCellData(xlfile, xlSheetName, 437, 7, actRow1List);

		Thread.sleep(1000);

		System.out.println(actRow1List);
		System.out.println(expRow1List);

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
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : 3");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 440, 6);
		excelReader.setCellData(xlfile, xlSheetName, 440, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actRow1List.equalsIgnoreCase(expRow1List) && actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 437, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 437, 8, resFail);
			return false;
		}
	}




	



	public boolean checkSavingMaterialReceiptNoteVoucher4WithGroupOneAccountOneAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 443, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 444, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 445, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		select1stRow_4thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 446, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 447, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrt0Txt));
		enter_purrt0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrt1Txt));
		enter_purrt1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrtGrp0Txt));
		enter_purrtGrp0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrtGrp1Txt));
		enter_purrtGrp1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 448, 5));
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 5); 

		String next5DaysDate=df.format(cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 5 Days>>>>>>>> : "+next5DaysDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next5DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		ArrayList<String> row1 = new ArrayList<String>();

		int count = voucherBodyRow1List.size();

		for (int j = 0; j < count; j++) 
		{
			String data = voucherBodyRow1List.get(j).getText();
			row1.add(data);
		}

		String actRow1List = row1.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 443, 6)+next5DaysDate+", ]";
		excelReader.setCellData(xlfile, xlSheetName, 443, 7, actRow1List);

		Thread.sleep(1000);

		System.out.println(actRow1List);
		System.out.println(expRow1List);

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
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : 4");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 446, 6);
		excelReader.setCellData(xlfile, xlSheetName, 446, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actRow1List.equalsIgnoreCase(expRow1List) && actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 443, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 443, 8, resFail);
			return false;
		}
	}







	public boolean checkSavingMaterialReceiptNoteVoucher5WithGroupOneAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 449, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 550, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 551, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		select1stRow_4thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 552, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 553, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrt0Txt));
		enter_purrt0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrt1Txt));
		enter_purrt1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrtGrp0Txt));
		enter_purrtGrp0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrtGrp1Txt));
		enter_purrtGrp1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 554, 5));
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 5); 

		String next5DaysDate=df.format(cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 5 Days>>>>>>>> : "+next5DaysDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next5DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		ArrayList<String> row1 = new ArrayList<String>();

		int count = voucherBodyRow1List.size();

		for (int j = 0; j < count; j++) 
		{
			String data = voucherBodyRow1List.get(j).getText();
			row1.add(data);
		}

		String actRow1List = row1.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 449, 6)+next5DaysDate+", ]";
		excelReader.setCellData(xlfile, xlSheetName, 449, 7, actRow1List);

		Thread.sleep(1000);

		System.out.println(actRow1List);
		System.out.println(expRow1List);

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
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : 5");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 452, 6);
		excelReader.setCellData(xlfile, xlSheetName, 452, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actRow1List.equalsIgnoreCase(expRow1List) && actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 449, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 449, 8, resFail);
			return false;
		}
	}








	public boolean checkSavingMaterialReceiptNoteVoucher6WithGroupTwoAccountOneAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 455, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 456, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 457, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		select1stRow_4thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 458, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 459, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrt0Txt));
		enter_purrt0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrt1Txt));
		enter_purrt1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrtGrp0Txt));
		enter_purrtGrp0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrtGrp1Txt));
		enter_purrtGrp1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 460, 5));
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 5); 

		String next5DaysDate=df.format(cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 5 Days>>>>>>>> : "+next5DaysDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next5DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		ArrayList<String> row1 = new ArrayList<String>();

		int count = voucherBodyRow1List.size();

		for (int j = 0; j < count; j++) 
		{
			String data = voucherBodyRow1List.get(j).getText();
			row1.add(data);
		}

		String actRow1List = row1.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 455, 6)+next5DaysDate+", ]";
		excelReader.setCellData(xlfile, xlSheetName, 455, 7, actRow1List);

		Thread.sleep(1000);

		System.out.println(actRow1List);
		System.out.println(expRow1List);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 6");
		expMessage.add("Saving in background.");*/
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Voucher saved successfully : 6");
		expMsg.add("Saving in background.");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 458, 6);
		excelReader.setCellData(xlfile, xlSheetName, 458, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actRow1List.equalsIgnoreCase(expRow1List) && actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 455, 5, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 455, 5, resFail);
			return false;
		}
	}







	public boolean checkSavingMaterialReceiptNoteVoucher7WithGroupTwoAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 461, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 462, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 463, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		select1stRow_4thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 464, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 465, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrt0Txt));
		enter_purrt0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrt1Txt));
		enter_purrt1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrtGrp0Txt));
		enter_purrtGrp0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_purrtGrp1Txt));
		enter_purrtGrp1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 466, 5));
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 5); 

		String next5DaysDate=df.format(cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 5 Days>>>>>>>> : "+next5DaysDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next5DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		ArrayList<String> row1 = new ArrayList<String>();

		int count = voucherBodyRow1List.size();

		for (int j = 0; j < count; j++) 
		{
			String data = voucherBodyRow1List.get(j).getText();
			row1.add(data);
		}

		String actRow1List = row1.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 461, 6)+next5DaysDate+", ]";
		excelReader.setCellData(xlfile, xlSheetName, 461, 7, actRow1List);

		Thread.sleep(1000);

		System.out.println(actRow1List);
		System.out.println(expRow1List);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 7");
		expMessage.add("Saving in background.");*/
		
		String actMessage = SavingInBackground(2);

		/*HashSet<String> expMsg = new HashSet();

		expMsg.add("Saving in background.");
		expMsg.add("Voucher saved successfully : 7");
		
		String expMessage = expMsg.toString();*/
		
		String expMessage = excelReader.getCellData(xlSheetName, 464, 6);
		excelReader.setCellData(xlfile, xlSheetName, 464, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actRow1List.equalsIgnoreCase(expRow1List) && actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 461, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 461, 8, resFail);
			return false;
		}
	}

	
	@FindBy(xpath="//*[@id='id_transentry_body_menu']/a[4]/label")
	public static WebElement duplicateRow;
	
	@FindBy(xpath="//*[@id='id_transactionentry_close']/a/span")
	public static WebElement pvVAT_CloseBtn;
	

public boolean checkSavingPurchaseVoucherVATVoucher1WithRMA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	getDriver().navigate().refresh();
	
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

	checkValidationMessage("");

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
	vendorAccountTxt.sendKeys("Vendor B");	
	Thread.sleep(2000);
	vendorAccountTxt.sendKeys(Keys.TAB);

	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	departmentTxt.sendKeys("INDIA");	
	Thread.sleep(2000);
	departmentTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
	placeOFSupplyTxt.click();
	placeOFSupplyTxt.sendKeys("Abu Dhabi");
	Thread.sleep(2000);
	placeOFSupplyTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
	jurisdictionTxt.click();
	jurisdictionTxt.sendKeys("DUBAI");
	Thread.sleep(2000);
	jurisdictionTxt.sendKeys(Keys.TAB);


	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
	enter_WarehouseTxt.sendKeys("HYDERABAD");
	Thread.sleep(2000);
	enter_WarehouseTxt.sendKeys(Keys.TAB);


	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.sendKeys("RMA FITEM");
	Thread.sleep(2000);
	enter_ItemTxt.sendKeys(Keys.TAB);	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
	enter_TaxCode.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
	enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	select1stRow_11thColumn.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	enter_Quantity.click();
	enter_Quantity.clear();
	enter_Quantity.sendKeys("1");


	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_13thColumn));
	select1stRow_13thColumn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	enter_Rate.click();
	enter_Rate.clear();
	enter_Rate.sendKeys("100");

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
	select1stRow_14thColumn.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
	enter_Gross.click();
	enter_Gross.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_16thColumn));
	select1stRow_16thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
	enter_PvVat.click();

	enter_PvVat.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
	enter_PvTaxable.click();
	enter_PvTaxable.sendKeys(Keys.TAB);

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaClearBtn));
	rmaClearBtn.click();
	Thread.sleep(1000);
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
	rmaSerialNumberTxtField.sendKeys("A");

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
	rmaAddBtn.click();

	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
	RMAPopupOkBtn.click();

	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_0stColumn));
	getAction().contextClick(select1stRow_0stColumn).build().perform();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(duplicateRow));
	getAction().moveToElement(duplicateRow);
	duplicateRow.click();
	
	//validation of two rows
	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	String actR1Warehouse = select1stRow_1stColumn.getText();
	String expR1Warehouse = "HYDERABAD";

	String actR1Item = select1stRow_2ndColumn.getText();
	String expR1Item = "RMA FITEM ";
	
	String actR1TaxCode = select1stRow_3rdColumn.getText();
	String expR1TaxCode = "Standard Rated Purchase - Recoverable";
	
	String actR1PurchaseAccount = select1stRow_4thColumn.getText();
	String expR1PurchaseAccount ="WA COGS ACC INV";
	
	String actR1Units = select1stRow_5thColumn.getText();
	String expR1Units ="Pcs";

	String actR1Quantity = select1stRow_11thColumn.getText();
	String expR1Quantity = "1.00";

	String actR1Rate = select1stRow_13thColumn.getText();
	String expR1Rate ="100.00";

	String actR1Gross = select1stRow_14thColumn.getText();
	String expR1Gross = "100.00";
	
	String actR1RMA = select1stRow_21stColumn.getText();
	String expR1RMA = "";
	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
	String actR2Warehouse = select2ndRow_1stColumn.getText();
	String expR2Warehouse = "HYDERABAD";

	String actR2Item = select2ndRow_2ndColumn.getText();
	String expR2Item = "RMA FITEM";
	
	String actR2TaxCode = select2ndRow_3rdColumn.getText();
	String expR2TaxCode = "Standard Rated Purchase - Recoverable";
	
	String actR2PurchaseAccount = select2ndRow_4thColumn.getText();
	String expR2PurchaseAccount ="WA COGS ACC INV";
	
	String actR2Units = select2ndRow_5thColumn.getText();
	String expR2Units ="Pcs";

	String actR2Quantity = select2ndRow_11thColumn.getText();
	String expR2Quantity = "1.00";

	String actR2Rate = select2ndRow_13thColumn.getText();
	String expR2Rate ="100.00";

	String actR2Gross = select2ndRow_14thColumn.getText();
	String expR2Gross = "100.00";
	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_17thColumn));
	select2ndRow_17thColumn.click();
	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
	enter_PvTaxable.click();
	enter_PvTaxable.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaClearBtn));
	rmaClearBtn.click();
	
	Thread.sleep(1000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
	rmaSerialNumberTxtField.sendKeys("A");

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
	rmaAddBtn.click();
	
	if(getIsAlertPresent())
	{
		getWaitForAlert();
		Thread.sleep(1000);
		String alertTxt1=getAlert().getText();
		System.err.println("Alert Text with same series of RMA    "+alertTxt1);
	}
	
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
	RMAPopupOkBtn.click();
	Thread.sleep(1000);
	
	if(getIsAlertPresent())
	{
		getWaitForAlert();
		Thread.sleep(1000);
		String alertTxt1=getAlert().getText();
		System.err.println("Alert Text with same series of RMA after clicking OK    "+alertTxt1);
	}

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvVAT_CloseBtn));
	pvVAT_CloseBtn.click();
	
	getWaitForAlert();
	Thread.sleep(1000);
	getAlert().accept();
	Thread.sleep(1000);
	
	
	
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
	//System.out.println("RMA             : "+actR2RMA				+"  Value Expected  "+expR2RMA);
	

	Thread.sleep(2000);
	
	
	if(actR1Warehouse.equalsIgnoreCase(actR2Warehouse) && actR1Item.equalsIgnoreCase(actR2Item)
		&&actR1TaxCode.equalsIgnoreCase(actR2TaxCode) && actR1PurchaseAccount.equalsIgnoreCase(actR2PurchaseAccount)
		&&actR1Units.equalsIgnoreCase(actR2Units)  && actR1Quantity.equalsIgnoreCase(actR2Quantity)
		&& actR1Rate.equalsIgnoreCase(actR2Rate) && actR1Gross.equalsIgnoreCase(actR2Gross)
		)
	{
		return true;
	}

	//
	else
	{
	return false;
	}
	

}



	public BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}		


}
