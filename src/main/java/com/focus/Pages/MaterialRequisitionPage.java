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

public class MaterialRequisitionPage extends BaseEngine 
{
	private static String xlfile = getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlSheetName = "MaterialRequisitionPage";
	
	public static boolean checkSavingMaterialRequisitionVoucher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(materialRequisitionVoucher));
		materialRequisitionVoucher.click();

		Thread.sleep(2000);

		checkDeleteLinkStatus();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		MRpurchaseAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 8, 5));
		Thread.sleep(2000);
		MRpurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		enter_Rate.sendKeys(Keys.TAB);			

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

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

		expMessage.add("Saving in background.");
		expMessage.add("Voucher saved successfully : 1");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 8, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 9, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actMessage.equals(expMessage))		*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 8, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 9, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

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








	public boolean checkVerifyingAndDeletingMaterialRequisitionVoucher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocumentNumber = documentNumberTxt.getAttribute("value");
		String expDocumentNumber = excelReader.getCellData(xlSheetName, 13, 6);
		excelReader.setCellData(xlfile, xlSheetName, 13, 7, actDocumentNumber);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		String actPurchaseAccount = MRpurchaseAccountTxt.getAttribute("value");
		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 14, 6);
		excelReader.setCellData(xlfile, xlSheetName, 14, 7, actPurchaseAccount);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 15, 6);
		excelReader.setCellData(xlfile, xlSheetName, 15, 7, actDepartment);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));

		String actR1C1Item =select1stRow_1stColumn.getText();
		String actR1C2Units=select1stRow_2ndColumn.getText();
		String actR1C5Qty  =select1stRow_3rdColumn.getText();
		String actR1C7Rate =select1stRow_4thColumn.getText();
		String actR1C8Gross=select1stRow_5thColumn.getText();

		String expR1C1Item =excelReader.getCellData(xlSheetName, 16, 6);
		excelReader.setCellData(xlfile, xlSheetName, 16, 7, actR1C1Item);
		
		String expR1C2Units=excelReader.getCellData(xlSheetName, 17, 6);
		excelReader.setCellData(xlfile, xlSheetName, 17, 7, actR1C2Units);
		
		String expR1C5Qty  =excelReader.getCellData(xlSheetName, 18, 6);
		excelReader.setCellData(xlfile, xlSheetName, 18, 7, actR1C5Qty);
		
		String expR1C7Rate =excelReader.getCellData(xlSheetName, 19, 6);
		excelReader.setCellData(xlfile, xlSheetName, 19, 7, actR1C7Rate);
		
		String expR1C8Gross=excelReader.getCellData(xlSheetName, 20, 6);
		excelReader.setCellData(xlfile, xlSheetName, 20, 7, actR1C8Gross);

		System.out.println("***************************************checkVerifyingAndDeletingMaterialRequisitionVoucher1****************************************");
		System.out.println("Document Number                                 : "+actDocumentNumber   +" Value Expected : "+expDocumentNumber);
		System.out.println("Purchase Account                                : "+actPurchaseAccount  +" Value Expected : "+expPurchaseAccount);
		System.out.println("Department                                      : "+actDepartment       +" Value Expected : "+expDepartment);
		System.out.println("Entry Page Save with Vendor Item Value Actual   : "+actR1C1Item         +" Value Expected : "+expR1C1Item);	
		System.out.println("Entry Page Save with Vendor Units Value Actual  : "+actR1C2Units        +" Value Expected : "+expR1C2Units);
		System.out.println("Entry Page Save with Vendor QTY Value Actual    : "+actR1C5Qty          +" Value Expected : "+expR1C5Qty);
		System.out.println("Entry Page Save with Vendor RATE Value Actual   : "+actR1C7Rate         +" Value Expected : "+expR1C7Rate);
		System.out.println("Entry Page Save with Vendor Gross Value Actual  : "+actR1C8Gross        +" Value Expected : "+expR1C8Gross);			

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		Thread.sleep(2000);
		String expMessage = excelReader.getCellData(xlSheetName, 21, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actMessage);

		Thread.sleep(2000);

		if(actDocumentNumber.equalsIgnoreCase(expDocumentNumber) && actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actR1C1Item.equalsIgnoreCase(expR1C1Item) && actR1C2Units.equalsIgnoreCase(expR1C2Units)
				&& actR1C5Qty.equalsIgnoreCase(expR1C5Qty) && actR1C7Rate.equalsIgnoreCase(expR1C7Rate) && actR1C8Gross.equalsIgnoreCase(expR1C8Gross)
				&& actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 13, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 13, 8, resFail);
			return false;
		}	
	}








	public static boolean checkReSavingMaterialRequisitionVoucher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		MRpurchaseAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 22, 5));
		Thread.sleep(2000);
		MRpurchaseAccountTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 23, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 24, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 25, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 26, 5));
		enter_Rate.sendKeys(Keys.TAB);			

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

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
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 22, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 23, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 22, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actMessage.equals(expMessage))		*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 22, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 23, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 22, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
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








	public boolean checkVerifyingAndUpdatingMaterialRequisitionVoucher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocumentNumber = documentNumberTxt.getAttribute("value");
		String expDocumentNumber = excelReader.getCellData(xlSheetName, 27, 6);
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actDocumentNumber);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		String actPurchaseAccount = MRpurchaseAccountTxt.getAttribute("value");
		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 28, 6);
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actPurchaseAccount);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 29, 6);
		excelReader.setCellData(xlfile, xlSheetName, 29, 7, actDepartment);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));

		String actR1C1Item =select1stRow_1stColumn.getText();
		String actR1C2Units=select1stRow_2ndColumn.getText();
		String actR1C5Qty  =select1stRow_3rdColumn.getText();
		String actR1C7Rate =select1stRow_4thColumn.getText();
		String actR1C8Gross=select1stRow_5thColumn.getText();

		String expR1C1Item =excelReader.getCellData(xlSheetName, 30, 6);
		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actR1C1Item);
		
		String expR1C2Units=excelReader.getCellData(xlSheetName, 31, 6);
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actR1C2Units);
		
		String expR1C5Qty  =excelReader.getCellData(xlSheetName, 32, 6);
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actR1C5Qty);
		
		String expR1C7Rate =excelReader.getCellData(xlSheetName, 33, 6);
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actR1C7Rate);
		
		String expR1C8Gross=excelReader.getCellData(xlSheetName, 34, 6);
		excelReader.setCellData(xlfile, xlSheetName, 34, 7, actR1C8Gross);
		
		System.out.println("***************************************checkVerifyingAndUpdatingMaterialRequisitionVoucher1****************************************");
		System.out.println("Document Number                                 : "+actDocumentNumber   +" Value Expected : "+expDocumentNumber);
		System.out.println("Purchase Account                                : "+actPurchaseAccount  +" Value Expected : "+expPurchaseAccount);
		System.out.println("Department                                      : "+actDepartment       +" Value Expected : "+expDepartment);
		System.out.println("Entry Page Save with Vendor Item Value Actual   : "+actR1C1Item         +" Value Expected : "+expR1C1Item);	
		System.out.println("Entry Page Save with Vendor Units Value Actual  : "+actR1C2Units        +" Value Expected : "+expR1C2Units);
		System.out.println("Entry Page Save with Vendor QTY Value Actual    : "+actR1C5Qty          +" Value Expected : "+expR1C5Qty);
		System.out.println("Entry Page Save with Vendor RATE Value Actual   : "+actR1C7Rate         +" Value Expected : "+expR1C7Rate);
		System.out.println("Entry Page Save with Vendor Gross Value Actual  : "+actR1C8Gross        +" Value Expected : "+expR1C8Gross);			

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 27, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 28, 5));
		enter_Rate.sendKeys(Keys.TAB);			

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		Thread.sleep(3000);
		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 1");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 35, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 36, 6));
		
		String expMessage = expMsg.toString();

		excelReader.setCellData(xlfile, xlSheetName, 35, 7, actMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 35, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 36, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 35, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		if(actDocumentNumber.equalsIgnoreCase(expDocumentNumber) && actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actR1C1Item.equalsIgnoreCase(expR1C1Item) && actR1C2Units.equalsIgnoreCase(expR1C2Units)
				&& actR1C5Qty.equalsIgnoreCase(expR1C5Qty) && actR1C7Rate.equalsIgnoreCase(expR1C7Rate) && actR1C8Gross.equalsIgnoreCase(expR1C8Gross)
				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 27, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 27, 8, resFail);
			return false;
		}	
	}









	public boolean checkVerifyingMaterialRequisitionVoucher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocumentNumber = documentNumberTxt.getAttribute("value");
		String expDocumentNumber = excelReader.getCellData(xlSheetName, 37, 6);
		excelReader.setCellData(xlfile, xlSheetName, 37, 7, actDocumentNumber);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		String actPurchaseAccount = MRpurchaseAccountTxt.getAttribute("value");
		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 38, 6);
		excelReader.setCellData(xlfile, xlSheetName, 38, 7, actPurchaseAccount);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 39, 6);
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actDepartment);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));

		String actR1C1Item =select1stRow_1stColumn.getText();
		String actR1C2Units=select1stRow_2ndColumn.getText();
		String actR1C5Qty  =select1stRow_3rdColumn.getText();
		String actR1C7Rate =select1stRow_4thColumn.getText();
		String actR1C8Gross=select1stRow_5thColumn.getText();

		String expR1C1Item =excelReader.getCellData(xlSheetName, 40, 6);
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, actR1C1Item);
		
		String expR1C2Units=excelReader.getCellData(xlSheetName, 41, 6);
		excelReader.setCellData(xlfile, xlSheetName, 41, 7, actR1C2Units);
		
		String expR1C5Qty  =excelReader.getCellData(xlSheetName, 42, 6);
		excelReader.setCellData(xlfile, xlSheetName, 43, 7, actR1C5Qty);
		
		String expR1C7Rate =excelReader.getCellData(xlSheetName, 43, 6);
		excelReader.setCellData(xlfile, xlSheetName, 43, 7, actR1C7Rate);
		
		String expR1C8Gross=excelReader.getCellData(xlSheetName, 44, 6);
		excelReader.setCellData(xlfile, xlSheetName, 45, 7, actR1C8Gross);
		
		System.out.println("***************************************checkVerifyingMaterialRequisitionVoucher1****************************************");
		System.out.println("Document Number                                 : "+actDocumentNumber   +" Value Expected : "+expDocumentNumber);
		System.out.println("Purchase Account                                : "+actPurchaseAccount  +" Value Expected : "+expPurchaseAccount);
		System.out.println("Department                                      : "+actDepartment       +" Value Expected : "+expDepartment);
		System.out.println("Entry Page Save with Vendor Item Value Actual   : "+actR1C1Item         +" Value Expected : "+expR1C1Item);	
		System.out.println("Entry Page Save with Vendor Units Value Actual  : "+actR1C2Units        +" Value Expected : "+expR1C2Units);
		System.out.println("Entry Page Save with Vendor QTY Value Actual    : "+actR1C5Qty          +" Value Expected : "+expR1C5Qty);
		System.out.println("Entry Page Save with Vendor RATE Value Actual   : "+actR1C7Rate         +" Value Expected : "+expR1C7Rate);
		System.out.println("Entry Page Save with Vendor Gross Value Actual  : "+actR1C8Gross        +" Value Expected : "+expR1C8Gross);			

		if(actDocumentNumber.equalsIgnoreCase(expDocumentNumber) && actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actR1C1Item.equalsIgnoreCase(expR1C1Item) && actR1C2Units.equalsIgnoreCase(expR1C2Units)
				&& actR1C5Qty.equalsIgnoreCase(expR1C5Qty) && actR1C7Rate.equalsIgnoreCase(expR1C7Rate) && actR1C8Gross.equalsIgnoreCase(expR1C8Gross))
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

	
	


	
	public static boolean checkSavingMaterialRequisitionVoucher2AndSuspendingFromEntryPage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
		new_newBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		MRpurchaseAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 45, 5));
		Thread.sleep(2000);
		MRpurchaseAccountTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 46, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 47, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 48, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 49, 5));
		enter_Rate.sendKeys(Keys.TAB);			

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

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

		expMessage.add("Voucher saved successfully : 2");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

	//	expMsg.add(excelReader.getCellData(xlSheetName, 45, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 46, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 45, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 45, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 46, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 45, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
		new_SuspendBtn.click();
		Thread.sleep(2000);

		String expSuspendMessage1=excelReader.getCellData(xlSheetName, 47, 6);
		String expSuspendMessage2=excelReader.getCellData(xlSheetName, 48, 6);

		String actSuspendMessage=checkValidationMessage(expSuspendMessage1);

		excelReader.setCellData(xlfile, xlSheetName, 47, 7, actSuspendMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
		Thread.sleep(2000);
		
		String actSuspendStatus = voucherHomeRow1SuspendedStatus.getText();
		String expSuspendStatus = excelReader.getCellData(xlSheetName, 49, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 49, 7, actSuspendStatus);
		
		System.out.println("Suspended Status : "+actSuspendStatus+"  Value Expected  "+expSuspendStatus);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2) && actSuspendMessage.startsWith(expSuspendMessage1) && actSuspendMessage.endsWith(expSuspendMessage2)
				&& actSuspendStatus.equalsIgnoreCase(expSuspendStatus))
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
	
	
	

	
	
	public boolean checkResavingMaterialRequisitionVoucher2AndSuspendingFromHomepage() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 50, 5);

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
		Thread.sleep(2000);

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

		//expMsg.add(excelReader.getCellData(xlSheetName, 50, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 51, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 50, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 50, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 51, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 50, 7, actMessage);
	
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
			String data =excelReader.getCellData(xlSheetName, 51, 5);

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
		
		String expSuspendMessage = excelReader.getCellData(xlSheetName, 52, 6);
		String actSuspendMessage = checkValidationMessage(expSuspendMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 52, 7, actSuspendMessage);
		
		Thread.sleep(2000);
		
		String actSuspendStatus = voucherHomeRow1SuspendedStatus.getText();
		String expSuspendStatus = excelReader.getCellData(xlSheetName, 53, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 53, 7, actSuspendStatus);
		
		System.out.println("Suspended Status : "+actSuspendStatus+"  Value Expected  "+expSuspendStatus);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2) && actSuspendMessage.equalsIgnoreCase(expSuspendMessage)
				&& actSuspendStatus.equalsIgnoreCase(expSuspendStatus))
		{
			excelReader.setCellData(xlfile, xlSheetName, 50, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 50, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	public boolean checkDeleteOptionInMaterialRequisitionHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 54, 5);

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

		String expAlertMessage=excelReader.getCellData(xlSheetName, 54, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 54, 7, actAlertMessage);

		System.out.println("Alert Message on Deleteing the Voucher :  "+actAlertMessage +"  "+expAlertMessage);

		Thread.sleep(3000);

		getAlert().accept();

		String expMessage = excelReader.getCellData(xlSheetName, 55, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 55, 7, actMessage);
		
		System.out.println("Message : "+actMessage+" Value Expected : "+expMessage);

		if(actAlertMessage.contains(expAlertMessage) && actMessage.equalsIgnoreCase(expMessage))
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
	
	
	
	
	

	public boolean checkVerifingDetailsOfSavedMaterialRequisitionVouchersInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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

		expVouchersList.add(excelReader.getCellData(xlSheetName, 56, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 57, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 58, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 59, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 60, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 61, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 62, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 63, 6));
		
		excelReader.setCellData(xlfile, xlSheetName, 56, 7, actVouchersList.toString());

		System.out.println("**********************checkVerifingDetailsOfSavedOpeningStocksVouchersInHomepage*********************");

		System.out.println(actVouchersList);
		System.out.println(expVouchersList);

		if (actVouchersList.equals(expVouchersList)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 56, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 56, 8, resFail);
			return false;
		}
	} 

















	public MaterialRequisitionPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
