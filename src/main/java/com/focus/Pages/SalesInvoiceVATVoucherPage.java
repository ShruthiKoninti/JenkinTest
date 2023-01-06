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
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xalan.xsltc.compiler.util.BooleanType;
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

public class SalesInvoiceVATVoucherPage extends BaseEngine 
{
	private static String xlfile = getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlSheetName = "SalesInvoiceVATVoucherPage";
	
	public static boolean checkVoucherSavingMessage(String docno) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		try
		{
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(errorMessage));
			String actErrorMessage=errorMessage.getText();
			String expErrorMessage="Voucher saved successfully";
			String expErrorMessage1=": "+docno;

			System.out.println("SavingMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage+" "+expErrorMessage1);

			if(actErrorMessage.startsWith(expErrorMessage) && actErrorMessage.endsWith(expErrorMessage1))
			{
				try
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
					errorMessageCloseBtn.click();

					return  true;
				}
				catch(Exception ee)
				{
					return true;
				}	
			}	
			else
			{
				return false;
			}	
		}
		catch(Exception e)
		{
			System.err.println("UNABLE TO COMPARE");
			return false;
		}
	}	









	public boolean checkSalesInvoiceVoucherToCheckLinkPopUpScreenWithCRTlPlusL() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
		financialTransactionSalesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
		salesInvoiceVATVoucher.click();

		Thread.sleep(2000);

		checkDeleteLinkStatus();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 8, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
		salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		Thread.sleep(2000);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		
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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 14, 5));
		enter_AQTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 15, 5));
		enter_FQTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		List<WebElement> rowCountInLinkScreen=LinkRowCount;

		int actLinkRowCountValue1=rowCountInLinkScreen.size();
		String actLinkRowCountValue=String.valueOf(rowCountInLinkScreen.size());
		
		//String actLinkRowCountValue=actLinkRowCountValue1.toString();
		boolean actsearchOnDropdown=searchOnDropdown.isDisplayed();
		boolean actcontainingDropdown=containingDropdown.isDisplayed();
		boolean actworkFlowSearchTxt=workFlowSearchTxt.isDisplayed();
		boolean actworkFlowCustomizeBtn=workFlowCustomizeBtn.isDisplayed();
		boolean actworkFlowOkBtn=workFlowOkBtn.isDisplayed();

		String actworkFlowRow11Column1=workFlowRow1Column1.getText();
		String actworkFlowRow11Column2=workFlowRow1Column2.getText();
		String actworkFlowRow11Column3=workFlowRow1Column3.getText();
		String actworkFlowRow11Column4=workFlowRow1Column4.getText();
		//String actworkFlowRow11Column5=workFlowRow1Column5.getText();
		
		
		String actworkFlowRow2Column1=workFlowRow2Column1.getText();
		String actworkFlowRow2Column2=workFlowRow2Column2.getText();
		String actworkFlowRow2Column3=workFlowRow2Column3.getText();
		String actworkFlowRow2Column4=workFlowRow2Column4.getText();
		//String actworkFlowRow2Column5=workFlowRow2Column5.getText();


		// Expected

		//int expLinkRowCountValue=Integer.getInteger(excelReader.getCellData(xlSheetName, 8, 6));
		String expLinkRowCountValue=excelReader.getCellData(xlSheetName, 8, 6);
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actLinkRowCountValue);
		
		boolean expsearchOnDropdown=Boolean.getBoolean(excelReader.getCellData(xlSheetName, 9, 6));
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, Boolean.toString(actsearchOnDropdown));
		
		boolean expcontainingDropdown=Boolean.getBoolean(excelReader.getCellData(xlSheetName, 10, 6));
		excelReader.setCellData(xlfile, xlSheetName, 10, 7, Boolean.toString(actcontainingDropdown));
		
		boolean expworkFlowSearchTxt=Boolean.getBoolean(excelReader.getCellData(xlSheetName, 11, 6));
		excelReader.setCellData(xlfile, xlSheetName, 11, 7, Boolean.toString(actworkFlowSearchTxt));
		
		boolean expworkFlowCustomizeBtn=Boolean.getBoolean(excelReader.getCellData(xlSheetName, 12, 6));
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, Boolean.toString(actworkFlowCustomizeBtn));
		
		boolean expworkFlowOkBtn=Boolean.getBoolean(excelReader.getCellData(xlSheetName, 13, 6));
		excelReader.setCellData(xlfile, xlSheetName, 13, 7, Boolean.toString(actworkFlowOkBtn));

		String expworkFlowRow11Column1=excelReader.getCellData(xlSheetName, 14, 6);
		excelReader.setCellData(xlfile, xlSheetName, 14, 7, actworkFlowRow11Column1);
		
		String expworkFlowRow11Column3=excelReader.getCellData(xlSheetName, 15, 6);
		excelReader.setCellData(xlfile, xlSheetName, 15, 7, actworkFlowRow11Column3);
		
		String expworkFlowRow11Column4=excelReader.getCellData(xlSheetName, 16, 6);
		excelReader.setCellData(xlfile, xlSheetName, 16, 7, actworkFlowRow11Column4);
		
	/*	String expworkFlowRow11Column5=excelReader.getCellData(xlSheetName, 17, 6);
		excelReader.setCellData(xlfile, xlSheetName, 17, 7, actworkFlowRow11Column5);*/
		
		String expworkFlowRow2Column1=excelReader.getCellData(xlSheetName, 18, 6);
		excelReader.setCellData(xlfile, xlSheetName, 18, 7, actworkFlowRow2Column1);
		
		String expworkFlowRow2Column3=excelReader.getCellData(xlSheetName, 19, 6);
		excelReader.setCellData(xlfile, xlSheetName, 19, 7, actworkFlowRow2Column3);
		
		String expworkFlowRow2Column4=excelReader.getCellData(xlSheetName, 20, 6);
		excelReader.setCellData(xlfile, xlSheetName, 20, 7, actworkFlowRow2Column4);
		
	/*	String expworkFlowRow2Column5=excelReader.getCellData(xlSheetName, 21, 6);
		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actworkFlowRow2Column4);

		boolean actRow1Suspend = false;

		boolean expRow1Suspend = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 22, 6));
		
		if (actworkFlowRow11Column5.isEmpty() || actworkFlowRow11Column5.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 16, 5))) 
		{
			actRow1Suspend = true;
		}

		excelReader.setCellData(xlfile, xlSheetName, 22, 7, Boolean.toString(actRow1Suspend));
		*/
		boolean actRow2Suspend = false;

		boolean expRow2Suspend = Boolean.getBoolean(excelReader.getCellData(xlSheetName, 23, 6));

	/*	if (actworkFlowRow2Column5.isEmpty() || actworkFlowRow2Column5.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 17, 5))) 
		{
			actRow2Suspend = true;
		}*/
		
		excelReader.setCellData(xlfile, xlSheetName, 23, 7, Boolean.toString(actRow2Suspend));

		System.out.println("LinkRowCountValue  		: "+actLinkRowCountValue		+" 	Value Expected : "+expLinkRowCountValue);

		System.out.println("searchOnDropdown  		: "+actsearchOnDropdown		    +" 	Value Expected : "+expsearchOnDropdown);
		System.out.println("containingDropdown  	: "+actcontainingDropdown		+" 	Value Expected : "+expcontainingDropdown);
		System.out.println("workFlowSearchTxt  		: "+actworkFlowSearchTxt		+" 	Value Expected : "+expworkFlowSearchTxt);
		System.out.println("workFlowCustomizeBtn  	: "+actworkFlowCustomizeBtn	    +" 	Value Expected : "+expworkFlowCustomizeBtn);
		System.out.println("workFlowOkBtn  			: "+actworkFlowOkBtn			+" 	Value Expected : "+expworkFlowOkBtn);

		System.out.println("workFlowRow11Column1  	: "+actworkFlowRow11Column1	+" 	Value Expected : "+expworkFlowRow11Column1);
		System.out.println("workFlowRow11Column3  	: "+actworkFlowRow11Column3	+" 	Value Expected : "+expworkFlowRow11Column3);
		System.out.println("workFlowRow11Column4  	: "+actworkFlowRow11Column4	+" 	Value Expected : "+expworkFlowRow11Column4);
		//System.out.println("workFlowRow11Column5  	: "+actworkFlowRow11Column5	+" 	Value Expected : "+expworkFlowRow11Column5);
		
		System.out.println("workFlowRow2Column1  	: "+actworkFlowRow2Column1	+" 	Value Expected : "+expworkFlowRow2Column1);
		System.out.println("workFlowRow2Column3  	: "+actworkFlowRow2Column3	+" 	Value Expected : "+expworkFlowRow2Column3);
		System.out.println("workFlowRow2Column4  	: "+actworkFlowRow2Column4	+" 	Value Expected : "+expworkFlowRow2Column4);
		//System.out.println("workFlowRow2Column5  	: "+actworkFlowRow2Column5	+" 	Value Expected : "+expworkFlowRow2Column5);

		if(actLinkRowCountValue.equalsIgnoreCase(expLinkRowCountValue)

				&& actsearchOnDropdown==expsearchOnDropdown && actcontainingDropdown==expcontainingDropdown && actworkFlowSearchTxt==expworkFlowSearchTxt
				&& actworkFlowCustomizeBtn==expworkFlowCustomizeBtn && actworkFlowOkBtn==expworkFlowOkBtn

				&& actworkFlowRow11Column1.equalsIgnoreCase(expworkFlowRow11Column1)
				&& actworkFlowRow11Column3.equalsIgnoreCase(expworkFlowRow11Column3)
				&& actworkFlowRow11Column4.equalsIgnoreCase(expworkFlowRow11Column4)
				/*&& actRow1Suspend==expRow1Suspend*/
				
				&& actworkFlowRow2Column1.equalsIgnoreCase(expworkFlowRow2Column1)
				&& actworkFlowRow2Column3.equalsIgnoreCase(expworkFlowRow2Column3)
				&& actworkFlowRow2Column4.equalsIgnoreCase(expworkFlowRow2Column4)
				&& actRow2Suspend==expRow2Suspend)
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
		
		
		
		
		
		
		
		
		
	public boolean checkSalesInvoiceVoucherSearchInLinkPopupScreenAndSaveWithBatchItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{	
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		if (workFlowRow1Column4.getText().equalsIgnoreCase(excelReader.getCellData(xlSheetName, 24, 5))) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowRadioBtnRowOne));
			workFlowRadioBtnRowOne.click();
		} 
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowRadioBtnRowTwo));
			workFlowRadioBtnRowTwo.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowOkBtn));
		workFlowOkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(Keys.END);
		enter_Rate.sendKeys(Keys.SHIFT, Keys.HOME);
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 25, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
		select1stRow_19thColumn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		batchPickOnFIFOIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
		batchOkIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		Thread.sleep(2000);
		
		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
/*		String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 24, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 25, 6)+docno);
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 24, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 25, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		//if(actMessage.equals(expMessage))
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






	public boolean checkSalesInvoiceVoucherLinewiseConsumingForBatchItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.click();
		customerAccountTxt.sendKeys(Keys.END);
		customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 26, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 27, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
		salesInvoiceVATPlaceOFSupply.click();
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.END);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SHIFT,Keys.HOME);
		salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 28, 5));
		Thread.sleep(2000);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(Keys.END);
		jurisdictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 29, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 30, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 31, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 32, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(4000);

		List<WebElement> rowCountInLinkScreen=LinkRowCount;

		int actLinkRowCountValue=rowCountInLinkScreen.size();

		boolean actsearchOnDropdown=searchOnDropdown.isDisplayed();
		boolean actcontainingDropdown=containingDropdown.isDisplayed();
		boolean actworkFlowSearchTxt=workFlowSearchTxt.isDisplayed();
		boolean actworkFlowCustomizeBtn=workFlowCustomizeBtn.isDisplayed();
		boolean actworkFlowOkBtn=workFlowOkBtn.isDisplayed();

		String actworkFlowRow11Column1=workFlowRow1Column1.getText();
		String actworkFlowRow11Column2=workFlowRow1Column2.getText();
		String actworkFlowRow11Column3=workFlowRow1Column3.getText();
		String actworkFlowRow11Column4=workFlowRow1Column4.getText();
		//String actworkFlowRow11Column5=workFlowRow1Column5.getText();


		// Expected

		int expLinkRowCountValue=Integer.parseInt(excelReader.getCellData(xlSheetName, 26, 6));
		excelReader.setCellData(xlfile, xlSheetName, 26, 7, Integer.toString(actLinkRowCountValue));
		
		boolean expsearchOnDropdown=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 27, 6));
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, Boolean.toString(actsearchOnDropdown));
		
		boolean expcontainingDropdown=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 28, 6));
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, Boolean.toString(actcontainingDropdown));
		
		boolean expworkFlowSearchTxt=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 29, 6));
		excelReader.setCellData(xlfile, xlSheetName, 29, 7, Boolean.toString(actworkFlowSearchTxt));
		
		boolean expworkFlowCustomizeBtn=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 30, 6));
		excelReader.setCellData(xlfile, xlSheetName, 30, 7, Boolean.toString(actworkFlowCustomizeBtn));
		
		boolean expworkFlowOkBtn=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 31, 6));
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, Boolean.toString(actworkFlowOkBtn));

		String expworkFlowRow11Column1=excelReader.getCellData(xlSheetName, 32, 6);
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actworkFlowRow11Column1);
		
		String expworkFlowRow11Column3=excelReader.getCellData(xlSheetName, 33, 6);
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actworkFlowRow11Column3);
		
		String expworkFlowRow11Column4=excelReader.getCellData(xlSheetName, 34, 6);
		excelReader.setCellData(xlfile, xlSheetName, 34, 7, actworkFlowRow11Column4);
		
		/*String expworkFlowRow11Column5=excelReader.getCellData(xlSheetName, 35, 6);
		excelReader.setCellData(xlfile, xlSheetName, 35, 7, actworkFlowRow11Column5);*/

		boolean actRow1Suspend = false;

		boolean expRow1Suspend = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 36, 6));

	/*	if (actworkFlowRow11Column5.isEmpty() || actworkFlowRow11Column5.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 33, 5))) 
		{
			actRow1Suspend = true;
		}

		excelReader.setCellData(xlfile, xlSheetName, 36, 7, Boolean.toString(actRow1Suspend));
		*/
		System.out.println("LinkRowCountValue  		: "+actLinkRowCountValue		+" 	Value Expected : "+expLinkRowCountValue);

		System.out.println("searchOnDropdown  		: "+actsearchOnDropdown		    +" 	Value Expected : "+expsearchOnDropdown);
		System.out.println("containingDropdown  	: "+actcontainingDropdown		+" 	Value Expected : "+expcontainingDropdown);
		System.out.println("workFlowSearchTxt  		: "+actworkFlowSearchTxt		+" 	Value Expected : "+expworkFlowSearchTxt);
		System.out.println("workFlowCustomizeBtn  	: "+actworkFlowCustomizeBtn	    +" 	Value Expected : "+expworkFlowCustomizeBtn);
		System.out.println("workFlowOkBtn  			: "+actworkFlowOkBtn			+" 	Value Expected : "+expworkFlowOkBtn);

		System.out.println("workFlowRow11Column1  	: "+actworkFlowRow11Column1	+" 	Value Expected : "+expworkFlowRow11Column1);
		System.out.println("workFlowRow11Column3  	: "+actworkFlowRow11Column3	+" 	Value Expected : "+expworkFlowRow11Column3);
		System.out.println("workFlowRow11Column4  	: "+actworkFlowRow11Column4	+" 	Value Expected : "+expworkFlowRow11Column4);
		//System.out.println("workFlowRow11Column5  	: "+actworkFlowRow11Column5	+" 	Value Expected : "+expworkFlowRow11Column5);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowRadioBtnRowOne));
		workFlowRadioBtnRowOne.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowOkBtn));
		workFlowOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		// Actual
		String actLSalesOrderAfterLinkSecltionLinewsie =select1stRow_12thColumn.getText();
		String actQtyAfterLinkSecltionLinewsie         =select1stRow_11thColumn.getText();
		String actReleaseQtyAfterLinkSecltionLinewsie  =select1stRow_13thColumn.getText();

		// Expected
		String expLSalesOrderAfterLinkSecltionLinewsie   = excelReader.getCellData(xlSheetName, 37, 6);
		excelReader.setCellData(xlfile, xlSheetName, 37, 7, actLSalesOrderAfterLinkSecltionLinewsie);
		
		String expQtyAfterLinkSecltionLinewsie           = excelReader.getCellData(xlSheetName, 38, 6);
		excelReader.setCellData(xlfile, xlSheetName, 38, 7, actQtyAfterLinkSecltionLinewsie);
		
		String expReleaseQtyAfterLinkSecltionLinewsie    = excelReader.getCellData(xlSheetName, 39, 6);
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actReleaseQtyAfterLinkSecltionLinewsie);
		
		System.out.println("LSalesOrderAfterLinkSecltionLinewsie  : "+actLSalesOrderAfterLinkSecltionLinewsie	+" 	Value Expected : "+expLSalesOrderAfterLinkSecltionLinewsie);
		System.out.println("QtyAfterLinkSecltionLinewsie  		  : "+actQtyAfterLinkSecltionLinewsie			+" 	Value Expected : "+expQtyAfterLinkSecltionLinewsie);
		System.out.println("ReleaseQtyAfterLinkSecltionLinewsie   : "+actReleaseQtyAfterLinkSecltionLinewsie	+" 	Value Expected : "+expReleaseQtyAfterLinkSecltionLinewsie);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 34, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
		select1stRow_19thColumn.click();	

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		batchPickOnFIFOIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
		batchOkIcon.click();

		if(getIsAlertPresent())
		{
			getAlert().accept();
		}

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

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
	/*	String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 40, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 41, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 40, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 41, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		if(actLinkRowCountValue==expLinkRowCountValue && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)
		
				&& actsearchOnDropdown==expsearchOnDropdown && actcontainingDropdown==expcontainingDropdown && actworkFlowSearchTxt==expworkFlowSearchTxt
				&& actworkFlowCustomizeBtn==expworkFlowCustomizeBtn && actworkFlowOkBtn==expworkFlowOkBtn

				&& actworkFlowRow11Column1.equalsIgnoreCase(expworkFlowRow11Column1)
				&& actworkFlowRow11Column3.equalsIgnoreCase(expworkFlowRow11Column3)
				&& actworkFlowRow11Column4.equalsIgnoreCase(expworkFlowRow11Column4)
				&& actRow1Suspend==expRow1Suspend

				&& actLSalesOrderAfterLinkSecltionLinewsie.equalsIgnoreCase(expLSalesOrderAfterLinkSecltionLinewsie) 
				&& actQtyAfterLinkSecltionLinewsie.equalsIgnoreCase(expQtyAfterLinkSecltionLinewsie)
				&& actReleaseQtyAfterLinkSecltionLinewsie.equalsIgnoreCase(expReleaseQtyAfterLinkSecltionLinewsie))
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









	public boolean checkSuspendOptionInSalesInvoiceEntryPage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
		new_SuspendBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		boolean actMessage = checkVoucherSavingMessage(docno);
		boolean expMessage = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 42, 6)); 
		excelReader.setCellData(xlfile, xlSheetName, 42, 7, Boolean.toString(actMessage));
		
		if(actMessage==expMessage)
		{
			excelReader.setCellData(xlfile, xlSheetName, 42, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 42, 8, resFail);
			return false;
		}
	}

	
	@FindBy(xpath="//*[@id='id_body_16777315']")
	
	public static WebElement enter_TaxCode;

	public boolean checkSalesInvoiceVoucherToCheckBinItemLinkPopUpScreenWithCRTlPlusL() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 43, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 44, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);	


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
		salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 45, 5));
		Thread.sleep(2000);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 46, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		
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

		enter_TaxCode.click();
		enter_TaxCode.sendKeys("Std Rate");
		enter_TaxCode.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 49, 5));
		enter_AQTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		
		List<WebElement> rowCountInLinkScreen=LinkRowCount;

		int actLinkRowCountValue=rowCountInLinkScreen.size();

		boolean actsearchOnDropdown=searchOnDropdown.isDisplayed();
		boolean actcontainingDropdown=containingDropdown.isDisplayed();
		boolean actworkFlowSearchTxt=workFlowSearchTxt.isDisplayed();
		boolean actworkFlowCustomizeBtn=workFlowCustomizeBtn.isDisplayed();
		boolean actworkFlowOkBtn=workFlowOkBtn.isDisplayed();

		String actworkFlowRow11Column1=workFlowRow1Column1.getText();
		String actworkFlowRow11Column2=workFlowRow1Column2.getText();
		String actworkFlowRow11Column3=workFlowRow1Column3.getText();
		String actworkFlowRow11Column4=workFlowRow1Column4.getText();
		String actworkFlowRow11Column5=workFlowRow1Column5.getText();
		
		String actworkFlowRow2Column1=workFlowRow2Column1.getText();
		String actworkFlowRow2Column2=workFlowRow2Column2.getText();
		String actworkFlowRow2Column3=workFlowRow2Column3.getText();
		String actworkFlowRow2Column4=workFlowRow2Column4.getText();
		String actworkFlowRow2Column5=workFlowRow2Column5.getText();


		// Expected

		int expLinkRowCountValue=Integer.parseInt(excelReader.getCellData(xlSheetName, 43, 6));
		excelReader.setCellData(xlfile, xlSheetName, 43, 7, Integer.toString(actLinkRowCountValue));

		boolean expsearchOnDropdown=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 44, 6));
		excelReader.setCellData(xlfile, xlSheetName, 44, 7, Boolean.toString(actsearchOnDropdown));
		
		boolean expcontainingDropdown=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 45, 6));
		excelReader.setCellData(xlfile, xlSheetName, 45, 7, Boolean.toString(actcontainingDropdown));
		
		boolean expworkFlowSearchTxt=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 46, 6));
		excelReader.setCellData(xlfile, xlSheetName, 46, 7, Boolean.toString(actworkFlowSearchTxt));
		
		boolean expworkFlowCustomizeBtn=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 47, 6));
		excelReader.setCellData(xlfile, xlSheetName, 47, 7, Boolean.toString(actworkFlowCustomizeBtn));
		
		boolean expworkFlowOkBtn=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 48, 6));
		excelReader.setCellData(xlfile, xlSheetName, 48, 7, Boolean.toString(actworkFlowOkBtn));

		String expworkFlowRow11Column1=excelReader.getCellData(xlSheetName, 49, 6);
		excelReader.setCellData(xlfile, xlSheetName, 49, 7, actworkFlowRow11Column1);
		
		String expworkFlowRow11Column3=excelReader.getCellData(xlSheetName, 50, 6);
		excelReader.setCellData(xlfile, xlSheetName, 50, 7, actworkFlowRow11Column3);
		
		String expworkFlowRow11Column4=excelReader.getCellData(xlSheetName, 51, 6);
		excelReader.setCellData(xlfile, xlSheetName, 51, 7, actworkFlowRow11Column4);
		
		String expworkFlowRow11Column5=excelReader.getCellData(xlSheetName, 52, 6);
		excelReader.setCellData(xlfile, xlSheetName, 52, 7, actworkFlowRow11Column5);
		
		String expworkFlowRow2Column1=excelReader.getCellData(xlSheetName, 53, 6);
		excelReader.setCellData(xlfile, xlSheetName, 53, 7, actworkFlowRow2Column1);
		
		String expworkFlowRow2Column3=excelReader.getCellData(xlSheetName, 54, 6);
		excelReader.setCellData(xlfile, xlSheetName, 54, 7, actworkFlowRow2Column3);
		
		String expworkFlowRow2Column4=excelReader.getCellData(xlSheetName, 55, 6);
		excelReader.setCellData(xlfile, xlSheetName, 55, 7, actworkFlowRow2Column4);
		
		String expworkFlowRow2Column5=excelReader.getCellData(xlSheetName, 56, 6);
		excelReader.setCellData(xlfile, xlSheetName, 56, 7, actworkFlowRow2Column5);

		boolean actRow1Suspend = false;

		boolean expRow1Suspend = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 57, 6));

		if (actworkFlowRow11Column5.isEmpty() || actworkFlowRow11Column5.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 50, 5))) 
		{
			actRow1Suspend = true;
		}

		excelReader.setCellData(xlfile, xlSheetName, 57, 7, Boolean.toString(actRow1Suspend));
		
		boolean actRow2Suspend = false;

		boolean expRow2Suspend = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 58, 6));

		if (actworkFlowRow2Column5.isEmpty() || actworkFlowRow2Column5.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 51, 5))) 
		{
			actRow2Suspend = true;
		}
		
		excelReader.setCellData(xlfile, xlSheetName, 58, 7, Boolean.toString(actRow2Suspend));

		System.out.println("LinkRowCountValue  		: "+actLinkRowCountValue		+" 	Value Expected : "+expLinkRowCountValue);

		System.out.println("searchOnDropdown  		: "+actsearchOnDropdown		    +" 	Value Expected : "+expsearchOnDropdown);
		System.out.println("containingDropdown  	: "+actcontainingDropdown		+" 	Value Expected : "+expcontainingDropdown);
		System.out.println("workFlowSearchTxt  		: "+actworkFlowSearchTxt		+" 	Value Expected : "+expworkFlowSearchTxt);
		System.out.println("workFlowCustomizeBtn  	: "+actworkFlowCustomizeBtn	    +" 	Value Expected : "+expworkFlowCustomizeBtn);
		System.out.println("workFlowOkBtn  			: "+actworkFlowOkBtn			+" 	Value Expected : "+expworkFlowOkBtn);

		System.out.println("workFlowRow11Column1  	: "+actworkFlowRow11Column1	+" 	Value Expected : "+expworkFlowRow11Column1);
		System.out.println("workFlowRow11Column3  	: "+actworkFlowRow11Column3	+" 	Value Expected : "+expworkFlowRow11Column3);
		System.out.println("workFlowRow11Column4  	: "+actworkFlowRow11Column4	+" 	Value Expected : "+expworkFlowRow11Column4);
		System.out.println("workFlowRow11Column5  	: "+actworkFlowRow11Column5	+" 	Value Expected : "+expworkFlowRow11Column5);
		
		System.out.println("workFlowRow2Column1  	: "+actworkFlowRow2Column1	+" 	Value Expected : "+expworkFlowRow2Column1);
		System.out.println("workFlowRow2Column3  	: "+actworkFlowRow2Column3	+" 	Value Expected : "+expworkFlowRow2Column3);
		System.out.println("workFlowRow2Column4  	: "+actworkFlowRow2Column4	+" 	Value Expected : "+expworkFlowRow2Column4);
		System.out.println("workFlowRow2Column5  	: "+actworkFlowRow2Column5	+" 	Value Expected : "+expworkFlowRow2Column5);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowRadioBtnRowOne));
		workFlowRadioBtnRowOne.click();
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowOkBtn));
		workFlowOkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 52, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_20thColumn));
		select1stRow_20thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
		binAutoAllocateBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		binOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();
		Thread.sleep(2000);

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

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 59, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 60, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 59, 7, actMessage);
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
		
		if(actMessage.equals(expMessage))*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 59, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 60, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 59, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
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



	
	
	


	public boolean checkSalesInvoiceVoucherLinewiseConsumingForBinItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.click();
		customerAccountTxt.sendKeys(Keys.END);
		customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 61, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 62, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
		salesInvoiceVATPlaceOFSupply.click();
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.END);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SHIFT,Keys.HOME);
		salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 63, 5));
		Thread.sleep(2000);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(Keys.END);
		jurisdictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 64, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 65, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 66, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 67, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(4000);

		List<WebElement> rowCountInLinkScreen=LinkRowCount;

		int actLinkRowCountValue=rowCountInLinkScreen.size();

		boolean actsearchOnDropdown=searchOnDropdown.isDisplayed();
		boolean actcontainingDropdown=containingDropdown.isDisplayed();
		boolean actworkFlowSearchTxt=workFlowSearchTxt.isDisplayed();
		boolean actworkFlowCustomizeBtn=workFlowCustomizeBtn.isDisplayed();
		boolean actworkFlowOkBtn=workFlowOkBtn.isDisplayed();

		String actworkFlowRow11Column1=workFlowRow1Column1.getText();
		String actworkFlowRow11Column2=workFlowRow1Column2.getText();
		String actworkFlowRow11Column3=workFlowRow1Column3.getText();
		String actworkFlowRow11Column4=workFlowRow1Column4.getText();
		String actworkFlowRow11Column5=workFlowRow1Column5.getText();

		// Expected

		int expLinkRowCountValue=Integer.parseInt(excelReader.getCellData(xlSheetName, 61, 6));
		excelReader.setCellData(xlfile, xlSheetName, 61, 7, Integer.toString(actLinkRowCountValue));

		boolean expsearchOnDropdown=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 62, 6));
		excelReader.setCellData(xlfile, xlSheetName, 62, 7, Boolean.toString(actsearchOnDropdown));
		
		boolean expcontainingDropdown=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 63, 6));
		excelReader.setCellData(xlfile, xlSheetName, 63, 7, Boolean.toString(actcontainingDropdown));
		
		boolean expworkFlowSearchTxt=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 64, 6));
		excelReader.setCellData(xlfile, xlSheetName, 64, 7, Boolean.toString(actworkFlowSearchTxt));
		
		boolean expworkFlowCustomizeBtn=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 65, 6));
		excelReader.setCellData(xlfile, xlSheetName, 65, 7, Boolean.toString(actworkFlowCustomizeBtn));
		
		boolean expworkFlowOkBtn=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 66, 6));
		excelReader.setCellData(xlfile, xlSheetName, 66, 7, Boolean.toString(actworkFlowOkBtn));

		String expworkFlowRow11Column1=excelReader.getCellData(xlSheetName, 67, 6);
		excelReader.setCellData(xlfile, xlSheetName, 67, 7, actworkFlowRow11Column1);
		
		String expworkFlowRow11Column3=excelReader.getCellData(xlSheetName, 68, 6);
		excelReader.setCellData(xlfile, xlSheetName, 68, 7, actworkFlowRow11Column3);
		
		String expworkFlowRow11Column4=excelReader.getCellData(xlSheetName, 69, 6);
		excelReader.setCellData(xlfile, xlSheetName, 69, 7, actworkFlowRow11Column4);
		
		String expworkFlowRow11Column5=excelReader.getCellData(xlSheetName, 70, 6);
		excelReader.setCellData(xlfile, xlSheetName, 70, 7, actworkFlowRow11Column5);

		boolean actRow1Suspend = false;

		boolean expRow1Suspend = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 71, 6));

		if (actworkFlowRow11Column5.isEmpty() || actworkFlowRow11Column5.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 68, 5))) 
		{
			actRow1Suspend = true;
		}
		
		excelReader.setCellData(xlfile, xlSheetName, 71, 7, Boolean.toString(expRow1Suspend));

		System.out.println("LinkRowCountValue  		: "+actLinkRowCountValue		+" 	Value Expected : "+expLinkRowCountValue);

		System.out.println("searchOnDropdown  		: "+actsearchOnDropdown		    +" 	Value Expected : "+expsearchOnDropdown);
		System.out.println("containingDropdown  	: "+actcontainingDropdown		+" 	Value Expected : "+expcontainingDropdown);
		System.out.println("workFlowSearchTxt  		: "+actworkFlowSearchTxt		+" 	Value Expected : "+expworkFlowSearchTxt);
		System.out.println("workFlowCustomizeBtn  	: "+actworkFlowCustomizeBtn	    +" 	Value Expected : "+expworkFlowCustomizeBtn);
		System.out.println("workFlowOkBtn  			: "+actworkFlowOkBtn			+" 	Value Expected : "+expworkFlowOkBtn);

		System.out.println("workFlowRow11Column1  	: "+actworkFlowRow11Column1	+" 	Value Expected : "+expworkFlowRow11Column1);
		System.out.println("workFlowRow11Column3  	: "+actworkFlowRow11Column3	+" 	Value Expected : "+expworkFlowRow11Column3);
		System.out.println("workFlowRow11Column4  	: "+actworkFlowRow11Column4	+" 	Value Expected : "+expworkFlowRow11Column4);
		System.out.println("workFlowRow11Column5  	: "+actworkFlowRow11Column5	+" 	Value Expected : "+expworkFlowRow11Column5);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowRadioBtnRowOne));
		workFlowRadioBtnRowOne.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowOkBtn));
		workFlowOkBtn.click();

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_LSalesOrders));
		enter_LSalesOrders.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		// Actual
		String actLSalesOrderAfterLinkSecltionLinewsie =select1stRow_12thColumn.getText();
		String actQtyAfterLinkSecltionLinewsie         =select1stRow_11thColumn.getText();
		String actReleaseQtyAfterLinkSecltionLinewsie  =select1stRow_13thColumn.getText();

		// Expected
		String expLSalesOrderAfterLinkSecltionLinewsie   = excelReader.getCellData(xlSheetName, 72, 6);
		excelReader.setCellData(xlfile, xlSheetName, 72, 7, actLSalesOrderAfterLinkSecltionLinewsie);
		
		String expQtyAfterLinkSecltionLinewsie           = excelReader.getCellData(xlSheetName, 73, 6);
		excelReader.setCellData(xlfile, xlSheetName, 73, 7, actQtyAfterLinkSecltionLinewsie);

		System.out.println("LSalesOrderAfterLinkSecltionLinewsie  : "+actLSalesOrderAfterLinkSecltionLinewsie	+" 	Value Expected : "+expLSalesOrderAfterLinkSecltionLinewsie);
		System.out.println("QtyAfterLinkSecltionLinewsie  		  : "+actQtyAfterLinkSecltionLinewsie			+" 	Value Expected : "+expQtyAfterLinkSecltionLinewsie);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 69, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_20thColumn));
		select1stRow_20thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
		binAutoAllocateBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		binOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		Thread.sleep(2000);
		
		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 74, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 75, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 74, 7, actMessage);

		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
		
		
*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 74, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 75, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 74, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		if(actLinkRowCountValue==expLinkRowCountValue && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)

				&& actsearchOnDropdown==expsearchOnDropdown && actcontainingDropdown==expcontainingDropdown && actworkFlowSearchTxt==expworkFlowSearchTxt
				&& actworkFlowCustomizeBtn==expworkFlowCustomizeBtn && actworkFlowOkBtn==expworkFlowOkBtn

				&& actworkFlowRow11Column1.equalsIgnoreCase(expworkFlowRow11Column1)
				&& actworkFlowRow11Column3.equalsIgnoreCase(expworkFlowRow11Column3)
				&& actworkFlowRow11Column4.equalsIgnoreCase(expworkFlowRow11Column4)
				&& actRow1Suspend==expRow1Suspend

				&& actLSalesOrderAfterLinkSecltionLinewsie.equalsIgnoreCase(expLSalesOrderAfterLinkSecltionLinewsie) 
				&& actQtyAfterLinkSecltionLinewsie.equalsIgnoreCase(expQtyAfterLinkSecltionLinewsie))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 61, 8, resPass);
			return true;
		}	
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 61, 8, resFail);
			return false;
		}
	}	







	public boolean checkSuspendOptionInSalesInvoiceHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
		financialTransactionSalesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
		salesInvoiceVATVoucher.click();

		Thread.sleep(2000);

		
		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 76, 5);
			if(voucherGridDocNo.get(i).getText().equalsIgnoreCase(data))
			{
				voucherGridIndexChkBox.get(i).click();
				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(suspendBtn));
		suspendBtn.click();

		String expSuspendMessage=excelReader.getCellData(xlSheetName, 76, 6);
		String actSuspendMessage=checkValidationMessage(expSuspendMessage);
		excelReader.setCellData(xlfile, xlSheetName, 73, 7, actSuspendMessage);
		
		System.out.println("SuspendMessage  :  "+actSuspendMessage +" Value Expected : "+expSuspendMessage);

		Thread.sleep(2000);

		if(actSuspendMessage.equalsIgnoreCase(expSuspendMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 73, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 73, 8, resFail);
			return false;
		}
	}




	public boolean checkSalesInvoiceVoucherToCheckRMAItemLinkPopUpScreenWithCRTlPlusL() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 77, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 78, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);	


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
		salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 79, 5));
		Thread.sleep(2000);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 80, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 81, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 82, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 83, 5));
		enter_AQTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		if (validationConfirmationMessage.getText().isEmpty()==false) 
		{
			errorMessageCloseBtn.click();
		}
		
		List<WebElement> rowCountInLinkScreen=LinkRowCount;

		int actLinkRowCountValue=rowCountInLinkScreen.size();

		boolean actsearchOnDropdown=searchOnDropdown.isDisplayed();
		boolean actcontainingDropdown=containingDropdown.isDisplayed();
		boolean actworkFlowSearchTxt=workFlowSearchTxt.isDisplayed();
		boolean actworkFlowCustomizeBtn=workFlowCustomizeBtn.isDisplayed();
		boolean actworkFlowOkBtn=workFlowOkBtn.isDisplayed();

		String actworkFlowRow11Column1=workFlowRow1Column1.getText();
		String actworkFlowRow11Column2=workFlowRow1Column2.getText();
		String actworkFlowRow11Column3=workFlowRow1Column3.getText();
		String actworkFlowRow11Column4=workFlowRow1Column4.getText();
		String actworkFlowRow11Column5=workFlowRow1Column5.getText();
		
		
		String actworkFlowRow2Column1=workFlowRow2Column1.getText();
		String actworkFlowRow2Column2=workFlowRow2Column2.getText();
		String actworkFlowRow2Column3=workFlowRow2Column3.getText();
		String actworkFlowRow2Column4=workFlowRow2Column4.getText();
		String actworkFlowRow2Column5=workFlowRow2Column5.getText();


		// Expected

		int expLinkRowCountValue=Integer.parseInt(excelReader.getCellData(xlSheetName, 77, 6));
		excelReader.setCellData(xlfile, xlSheetName, 77, 7, Integer.toString(actLinkRowCountValue));

		boolean expsearchOnDropdown=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 78, 6));
		excelReader.setCellData(xlfile, xlSheetName, 78, 7, Boolean.toString(actsearchOnDropdown));
		
		boolean expcontainingDropdown=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 79, 6));
		excelReader.setCellData(xlfile, xlSheetName, 79, 7, Boolean.toString(actcontainingDropdown));
		
		boolean expworkFlowSearchTxt=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 80, 6));
		excelReader.setCellData(xlfile, xlSheetName, 80, 7, Boolean.toString(actworkFlowSearchTxt));
		
		boolean expworkFlowCustomizeBtn=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 81, 6));
		excelReader.setCellData(xlfile, xlSheetName, 81, 7, Boolean.toString(actworkFlowCustomizeBtn));
		
		boolean expworkFlowOkBtn=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 82, 6));
		excelReader.setCellData(xlfile, xlSheetName, 82, 7, Boolean.toString(actworkFlowOkBtn));
		

		String expworkFlowRow11Column1=excelReader.getCellData(xlSheetName, 83, 6);
		excelReader.setCellData(xlfile, xlSheetName, 83, 7, actworkFlowRow11Column1);
		
		String expworkFlowRow11Column3=excelReader.getCellData(xlSheetName, 84, 6);
		excelReader.setCellData(xlfile, xlSheetName, 84, 7, actworkFlowRow11Column3);
		
		String expworkFlowRow11Column4=excelReader.getCellData(xlSheetName, 85, 6);
		excelReader.setCellData(xlfile, xlSheetName, 85, 7, actworkFlowRow11Column4);
		
		String expworkFlowRow11Column5=excelReader.getCellData(xlSheetName, 86, 6);
		excelReader.setCellData(xlfile, xlSheetName, 86, 7, actworkFlowRow11Column5);
		
		String expworkFlowRow2Column1=excelReader.getCellData(xlSheetName, 87, 6);
		excelReader.setCellData(xlfile, xlSheetName, 87, 7, actworkFlowRow2Column1);
		
		String expworkFlowRow2Column3=excelReader.getCellData(xlSheetName, 88, 6);
		excelReader.setCellData(xlfile, xlSheetName, 88, 7, actworkFlowRow2Column3);
		
		String expworkFlowRow2Column4=excelReader.getCellData(xlSheetName, 89, 6);
		excelReader.setCellData(xlfile, xlSheetName, 89, 7, actworkFlowRow2Column4);
		
		String expworkFlowRow2Column5=excelReader.getCellData(xlSheetName, 90, 6);
		excelReader.setCellData(xlfile, xlSheetName, 90, 7, actworkFlowRow2Column5);

		boolean actRow1Suspend = false;

		boolean expRow1Suspend = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 91, 6));

		if (actworkFlowRow11Column5.isEmpty() || actworkFlowRow11Column5.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 84, 5))) 
		{
			actRow1Suspend = true;
		}
		
		excelReader.setCellData(xlfile, xlSheetName, 91, 7, Boolean.toString(actRow1Suspend));
		
		boolean actRow2Suspend = false;

		boolean expRow2Suspend = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 92, 6));

		if (actworkFlowRow2Column5.isEmpty() || actworkFlowRow2Column5.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 85, 5))) 
		{
			actRow2Suspend = true;
		}
		
		excelReader.setCellData(xlfile, xlSheetName, 92, 7, Boolean.toString(actRow2Suspend));

		System.out.println("LinkRowCountValue  		: "+actLinkRowCountValue		+" 	Value Expected : "+expLinkRowCountValue);

		System.out.println("searchOnDropdown  		: "+actsearchOnDropdown		    +" 	Value Expected : "+expsearchOnDropdown);
		System.out.println("containingDropdown  	: "+actcontainingDropdown		+" 	Value Expected : "+expcontainingDropdown);
		System.out.println("workFlowSearchTxt  		: "+actworkFlowSearchTxt		+" 	Value Expected : "+expworkFlowSearchTxt);
		System.out.println("workFlowCustomizeBtn  	: "+actworkFlowCustomizeBtn	    +" 	Value Expected : "+expworkFlowCustomizeBtn);
		System.out.println("workFlowOkBtn  			: "+actworkFlowOkBtn			+" 	Value Expected : "+expworkFlowOkBtn);

		System.out.println("workFlowRow11Column1  	: "+actworkFlowRow11Column1	+" 	Value Expected : "+expworkFlowRow11Column1);
		System.out.println("workFlowRow11Column3  	: "+actworkFlowRow11Column3	+" 	Value Expected : "+expworkFlowRow11Column3);
		System.out.println("workFlowRow11Column4  	: "+actworkFlowRow11Column4	+" 	Value Expected : "+expworkFlowRow11Column4);
		System.out.println("workFlowRow11Column5  	: "+actworkFlowRow11Column5	+" 	Value Expected : "+expworkFlowRow11Column5);
		
		System.out.println("workFlowRow2Column1  	: "+actworkFlowRow2Column1	+" 	Value Expected : "+expworkFlowRow2Column1);
		System.out.println("workFlowRow2Column3  	: "+actworkFlowRow2Column3	+" 	Value Expected : "+expworkFlowRow2Column3);
		System.out.println("workFlowRow2Column4  	: "+actworkFlowRow2Column4	+" 	Value Expected : "+expworkFlowRow2Column4);
		System.out.println("workFlowRow2Column5  	: "+actworkFlowRow2Column5	+" 	Value Expected : "+expworkFlowRow2Column5);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowRadioBtnRowTwo));
		workFlowRadioBtnRowTwo.click();
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowOkBtn));
		workFlowOkBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 86, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_20thColumn));
		select1stRow_20thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaoutwardExpandBtn));
		rmaoutwardExpandBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAHeaderChkBox));
		searchRMAHeaderChkBox.click();

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
		
		Thread.sleep(2000);


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

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 93, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 94, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 93, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actMessage.equals(expMessage))*/
		
		
	String expMessage1 = excelReader.getCellData(xlSheetName, 93, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 94, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 93, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 77, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 77, 8, resFail);
			return false;
		}
	}


	
	
	
	
	

	public boolean checkSalesInvoiceVoucherLinewiseConsumingForRmaItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		if(validationConfirmationMessage.getText().isEmpty()==false)
		{
			errorMessageCloseBtn.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.click();
		customerAccountTxt.sendKeys(Keys.END);
		customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 95, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 96, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
		salesInvoiceVATPlaceOFSupply.click();
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.END);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SHIFT,Keys.HOME);
		salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 97, 5));
		Thread.sleep(2000);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(Keys.END);
		jurisdictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 98, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 99, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 100, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 101, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(Keys.TAB);



		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		List<WebElement> rowCountInLinkScreen=LinkRowCount;

		int actLinkRowCountValue=rowCountInLinkScreen.size();


		boolean actsearchOnDropdown=searchOnDropdown.isDisplayed();
		boolean actcontainingDropdown=containingDropdown.isDisplayed();
		boolean actworkFlowSearchTxt=workFlowSearchTxt.isDisplayed();
		boolean actworkFlowCustomizeBtn=workFlowCustomizeBtn.isDisplayed();
		boolean actworkFlowOkBtn=workFlowOkBtn.isDisplayed();

		String actworkFlowRow11Column1=workFlowRow1Column1.getText();
		String actworkFlowRow11Column2=workFlowRow1Column2.getText();
		String actworkFlowRow11Column3=workFlowRow1Column3.getText();
		String actworkFlowRow11Column4=workFlowRow1Column4.getText();
		String actworkFlowRow11Column5=workFlowRow1Column5.getText();

		// Expected

		int expLinkRowCountValue=Integer.parseInt(excelReader.getCellData(xlSheetName, 95, 6));
		excelReader.setCellData(xlfile, xlSheetName, 95, 7, Integer.toString(actLinkRowCountValue));

		boolean expsearchOnDropdown=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 96, 6));
		excelReader.setCellData(xlfile, xlSheetName, 96, 7, Boolean.toString(actsearchOnDropdown));
		
		boolean expcontainingDropdown=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 97, 6));
		excelReader.setCellData(xlfile, xlSheetName, 97, 7, Boolean.toString(actcontainingDropdown));
		
		boolean expworkFlowSearchTxt=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 98, 6));
		excelReader.setCellData(xlfile, xlSheetName, 98, 7, Boolean.toString(actworkFlowSearchTxt));
		
		boolean expworkFlowCustomizeBtn=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 99, 6));
		excelReader.setCellData(xlfile, xlSheetName, 99, 7, Boolean.toString(actworkFlowCustomizeBtn));
		
		boolean expworkFlowOkBtn=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 100, 6));
		excelReader.setCellData(xlfile, xlSheetName, 100, 7, Boolean.toString(actworkFlowOkBtn));

		String expworkFlowRow11Column1=excelReader.getCellData(xlSheetName, 101, 6);
		excelReader.setCellData(xlfile, xlSheetName, 101, 7, actworkFlowRow11Column1);
		
		String expworkFlowRow11Column3=excelReader.getCellData(xlSheetName, 102, 6);
		excelReader.setCellData(xlfile, xlSheetName, 102, 7, actworkFlowRow11Column3);
		
		String expworkFlowRow11Column4=excelReader.getCellData(xlSheetName, 103, 6);
		excelReader.setCellData(xlfile, xlSheetName, 104, 7, actworkFlowRow11Column4);
		
		String expworkFlowRow11Column5=excelReader.getCellData(xlSheetName, 104, 6);
		excelReader.setCellData(xlfile, xlSheetName, 105, 7, actworkFlowRow11Column5);

		boolean actRow1Suspend = false;
		boolean expRow1Suspend = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 105, 6));

		if (actworkFlowRow11Column5.isEmpty() || actworkFlowRow11Column5.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 102, 5))) 
		{
			actRow1Suspend = true;
		}

		excelReader.setCellData(xlfile, xlSheetName, 105, 7, Boolean.toString(actRow1Suspend));

		System.out.println("LinkRowCountValue  		: "+actLinkRowCountValue		+" 	Value Expected : "+expLinkRowCountValue);

		System.out.println("searchOnDropdown  		: "+actsearchOnDropdown		    +" 	Value Expected : "+expsearchOnDropdown);
		System.out.println("containingDropdown  	: "+actcontainingDropdown		+" 	Value Expected : "+expcontainingDropdown);
		System.out.println("workFlowSearchTxt  		: "+actworkFlowSearchTxt		+" 	Value Expected : "+expworkFlowSearchTxt);
		System.out.println("workFlowCustomizeBtn  	: "+actworkFlowCustomizeBtn	    +" 	Value Expected : "+expworkFlowCustomizeBtn);
		System.out.println("workFlowOkBtn  			: "+actworkFlowOkBtn			+" 	Value Expected : "+expworkFlowOkBtn);

		System.out.println("workFlowRow11Column1  	: "+actworkFlowRow11Column1	+" 	Value Expected : "+expworkFlowRow11Column1);
		System.out.println("workFlowRow11Column3  	: "+actworkFlowRow11Column3	+" 	Value Expected : "+expworkFlowRow11Column3);
		System.out.println("workFlowRow11Column4  	: "+actworkFlowRow11Column4	+" 	Value Expected : "+expworkFlowRow11Column4);
		System.out.println("workFlowRow11Column5  	: "+actworkFlowRow11Column5	+" 	Value Expected : "+expworkFlowRow11Column5);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowRadioBtnRowOne));
		workFlowRadioBtnRowOne.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowOkBtn));
		workFlowOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(Keys.END);
		enter_Quantity.sendKeys(Keys.SHIFT, Keys.HOME);
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 103, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		// Actual
		String actLSalesOrderAfterLinkSecltionLinewsie =select1stRow_12thColumn.getText();
		String actQtyAfterLinkSecltionLinewsie         =select1stRow_11thColumn.getText();
		String actReleaseQtyAfterLinkSecltionLinewsie  =select1stRow_13thColumn.getText();

		// Expected
		String expLSalesOrderAfterLinkSecltionLinewsie   = excelReader.getCellData(xlSheetName, 106, 6);
		excelReader.setCellData(xlfile, xlSheetName, 106, 7, actLSalesOrderAfterLinkSecltionLinewsie);
		
		String expQtyAfterLinkSecltionLinewsie           = excelReader.getCellData(xlSheetName, 107, 6);
		excelReader.setCellData(xlfile, xlSheetName, 107, 7, actQtyAfterLinkSecltionLinewsie);
		
		String expReleaseQtyAfterLinkSecltionLinewsie    = excelReader.getCellData(xlSheetName, 108, 6);
		excelReader.setCellData(xlfile, xlSheetName, 108, 7, actReleaseQtyAfterLinkSecltionLinewsie);
		
		System.out.println("LSalesOrderAfterLinkSecltionLinewsie  : "+actLSalesOrderAfterLinkSecltionLinewsie	+" 	Value Expected : "+expLSalesOrderAfterLinkSecltionLinewsie);
		System.out.println("QtyAfterLinkSecltionLinewsie  		  : "+actQtyAfterLinkSecltionLinewsie			+" 	Value Expected : "+expQtyAfterLinkSecltionLinewsie);
		System.out.println("ReleaseQtyAfterLinkSecltionLinewsie   : "+actReleaseQtyAfterLinkSecltionLinewsie	+" 	Value Expected : "+expReleaseQtyAfterLinkSecltionLinewsie);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_20thColumn));
		select1stRow_20thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
		enter_RMA.sendKeys(Keys.SPACE);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaoutwardExpandBtn));
		rmaoutwardExpandBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAHeaderChkBox));
		searchRMAHeaderChkBox.click();

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

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 109, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 110, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 109, 7, actMessage);
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 109, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 110, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 109, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		if(actLinkRowCountValue==expLinkRowCountValue && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)

				&& actsearchOnDropdown==expsearchOnDropdown && actcontainingDropdown==expcontainingDropdown && actworkFlowSearchTxt==expworkFlowSearchTxt
				&& actworkFlowCustomizeBtn==expworkFlowCustomizeBtn && actworkFlowOkBtn==expworkFlowOkBtn

				&& actworkFlowRow11Column1.equalsIgnoreCase(expworkFlowRow11Column1)
				&& actworkFlowRow11Column3.equalsIgnoreCase(expworkFlowRow11Column3)
				&& actworkFlowRow11Column4.equalsIgnoreCase(expworkFlowRow11Column4)
				&& actRow1Suspend==expRow1Suspend

				&& actLSalesOrderAfterLinkSecltionLinewsie.equalsIgnoreCase(expLSalesOrderAfterLinkSecltionLinewsie) 
				&& actQtyAfterLinkSecltionLinewsie.equalsIgnoreCase(expQtyAfterLinkSecltionLinewsie)
				&& actReleaseQtyAfterLinkSecltionLinewsie.equalsIgnoreCase(expReleaseQtyAfterLinkSecltionLinewsie))
		{
			excelReader.setCellData(xlfile, xlSheetName, 95, 8, resPass);
			return true;
		}	
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 95, 8, resFail);
			return false;
		}
	}	

	
	
	
	


	public boolean checkDeleteOptionInSalesInvoiceEntryPage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded Successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();
		Thread.sleep(2000);

		String actAlertMessage=getAlert().getText();

		String expAlertMessage=excelReader.getCellData(xlSheetName, 111, 6);

		excelReader.setCellData(xlfile, xlSheetName, 111, 7, actAlertMessage);
		
		System.out.println("Alert Message on Deleteing the Voucher :  "+actAlertMessage +"  "+expAlertMessage);

		getAlert().accept();

		String expMessage = excelReader.getCellData(xlSheetName, 112, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 112, 7, actMessage);

		System.out.println("Message : "+actMessage+" Value Expected : "+expMessage);

		Thread.sleep(2000);

		if(actAlertMessage.contains(expAlertMessage) && actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 111, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 111, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	


	public boolean checkSalesInvoiceVoucherToCheckStockItemLinkPopUpScreenWithCRTlPlusL() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 113, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 114, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);	


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
		salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 115, 5));
		Thread.sleep(2000);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 116, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 117, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 118, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 119, 5));
		enter_AQTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		List<WebElement> rowCountInLinkScreen=LinkRowCount;

		int actLinkRowCountValue=rowCountInLinkScreen.size();


		boolean actsearchOnDropdown=searchOnDropdown.isDisplayed();
		boolean actcontainingDropdown=containingDropdown.isDisplayed();
		boolean actworkFlowSearchTxt=workFlowSearchTxt.isDisplayed();
		boolean actworkFlowCustomizeBtn=workFlowCustomizeBtn.isDisplayed();
		boolean actworkFlowOkBtn=workFlowOkBtn.isDisplayed();

		String actworkFlowRow11Column1=workFlowRow1Column1.getText();
		String actworkFlowRow11Column2=workFlowRow1Column2.getText();
		String actworkFlowRow11Column3=workFlowRow1Column3.getText();
		String actworkFlowRow11Column4=workFlowRow1Column4.getText();
		String actworkFlowRow11Column5=workFlowRow1Column5.getText();

		// Expected

		int expLinkRowCountValue=Integer.parseInt(excelReader.getCellData(xlSheetName, 113, 6));
		excelReader.setCellData(xlfile, xlSheetName, 113, 7, Integer.toString(actLinkRowCountValue));

		boolean expsearchOnDropdown=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 114, 6));
		excelReader.setCellData(xlfile, xlSheetName, 114, 7, Boolean.toString(actsearchOnDropdown));
		
		boolean expcontainingDropdown=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 115, 6));
		excelReader.setCellData(xlfile, xlSheetName, 115, 7, Boolean.toString(actcontainingDropdown));
		
		boolean expworkFlowSearchTxt=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 116, 6));
		excelReader.setCellData(xlfile, xlSheetName, 116, 7, Boolean.toString(actworkFlowSearchTxt));
		
		boolean expworkFlowCustomizeBtn=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 117, 6));
		excelReader.setCellData(xlfile, xlSheetName, 117, 7, Boolean.toString(actworkFlowCustomizeBtn));
		
		boolean expworkFlowOkBtn=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 118, 6));
		excelReader.setCellData(xlfile, xlSheetName, 118, 7, Boolean.toString(actworkFlowOkBtn));

		String expworkFlowRow11Column1=excelReader.getCellData(xlSheetName, 119, 6);
		excelReader.setCellData(xlfile, xlSheetName, 119, 7, actworkFlowRow11Column1);
		
		String expworkFlowRow11Column3=excelReader.getCellData(xlSheetName, 120, 6);
		excelReader.setCellData(xlfile, xlSheetName, 120, 7, actworkFlowRow11Column3);
		
		String expworkFlowRow11Column4=excelReader.getCellData(xlSheetName, 121, 6);
		excelReader.setCellData(xlfile, xlSheetName, 121, 7, actworkFlowRow11Column4);
		
		String expworkFlowRow11Column5=excelReader.getCellData(xlSheetName, 122, 6);
		excelReader.setCellData(xlfile, xlSheetName, 122, 7, actworkFlowRow11Column5);

		boolean actRow1Suspend = false;
		boolean expRow1Suspend = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 123, 6));

		if (actworkFlowRow11Column5.isEmpty() || actworkFlowRow11Column5.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 120, 5))) 
		{
			actRow1Suspend = true;
		}

		excelReader.setCellData(xlfile, xlSheetName, 123, 7, Boolean.toString(actRow1Suspend));

		System.out.println("LinkRowCountValue  		: "+actLinkRowCountValue		+" 	Value Expected : "+expLinkRowCountValue);

		System.out.println("searchOnDropdown  		: "+actsearchOnDropdown		    +" 	Value Expected : "+expsearchOnDropdown);
		System.out.println("containingDropdown  	: "+actcontainingDropdown		+" 	Value Expected : "+expcontainingDropdown);
		System.out.println("workFlowSearchTxt  		: "+actworkFlowSearchTxt		+" 	Value Expected : "+expworkFlowSearchTxt);
		System.out.println("workFlowCustomizeBtn  	: "+actworkFlowCustomizeBtn	    +" 	Value Expected : "+expworkFlowCustomizeBtn);
		System.out.println("workFlowOkBtn  			: "+actworkFlowOkBtn			+" 	Value Expected : "+expworkFlowOkBtn);

		System.out.println("workFlowRow11Column1  	: "+actworkFlowRow11Column1	+" 	Value Expected : "+expworkFlowRow11Column1);
		System.out.println("workFlowRow11Column3  	: "+actworkFlowRow11Column3	+" 	Value Expected : "+expworkFlowRow11Column3);
		System.out.println("workFlowRow11Column4  	: "+actworkFlowRow11Column4	+" 	Value Expected : "+expworkFlowRow11Column4);
		System.out.println("workFlowRow11Column5  	: "+actworkFlowRow11Column5	+" 	Value Expected : "+expworkFlowRow11Column5);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowRadioBtnRowOne));
		workFlowRadioBtnRowOne.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowOkBtn));
		workFlowOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 121, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);


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

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 124, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 125, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 124, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actMessage.equals(expMessage))*/
		
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 124, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 125, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 124, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 113, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 113, 8, resFail);
			return false;
		}
	}

	
	
	
	
	

	public boolean checkSalesInvoiceVoucherLinewiseConsumingForStockItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.click();
		customerAccountTxt.sendKeys(Keys.END);
		customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 126, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 127, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
		salesInvoiceVATPlaceOFSupply.click();
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.END);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SHIFT,Keys.HOME);
		salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 128, 5));
		Thread.sleep(2000);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(Keys.END);
		jurisdictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 129, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 130, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 131, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 132, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		List<WebElement> rowCountInLinkScreen=LinkRowCount;

		int actLinkRowCountValue=rowCountInLinkScreen.size();

		boolean actsearchOnDropdown=searchOnDropdown.isDisplayed();
		boolean actcontainingDropdown=containingDropdown.isDisplayed();
		boolean actworkFlowSearchTxt=workFlowSearchTxt.isDisplayed();
		boolean actworkFlowCustomizeBtn=workFlowCustomizeBtn.isDisplayed();
		boolean actworkFlowOkBtn=workFlowOkBtn.isDisplayed();

		String actworkFlowRow11Column1=workFlowRow1Column1.getText();
		String actworkFlowRow11Column2=workFlowRow1Column2.getText();
		String actworkFlowRow11Column3=workFlowRow1Column3.getText();
		String actworkFlowRow11Column4=workFlowRow1Column4.getText();
		String actworkFlowRow11Column5=workFlowRow1Column5.getText();


		// Expected

		int expLinkRowCountValue=Integer.parseInt(excelReader.getCellData(xlSheetName, 126, 6));
		excelReader.setCellData(xlfile, xlSheetName, 126, 7, Integer.toString(actLinkRowCountValue));

		boolean expsearchOnDropdown=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 127, 6));
		excelReader.setCellData(xlfile, xlSheetName, 127, 7, Boolean.toString(actsearchOnDropdown));
		
		boolean expcontainingDropdown=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 128, 6));
		excelReader.setCellData(xlfile, xlSheetName, 128, 7, Boolean.toString(actcontainingDropdown));
		
		boolean expworkFlowSearchTxt=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 129, 6));
		excelReader.setCellData(xlfile, xlSheetName, 129, 7, Boolean.toString(actworkFlowSearchTxt));
		
		boolean expworkFlowCustomizeBtn=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 130, 6));
		excelReader.setCellData(xlfile, xlSheetName, 130, 7, Boolean.toString(actworkFlowCustomizeBtn));
		
		boolean expworkFlowOkBtn=Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 131, 6));
		excelReader.setCellData(xlfile, xlSheetName, 131, 7, Boolean.toString(actworkFlowOkBtn));

		String expworkFlowRow11Column1=excelReader.getCellData(xlSheetName, 132, 6);
		excelReader.setCellData(xlfile, xlSheetName, 132, 7, actworkFlowRow11Column1);
		
		String expworkFlowRow11Column3=excelReader.getCellData(xlSheetName, 133, 6);
		excelReader.setCellData(xlfile, xlSheetName, 133, 7, actworkFlowRow11Column3);
		
		String expworkFlowRow11Column4=excelReader.getCellData(xlSheetName, 134, 6);
		excelReader.setCellData(xlfile, xlSheetName, 134, 7, actworkFlowRow11Column4);
		
		String expworkFlowRow11Column5=excelReader.getCellData(xlSheetName, 135, 6);
		excelReader.setCellData(xlfile, xlSheetName, 135, 7, actworkFlowRow11Column5);

		boolean actRow1Suspend = false;
		boolean expRow1Suspend = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 136, 6));

		if (actworkFlowRow11Column5.isEmpty() || actworkFlowRow11Column5.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 133, 5))) 
		{
			actRow1Suspend = true;
		}

		excelReader.setCellData(xlfile, xlSheetName, 136, 7, Boolean.toString(actRow1Suspend));

		System.out.println("LinkRowCountValue  		: "+actLinkRowCountValue		+" 	Value Expected : "+expLinkRowCountValue);

		System.out.println("searchOnDropdown  		: "+actsearchOnDropdown		    +" 	Value Expected : "+expsearchOnDropdown);
		System.out.println("containingDropdown  	: "+actcontainingDropdown		+" 	Value Expected : "+expcontainingDropdown);
		System.out.println("workFlowSearchTxt  		: "+actworkFlowSearchTxt		+" 	Value Expected : "+expworkFlowSearchTxt);
		System.out.println("workFlowCustomizeBtn  	: "+actworkFlowCustomizeBtn	    +" 	Value Expected : "+expworkFlowCustomizeBtn);
		System.out.println("workFlowOkBtn  			: "+actworkFlowOkBtn			+" 	Value Expected : "+expworkFlowOkBtn);

		System.out.println("workFlowRow11Column1  	: "+actworkFlowRow11Column1	+" 	Value Expected : "+expworkFlowRow11Column1);
		System.out.println("workFlowRow11Column3  	: "+actworkFlowRow11Column3	+" 	Value Expected : "+expworkFlowRow11Column3);
		System.out.println("workFlowRow11Column4  	: "+actworkFlowRow11Column4	+" 	Value Expected : "+expworkFlowRow11Column4);
		System.out.println("workFlowRow11Column5  	: "+actworkFlowRow11Column5	+" 	Value Expected : "+expworkFlowRow11Column5);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowRadioBtnRowOne));
		workFlowRadioBtnRowOne.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workFlowOkBtn));
		workFlowOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		// Actual
		String actLSalesOrderAfterLinkSecltionLinewsie =select1stRow_12thColumn.getText();
		String actQtyAfterLinkSecltionLinewsie         =select1stRow_11thColumn.getText();
		String actReleaseQtyAfterLinkSecltionLinewsie  =select1stRow_13thColumn.getText();

		// Expected
		String expLSalesOrderAfterLinkSecltionLinewsie   = excelReader.getCellData(xlSheetName, 137, 6);
		excelReader.setCellData(xlfile, xlSheetName, 137, 7, actLSalesOrderAfterLinkSecltionLinewsie);

		String expQtyAfterLinkSecltionLinewsie           = excelReader.getCellData(xlSheetName, 138, 6);
		excelReader.setCellData(xlfile, xlSheetName, 138, 7, actQtyAfterLinkSecltionLinewsie);
		
		String expReleaseQtyAfterLinkSecltionLinewsie    = excelReader.getCellData(xlSheetName, 139, 6);
		excelReader.setCellData(xlfile, xlSheetName, 139, 7, actReleaseQtyAfterLinkSecltionLinewsie);
		
		System.out.println("LSalesOrderAfterLinkSecltionLinewsie  : "+actLSalesOrderAfterLinkSecltionLinewsie	+" 	Value Expected : "+expLSalesOrderAfterLinkSecltionLinewsie);
		System.out.println("QtyAfterLinkSecltionLinewsie  		  : "+actQtyAfterLinkSecltionLinewsie			+" 	Value Expected : "+expQtyAfterLinkSecltionLinewsie);
		System.out.println("ReleaseQtyAfterLinkSecltionLinewsie   : "+actReleaseQtyAfterLinkSecltionLinewsie	+" 	Value Expected : "+expReleaseQtyAfterLinkSecltionLinewsie);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 134, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

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

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 140, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 141, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 140, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 140, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 141, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 140, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

	//	if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();

		Thread.sleep(2000);

		if(actLinkRowCountValue==expLinkRowCountValue && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)

				&& actsearchOnDropdown==expsearchOnDropdown && actcontainingDropdown==expcontainingDropdown && actworkFlowSearchTxt==expworkFlowSearchTxt
				&& actworkFlowCustomizeBtn==expworkFlowCustomizeBtn && actworkFlowOkBtn==expworkFlowOkBtn

				&& actworkFlowRow11Column1.equalsIgnoreCase(expworkFlowRow11Column1)
				&& actworkFlowRow11Column3.equalsIgnoreCase(expworkFlowRow11Column3)
				&& actworkFlowRow11Column4.equalsIgnoreCase(expworkFlowRow11Column4)
				&& actRow1Suspend==expRow1Suspend

				&& actLSalesOrderAfterLinkSecltionLinewsie.equalsIgnoreCase(expLSalesOrderAfterLinkSecltionLinewsie) 
				&& actQtyAfterLinkSecltionLinewsie.equalsIgnoreCase(expQtyAfterLinkSecltionLinewsie)
				&& actReleaseQtyAfterLinkSecltionLinewsie.equalsIgnoreCase(expReleaseQtyAfterLinkSecltionLinewsie))
		{
			excelReader.setCellData(xlfile, xlSheetName, 126, 8, resPass);
			return true;
		}	
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 126, 8, resFail);
			return false;
		}
	}	


	
	
	public boolean checkDeleteOptionInSalesInvoiceHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 142, 5);

			if(voucherGridDocNo.get(i).getText().equalsIgnoreCase(data))
			{
				voucherGridIndexChkBox.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHomeDeleteBtn));
		voucherHomeDeleteBtn.click();

		getWaitForAlert();

		String actAlertMessage=getAlert().getText();

		String expAlertMessage=excelReader.getCellData(xlSheetName, 142, 6);

		excelReader.setCellData(xlfile, xlSheetName, 142, 7, actAlertMessage);

		System.out.println("Alert Message on Deleteing the Voucher :  "+actAlertMessage +"  "+expAlertMessage);

		getAlert().accept();

		String expMessage = excelReader.getCellData(xlSheetName, 143, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 143, 7, actMessage);

		System.out.println("Message : "+actMessage+" Value Expected : "+expMessage);

		if(actAlertMessage.contains(expAlertMessage) && actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 142, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 142, 8, resFail);
			return false;
		}
	}








	public boolean checkSalesInvoiceVoucherSavingSuspenigAndDeletingThroughConvertOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
		financialTransactionSalesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
		salesInvoiceVATVoucher.click();

		Thread.sleep(2000);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homepagePendingSalesOrders));
		homepagePendingSalesOrders.click();

		Thread.sleep(3000);

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 144, 5);
			if(voucherGridDocNo.get(i).getText().equalsIgnoreCase(data))
			{
				voucherGridIndexChkBox.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertBtn));
		convertBtn.click();

		checkValidationMessage("");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 145, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);
		
		getWaitForAlert();
		getAlert().accept();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
		salesInvoiceVATPlaceOFSupply.click();
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.END);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SHIFT,Keys.HOME);
		salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 146, 5));
		Thread.sleep(2000);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(Keys.END);
		jurisdictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 147, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);



		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 148, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		Thread.sleep(2000);

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 144, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 145, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 144, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 144, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 145, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 144, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

	//	if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();
		Thread.sleep(2000);

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno2=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
		new_SuspendBtn.click();

		boolean SuspendMessage=checkVoucherSavingMessage(docno2);

		String actSuspendMessage=Boolean.toString(SuspendMessage);
		String expSuspendMessage=excelReader.getCellData(xlSheetName, 146, 6);
		excelReader.setCellData(xlfile, xlSheetName, 146, 7, actSuspendMessage);

		System.out.println("SuspendMessage  :  "+actSuspendMessage +" Value Expected : "+expSuspendMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();

		String actAlertMessage=getAlert().getText();

		String expAlertMessage=excelReader.getCellData(xlSheetName, 147, 6);

		excelReader.setCellData(xlfile, xlSheetName, 147, 7, actAlertMessage);

		System.out.println("Alert Message on Deleteing the Voucher :  "+actAlertMessage +"  "+expAlertMessage);

		getAlert().accept();

		String expDeleteMessage = excelReader.getCellData(xlSheetName, 148, 6);
		String actDeleteMessage = checkValidationMessage(expDeleteMessage);
		excelReader.setCellData(xlfile, xlSheetName, 148, 7, actDeleteMessage);

		System.out.println("Message : "+actMessage+" Value Expected : "+expMessage1);

		Thread.sleep(3000);

		if( actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2) && actSuspendMessage.equals(expSuspendMessage) 
				&& actAlertMessage.equalsIgnoreCase(expAlertMessage) && actDeleteMessage.equalsIgnoreCase(expDeleteMessage))
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




	
	


	public boolean checkSavingSalesInvoiceWithhAllItemsOfTransferedStockInSingleVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.click();
		customerAccountTxt.sendKeys(Keys.END);
		customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 149, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 150, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
		salesInvoiceVATPlaceOFSupply.click();
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.END);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SHIFT,Keys.HOME);
		salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 151, 5));
		Thread.sleep(2000);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(Keys.END);
		jurisdictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 152, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 153, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 154, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 155, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 156, 5));
		enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 157, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
		select1stRow_19thColumn.click();	

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
		batchPickOnFIFOIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
		batchOkIcon.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 158, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_8thColumn));
		select2ndRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 159, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_11thColumn));
		select2ndRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_14thColumn));
		select2ndRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 160, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_20thColumn));
		select2ndRow_20thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();

		Thread.sleep(7000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
		binAutoAllocateBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		binOkBtn.click();

		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
		select3rdRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 161, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_8thColumn));
		select3rdRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 162, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_11thColumn));
		select3rdRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_14thColumn));
		select3rdRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 163, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_20thColumn));
		select3rdRow_20thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaoutwardExpandBtn));
		rmaoutwardExpandBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAHeaderChkBox));
		searchRMAHeaderChkBox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
		searchRMAOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaoutwardOkBtn));
		rmaoutwardOkBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_1stColumn));
		select4thRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 164, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_5thColumn));
		select4thRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_8thColumn));
		select4thRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 165, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_11thColumn));
		select4thRow_11thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_14thColumn));
		select4thRow_14thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 166, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

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

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 149, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 150, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 149, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actMessage.equals(expMessage))*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 149, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 150, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 150, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 149, 8, resPass);
			return true;
		}	
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 149, 8, resFail);
			return false;
		}
	}	






	public boolean checkVerifingDetailsOfSavedSalesInvoiceVATVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherAll));
		voucherAll.click();

		Thread.sleep(1000);

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

			if (data.isEmpty()==false && i!=9 && i!=10 && i!=23 && i!=24 && i!=37 && i!=38 && i!=51 && i!=52 && i!=65 && i!=66 && i!=79 && i!=80 && i!=93 && i!=94) 
			{
				System.out.println(data);

				actVouchersList.add(data);
			}
		}


		ArrayList<String>  expVouchersList= new ArrayList<String>(); 

		expVouchersList.add(excelReader.getCellData(xlSheetName, 167, 5));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 168, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 169, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 170, 5));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 171, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 172, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 173, 5));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 174, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 175, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 176, 5));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 177, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 178, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 179, 5));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 180, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 181, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 182, 5));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 183, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 184, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 185, 5));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 186, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 187, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 188, 5));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 189, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 190, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 191, 5));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 192, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 193, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 194, 5));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 195, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 196, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 197, 5));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 198, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 199, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 200, 5));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 201, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 202, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 203, 5));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 204, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 205, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 206, 5));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 207, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 208, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 209, 5));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 210, 5));
		
		excelReader.setCellData(xlfile, xlSheetName, 167, 7, actVouchersList.toString());

		System.out.println("**********************checkVerifingDetailsOfSavedSalesInvoiceVATVoucherInHomepage*********************");

		System.out.println(actVouchersList);
		System.out.println(expVouchersList);

		if (actVouchersList.equals(expVouchersList)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 167, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 167, 8, resFail);
			return false;
		}
	}









	public boolean checkSavingSalesReturnsVoucherWithConvertOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
		financialTransactionSalesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesReturnsVoucher));
		salesReturnsVoucher.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pendingSalesInvoicesView));
		pendingSalesInvoicesView.click();

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 211, 5);
			if(voucherGridDocNo.get(i).getText().equalsIgnoreCase(data))
			{
				voucherGridIndexChkBox.get(i).click();
				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertBtn));
		convertBtn.click();

		checkValidationMessage("");

		Thread.sleep(3000);

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

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 211, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 212, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 211, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
		
		if(actMessage.equals(expMessage))*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 211, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 212, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 211, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 211, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 211, 8, resFail);
			return false;
		}
	}








	public boolean checkVerifingDetailsOfSavedSalesReturnsVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherAll));
		voucherAll.click();

		Thread.sleep(1000);

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

		expVouchersList.add(excelReader.getCellData(xlSheetName, 213, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 214, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 215, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 126, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 217, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 218, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 219, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 220, 6));
		
		excelReader.setCellData(xlfile, xlSheetName, 213, 7, actVouchersList.toString());

		System.out.println("**********************checkVerifingDetailsOfSavedSalesReturnsVoucherInHomepage*********************");

		System.out.println(actVouchersList);
		System.out.println(expVouchersList);

		if (actVouchersList.equals(expVouchersList)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 213, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 213, 8, resFail);
			return false;
		}
	}






	public boolean checkPrintPdfOfSalesInvoiceVAT() throws Exception, EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
		financialTransactionSalesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
		salesInvoiceVATVoucher.click();

		Thread.sleep(2000);

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

		Thread.sleep(2000);

		getDriver().switchTo().window(openTabs.get(2)).close();
		Thread.sleep(2000);
		getDriver().switchTo().window(openTabs.get(1)).close();
		Thread.sleep(2000);
		getDriver().switchTo().window(openTabs.get(0));

		String actPDF = "C:\\Users\\D.n\\Downloads\\SalesInvoice.pdf";
		String expPDF = "D:\\FocusSanity\\FocusAI\\autoIt\\Prints\\ExpectedSalesInvoice.pdf";

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

	
	
	@FindBy(xpath="//*[@id='tblFooterReportRender']/div[2]")
	public static WebElement footerDetailsTxt;
	
	@FindBy(xpath="//*[@id='id_footer']/strong")
	public static WebElement footerCompanyTxt;
	
	

	public boolean checkVoucherFooterDetailsatHomePage() throws InterruptedException
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
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
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(footerDetailsTxt));
		String actText=footerDetailsTxt.getText();
		
		String s1=actText.substring(0, 17);
			
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

		if(actText.startsWith(s2) && actText.endsWith(s3) && actGetLoginCompanyNameInformation.equalsIgnoreCase(expCompNameatFooter))
		{
		
			return true;
		}
		else
		{
			return false;
		}
	}



////////////////////////
	@FindBy(xpath="//tbody[@id='id_header_4_table_body']/tr/td[2]")
	private static List<WebElement> customerAccountListCount; 
	
	public boolean checkSearchingCustomerandItemsListwithoutAstrickSigninSalesInvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
		financialTransactionSalesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
		salesInvoiceVATVoucher.click();

		Thread.sleep(2000);

		checkDeleteLinkStatus();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.click();
		customerAccountTxt.sendKeys("er");
		
		
		ArrayList<String>  actCustomerAccountList= new ArrayList<String>(); 
		int customercount=customerAccountListCount.size();

		System.err.println(customercount);

		for(int i=0 ; i < customercount ;i++)
		{
			String data=customerAccountListCount.get(i).getText();

			actCustomerAccountList.add(data);
		}

		boolean actCustList=actCustomerAccountList.isEmpty();
		boolean expCustList=true;
		
		System.out.println(actCustList);
		System.out.println(expCustList);
		
		
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
		if(actCustList==expCustList  && actItemsList==expItemsList)
		{
			return true;
		} 
		else 
		{
			return false;
		}

			
	}
	

	
	public boolean checkSearchingofVendorandItemsListusingAstricksigninSalesInvoiceVAT() throws InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.click();
		customerAccountTxt.sendKeys("*er");
		
		
		ArrayList<String>  actCustomerAccountList= new ArrayList<String>(); 
		int customercount=customerAccountListCount.size();

		System.err.println(customercount);

		for(int i=0 ; i < customercount ;i++)
		{
			String data=customerAccountListCount.get(i).getText();

			actCustomerAccountList.add(data);
		}

		String actCustomersList=null;
		
		if(actCustomerAccountList.isEmpty()==false)
		{
			
			actCustomersList=actCustomerAccountList.toString();
		}
		
		String expCustomersList="[Customer A, Customer B, Customer C, Customer Display CD For Each Account One, Customer Display CD For Each Account Three, Customer Display CD For Each Account Two, Customer Update]";
		
		vendorAccountTxt.sendKeys(Keys.TAB);

			
		System.out.println(actCustomersList);
		System.out.println(expCustomersList);
		
		
		
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
		if(actCustomersList.equalsIgnoreCase(expCustomersList)  && actitemsArry.equalsIgnoreCase(expItemsArry))
		{
			return true;
		} 
		else 
		{
			return false;
		}
	}

	



	


	public SalesInvoiceVATVoucherPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
