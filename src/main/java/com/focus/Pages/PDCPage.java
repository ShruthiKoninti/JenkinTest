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

public class PDCPage extends BaseEngine
{
	private static String xlfile = getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlSheetName = "PDCPage";
	
	
	@FindBy(xpath="//*[@id='id_search_menu']/input")
	private static WebElement searchMenuTxt;
	
	public boolean checkPDCOptionsUnderSettings() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(preferencesBtn));
		preferencesBtn.click();
		//Thread.sleep(2000);

	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTxt));
		searchMenuTxt.click();
		searchMenuTxt.sendKeys("Configure Transactions");
		searchMenuTxt.sendKeys(Keys.ENTER);
		
	*/
		Thread.sleep(3000);
		
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDCBtn));
		getAction().moveToElement(PDCBtn).build().perform();
		PDCBtn.click();
		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(postDatedChequeChkbox));
		if (postDatedChequeChkboxIsSelected.isSelected()==false)
		{
			postDatedChequeChkbox.click();
		}

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(displayLedgerAndBalanceChkbox));
		if (displayLedgerAndBalanceChkboxIsSelected.isSelected()==true)
		{
			displayLedgerAndBalanceChkbox.click();
		}

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		getWaitForAlert();

		getAlert().accept();

		String expValidationMessage=excelReader.getCellData(xlSheetName, 8, 6);

		String actValidationMessage=checkValidationMessage(expValidationMessage);

		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actValidationMessage);
		
		if (actValidationMessage.equalsIgnoreCase(expValidationMessage)) 
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



	




	public boolean checkConvertedMaturePDCScreenWithNoPDCExists() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getDriver().navigate().refresh();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(autoPostingMenu));
		autoPostingMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCSMenu));
		convertMaturedPDCSMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridSelectAllOption));
		gridSelectAllOption.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCsOkIcon));
		convertMaturedPDCsOkIcon.click();

		/*String expValidationMessage="No PDC's exists";

		String actValidationMessage=checkValidationMessage(expValidationMessage);*/
		
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcVoucherCancelIcon));
		
		boolean actpdcVoucherCancelIcon = pdcVoucherCancelIcon.isDisplayed();
		
		boolean exppdcVoucherCancelIcon = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 9, 6));
		
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, Boolean.toString(actpdcVoucherCancelIcon));
		
		System.out.println("PDC Voucher Cancel Icon : "+actpdcVoucherCancelIcon+"  Value Expected  "+exppdcVoucherCancelIcon);
				
		if(actpdcVoucherCancelIcon==exppdcVoucherCancelIcon /*&& actValidationMessage.equalsIgnoreCase(expValidationMessage)*/)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcVoucherCancelIcon));
			pdcVoucherCancelIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 9, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcVoucherCancelIcon));
			pdcVoucherCancelIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 9, 8, resFail);
			return false;
		}
	}





	//PostDatedReceipts

	public boolean checkpostDatedRecepitsCustomizationOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDRVAT));
		PDRVAT.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();


		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));
		toggleBtn.click();

		//Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingBtn));
		settingBtn.click();

		//Thread.sleep(2000);
		settingMiscellaneousTab.click();

		//Thread.sleep(2000);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateFAChkbox));
		    if(updateFAChkbox.isSelected()==false)
		    {
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateFAChkbox));
			    updateFAChkbox.click();


		    }*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingUpdateIcon));
		settingUpdateIcon.click();

		String expValidationMessage="Data saved successfully";

		String actValidationMessage=checkValidationMessage(expValidationMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingCloseIcon));
		settingCloseIcon.click();

		if(actValidationMessage.equalsIgnoreCase(expValidationMessage))
		{
			System.out.println("****Test Pass : Data Updated Successfully in Header Feild OF Post Dated Receipts *********");
			return true;
		}
		else
		{
			System.err.println("****Test Fail : Data  Not Updated Successfully*********");
			return false;
		}

	}



	public boolean checkSavingVoucherToPostDatedReceipts() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getDriver().navigate().refresh();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDRVAT));
		PDRVAT.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCashBankAccountTxt));
		newCashBankAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		Thread.sleep(2000);
		newCashBankAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDRVATPlaceOfSupplyTXt));
		PDRVATPlaceOfSupplyTXt.click();
		PDRVATPlaceOfSupplyTXt.sendKeys(Keys.END);
		PDRVATPlaceOfSupplyTXt.sendKeys(Keys.SHIFT,Keys.HOME);
		PDRVATPlaceOfSupplyTXt.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		Thread.sleep(2000);
		PDRVATPlaceOfSupplyTXt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDRVAT_JuridictionTxt));
		PDRVAT_JuridictionTxt.sendKeys(Keys.END);
		PDRVAT_JuridictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		PDRVAT_JuridictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
		Thread.sleep(2000);
		PDRVAT_JuridictionTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 14, 5));
		Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enterTaxcode));
		enterTaxcode.click();
		enterTaxcode.sendKeys(Keys.SHIFT,Keys.HOME);
		enterTaxcode.sendKeys(excelReader.getCellData(xlSheetName, 15, 5));
		Thread.sleep(2000);
		enterTaxcode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Amount));
		enter_Amount.click();
		enter_Amount.clear();
		enter_Amount.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));
		enter_Amount.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPartyName));
		String actPartyName = billRefPartyName.getText();
		String expPartyName = excelReader.getCellData(xlSheetName, 10, 6);
		excelReader.setCellData(xlfile, xlSheetName, 10, 7, actPartyName);
		
		System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		String actbillRefNewReferenceTxt=billRefNewReferenceTxt.getAttribute("value");
		String expbillRefNewReferenceTxt=excelReader.getCellData(xlSheetName, 11, 6);
		excelReader.setCellData(xlfile, xlSheetName, 11, 7, actbillRefNewReferenceTxt);

		System.err.println(" billRefNewReferenceTxt : "+actbillRefNewReferenceTxt +" Value Expected : "+expbillRefNewReferenceTxt);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		//Thread.sleep(2000);

		/*HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/

		//String actMessage = SavingInBackground(2);
		/*String actMessage = SavingInBackground(1);
		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 12, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 13, 6)+docno);
		
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actMessage);
		
		String expMessage = expMsg.toString();
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 12, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 13, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		
		if(actbillRefNewReferenceTxt.equalsIgnoreCase(expbillRefNewReferenceTxt) && 
				actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 10, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 10, 8, resFail);
			return false;
		}
	}

	
	
	
	

	public boolean checkSavedVoucherInPDRVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 17, 6);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dateTxt));
		dateTxt.click();

		String actDate=dateTxt.getAttribute("value");
		String expDate=maturityDateTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashBankAccountTxt));
		cashBankAccountTxt.click();

		String actAccount=cashBankAccountTxt.getAttribute("value");
		String expAccount=excelReader.getCellData(xlSheetName, 18, 6);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.click();

		String actCurrency=voucherHeaderCurrency.getAttribute("value");
		String expCurrency=excelReader.getCellData(xlSheetName, 19, 6);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderExchangeRate));
		voucherHeaderExchangeRate.click();

		String actExchangeRate=voucherHeaderExchangeRate.getAttribute("value");
		String expExchangeRate=excelReader.getCellData(xlSheetName, 20, 6);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();

		String actdepartment=departmentTxt.getAttribute("value");
		String expdepartment=excelReader.getCellData(xlSheetName, 21, 6);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderLocalExchangeRate));
		voucherHeaderLocalExchangeRate.click();

		String actLocalExchangeRate=voucherHeaderLocalExchangeRate.getAttribute("value");
		String expLocalExchangeRate=excelReader.getCellData(xlSheetName, 22, 6);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actGridAccount=select1stRow_1stColumn.getText();
		String expGridAccount=excelReader.getCellData(xlSheetName, 23, 6);

		String actTaxCode=select1stRow_2ndColumn.getText();
		String expTaxCode=excelReader.getCellData(xlSheetName, 24, 6);

		String actAmount=select1stRow_3rdColumn.getText();
		String expAmount=excelReader.getCellData(xlSheetName, 25, 6);

		System.out.println("**************************************checkSavedVoucherInPDRVAT******************");

		System.out.println("DocNo              : "+actDocNo           +" Value Expected  : "+expDocNo);
		System.out.println("Date               : "+actDate            +" Value Expected  : "+expDate);
		System.out.println("Account            : "+actAccount         +" Value Expected  : "+expAccount);
		System.out.println("Currency           :"+actCurrency         +" Value Expected  : "+expCurrency);
		System.out.println("ExchangeRate       :"+actExchangeRate     +" Value Expected  : "+expExchangeRate);
		System.out.println("department         :"+actdepartment       +" Value Expected  : "+expdepartment);
		System.out.println("LocalExchangeRate  :"+actLocalExchangeRate+" Value Expected  :"+expLocalExchangeRate);
		System.out.println("TaxCode           :"+actTaxCode           +" Value Expected  : "+expTaxCode);
		System.out.println("GridAccount        :"+actGridAccount      +" Value Expected  : "+expGridAccount);
		System.out.println("Amount             :"+actAmount           +" Value Expected  : "+expAmount);

		if (actDocNo.equalsIgnoreCase(expDocNo) && actDate.equalsIgnoreCase(expDate) && actAccount.equalsIgnoreCase(expAccount) &&
				actCurrency.equalsIgnoreCase(expCurrency)&& actExchangeRate.equalsIgnoreCase(expExchangeRate)&& 
				actdepartment.equalsIgnoreCase(expdepartment) && actLocalExchangeRate.equalsIgnoreCase(expLocalExchangeRate) && 
				actTaxCode.equalsIgnoreCase(expTaxCode)&&
				actGridAccount.equalsIgnoreCase(expGridAccount) && actAmount.equalsIgnoreCase(expAmount) ) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 17, 8, resPass);
			return true;
		}
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 17, 8, resFail);
			return false;
		}
	}



	
	

	public boolean checkEditingVoucherAndSavingInPDRVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDRVATChequeNoTxt));
		PDRVATChequeNoTxt.click();
		PDRVATChequeNoTxt.sendKeys(excelReader.getCellData(xlSheetName, 26, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actPDRVATChequeNoTxt=PDRVATChequeNoTxt.getAttribute("value");
		String expPDRVATChequeNoTxt=excelReader.getCellData(xlSheetName, 26, 6);
		excelReader.setCellData(xlfile, xlSheetName, 26, 7, actPDRVATChequeNoTxt);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");
		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();
		Thread.sleep(2000);
		//enter_Amount.sendKeys(Keys.CLEAR);
		//enter_Amount.sendKeys("6");
		enter_Amount.sendKeys(excelReader.getCellData(xlSheetName, 27, 5));
		enter_Amount.sendKeys(Keys.TAB);
		

		String actAmount=select1stRow_3rdColumn.getText();
		String expAmount=excelReader.getCellData(xlSheetName, 27, 6);
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actAmount);
		
		Thread.sleep(2000);

		System.out.println("***************************checkEditingVoucherAfterSavingInPDRVAT*************");

		System.out.println("PDRVATChequeNoTxt           :"+actPDRVATChequeNoTxt +" Value Expected  : "+expPDRVATChequeNoTxt);
		System.out.println("Amount                       :"+actAmount           +" Value Expected  : "+expAmount);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		//Thread.sleep(2000);

		/*HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
		//String actMessage = SavingInBackground(2);
		/*String actMessage = SavingInBackground(1);
		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 28, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 28, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
String expMessage1 = excelReader.getCellData(xlSheetName, 28, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 29, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		if (actAmount.equalsIgnoreCase(expAmount)  && actPDRVATChequeNoTxt.equalsIgnoreCase(expPDRVATChequeNoTxt) && 
				actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
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


	@FindBy(xpath="//*[@id='mytableoptReversePosting']")
	public static WebElement bankType;
	




	public boolean checkSavedVouchersInPDCVoucherScreenWithOutPostOnDate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(autoPostingMenu));
		autoPostingMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCSMenu));
		convertMaturedPDCSMenu.click();
		
		Thread.sleep(2000);

		int pdcDOClistCount=pdcDOClist.size();

		for (int i = 0; i < pdcDOClistCount; i++) 
		{
			String data=pdcDOClist.get(i).getText();
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 30, 5)))
			{
				pdcCheckBoxlist.get(i).click();
			}
		}

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(bankType));
		bankType.click();
		bankType.sendKeys("HDFC");
		bankType.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCsOkIcon));
		convertMaturedPDCsOkIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcGridRow1Chkbox));
		pdcGridRow1Chkbox.click();

		String actNumber=pdcGridRow1Col3.getText();
		String expNumber=excelReader.getCellData(xlSheetName, 30, 6);
		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actNumber);

		String actMaturityDate=pdcGridRow1Col4.getText();
		DateFormat df = new SimpleDateFormat("dd-MM-yyyy");
		Date date=new Date();
		String CurrentDate=df.format(date);
		String expMaturityDate=CurrentDate;

		String actName=pdcGridRow1Col5.getText();
		String expName=excelReader.getCellData(xlSheetName, 31, 6);
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actName);

		String actAmount=pdcGridRow1Col7.getText();
		String expAmount=excelReader.getCellData(xlSheetName, 32, 6);
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actAmount);

		String actReassignedVouNo=pdcGridRow1Col8.getText();
		String expReassignedVouNo=excelReader.getCellData(xlSheetName, 33, 6);
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actReassignedVouNo);
		
		System.out.println("******************checkSavedVouchersInPDCVoucherScreenWithOUTPostOnDate********************************");
		System.out.println(" Number         : "+actNumber           +" Value Expected  : "+expNumber);
		System.out.println("Name            : "+actName             +" Value Expected  : "+expName);
		System.out.println("Amount          : "+actAmount           +" Value Expected  : "+expAmount);
		System.out.println("ReassignedVouNo : " +actReassignedVouNo +" Value Expected  : "+expReassignedVouNo);

		if (actNumber.equalsIgnoreCase(expNumber) && actMaturityDate.equalsIgnoreCase(expMaturityDate)&&
				actName.equalsIgnoreCase(expName) &&actAmount.equalsIgnoreCase(expAmount) &&  actReassignedVouNo.equalsIgnoreCase(expReassignedVouNo))
		{
			excelReader.setCellData(xlfile, xlSheetName, 30, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 30, 8, resFail);
			return false;
		}
	}


	
	
	


	public boolean checkConvertingVouchersInPDCVoucherScreenWithOutPostOnDate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		Thread.sleep(1999);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcVoucherOkIcon));
		pdcVoucherOkIcon.click();

		String expValidationMessage=excelReader.getCellData(xlSheetName, 34, 6);

		String actValidationMessage=checkValidationMessage(expValidationMessage);

		excelReader.setCellData(xlfile, xlSheetName, 34, 7, actValidationMessage);
		
		if (actValidationMessage.equalsIgnoreCase(expValidationMessage) )
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





	//Receipts

	public boolean checkAddExtraFieldPDCInHeaderFieldOFReceipts() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receiptsVoucher));
		receiptsVoucher.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsBtn));
		settingsBtn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTab));
		editLayoutTab.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutAddFieldsBtn));
		editLayoutAddFieldsBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCaptionTxt));
		editLayoutCaptionTxt.click();
		editLayoutCaptionTxt.sendKeys(excelReader.getCellData(xlSheetName, 35, 5));
		editLayoutCaptionTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutApplyBtn));
		editLayoutApplyBtn.click();

		checkValidationMessage(" ");

		Thread.sleep(3000);

		for (int i = 0; i < 5; i++) 
		{
			if (getIsAlertPresent()==true) 
			{
				getAlert().accept();
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingUpdateIcon));
		settingUpdateIcon.click();

		for (int i = 0; i < 5; i++) 
		{
			if (getIsAlertPresent()==true) 
			{
				getAlert().accept();
			}
		}

		String expValidationMessage=excelReader.getCellData(xlSheetName, 35, 6);

		String actValidationMessage=checkValidationMessage(expValidationMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 35, 7, actValidationMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
		settings_closeBtn.click();

		if (actValidationMessage.equalsIgnoreCase(expValidationMessage) )
		{
			excelReader.setCellData(xlfile, xlSheetName, 35, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 35, 8, resFail);
			return false;
		}

	}





	



	public boolean checkCovertingVoucherInPDC() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(autoPostingMenu));
		autoPostingMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCSMenu));
		convertMaturedPDCSMenu.click();
		
		Thread.sleep(2000);

		int pdcDOClistCount=pdcDOClist.size();

		for (int i = 0; i < pdcDOClistCount; i++) 
		{
			String data=pdcDOClist.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 36, 5)))
			{
				pdcCheckBoxlist.get(i).click();

			}
		}
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(bankType));
		bankType.click();
		bankType.sendKeys("HDFC");
		Thread.sleep(2000);
*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCsOkIcon));
		convertMaturedPDCsOkIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcGridRow1Chkbox));
		pdcGridRow1Chkbox.click();

		boolean getGridSelectRow1Chkbox=pdcGridRow1Chkbox.isSelected();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcVoucherOkIcon));
		pdcVoucherOkIcon.click();

		String expValidationMessage=excelReader.getCellData(xlSheetName, 36, 6);

		String actValidationMessage=checkValidationMessage(expValidationMessage);

		excelReader.setCellData(xlfile, xlSheetName, 36, 7, actValidationMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcVoucherCancelIcon));
		pdcVoucherCancelIcon.click();

		if (getGridSelectRow1Chkbox==true && actValidationMessage.equalsIgnoreCase(expValidationMessage))
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



	
	
	
	

	public boolean checkConvertedPDCVoucherInReceipts() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receiptsVoucher));
		receiptsVoucher.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
		getAction().doubleClick(grid_ChkBox1).click().build().perform();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocno             = documentNumberTxt.getAttribute("value");
		String actVouDate           = dateTxt.getAttribute("value");
		String actDepartment        = departmentTxt.getAttribute("value");
		String actChequeNo          = receipts_ChequeNoTxt.getAttribute("value");
		String actPDCNo             = voucherHeaderPDCNOTxt.getAttribute("value");
		String actAccount           = select1stRow_1stColumn.getText();
		String actAmount            = select1stRow_2ndColumn.getText();
		String actref               = select1stRow_3rdColumn.getText();

		String expDocno              = excelReader.getCellData(xlSheetName, 37, 6);
		excelReader.setCellData(xlfile, xlSheetName, 37, 7, actDocno);
		
		String expCurrency           = excelReader.getCellData(xlSheetName, 38, 6);
		excelReader.setCellData(xlfile, xlSheetName, 38, 7, expCurrency);
		
		String expDepartment         = excelReader.getCellData(xlSheetName, 39, 6);
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actDepartment);
		
		String expPDCNo              = excelReader.getCellData(xlSheetName, 40, 6);
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, actPDCNo);
		
		String expChequeNo           = excelReader.getCellData(xlSheetName, 41, 6);
		excelReader.setCellData(xlfile, xlSheetName, 41, 7, actChequeNo);
		
		String expAccount            = excelReader.getCellData(xlSheetName, 42, 6);
		excelReader.setCellData(xlfile, xlSheetName, 42, 7, actAccount);
		
		String expAmount             = excelReader.getCellData(xlSheetName, 43, 6);
		excelReader.setCellData(xlfile, xlSheetName, 43, 7, actAmount);
		
		String expref                = excelReader.getCellData(xlSheetName, 44, 6);
		excelReader.setCellData(xlfile, xlSheetName, 44, 7, actref);
		
		String actFooterAmt   = recepitsFooterAmt.getText();
		String expFooterAmt   = excelReader.getCellData(xlSheetName, 45, 6);
		excelReader.setCellData(xlfile, xlSheetName, 45, 7, actFooterAmt);

		System.out.println("Entry Page Document Number    "+actDocno       +"  value Expected  "+expDocno);
		System.out.println("Entry Page ChequeNoTxt        "+actChequeNo    +"  value Expected  "+expChequeNo);
		System.out.println("Entry Page Department         "+actDepartment  +"  value Expected  "+expDepartment);
		System.out.println("Entry Page Account            "+actAccount     +"  value Expected  "+expAccount);
		System.out.println("Entry Page Amount             "+actAmount      +"  value Expected  "+expAmount);
		System.out.println("Entry Page Reference          "+actref         +"  value Expected  "+expref);
		System.out.println("Entry Page  PDC Number        "+actPDCNo       +"  value Expected  "+expPDCNo);
		System.out.println("Entry Page Footer  Amount     "+actFooterAmt   +"  Value Expected  "+expFooterAmt);

		if (actDocno.equalsIgnoreCase(expDocno) && actDepartment.equalsIgnoreCase(expDepartment) && 
				actAccount.equalsIgnoreCase(expAccount) && actChequeNo.equalsIgnoreCase(expChequeNo) && actref.startsWith(expref) && actAmount.equalsIgnoreCase(expAmount) && 
				actFooterAmt.equalsIgnoreCase(expFooterAmt) && actPDCNo.equalsIgnoreCase(expPDCNo) )
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




	public boolean checkDeleteOfConvertedVoucherinRecepits() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();

		getAlert().accept();

		String expValidationMessage=excelReader.getCellData(xlSheetName, 46, 6);

		String actValidationMessage=checkValidationMessage(expValidationMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 46, 7, actValidationMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();

		if (actValidationMessage.equalsIgnoreCase(expValidationMessage)) 
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

	
	
	
	
	

	public boolean checkClickOnConvertedPDCVouchersAfterDeletionOfConvertedVoucherInRecepits() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsAutoPostingsMenu));
		financialsTransactionsAutoPostingsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCs));
		convertMaturedPDCs.click();
		
		Thread.sleep(2000);

		int pdcDOClistCount=pdcDOClist.size();

		for (int i = 0; i < pdcDOClistCount; i++) 
		{
			String data=pdcDOClist.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 47, 5)))
			{
				pdcCheckBoxlist.get(i).click();
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCsOkIcon));
		convertMaturedPDCsOkIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcGridRow1Col3));
		String actRow1Number=pdcGridRow1Col3.getText();
		String expRow1Number=excelReader.getCellData(xlSheetName, 47, 6);
		excelReader.setCellData(xlfile, xlSheetName, 47, 7, actRow1Number);

		if (actRow1Number.equalsIgnoreCase(expRow1Number))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcGridRow1Chkbox));
			pdcGridRow1Chkbox.click();

			boolean getGridSelectRow1Chkbox=pdcGridRow1Chkbox.isSelected();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcVoucherOkIcon));
			pdcVoucherOkIcon.click();

			checkValidationMessage("");
			
			excelReader.setCellData(xlfile, xlSheetName, 47, 8, resPass);

			return true;
		}
		else
		{
			System.err.println(" Test Fail: Deleted Voucher in Receipts  is Not   Appered in PDC Voucher Screen ");
			excelReader.setCellData(xlfile, xlSheetName, 47, 8, resFail);
			return false;
		}

	}


	  

	


	public boolean checkUpdateinPDCUnderSettingMenuForEnableReversePostingInRecepits() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(preferencesBtn));
		preferencesBtn.click();
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTxt));
		searchMenuTxt.click();
		searchMenuTxt.sendKeys("Configure Transactions");
		searchMenuTxt.sendKeys(Keys.ENTER);
		
	*/
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDCMenu));
		PDCMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcReseverseAccountpostingOnPDCConversionChkbox));
		if (pdcReseverseAccountpostingOnPDCConversionChkboxIsSelected.isSelected()==false) 
		{
			pdcReseverseAccountpostingOnPDCConversionChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		getWaitForAlert();

		String actAlert = getAlert().getText();
		String expAlert= excelReader.getCellData(xlSheetName, 48, 6);
		excelReader.setCellData(xlfile, xlSheetName, 48, 7, actAlert);

		getAlert().accept();

		String expValidationMessage=excelReader.getCellData(xlSheetName, 49, 6);

		String actValidationMessage=checkValidationMessage(expValidationMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 49, 7, actValidationMessage);

		System.out.println(" Alert Present     :  "+actAlert          +" Value Expected  :  "+expAlert);

		boolean actpdcPostDatedChequeChkbox                         = postDatedChequeChkboxIsSelected.isSelected();
		boolean actpdcReseverseAccountpostingOnPDCConversionChkbox  = pdcReseverseAccountpostingOnPDCConversionChkboxIsSelected.isSelected();
		boolean actpdcAutoCheckforMaturedChkbox                     = pdcAutoCheckforMaturedChkbox.isDisplayed();
		boolean actpdcDisplayinLedgerChkbox                         = displayLedgerAndBalanceChkbox.isDisplayed();
		boolean actpdcCheckDiscountingTxt                           = pdcCheckDiscountingTxt.isDisplayed();
		boolean actpdcChequeDiscountingVoucherType                  = pdcChequeDiscountingVoucherType.isDisplayed();

		boolean exppdcPostDatedChequeChkbox                         = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 50, 6));
		excelReader.setCellData(xlfile, xlSheetName, 50, 7, Boolean.toString(actpdcPostDatedChequeChkbox));
		
		boolean exppdcReseverseAccountpostingOnPDCConversionChkbox  = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 51, 6));
		excelReader.setCellData(xlfile, xlSheetName, 51, 7, Boolean.toString(actpdcReseverseAccountpostingOnPDCConversionChkbox));
		
		boolean exppdcAutoCheckforMaturedChkbox                     = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 52, 6));
		excelReader.setCellData(xlfile, xlSheetName, 52, 7, Boolean.toString(actpdcAutoCheckforMaturedChkbox));
		
		boolean exppdcDisplayinLedgerChkbox                         = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 53, 6));
		excelReader.setCellData(xlfile, xlSheetName, 53, 7, Boolean.toString(actpdcDisplayinLedgerChkbox));
		
		boolean exppdcCheckDiscountingTxt                           = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 54, 6));
		excelReader.setCellData(xlfile, xlSheetName, 54, 7, Boolean.toString(actpdcCheckDiscountingTxt));
		
		boolean exppdcChequeDiscountingVoucherType                  = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 55, 6));
		excelReader.setCellData(xlfile, xlSheetName, 55, 7, Boolean.toString(actpdcChequeDiscountingVoucherType));
		
		System.out.println("pdcPostDatedChequeChkbox                        : "+actpdcPostDatedChequeChkbox                         +"  "+exppdcPostDatedChequeChkbox);
		System.out.println("pdcReseverseAccountpostingOnPDCConversionChkbox : "+actpdcReseverseAccountpostingOnPDCConversionChkbox  +"  "+exppdcReseverseAccountpostingOnPDCConversionChkbox);
		System.out.println("pdcAutoCheckforMaturedChkbox                    : "+actpdcAutoCheckforMaturedChkbox                     +"  "+exppdcAutoCheckforMaturedChkbox);
		System.out.println("pdcDisplayinLedgerChkbox                        : "+actpdcDisplayinLedgerChkbox                         +"  "+exppdcDisplayinLedgerChkbox);
		System.out.println("pdcCheckDiscountingTxt                          : "+actpdcCheckDiscountingTxt                           +"  "+exppdcCheckDiscountingTxt);
		System.out.println("pdcChequeDiscountingVoucherType                 : "+actpdcChequeDiscountingVoucherType                  +"  "+exppdcChequeDiscountingVoucherType);

		if (actpdcPostDatedChequeChkbox==exppdcPostDatedChequeChkbox && actpdcReseverseAccountpostingOnPDCConversionChkbox==exppdcReseverseAccountpostingOnPDCConversionChkbox &&
				actpdcAutoCheckforMaturedChkbox==exppdcAutoCheckforMaturedChkbox && actpdcDisplayinLedgerChkbox==exppdcDisplayinLedgerChkbox && actpdcCheckDiscountingTxt==exppdcCheckDiscountingTxt &&
				actpdcChequeDiscountingVoucherType==exppdcChequeDiscountingVoucherType && actValidationMessage.equalsIgnoreCase(expValidationMessage) && actAlert.equalsIgnoreCase(expAlert)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 48, 8, resPass);
			return true;
		} 
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 48, 8, resFail);
			return false;
		}
	}
	
	
	@FindBy(xpath="//*[@id='ddlCompany']")
	private static WebElement companyList;

	

	public static boolean checkLogoutAndLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(downArrow));
		downArrow.click();
		Thread.sleep(2000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(3000);

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt="su";

		String pawslt="su";
		
		String compname="Automation Company";

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);

		lp.clickOnSignInBtn();
		
		Thread.sleep(2000);
		
		Select s1=new Select(companyList);
        s1.selectByValue("36");
	
		
		reLogin(unamelt, pawslt, compname);

		//checkRefershPopOnlogin();

		//checkPopUpWindow();

		Thread.sleep(6000);

		String actUserInfo=userNameDisplay.getText();

		System.out.println("User Info  : "+actUserInfo);

		System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String getLoginCompanyName=getCompanyTxt.substring(0, 19);
		System.out.println("company name  :  "+ getLoginCompanyName);
		companyLogo.click();


		String expUserInfo           ="su";
		
		String expLoginCompanyName   ="Automation Company";
		;
		String actDashboard = Boolean.toString(dashboard.isDisplayed());
		String expDashboard = "true";

		System.out.println("**********************************************checkLogoutAndLoginWithSU*********************************************");

		System.out.println("User Info               : "+actUserInfo         +" Value Expected "+expUserInfo);
		System.out.println("Login Company           : "+getLoginCompanyName +" Value Expected "+expLoginCompanyName);
		System.out.println("Dashboard is Displaying : "+actDashboard        +" Value Expected "+expDashboard );

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName) && actDashboard.equalsIgnoreCase(expDashboard)) 
		{
			
			return true;
		}
		else
		{
			
			return false;
		}
	}

	@FindBy(xpath = "//*[@id='errmsgDiv']")
    public static WebElement loginPageErrorMesg;
	
	 public static void reLogin(String unamelt, String pawslt,String compname)
	    {
	    	LoginPage lp=new LoginPage(getDriver());
	            try {
	                    
	                    if(loginPageErrorMesg.isDisplayed())
	                    {
	                            
	                            Thread.sleep(1999);
	                            username.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	                            
	                            lp.enterUserName(unamelt);

	                            Thread.sleep(2000);
	                            
	                            password.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);

	                            lp.enterPassword(pawslt);
	                            
	                            Select oSelect = new Select(companyDropDownList);
	                            oSelect.selectByValue("36");

	                           /* List <WebElement> elementCount = oSelect.getOptions();

	                            int cqSize = elementCount.size();

	                            System.out.println("CompanyDropdownList Count :"+cqSize);

	                            int i;

	                            for(i=0; i<elementCount.size(); i++) 
	                            {

	                                    elementCount.get(i).getText();

	                                    String optionName = elementCount.get(i).getText();
	                                    if(optionName.toUpperCase().startsWith(compname.toUpperCase()))
	                                    {
	                                            System.out.println("q"+elementCount.get(i).getText());
	                                            elementCount.get(i).click();
	                                    }

	                            }*/

	                           lp.clickOnSignInBtn();
	                            
	                            
	                    }
	                    
	            } catch (Exception e) {
	                    // TODO: handle exception
	            }
	            
	            
	    }


	



	public boolean checkCreationOfNewCashAndBankAccountType() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDRVAT));
		PDRVAT.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("");

		cashBankAccountTxt.click();
		cashBankAccountTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		cashBankAccountTxt.sendKeys(Keys.SPACE);
		cashBankAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 56, 5));
		Thread.sleep(2000);
		cashBankAccountTxt.sendKeys(Keys.TAB);

		Select accountType=new Select(accountTypeDropdown);
		accountType.selectByIndex(1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(popSaveBtn));
		popSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashBankAccountTxt));
		cashBankAccountTxt.click();
		cashBankAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		cashBankAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 57, 5));
		Thread.sleep(2000);
		cashBankAccountTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 58, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDRVATPlaceOfSupplyTXt));
		PDRVATPlaceOfSupplyTXt.click();
		PDRVATPlaceOfSupplyTXt.sendKeys(Keys.END);
		PDRVATPlaceOfSupplyTXt.sendKeys(Keys.SHIFT,Keys.HOME);
		PDRVATPlaceOfSupplyTXt.sendKeys(excelReader.getCellData(xlSheetName, 59, 5));
		Thread.sleep(2000);
		PDRVATPlaceOfSupplyTXt.sendKeys(Keys.TAB);

		PDRVAT_JuridictionTxt.sendKeys(Keys.END);
		PDRVAT_JuridictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		PDRVAT_JuridictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 60, 5));
		Thread.sleep(2000);
		PDRVAT_JuridictionTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDRVATChequeNoTxt));
		PDRVATChequeNoTxt.click();
		PDRVATChequeNoTxt.sendKeys(excelReader.getCellData(xlSheetName, 61, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		enter_AccountTxt.click();
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 62, 5));
		Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);

		enterTaxcode.click();
		enterTaxcode.sendKeys(Keys.SHIFT,Keys.HOME);
		enterTaxcode.sendKeys(excelReader.getCellData(xlSheetName, 63, 5));
		Thread.sleep(2000);

		enterTaxcode.sendKeys(Keys.TAB);

		enter_Amount.click();
		enter_Amount.clear();
		enter_Amount.sendKeys(excelReader.getCellData(xlSheetName, 64, 5));
		enter_Amount.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPartyName));
		String actPartyName = billRefPartyName.getText();
		String expPartyName = excelReader.getCellData(xlSheetName, 56, 6);
		excelReader.setCellData(xlfile, xlSheetName, 56, 7, actPartyName);

		System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		String actbillRefNewReferenceTxt=billRefNewReferenceTxt.getAttribute("value");
		String expbillRefNewReferenceTxt=excelReader.getCellData(xlSheetName, 57, 6);
		excelReader.setCellData(xlfile, xlSheetName, 57, 7, actbillRefNewReferenceTxt);
		
		System.err.println(" billRefNewReferenceTxt : "+actbillRefNewReferenceTxt +" Value Expected : "+expbillRefNewReferenceTxt);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 58, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 59, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 58, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		
String expMessage1 = excelReader.getCellData(xlSheetName, 58, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 59, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 58, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		if(actbillRefNewReferenceTxt.equalsIgnoreCase(expbillRefNewReferenceTxt) && 
				actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 56, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 56, 8, resFail);
			return false;
		}
	}
	
	
	


	public boolean checkSavingVoucherInPostDatedReceiptsWithCustomerC() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashBankAccountTxt));
		previousBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));
		toggleBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(copytoClipboardBtn));
		copytoClipboardBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nextBtn));
		nextBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));
		toggleBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pastefromClipboardBtn));
		pastefromClipboardBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		select1stRow_4thColumn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPartyName));
		String actPartyName = billRefPartyName.getText();
		String expPartyName = excelReader.getCellData(xlSheetName, 65, 6);
		excelReader.setCellData(xlfile, xlSheetName, 65, 7, actPartyName);

		System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		String actbillRefNewReferenceTxt=billRefNewReferenceTxt.getAttribute("value");
		String expbillRefNewReferenceTxt=excelReader.getCellData(xlSheetName, 66, 6);
		excelReader.setCellData(xlfile, xlSheetName, 66, 7, actbillRefNewReferenceTxt);

		System.err.println(" billRefNewReferenceTxt : "+actbillRefNewReferenceTxt +" Value Expected : "+expbillRefNewReferenceTxt);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 67, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 68, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 67, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
String expMessage1 = excelReader.getCellData(xlSheetName, 67, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 68, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 67, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		if(actbillRefNewReferenceTxt.equalsIgnoreCase(expbillRefNewReferenceTxt) && 
				actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 65, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 65, 8, resFail);
			return false;
		}
	}


	


	public boolean checkDeleteOptionInPDRVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));	
		previousBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();

		getAlert().accept();

		String expValidationMessage=excelReader.getCellData(xlSheetName, 69, 6);

		String actValidationMessage=checkValidationMessage(expValidationMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 69, 7, actValidationMessage);

		if (actValidationMessage.equalsIgnoreCase(expValidationMessage)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 69, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 69, 8, resFail);
			return false;
		}
	}








	public boolean checkCovertingVoucherInPDCWithReversePosting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(autoPostingMenu));
		autoPostingMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCSMenu));
		convertMaturedPDCSMenu.click();
		
		Thread.sleep(2000);

		int pdcDOClistCount=pdcDOClist.size();

		for (int i = 0; i < pdcDOClistCount; i++) 
		{
			String data=pdcDOClist.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 70, 5)))
			{
				pdcCheckBoxlist.get(i).click();
				
				pdcBanklist.get(i).click();

				gridEnterBank.click();

				gridEnterBank.sendKeys(excelReader.getCellData(xlSheetName, 71, 5));
			}
		}

		gridEnterBank.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCsOkIcon));
		convertMaturedPDCsOkIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcGridRow1Chkbox));
		pdcGridRow1Chkbox.click();

		boolean getGridSelectRow1Chkbox=pdcGridRow1Chkbox.isSelected();

		String actRow1NumberBeforeConverting=pdcGridRow1Col3.getText();
		String expRow1NumberBeforeConverting=excelReader.getCellData(xlSheetName, 70, 6);
		excelReader.setCellData(xlfile, xlSheetName, 70, 7, actRow1NumberBeforeConverting);

		String actNameCol=pdcGridRow1Col5.getText();
		String expNameCol=excelReader.getCellData(xlSheetName, 71, 6);
		excelReader.setCellData(xlfile, xlSheetName, 71, 7, actNameCol);

		System.out.println("actNameCol   :"+actNameCol +" Value Expected  : "+expNameCol);

		String actBankNameCol=pdcGridRow1Col6.getText();
		String expBankNameCol=excelReader.getCellData(xlSheetName, 72, 6);
		excelReader.setCellData(xlfile, xlSheetName, 72, 7, actBankNameCol);

		System.out.println("actBankNameCol   :"+actBankNameCol +" Value Expected  : "+expBankNameCol);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcVoucherOkIcon));
		pdcVoucherOkIcon.click();

		String expValidationMessage=excelReader.getCellData(xlSheetName, 73, 6);

		String actValidationMessage=checkValidationMessage(expValidationMessage);

		excelReader.setCellData(xlfile, xlSheetName, 73, 7, actValidationMessage);
		
		System.out.println(" Validation Message   : "+actValidationMessage+"  Value Expected : "+expValidationMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcVoucherFilterBtn));
		pdcVoucherFilterBtn.click();

		System.out.println("BeforeConvertingMaturityRow1Date----->" +actRow1NumberBeforeConverting     +"Value Expected :"+expRow1NumberBeforeConverting);

		if (getGridSelectRow1Chkbox==true && actValidationMessage.equalsIgnoreCase(expValidationMessage) &&
				actNameCol.equalsIgnoreCase(expNameCol) && actBankNameCol.equalsIgnoreCase(expBankNameCol))
		{
			excelReader.setCellData(xlfile, xlSheetName, 70, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 70, 8, resFail);
			return false;
		}
	}



	
	

	public boolean checkSavedCovertedVoucherInReceiptsAfterReversePosting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receiptsVoucher));
		receiptsVoucher.click();
		
		Thread.sleep(2000);

		int count = grid_VoucherNoList.size();

		for (int i = 0; i < count; i++) 
		{
			String voucherNo = grid_VoucherNoList.get(i).getText();

			if (voucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 74, 5))) 
			{
				getAction().doubleClick(grid_VoucherNoList.get(i)).build().perform();
				break;
			}
		}

		checkValidationMessage("Voucher loaded Successfully");
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocno             = documentNumberTxt.getAttribute("value");
		String actVouDate           = dateTxt.getAttribute("value");
		String actDepartment        = departmentTxt.getAttribute("value");
		String actChequeNo          = receipts_ChequeNoTxt.getAttribute("value");
		String actPDCNo             = voucherHeaderPDCNOTxt.getAttribute("value");
		String actAccount           = select1stRow_1stColumn.getText();
		String actAmount            = select1stRow_2ndColumn.getText();

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		Date date=new Date();
		String expDate=df.format(date);
		System.out.println("expDate   :"+expDate);

		String expDocno              =excelReader.getCellData(xlSheetName, 74, 6);
		excelReader.setCellData(xlfile, xlSheetName, 74, 7, actDocno);

		String expCurrency           =excelReader.getCellData(xlSheetName, 75, 6);
		excelReader.setCellData(xlfile, xlSheetName, 75, 7, expCurrency);
		
		String expDepartment         =excelReader.getCellData(xlSheetName, 76, 6);
		excelReader.setCellData(xlfile, xlSheetName, 76, 7, actDepartment);
		
		String expPDCNo              =excelReader.getCellData(xlSheetName, 77, 6);
		excelReader.setCellData(xlfile, xlSheetName, 77, 7, actPDCNo);
		
		String expChequeNo           =excelReader.getCellData(xlSheetName, 78, 6);
		excelReader.setCellData(xlfile, xlSheetName, 78, 7, actChequeNo);
		
		String expAccount            =excelReader.getCellData(xlSheetName, 79, 6);
		excelReader.setCellData(xlfile, xlSheetName, 79, 7, actAccount);
		
		String expAmount             =excelReader.getCellData(xlSheetName, 80, 6);
		excelReader.setCellData(xlfile, xlSheetName, 80, 7, actAmount);
		
		String actFooterAmt   =recepitsFooterAmt.getText();
		String expFooterAmt   =excelReader.getCellData(xlSheetName, 81, 6);
		excelReader.setCellData(xlfile, xlSheetName, 81, 7, actFooterAmt);

		System.out.println("Entry Page Document Number    "+actDocno       +"  value Expected  "+expDocno);
		System.out.println("Entry Page Voucher Date       "+actVouDate     +"  value Expected  "+expDate);
		System.out.println("Entry Page ChequeNoTxt        "+actChequeNo    +"  value Expected  "+expChequeNo);
		System.out.println("Entry Page Department         "+actDepartment  +"  value Expected  "+expDepartment);
		System.out.println("Entry Page Account            "+actAccount     +"  value Expected  "+expAccount);
		System.out.println("Entry Page Amount             "+actAmount      +"  value Expected  "+expAmount);
		System.out.println("Entry Page  PDC Number        "+actPDCNo       +"  value Expected  "+expPDCNo);
		System.out.println("Entry Page Footer  Amount     "+actFooterAmt   +"  Value Expected  "+expFooterAmt);


		if (actDocno.equalsIgnoreCase(expDocno) && actVouDate.equalsIgnoreCase(expDate) &&actDepartment.equalsIgnoreCase(expDepartment) && 
				actAccount.equalsIgnoreCase(expAccount) && actChequeNo.equalsIgnoreCase(expChequeNo) && 
				actAmount.equalsIgnoreCase(expAmount) && 
				actFooterAmt.equalsIgnoreCase(expFooterAmt) && actPDCNo.equalsIgnoreCase(expPDCNo))
		{
			excelReader.setCellData(xlfile, xlSheetName, 74, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 74, 8, resFail);
			return false;
		}
	}






	public boolean checkUpdateinPDCUnderSettingMenuForDisableEnableReversePostingInRecepits() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(preferencesBtn));
		preferencesBtn.click();
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTxt));
		searchMenuTxt.click();
		searchMenuTxt.sendKeys("Configure Transactions");
		searchMenuTxt.sendKeys(Keys.ENTER);
		
	*/
		Thread.sleep(3000);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDCMenu));
		PDCMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postDatedChequeChkbox));
		postDatedChequeChkbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		getWaitForAlert();

		String actAlert = getAlert().getText();
		String expAlert= excelReader.getCellData(xlSheetName, 82, 6);
		excelReader.setCellData(xlfile, xlSheetName, 82, 7, actAlert);
		
		getAlert().accept();

		String expValidationMessage=excelReader.getCellData(xlSheetName, 83, 6);

		String actValidationMessage=checkValidationMessage(expValidationMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 83, 7, actValidationMessage);

		System.out.println(" Alert Present     :  "+actAlert          +" Value Expected  :  "+expAlert);

		if (actValidationMessage.equalsIgnoreCase(expValidationMessage) && actAlert.equalsIgnoreCase(expAlert)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 82, 8, resPass);
			return true;
		} 
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 82, 8, resFail);
			return false;
		}
	}



	




	public boolean checkConvertMaturedPDCScreenForPDPVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(autoPostingMenu));
		autoPostingMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCSMenu));
		convertMaturedPDCSMenu.click();
		
		Thread.sleep(2000);

		int pdcDOClistCount=pdcDOClist.size();

		for (int i = 0; i < pdcDOClistCount; i++) 
		{
			String data=pdcDOClist.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 84, 5)))
			{
				pdcCheckBoxlist.get(i).click();
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCsOkIcon));
		convertMaturedPDCsOkIcon.click();

		/*String expValidationMessage="No PDC's exists";

		String actValidationMessage=checkValidationMessage(expValidationMessage);*/
		
		Thread.sleep(2000);

		boolean actpdcVoucherCancelIcon = pdcVoucherCancelIcon.isDisplayed();
		boolean exppdcVoucherCancelIcon = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 84, 6)); 
		excelReader.setCellData(xlfile, xlSheetName, 84, 7, Boolean.toString(actpdcVoucherCancelIcon));
		
		System.out.println("PDC Voucher Cancel Icon : "+actpdcVoucherCancelIcon+"  Value Expected "+exppdcVoucherCancelIcon);
		
		if(actpdcVoucherCancelIcon==exppdcVoucherCancelIcon /*&& actValidationMessage.equalsIgnoreCase(expValidationMessage)*/)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcVoucherCancelIcon));
			pdcVoucherCancelIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 84, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcVoucherCancelIcon));
			pdcVoucherCancelIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 84, 8, resFail);
			return false;
		}
	}


	
	
	//PostDatedPayments

	public boolean CheckSavingVoucherInPDPVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDPVAT));
		PDPVAT.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCashBankAccountTxt));
		newCashBankAccountTxt.click();
		newCashBankAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 85, 5));
		Thread.sleep(2000);
		newCashBankAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 86, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDRVATPlaceOfSupplyTXt));
		PDRVATPlaceOfSupplyTXt.click();
		PDRVATPlaceOfSupplyTXt.sendKeys(Keys.END);
		PDRVATPlaceOfSupplyTXt.sendKeys(Keys.SHIFT,Keys.HOME);
		PDRVATPlaceOfSupplyTXt.sendKeys(excelReader.getCellData(xlSheetName, 87, 5));
		Thread.sleep(2000);
		PDRVATPlaceOfSupplyTXt.sendKeys(Keys.TAB);

		PDRVAT_JuridictionTxt.sendKeys(Keys.END);
		PDRVAT_JuridictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		PDRVAT_JuridictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 88, 5));
		Thread.sleep(2000);
		PDRVAT_JuridictionTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvpVATChequeTXT));
		pvpVATChequeTXT.click();
		pvpVATChequeTXT.sendKeys(excelReader.getCellData(xlSheetName, 89, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		enter_AccountTxt.click();
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 90, 5));
		Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enterPVPVATTaxcode));
		enterPVPVATTaxcode.click();
		enterPVPVATTaxcode.sendKeys(Keys.SHIFT,Keys.HOME);
		enterPVPVATTaxcode.sendKeys(Keys.BACK_SPACE);
		enterPVPVATTaxcode.sendKeys(excelReader.getCellData(xlSheetName, 91, 5));

		/* int taxcodelistCount=PDPVATtaxcodelist.size();

			System.err.println(taxcodelistCount);

			for(int i=0 ; i < taxcodelistCount ;i++)
			{
				String data=PDPVATtaxcodelist.get(i).getText();

				if(data.equalsIgnoreCase("Std Rate"))
				{
					PDPVATtaxcodelist.get(i).click();


					break;
				}
			}*/
		
		Thread.sleep(2000);
		
		enterPVPVATTaxcode.sendKeys(Keys.TAB);

		enter_Amount.click();
		enter_Amount.clear();
		enter_Amount.sendKeys(excelReader.getCellData(xlSheetName, 92, 5));
		enter_Amount.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPartyName));
		String actPartyName = billRefPartyName.getText();
		String expPartyName = excelReader.getCellData(xlSheetName, 85, 6);
		excelReader.setCellData(xlfile, xlSheetName, 85, 6, actPartyName);

		System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);

		int Adjustbills=billRefAdjustBillsGridList.size();

		String actAdjustbills=Integer.toString(Adjustbills);

		String expAdjustbills=excelReader.getCellData(xlSheetName, 86, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 86, 7, actAdjustbills);

		int billwiseAdjustBillsDocListcount=billwiseAdjustBillsDocList.size();

		HashSet<String> actbillwiseAdjustBillsDocList  = new HashSet<String>();

		for (int i = 0; i < billwiseAdjustBillsDocListcount; i++) 
		{
			String data	= billwiseAdjustBillsDocList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 93, 5)))
			{
				billwiseAdjustBillsDocList.get(i).click();
			}
		}


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 87, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 88, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 87, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
String expMessage1 = excelReader.getCellData(xlSheetName, 87, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 88, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 87, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))


		if(actAdjustbills.equalsIgnoreCase(expAdjustbills) && actPartyName.equalsIgnoreCase(expPartyName) && 
				actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 85, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 85, 8, resFail);
			return false;
		}
	}

	





	public boolean checkSavingNewVoucherInPDPVATOnFutureDateWithCurrencyAEDAndDepartmentINDIA() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dateTxt));
		dateTxt.click();
		dateTxt.click();

		dateTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		dateTxt.sendKeys(Keys.BACK_SPACE);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		Date date=new Date();

		String docdate = df.format(date);  
		Calendar c = Calendar.getInstance();
		c.add(Calendar.DAY_OF_MONTH, 21);  

		String FilterDate = df.format(c.getTime());  

		System.out.println("FilterDate  : "+FilterDate);

		dateTxt.sendKeys(FilterDate);

		dateTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCashBankAccountTxt));
		newCashBankAccountTxt.click();
		newCashBankAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 94, 5));
		Thread.sleep(2000);
		newCashBankAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.click();
		voucherHeaderCurrency.sendKeys(Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 95, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 96, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDRVATPlaceOfSupplyTXt));
		PDRVATPlaceOfSupplyTXt.click();
		PDRVATPlaceOfSupplyTXt.sendKeys(Keys.END);
		PDRVATPlaceOfSupplyTXt.sendKeys(Keys.SHIFT,Keys.HOME);
		PDRVATPlaceOfSupplyTXt.sendKeys(excelReader.getCellData(xlSheetName, 97, 5));
		Thread.sleep(2000);
		PDRVATPlaceOfSupplyTXt.sendKeys(Keys.TAB);

		PDRVAT_JuridictionTxt.sendKeys(Keys.END);
		PDRVAT_JuridictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);

		PDRVAT_JuridictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 98, 5));
		Thread.sleep(2000);
		PDRVAT_JuridictionTxt.sendKeys(Keys.TAB);

		pvpVATChequeTXT.click();
		pvpVATChequeTXT.sendKeys(excelReader.getCellData(xlSheetName, 99, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		enter_AccountTxt.click();
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 100, 5));
		Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);

		enterPVPVATTaxcode.click();
		enterPVPVATTaxcode.sendKeys(Keys.SHIFT,Keys.HOME);
		enterPVPVATTaxcode.sendKeys(excelReader.getCellData(xlSheetName, 101, 5));
		Thread.sleep(2000);
		enterPVPVATTaxcode.sendKeys(Keys.TAB);

		enter_Amount.click();
		enter_Amount.clear();
		enter_Amount.sendKeys(excelReader.getCellData(xlSheetName, 102, 5));
		enter_Amount.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPartyName));
		String actPartyName = billRefPartyName.getText();
		String expPartyName = excelReader.getCellData(xlSheetName, 94, 6);
		excelReader.setCellData(xlfile, xlSheetName, 94, 7, actPartyName);

		System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);

		int Adjustbills=billRefAdjustBillsGridList.size();

		String actAdjustbills=Integer.toString(Adjustbills);

		String expAdjustbills=excelReader.getCellData(xlSheetName, 95, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 95, 7, actAdjustbills);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		String actBillNewReference                           = billRefNewReferenceTxt.getAttribute("value");
		String actBillTransactionCurrency                    = billRefTransactionCurency.getText();
		String actBillBaseCurrency                           = billRefBaseCurrency.getText();
		String actBillLocalCurrency                          = localCurrencyDhs.getText();
		String actBillBalanceNewRefAmount                    = balanceNewReferenceAmt.getText();
		String actbillRefAdjustAmountInTransCurency          = billRefAdjustAmountInTransCurency.getText();
		String actbillRefBalanceAmountAdjustInTrnasCurrency  = billRefBalanceAmountAdjustInTrnasCurrency.getText();
		
		String expBillNewReference                      =excelReader.getCellData(xlSheetName, 96, 6);
		excelReader.setCellData(xlfile, xlSheetName, 96, 7, actBillNewReference);
		
		String expBillTransactionCurrency               =excelReader.getCellData(xlSheetName, 97, 6);
		excelReader.setCellData(xlfile, xlSheetName, 97, 7, actBillTransactionCurrency);
		
		String expBillBaseCurrency                      =excelReader.getCellData(xlSheetName, 98, 6);
		excelReader.setCellData(xlfile, xlSheetName, 98, 7, actBillBaseCurrency);
		
		String expBillLocalCurrency                     =excelReader.getCellData(xlSheetName, 99, 6);
		excelReader.setCellData(xlfile, xlSheetName, 99, 7, actBillLocalCurrency);

		String expBillBalanceNewRefAmount               =excelReader.getCellData(xlSheetName, 100, 6);
		excelReader.setCellData(xlfile, xlSheetName, 100, 7, actBillBalanceNewRefAmount);
		
		String expbillRefAdjustAmountInTransCurency         =excelReader.getCellData(xlSheetName, 101, 6);
		excelReader.setCellData(xlfile, xlSheetName, 101, 7, actbillRefAdjustAmountInTransCurency);
		
		String expbillRefBalanceAmountAdjustInTrnasCurrency =excelReader.getCellData(xlSheetName, 102, 6);
		excelReader.setCellData(xlfile, xlSheetName, 102, 7, actbillRefBalanceAmountAdjustInTrnasCurrency);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridOrginalAmtRow1));
		String actgridOrginalAmtRow1     = gridOrginalAmtRow1.getText();
		String actgridBalanceAmtRow1     = gridBalanceAmtRow1.getText();
		String actgridAdjustmentAmtRow1  = gridAdjustmentAmtRow1.getText();

		String expgridOrginalAmtRow1      =excelReader.getCellData(xlSheetName, 103, 6);
		excelReader.setCellData(xlfile, xlSheetName, 103, 7, actgridOrginalAmtRow1);
		
		String expgridBalanceAmtRow1      =excelReader.getCellData(xlSheetName, 104, 6);
		excelReader.setCellData(xlfile, xlSheetName, 104, 7, actgridBalanceAmtRow1);
		
		String expgridAdjustmentAmtRow1   =excelReader.getCellData(xlSheetName, 105, 6);
		excelReader.setCellData(xlfile, xlSheetName, 105, 7, actgridAdjustmentAmtRow1);


		System.out.println("*********************************************************************************************************");

		System.out.println("Bill reference Adjustment Bills  :"+ actAdjustbills+"                          "+ "expDate :"+expAdjustbills);
		System.out.println("actBillNewReference :             "+ actBillNewReference +"                    "+ "expBillNewReference :"+expBillNewReference);
		System.out.println("actBillTransactionCurrency       :"+ actBillTransactionCurrency + "            " + "expBillTransactionCurrency :"+expBillTransactionCurrency);
		System.out.println("actBillBaseCurrency :             "+ actBillBaseCurrency + "                   " +"expBillBaseCurrency :"+expBillBaseCurrency);
		System.out.println("actBillLocalCurrency :            "+ actBillLocalCurrency+ "                   " +"expBillLocalCurrency :"+expBillLocalCurrency);
		System.out.println("actBillBalanceNewRefAmount :      "+ actBillBalanceNewRefAmount + "            " + "expBillBalanceNewRefAmount :"+expBillBalanceNewRefAmount);

		System.out.println("actbillRefAdjustAmountInTransCurency :"+ actbillRefAdjustAmountInTransCurency+ "       " +"expbillRefAdjustAmountInTransCurency :"+expbillRefAdjustAmountInTransCurency);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrency :"+ actbillRefBalanceAmountAdjustInTrnasCurrency + "       " +"expbillRefBalanceAmountAdjustInTrnasCurrency :"+expbillRefBalanceAmountAdjustInTrnasCurrency);

		System.out.println("actgridOrginalAmtRow1 :" +actgridOrginalAmtRow1 +"       " +"expgridOrginalAmtRow1 :" +expgridOrginalAmtRow1);


		System.out.println("actgridBalanceAmtRow1 :" +actgridBalanceAmtRow1 +"       " +"expgridBalanceAmtRow1 :" +expgridBalanceAmtRow1);


		System.out.println("actgridAdjustmentAmtRow1 :" +actgridAdjustmentAmtRow1 +"       " +"expgridAdjustmentAmtRow1:" +expgridAdjustmentAmtRow1);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/

		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 106, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 107, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 106, 7, actMessage);
		
		String expMessage = expMsg.toString();
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
String expMessage1 = excelReader.getCellData(xlSheetName, 106, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 107, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 106, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		if (actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2) && actAdjustbills.equalsIgnoreCase(expAdjustbills) && actBillNewReference.equalsIgnoreCase(expBillNewReference) && actBillTransactionCurrency.equalsIgnoreCase(expBillTransactionCurrency) && 
				actBillBaseCurrency.equalsIgnoreCase(expBillBaseCurrency) && actBillLocalCurrency.equalsIgnoreCase(expBillLocalCurrency) && 
				actBillBalanceNewRefAmount.equalsIgnoreCase(expBillBalanceNewRefAmount) && 
				actbillRefAdjustAmountInTransCurency.equalsIgnoreCase(expbillRefAdjustAmountInTransCurency) &&
				actbillRefBalanceAmountAdjustInTrnasCurrency.equalsIgnoreCase(expbillRefBalanceAmountAdjustInTrnasCurrency))
		{
			excelReader.setCellData(xlfile, xlSheetName, 94, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 94, 8, resFail);
			return false;
		}
	}
	
	



	public boolean checkSavedVoucherInPDPVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 108, 6);
		excelReader.setCellData(xlfile, xlSheetName, 108, 7, actDocNo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dateTxt));
		dateTxt.click();

		String actDate=dateTxt.getAttribute("value");
		String expDate=maturityDateTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashBankAccountTxt));
		cashBankAccountTxt.click();

		String actAccount=cashBankAccountTxt.getAttribute("value");
		String expAccount=excelReader.getCellData(xlSheetName, 109, 6);
		excelReader.setCellData(xlfile, xlSheetName, 109, 7, actAccount);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.click();

		String actCurrency=voucherHeaderCurrency.getAttribute("value");
		String expCurrency=excelReader.getCellData(xlSheetName, 110, 6);
		excelReader.setCellData(xlfile, xlSheetName, 110, 7, actCurrency);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderExchangeRate));
		voucherHeaderExchangeRate.click();

		String actExchangeRate=voucherHeaderExchangeRate.getAttribute("value");
		String expExchangeRate=excelReader.getCellData(xlSheetName, 111, 6);
		excelReader.setCellData(xlfile, xlSheetName, 111, 7, actExchangeRate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();

		String actdepartment=departmentTxt.getAttribute("value");
		String expdepartment=excelReader.getCellData(xlSheetName, 112, 6);
		excelReader.setCellData(xlfile, xlSheetName, 112, 7, actdepartment);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderLocalExchangeRate));
		voucherHeaderLocalExchangeRate.click();

		String actLocalExchangeRate=voucherHeaderLocalExchangeRate.getAttribute("value");
		String expLocalExchangeRate=excelReader.getCellData(xlSheetName, 113, 6);
		excelReader.setCellData(xlfile, xlSheetName, 114, 7, actLocalExchangeRate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actGridAccount=select1stRow_1stColumn.getText();
		String expGridAccount=excelReader.getCellData(xlSheetName, 114, 6);
		excelReader.setCellData(xlfile, xlSheetName, 114, 7, actGridAccount);

		String actTaxCode=select1stRow_2ndColumn.getText();
		String expTaxCode=excelReader.getCellData(xlSheetName, 115, 6);
		excelReader.setCellData(xlfile, xlSheetName, 115, 7, actTaxCode);

		String actAmount=select1stRow_3rdColumn.getText();
		String expAmount=excelReader.getCellData(xlSheetName, 116, 6);
		excelReader.setCellData(xlfile, xlSheetName, 116, 7, actAmount);

		System.out.println("**************************************checkSavedVoucherInPDPVAT******************");

		System.out.println("DocNo              : "+actDocNo           +" Value Expected  : "+expDocNo);
		System.out.println("Date               : "+actDate            +" Value Expected  : "+expDate);
		System.out.println("Account            : "+actAccount         +" Value Expected  : "+expAccount);
		System.out.println("Currency           :"+actCurrency         +" Value Expected  : "+expCurrency);
		System.out.println("ExchangeRate       :"+actExchangeRate     +" Value Expected  : "+expExchangeRate);
		System.out.println("department         :"+actdepartment       +" Value Expected  : "+expdepartment);
		System.out.println("LocalExchangeRate  :"+actLocalExchangeRate+" Value Expected  :"+expLocalExchangeRate);
		System.out.println("TaxCode           :"+actTaxCode           +" Value Expected  : "+expTaxCode);
		System.out.println("GridAccount        :"+actGridAccount      +" Value Expected  : "+expGridAccount);
		System.out.println("Amount             :"+actAmount           +" Value Expected  : "+expAmount);

		if (actDocNo.equalsIgnoreCase(expDocNo) && actDate.equalsIgnoreCase(expDate) && actAccount.equalsIgnoreCase(expAccount) &&
				actCurrency.equalsIgnoreCase(expCurrency)&& actExchangeRate.equalsIgnoreCase(expExchangeRate)&& 
				actdepartment.equalsIgnoreCase(expdepartment) && actLocalExchangeRate.equalsIgnoreCase(expLocalExchangeRate) && 
				actTaxCode.equalsIgnoreCase(expTaxCode)&&
				actGridAccount.equalsIgnoreCase(expGridAccount) && actAmount.equalsIgnoreCase(expAmount) ) 
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

	
	
	
	
	

	public boolean checkEditingVoucherAndSavingInPDPVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvpVATChequeTXT));
		pvpVATChequeTXT.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actPDRVATChequeNoTxt=pvpVATChequeTXT.getAttribute("value");
		String expPDRVATChequeNoTxt=excelReader.getCellData(xlSheetName, 117, 6);
		excelReader.setCellData(xlfile, xlSheetName, 117, 7, actPDRVATChequeNoTxt);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();
		enter_Amount.sendKeys(Keys.CLEAR);
		enter_Amount.sendKeys(excelReader.getCellData(xlSheetName, 117, 5));
		enter_Amount.sendKeys(Keys.TAB);

		String actAmount=select1stRow_3rdColumn.getText();
		String expAmount=excelReader.getCellData(xlSheetName, 118, 6);
		excelReader.setCellData(xlfile, xlSheetName, 118, 7, actAmount);

		System.out.println("***************************checkEditingVoucherAfterSavingInPDRVAT*************");

		System.out.println("PDRVATChequeNoTxt           :"+actPDRVATChequeNoTxt +" Value Expected  : "+expPDRVATChequeNoTxt);
		System.out.println("Amount                       :"+actAmount           +" Value Expected  : "+expAmount);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 119, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 120, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 119, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
String expMessage1 = excelReader.getCellData(xlSheetName, 119, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 120, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 119, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		if (actAmount.equalsIgnoreCase(expAmount)  && actPDRVATChequeNoTxt.equalsIgnoreCase(expPDRVATChequeNoTxt) && 
				actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 117, 8, resPass);
			return true;
		}
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 117, 8, resFail);
			return false;
		}
	}





	public boolean checkDeleteOptionInPDPVAT() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));	
		previousBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();

		getAlert().accept();

		String expValidationMessage=excelReader.getCellData(xlSheetName, 121, 6);

		String actValidationMessage=checkValidationMessage(expValidationMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 121, 7, actValidationMessage);

		if (actValidationMessage.equalsIgnoreCase(expValidationMessage)) 
		{
			System.out.println(" Test Pass: Converted Voucher Deletion Successfully ");	
			excelReader.setCellData(xlfile, xlSheetName, 121, 8, resPass);
			return true;
		}
		else
		{
			System.out.println(" Test Fail: Converted Voucher Deletion Successfully ");	
			excelReader.setCellData(xlfile, xlSheetName, 121, 8, resFail);
			return false;
		}
	}





	public boolean CheckSavedVoucherInPDPVATInCovertedPDCScreen() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(autoPostingMenu));
		autoPostingMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCSMenu));
		convertMaturedPDCSMenu.click();
		
		Thread.sleep(2000);

		int pdcDOClistCount=pdcDOClist.size();

		for (int i = 0; i < pdcDOClistCount; i++) 
		{
			String data=pdcDOClist.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 122, 5)))
			{
				pdcCheckBoxlist.get(i).click();

			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCsOkIcon));
		convertMaturedPDCsOkIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcGridRow1Chkbox));
		pdcGridRow1Chkbox.click();

		String actNumber=pdcGridRow1Col3.getText();
		String expNumber=excelReader.getCellData(xlSheetName, 122, 6);
		excelReader.setCellData(xlfile, xlSheetName, 122, 7, actNumber);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcGridRow1Col4));
		String actMaturityDate=pdcGridRow1Col4.getText();
		DateFormat df = new SimpleDateFormat("dd-MM-yyyy");
		Date date=new Date();
		String CurrentDate=df.format(date);
		String expMaturityDate=CurrentDate;

		System.out.println("CurrentDate      : "+CurrentDate);
		System.out.println("Maturity Date    : "+expMaturityDate);

		String actName=pdcGridRow1Col5.getText();
		String expName=excelReader.getCellData(xlSheetName, 123, 6);
		excelReader.setCellData(xlfile, xlSheetName, 123, 7, actName);

		String actAmount=pdcGridRow1Col7.getText();
		String expAmount=excelReader.getCellData(xlSheetName, 124, 6);
		excelReader.setCellData(xlfile, xlSheetName, 124, 7, actAmount);

		String actReassignedVouNo=pdcGridRow1Col8.getText();
		String expReassignedVouNo=excelReader.getCellData(xlSheetName, 125, 6);
		excelReader.setCellData(xlfile, xlSheetName, 125, 7, actReassignedVouNo);

		System.out.println("******************CheckSavedVoucherInPDPVATInCovertedPDCScreen********************************");
		System.out.println(" Number         : "+actNumber           +" Value Expected  : "+expNumber);
		System.out.println("Name            : "+actName             +" Value Expected  : "+expName);
		System.out.println("Amount          : "+actAmount           +" Value Expected  : "+expAmount);
		System.out.println("ReassignedVouNo : " +actReassignedVouNo +" Value Expected  : "+expReassignedVouNo);

		if (actNumber.equalsIgnoreCase(expNumber) && actMaturityDate.equalsIgnoreCase(expMaturityDate)&&
				actName.equalsIgnoreCase(expName) &&actAmount.equalsIgnoreCase(expAmount) &&  actReassignedVouNo.equalsIgnoreCase(expReassignedVouNo))
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



	

	public boolean checkConvertingVouchersPDPInPDCVoucherScreenWithOutPostOnDate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcVoucherOkIcon));
		pdcVoucherOkIcon.click();

		String expValidationMessage=excelReader.getCellData(xlSheetName, 126, 6);
		String actValidationMessage=checkValidationMessage(expValidationMessage);
		excelReader.setCellData(xlfile, xlSheetName, 126, 7, actValidationMessage);
		
		if (actValidationMessage.equalsIgnoreCase(expValidationMessage))
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





	//Payments

	public boolean checkAddExtraFieldPDCInHeaderFieldOFPayments() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(paymentsVoucher));
		paymentsVoucher.click();
		
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsBtn));
		settingsBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTab));
		editLayoutTab.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutAddFieldsBtn));
		editLayoutAddFieldsBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCaptionTxt));
		editLayoutCaptionTxt.click();
		editLayoutCaptionTxt.sendKeys(excelReader.getCellData(xlSheetName, 127, 5));
		editLayoutCaptionTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutApplyBtn));
		editLayoutApplyBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
		errorMessageCloseBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingUpdateIcon));
		settingUpdateIcon.click();

		for (int i = 0; i < 5; i++) 
		{
			if (getIsAlertPresent()==true) 
			{
				getAlert().accept();
			}
		}

		String expValidationMessage=excelReader.getCellData(xlSheetName, 127, 6);

		String actValidationMessage=checkValidationMessage(expValidationMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 127, 7, actValidationMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
		settings_closeBtn.click();

		if (actValidationMessage.equalsIgnoreCase(expValidationMessage) )
		{
			excelReader.setCellData(xlfile, xlSheetName, 127, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 127, 8, resFail);
			return false;
		}
	}

	
	


	public boolean checkConvertingPDPVoucherInPDCScreenAfterExtraField() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(autoPostingMenu));
		autoPostingMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCSMenu));
		convertMaturedPDCSMenu.click();
		
		Thread.sleep(2000);

		int pdcDOClistCount=pdcDOClist.size();

		for (int i = 0; i < pdcDOClistCount; i++) 
		{
			String data=pdcDOClist.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 128, 5)))
			{
				pdcCheckBoxlist.get(i).click();
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCsOkIcon));
		convertMaturedPDCsOkIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcGridRow1Chkbox));
		pdcGridRow1Chkbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcVoucherOkIcon));
		pdcVoucherOkIcon.click();

		String expValidationMessage=excelReader.getCellData(xlSheetName, 128, 6);

		String actValidationMessage=checkValidationMessage(expValidationMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 128, 7, actValidationMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcVoucherCancelIcon));
		pdcVoucherCancelIcon.click();

		if ( actValidationMessage.equalsIgnoreCase(expValidationMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 128, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 128, 8, resFail);
			return false;
		}
	}

	
	
	

	public boolean checkConvertedPDCVoucherInPayments() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(paymentsVoucher));
		paymentsVoucher.click();
		
		Thread.sleep(2000);

		int count = grid_VoucherNoList.size();

		for (int i = 0; i < count; i++) 
		{
			String voucherNo = grid_VoucherNoList.get(i).getText();

			if (voucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 129, 5))) 
			{
				getAction().doubleClick(grid_VoucherNoList.get(i)).build().perform();
				break;
			}
		}

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocno             = documentNumberTxt.getAttribute("value");
		String actVouDate           = dateTxt.getAttribute("value");
		String actDepartment        = departmentTxt.getAttribute("value");
		String actChequeNo          = payments_ChequeNoTxt.getAttribute("value");
		String actAccount           = select1stRow_1stColumn.getText();
		String actAmount            = select1stRow_2ndColumn.getText();


		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		Date date1=new Date();

		Calendar c = Calendar.getInstance();
		c.add(Calendar.DAY_OF_MONTH, 21);  
		String expDate = df.format(c.getTime());  

		String expDocno              =excelReader.getCellData(xlSheetName, 129, 6);
		excelReader.setCellData(xlfile, xlSheetName, 129, 7, actDocno);
		
		String expDepartment         =excelReader.getCellData(xlSheetName, 130, 6);
		excelReader.setCellData(xlfile, xlSheetName, 130, 7, actDepartment);
		
		String expPDCNo              =excelReader.getCellData(xlSheetName, 131, 6);
		excelReader.setCellData(xlfile, xlSheetName, 131, 7, expPDCNo);
		
		String expChequeNo           =excelReader.getCellData(xlSheetName, 132, 6);
		excelReader.setCellData(xlfile, xlSheetName, 132, 7, actChequeNo);
		
		String expAccount            =excelReader.getCellData(xlSheetName, 133, 6);
		excelReader.setCellData(xlfile, xlSheetName, 133, 7, actAccount);
		
		String expAmount             =excelReader.getCellData(xlSheetName, 134, 6);
		excelReader.setCellData(xlfile, xlSheetName, 134, 7, actAmount);
		
		String actFooterAmt   =recepitsFooterAmt.getText();
		String expFooterAmt   =excelReader.getCellData(xlSheetName, 135, 6);
		excelReader.setCellData(xlfile, xlSheetName, 135, 7, actFooterAmt);

		System.out.println("Entry Page Document Number    "+actDocno       +"  value Expected  "+expDocno);
		System.out.println("Entry Page Voucher Date       "+actVouDate     +"  value Expected  "+expDate);
		System.out.println("Entry Page ChequeNoTxt        "+actChequeNo    +"  value Expected  "+expChequeNo);
		System.out.println("Entry Page Department         "+actDepartment  +"  value Expected  "+expDepartment);
		System.out.println("Entry Page Account            "+actAccount     +"  value Expected  "+expAccount);
		System.out.println("Entry Page Amount             "+actAmount      +"  value Expected  "+expAmount);
		System.out.println("Entry Page Footer  Amount     "+actFooterAmt   +"  Value Expected  "+expFooterAmt);

		if (actDocno.equalsIgnoreCase(expDocno) /*&& actVouDate.equalsIgnoreCase(expDate)*/ &&actDepartment.equalsIgnoreCase(expDepartment) && 
				actAccount.equalsIgnoreCase(expAccount) && actChequeNo.equalsIgnoreCase(expChequeNo)  && actAmount.equalsIgnoreCase(expAmount) && 
				actFooterAmt.equalsIgnoreCase(expFooterAmt))
		{
			excelReader.setCellData(xlfile, xlSheetName, 129, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 129, 8, resFail);
			return false;
		}
	}





	public boolean checkUpdateinPDCUnderSettingMenuForEnableReversePostingInPayments() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(preferencesBtn));
		preferencesBtn.click();
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTxt));
		searchMenuTxt.click();
		searchMenuTxt.sendKeys("Configure Transactions");
		searchMenuTxt.sendKeys(Keys.ENTER);
		
	*/
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDCMenu));
		PDCMenu.click();
		
		if (postDatedChequeChkboxIsSelected.isSelected()==false)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postDatedChequeChkbox));
			postDatedChequeChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcReseverseAccountpostingOnPDCConversionChkbox));
		pdcReseverseAccountpostingOnPDCConversionChkbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		getWaitForAlert();

		String actAlert = getAlert().getText();
		String expAlert= excelReader.getCellData(xlSheetName, 136, 6);
		excelReader.setCellData(xlfile, xlSheetName, 136, 7, actAlert);
		
		getAlert().accept();

		String expValidationMessage=excelReader.getCellData(xlSheetName, 137, 6);

		String actValidationMessage=checkValidationMessage(expValidationMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 137, 7, actValidationMessage);

		System.out.println("Alert Present     :  "+actAlert        +" Value Expected  :  "+expAlert);

		boolean actpdcPostDatedChequeChkbox                         = postDatedChequeChkboxIsSelected.isSelected();
		boolean actpdcReseverseAccountpostingOnPDCConversionChkbox  = pdcReseverseAccountpostingOnPDCConversionChkbox.isSelected();
		boolean actpdcAutoCheckforMaturedChkbox                     = pdcAutoCheckforMaturedChkbox.isDisplayed();
		boolean actpdcDisplayinLedgerChkbox                         = displayLedgerAndBalanceChkbox.isDisplayed();
		boolean actpdcCheckDiscountingTxt                           = pdcCheckDiscountingTxt.isDisplayed();
		boolean actpdcChequeDiscountingVoucherType                  = pdcChequeDiscountingVoucherType.isDisplayed();

		boolean exppdcPostDatedChequeChkbox                         = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 138, 6));
		excelReader.setCellData(xlfile, xlSheetName, 138, 7, Boolean.toString(actpdcPostDatedChequeChkbox));
		
		boolean exppdcReseverseAccountpostingOnPDCConversionChkbox  = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 139, 6));
		excelReader.setCellData(xlfile, xlSheetName, 139, 7, Boolean.toString(actpdcReseverseAccountpostingOnPDCConversionChkbox));
		
		boolean exppdcAutoCheckforMaturedChkbox                     = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 140, 6));
		excelReader.setCellData(xlfile, xlSheetName, 140, 7, Boolean.toString(actpdcAutoCheckforMaturedChkbox));
		
		boolean exppdcDisplayinLedgerChkbox                         = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 141, 6));
		excelReader.setCellData(xlfile, xlSheetName, 141, 7, Boolean.toString(actpdcDisplayinLedgerChkbox));
		
		boolean exppdcCheckDiscountingTxt                           = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 142, 6));
		excelReader.setCellData(xlfile, xlSheetName, 142, 7, Boolean.toString(actpdcCheckDiscountingTxt));
		
		boolean exppdcChequeDiscountingVoucherType                  = Boolean.parseBoolean(excelReader.getCellData(xlSheetName, 143, 6));
		excelReader.setCellData(xlfile, xlSheetName, 143, 7, Boolean.toString(actpdcChequeDiscountingVoucherType));
		
		System.out.println("pdcPostDatedChequeChkbox                        : "+actpdcPostDatedChequeChkbox                         +"  "+exppdcPostDatedChequeChkbox);
		System.out.println("pdcReseverseAccountpostingOnPDCConversionChkbox : "+actpdcReseverseAccountpostingOnPDCConversionChkbox  +"  "+exppdcReseverseAccountpostingOnPDCConversionChkbox);
		System.out.println("pdcAutoCheckforMaturedChkbox                    : "+actpdcAutoCheckforMaturedChkbox                     +"  "+exppdcAutoCheckforMaturedChkbox);
		System.out.println("pdcDisplayinLedgerChkbox                        : "+actpdcDisplayinLedgerChkbox                         +"  "+exppdcDisplayinLedgerChkbox);
		System.out.println("pdcCheckDiscountingTxt                          : "+actpdcCheckDiscountingTxt                           +"  "+exppdcCheckDiscountingTxt);
		System.out.println("pdcChequeDiscountingVoucherType                 : "+actpdcChequeDiscountingVoucherType                  +"  "+exppdcChequeDiscountingVoucherType);

		if (actpdcPostDatedChequeChkbox==exppdcPostDatedChequeChkbox && actpdcReseverseAccountpostingOnPDCConversionChkbox==exppdcReseverseAccountpostingOnPDCConversionChkbox &&
				actpdcAutoCheckforMaturedChkbox==exppdcAutoCheckforMaturedChkbox && actpdcDisplayinLedgerChkbox==exppdcDisplayinLedgerChkbox && actpdcCheckDiscountingTxt==exppdcCheckDiscountingTxt &&
				actpdcChequeDiscountingVoucherType==exppdcChequeDiscountingVoucherType && actValidationMessage.equalsIgnoreCase(expValidationMessage) && actAlert.equalsIgnoreCase(expAlert)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 136, 8, resPass);
			return true;
		} 
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 136, 8, resFail);
			return false;
		}
	}




	public boolean checkSavingVoucherInPDPVATWithVendorB() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDPVAT));
		PDPVAT.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCashBankAccountTxt));
		newCashBankAccountTxt.click();
		newCashBankAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 144, 5));
		Thread.sleep(2000);
		newCashBankAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 145, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PDRVATPlaceOfSupplyTXt));
		PDRVATPlaceOfSupplyTXt.click();
		PDRVATPlaceOfSupplyTXt.sendKeys(Keys.END);
		PDRVATPlaceOfSupplyTXt.sendKeys(Keys.SHIFT,Keys.HOME);
		PDRVATPlaceOfSupplyTXt.sendKeys(excelReader.getCellData(xlSheetName, 146, 5));
		Thread.sleep(2000);
		PDRVATPlaceOfSupplyTXt.sendKeys(Keys.TAB);

		PDRVAT_JuridictionTxt.sendKeys(Keys.END);
		PDRVAT_JuridictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		PDRVAT_JuridictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 147, 5));
		Thread.sleep(2000);
		PDRVAT_JuridictionTxt.sendKeys(Keys.TAB);

		pvpVATChequeTXT.click();
		pvpVATChequeTXT.sendKeys(excelReader.getCellData(xlSheetName, 148, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		enter_AccountTxt.click();
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 149, 5));
		Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);

		enterPVPVATTaxcode.click();
		enterPVPVATTaxcode.sendKeys(Keys.SHIFT,Keys.HOME);
		enterPVPVATTaxcode.sendKeys(Keys.SPACE);
		enterPVPVATTaxcode.sendKeys(excelReader.getCellData(xlSheetName, 150, 5));

		Thread.sleep(2000);
		
		enterPVPVATTaxcode.sendKeys(Keys.TAB);

		enter_Amount.click();
		enter_Amount.clear();
		enter_Amount.sendKeys(excelReader.getCellData(xlSheetName, 151, 5));
		enter_Amount.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPartyName));
		String actPartyName = billRefPartyName.getText();
		String expPartyName = excelReader.getCellData(xlSheetName, 144, 6);
		excelReader.setCellData(xlfile, xlSheetName, 144, 7, actPartyName);

		System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);

		int Adjustbills=billRefAdjustBillsGridList.size();

		String actAdjustbills=Integer.toString(Adjustbills);

		String expAdjustbills=excelReader.getCellData(xlSheetName, 145, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 145, 7, actAdjustbills);

		int billwiseAdjustBillsDocListcount=billwiseAdjustBillsDocList.size();

		HashSet<String> actbillwiseAdjustBillsDocList  = new HashSet<String>();

		for (int i = 0; i < billwiseAdjustBillsDocListcount; i++) 
		{
			String data		= billwiseAdjustBillsDocList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 152, 5)))
			{
				billwiseAdjustBillsDocList.get(i).click();
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		String actBillNewReference                           = billRefNewReferenceTxt.getAttribute("value");
		String actBillTransactionCurrency                    = billRefTransactionCurency.getText();
		String actBillBaseCurrency                           = billRefBaseCurrency.getText();
		String actBillLocalCurrency                          = localCurrencyDhs.getText();
		String actBillBalanceNewRefAmount                    = balanceNewReferenceAmt.getText();
		String actbillRefAdjustAmountInTransCurency          = billRefAdjustAmountInTransCurency.getText();
		String actbillRefBalanceAmountAdjustInTrnasCurrency  = billRefBalanceAmountAdjustInTrnasCurrency.getText();
		
		String expBillNewReference                      =excelReader.getCellData(xlSheetName, 146, 6);
		excelReader.setCellData(xlfile, xlSheetName, 146, 7, actBillNewReference);
		
		String expBillTransactionCurrency               =excelReader.getCellData(xlSheetName, 147, 6);
		excelReader.setCellData(xlfile, xlSheetName, 147, 7, actBillTransactionCurrency);
		
		String expBillBaseCurrency                      =excelReader.getCellData(xlSheetName, 148, 6);
		excelReader.setCellData(xlfile, xlSheetName, 148, 7, actBillBaseCurrency);
		
		String expBillLocalCurrency                     =excelReader.getCellData(xlSheetName, 149, 6);
		excelReader.setCellData(xlfile, xlSheetName, 149, 7, actBillLocalCurrency);
		
		String expBillBalanceNewRefAmount               =excelReader.getCellData(xlSheetName, 150, 6);
		excelReader.setCellData(xlfile, xlSheetName, 150, 7, actBillBalanceNewRefAmount);
		
		String expbillRefAdjustAmountInTransCurency         =excelReader.getCellData(xlSheetName, 151, 6);
		excelReader.setCellData(xlfile, xlSheetName, 151, 7, actbillRefAdjustAmountInTransCurency);
		
		String expbillRefBalanceAmountAdjustInTrnasCurrency =excelReader.getCellData(xlSheetName, 152, 6);
		excelReader.setCellData(xlfile, xlSheetName, 152, 7, actbillRefBalanceAmountAdjustInTrnasCurrency);
		
		System.out.println("*********************************************************************************************************");

		System.out.println("Bill reference Adjustment Bills  :"+ actAdjustbills+"                          "+ "expBills :"+expAdjustbills);
		System.out.println("actBillNewReference :             "+ actBillNewReference +"                    "+ "expBillNewReference :"+expBillNewReference);
		System.out.println("actBillTransactionCurrency       :"+ actBillTransactionCurrency + "            " + "expBillTransactionCurrency :"+expBillTransactionCurrency);
		System.out.println("actBillBaseCurrency :             "+ actBillBaseCurrency + "                   " +"expBillBaseCurrency :"+expBillBaseCurrency);
		System.out.println("actBillLocalCurrency :            "+ actBillLocalCurrency+ "                   " +"expBillLocalCurrency :"+expBillLocalCurrency);
		System.out.println("actBillBalanceNewRefAmount :      "+ actBillBalanceNewRefAmount + "            " + "expBillBalanceNewRefAmount :"+expBillBalanceNewRefAmount);

		System.out.println("actbillRefAdjustAmountInTransCurency :"+ actbillRefAdjustAmountInTransCurency+ "       " +"expbillRefAdjustAmountInTransCurency :"+expbillRefAdjustAmountInTransCurency);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrency :"+ actbillRefBalanceAmountAdjustInTrnasCurrency + "       " +"expbillRefBalanceAmountAdjustInTrnasCurrency :"+expbillRefBalanceAmountAdjustInTrnasCurrency);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : "+docno);
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 153, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 154, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 153, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
String expMessage1 = excelReader.getCellData(xlSheetName, 153, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 154, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 153, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		if (actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2) &&  actAdjustbills.equalsIgnoreCase(expAdjustbills)  && actBillNewReference.equalsIgnoreCase(expBillNewReference) && actBillTransactionCurrency.equalsIgnoreCase(expBillTransactionCurrency) && 
				actBillBaseCurrency.equalsIgnoreCase(expBillBaseCurrency) && actBillLocalCurrency.equalsIgnoreCase(expBillLocalCurrency) && 
				actBillBalanceNewRefAmount.equalsIgnoreCase(expBillBalanceNewRefAmount) && 
				actbillRefAdjustAmountInTransCurency.equalsIgnoreCase(expbillRefAdjustAmountInTransCurency) &&
				actbillRefBalanceAmountAdjustInTrnasCurrency.equalsIgnoreCase(expbillRefBalanceAmountAdjustInTrnasCurrency))
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


	
	
	
	
	public boolean checkCovertingPDPVoucherInPDCWithReversePosting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(autoPostingMenu));
		autoPostingMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCSMenu));
		convertMaturedPDCSMenu.click();
		
		Thread.sleep(2000);

		int pdcDOClistCount=pdcDOClist.size();

		for (int i = 0; i < pdcDOClistCount; i++) 
		{

			String data=pdcDOClist.get(i).getText();
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 155, 5)))
			{
				pdcCheckBoxlist.get(i).click();

			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCsOkIcon));
		convertMaturedPDCsOkIcon.click();

		String expValidationMessage=excelReader.getCellData(xlSheetName, 155, 6);

		String actValidationMessage=checkValidationMessage(expValidationMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 155, 7, actValidationMessage);

		for (int i = 0; i < pdcDOClistCount; i++) 
		{
			String data=pdcDOClist.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 156, 5)))
			{
				pdcBanklist.get(i).click();

				gridEnterBank.click();

				gridEnterBank.sendKeys(excelReader.getCellData(xlSheetName, 157, 5));
				
				Thread.sleep(2000);
			}
		}

		gridEnterBank.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertMaturedPDCsOkIcon));
		convertMaturedPDCsOkIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcGridRow1Col3));
		String actRow1Number          = pdcGridRow1Col3.getText();
		String actMaturityRow1Date    = pdcGridRow1Col4.getText();
		String actRow1Name            = pdcGridRow1Col5.getText();
		String actRow1AppliedBank     = pdcGridRow1Col6.getText();
		String actRow1Amount          = pdcGridRow1Col7.getText();
		String actRow1ReassignedVouNo = pdcGridRow1Col8.getText();
		String actRow1ChequeNo        = pdcGridRow1Col9.getText();

		String expRow1Number           =excelReader.getCellData(xlSheetName, 156, 6);
		excelReader.setCellData(xlfile, xlSheetName, 156, 7, actRow1Number);
		
		String expRow1Name             =excelReader.getCellData(xlSheetName, 157, 6);
		excelReader.setCellData(xlfile, xlSheetName, 157, 7, actRow1Name);
		
		String expRow1Amount           =excelReader.getCellData(xlSheetName, 158, 6);
		excelReader.setCellData(xlfile, xlSheetName, 158, 7, actRow1Amount);
		
		String expRow1AppliedBank      =excelReader.getCellData(xlSheetName, 159, 6);
		excelReader.setCellData(xlfile, xlSheetName, 159, 7, actRow1AppliedBank);
		
		String expRow1ReassignedVouNo  =excelReader.getCellData(xlSheetName, 160, 6);
		excelReader.setCellData(xlfile, xlSheetName, 160, 7, actRow1ReassignedVouNo);
		
		String expRow1ChequeNo         =excelReader.getCellData(xlSheetName, 161, 6);
		excelReader.setCellData(xlfile, xlSheetName, 161, 7, actRow1ChequeNo);

		System.out.println("Row1Number----------->" +actRow1Number           +"Value Expected :"+expRow1Number);
		System.out.println("Row1Name------------->" + actRow1Name            +"Value Expected :"+expRow1Name);
		System.out.println("Row1Amount----------->" +actRow1Amount           +"Value Expected :"+expRow1Amount);
		System.out.println("Row1ReassignedVouNo-->" + actRow1ReassignedVouNo +"Value Expected :"+expRow1ReassignedVouNo);
		System.out.println("Row1ChequeNo--------->" + actRow1ChequeNo        +"Value Expected :"+expRow1ChequeNo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcGridRow1Chkbox));
		pdcGridRow1Chkbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pdcVoucherOkIcon));
		pdcVoucherOkIcon.click();

		String expValidationMessage1=excelReader.getCellData(xlSheetName, 162, 6);

		String actValidationMessage1=checkValidationMessage(expValidationMessage1);
		
		excelReader.setCellData(xlfile, xlSheetName, 162, 7, actValidationMessage1);

		if ( actRow1Number.equalsIgnoreCase(expRow1Number)  && actRow1Name.equalsIgnoreCase(expRow1Name) &&
				actRow1Amount.equalsIgnoreCase(expRow1Amount) && actRow1ReassignedVouNo.equalsIgnoreCase(expRow1ReassignedVouNo) && 
				actRow1ChequeNo.equalsIgnoreCase(expRow1ChequeNo)  && actValidationMessage1.equalsIgnoreCase(expValidationMessage1))
		{
			excelReader.setCellData(xlfile, xlSheetName, 155, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 155, 8, resFail);
			return false;
		}
	}




	
	

	
	public boolean checkSavedCovertedVoucherInPaymentsAfterReversePosting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(paymentsVoucher));
		paymentsVoucher.click();
		
		Thread.sleep(2000);

		int count = grid_VoucherNoList.size();

		for (int i = 0; i < count; i++) 
		{
			String voucherNo = grid_VoucherNoList.get(i).getText();

			if (voucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 163, 5))) 
			{
				getAction().doubleClick(grid_VoucherNoList.get(i)).build().perform();
				break;
			}
		}

		checkValidationMessage("Voucher Loading Message");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocno             = documentNumberTxt.getAttribute("value");
		String actVouDate           = dateTxt.getAttribute("value");
		String actDepartment        = departmentTxt.getAttribute("value");
		String actChequeNo          = payments_ChequeNoTxt.getAttribute("value");
		String actAccount           = select1stRow_1stColumn.getText();
		String actAmount            = select1stRow_2ndColumn.getText();

		String expDocno              =excelReader.getCellData(xlSheetName, 163, 6);
		excelReader.setCellData(xlfile, xlSheetName, 163, 7, actDocno);
		
		String expDepartment         =excelReader.getCellData(xlSheetName, 164, 6);
		excelReader.setCellData(xlfile, xlSheetName, 164, 7, actDepartment);
		
		String expChequeNo           =excelReader.getCellData(xlSheetName, 165, 6);
		excelReader.setCellData(xlfile, xlSheetName, 165, 7, actChequeNo);
		
		String expAccount            =excelReader.getCellData(xlSheetName, 166, 6);
		excelReader.setCellData(xlfile, xlSheetName, 166, 7, actAccount);
		
		String expAmount             =excelReader.getCellData(xlSheetName, 167, 6);
		excelReader.setCellData(xlfile, xlSheetName, 167, 7, actAmount);

		String actFooterAmt   =recepitsFooterAmt.getText();
		String expFooterAmt   =excelReader.getCellData(xlSheetName, 168, 6);
		excelReader.setCellData(xlfile, xlSheetName, 168, 7, actFooterAmt);
		
		System.out.println("Entry Page Document Number    "+actDocno       +"  value Expected  "+expDocno);
		System.out.println("Entry Page ChequeNoTxt        "+actChequeNo    +"  value Expected  "+expChequeNo);
		System.out.println("Entry Page Department         "+actDepartment  +"  value Expected  "+expDepartment);
		System.out.println("Entry Page Account            "+actAccount     +"  value Expected  "+expAccount);
		System.out.println("Entry Page Amount             "+actAmount      +"  value Expected  "+expAmount);
		System.out.println("Entry Page Footer  Amount     "+actFooterAmt   +"  Value Expected  "+expFooterAmt);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeCloseBtn));
		homeCloseBtn.click();

		if (actDocno.equalsIgnoreCase(expDocno) &&actDepartment.equalsIgnoreCase(expDepartment) && 
				actAccount.equalsIgnoreCase(expAccount) && actChequeNo.equalsIgnoreCase(expChequeNo) && 
				actAmount.equalsIgnoreCase(expAmount) && 
				actFooterAmt.equalsIgnoreCase(expFooterAmt))
		{
			excelReader.setCellData(xlfile, xlSheetName, 163, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 163, 8, resFail);
			return false;
		}
	}








	public PDCPage(WebDriver driver)
	{
		PageFactory.initElements(driver, this);	
	}



}
