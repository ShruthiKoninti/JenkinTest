package com.focus.Pages;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

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

public class ReceiptsVATVoucherPage extends BaseEngine 
{
	private static String xlfile = getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlSheetName = "ReceiptsVATVoucherPage";
	
	public boolean checkSavingRecepitsVATVocher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATVoucher));
		recepitsVATVoucher.click();

		Thread.sleep(2000);

	//	checkDeleteLinkStatus();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCashBankAccountTxt));
		newCashBankAccountTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
		Thread.sleep(1000);
		newCashBankAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 8, 5));
		Thread.sleep(2000);
		newCashBankAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(Keys.SHIFT,Keys.HOME,Keys.BACK_SPACE);
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeofSupplyTxt));
		placeofSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		Thread.sleep(2000);
		placeofSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enterReceiptsVATTaxCode));	
		enterReceiptsVATTaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Amount));
		enter_Amount.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
		enter_Amount.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPartyName));

		String actPartyName = billRefPartyName.getText();
		String expPartyName = "Customer A (122-001)";

		//System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);

		int Adjustbills=billRefAdjustBillsGridList.size();

		String actAdjustbills=Integer.toString(Adjustbills);

		String expAdjustbills="5";
		

	//	System.err.println("actAdjustbills : "+actAdjustbills +" Value Expected  : "+expAdjustbills);

		
		int billwiseAdjustBillsDocListcount=billwiseAdjustBillsDocList.size();

		HashSet<String> actbillwiseAdjustBillsDocList  = new HashSet<String>();

		for (int i = 0; i < billwiseAdjustBillsDocListcount; i++) 
		{
			String data		                              = billwiseAdjustBillsDocList.get(i).getText();
			
			actbillwiseAdjustBillsDocList.add(data);
		}

		String actDocumentNumberText=actbillwiseAdjustBillsDocList.toString();

		String expDocumentNumberText="[NDT50:1, NDT50:4, NDT50:5, NDT50:6, NDT50:7]";

		System.out.println("actDocumentNumberText   : "+actDocumentNumberText);
		System.out.println("expDocumentNumberText   : "+expDocumentNumberText);

		Thread.sleep(2000);			

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billrefAdjuBills1stChkbox));
		billrefAdjuBills1stChkbox.click();
		
		Thread.sleep(2000);

		System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);

		System.err.println("actAdjustbills : "+actAdjustbills +" Value Expected  : "+expAdjustbills);

		Thread.sleep(2000);

		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("arguments[0].scrollIntoView();", gridAdjustmentAmtRow1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		
		String actBillNewReference                           = billRefNewReferenceTxt.getAttribute("value");
		String actBillTransactionCurrency                    = billRefTransactionCurency.getText();
		String actBillBaseCurrency                           = billRefBaseCurrency.getText();
		String actBillLocalCurrency                          = localCurrencyDhs.getText();
		String actBillBalanceNewRefAmount                    = balanceNewReferenceAmt.getText();
		String actbillRefAdjustAmountInTransCurency          = billRefAdjustAmountInTransCurency.getText();
		String actbillRefBalanceAmountAdjustInTrnasCurrency  = billRefBalanceAmountAdjustInTrnasCurrency.getText();
		
		String expBillNewReference                      =	"0.00";
		
		String expBillTransactionCurrency               =	"50.00";
		
		String expBillBaseCurrency                      = 	"50.00";
		
		
		String expBillLocalCurrency                     =	"3.50";
		
		String expBillBalanceNewRefAmount               =	"0.00";
		
		String expbillRefAdjustAmountInTransCurency     =	"50.00";
		
		String expbillRefBalanceAmountAdjustInTrnasCurrency =	"0.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefGridFirstRowAdjustmentAmtTxt));
		billRefGridFirstRowAdjustmentAmtTxt.click();

		
		getAction().moveToElement(billRefPickIcon).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));

		String actBillNewReferencePick                           = billRefNewReferenceTxt.getAttribute("value");
		String actBillTransactionCurrencyPick                    = billRefTransactionCurency.getText();
		String actBillBaseCurrencyPick                           = billRefBaseCurrency.getText();
		String actBillLocalCurrencyPick                          = localCurrencyDhs.getText();
		String actBillBalanceNewRefAmountPick                    = balanceNewReferenceAmt.getText();
		String actbillRefAdjustAmountInTransCurencyPick          = billRefAdjustAmountInTransCurency.getText();
		String actbillRefBalanceAmountAdjustInTrnasCurrencyPick  = billRefBalanceAmountAdjustInTrnasCurrency.getText();

		String actbreakUpDetailsAccountPick              = breakUpDetailsAccount.getText();
		String actbreakUpDetailsItemPick                 = breakUpDetailsItem.getText(); 
		String actbreakUpDetailsDepartmentPick           = breakUpDetailsDepartment.getText();
		//String actconversationRateBaseCurrencyRatePick   = conversationRateBaseCurrencyRate.getText();
		String actconversationRateLocalCurrencyRatePick  = conversationRateLocalCurrencyRate.getText();

		String actasOnEntryDateTransAmtPick              = asOnEntryDateTransAmt.getText();
		String actasOnEntryDateBaseConcersationRatePick  = asOnEntryDateBaseConcersationRate.getText();
		String actasOnEntryDateBaseAmountPick            = asOnEntryDateBaseAmount.getText();
		String actasOnEntryDateLocConversationRatePick   = asOnEntryDateLocConversationRate.getText();
		String actasOnEntryDateAmtPick                   = asOnEntryDateAmt.getText();

		String actbalOnAdjstDateTransAmtPick             = balOnAdjstDateTransAmt.getText();
		String actbalOnAdjstDateBasrConversionRatePick   = balOnAdjstDateBasrConversionRate.getText();
		String actbalOnAdjstDateBaseAmountPick           = balOnAdjstDateBaseAmount.getText();
		String actbalOnAdjstDateLocalConversionRatePick  = balOnAdjstDateLocalConversionRate.getText();
		String actbalOnAdjstDateAmtPick                  = balOnAdjstDateAmt.getText();

		String actadjustmentsAmount1Pick                 = adjustmentsAmount1.getText();
		String actadjustmentsAmount2Pick                 = adjustmentsAmount2.getText();
		String actadjustmentsAmount3Pick                 = adjustmentsAmount3.getText();
		String actadjustmentsAmount4Pick                 = adjustmentsAmount4.getText();

		String actexchangeGainLossForBaseCurrencyPick    = exchangeGainLossForBaseCurrency.getText();	
		String actexchangeGainLossForLocalCurrencyPick   = exchangeGainLossForLocalCurrency.getText();

		String expBillNewReferencePick                          = "0.00";
		
		String expBillTransactionCurrencyPick                   = "50";
		
		String expBillBaseCurrencyPick                          = "50";
		
		String expBillLocalCurrencyPick                         = "3.50";
		
		String expBillBalanceNewRefAmountPick                   = "0.00";
		
		String expbillRefAdjustAmountInTransCurencyPick         = "50.00";
		
		String expbillRefBalanceAmountAdjustInTrnasCurrencyPick = "0.00";
		
		String expbreakUpDetailsAccountPick             = "071-001";
		
		
		String expbreakUpDetailsDepartmentPick          ="DUBAI";
		
		
		String expconversationRateLocalCurrencyRatePick =	"0.0700000000";
		

		String expasOnEntryDateTransAmtPick             =	"100.00";
		
				
		String expasOnEntryDateBaseConcersationRatePick =	"1.0000000000";
		
		String expasOnEntryDateBaseAmountPick           =	"100.00";
		
		String expasOnEntryDateLocConversationRatePick  =	"0.0700000000";
		
		String expasOnEntryDateAmtPick                  =	"7.00";
		
		String expbalOnAdjstDateTransAmtPick            =	"100.00";
		
		String expbalOnAdjstDateBasrConversionRatePick  =	"1";
		
		String expbalOnAdjstDateBaseAmountPick          =	"100.00";
		
		String expbalOnAdjstDateLocalConversionRatePick =	"0.0700000000";
		
		String expbalOnAdjstDateAmtPick                 = 	"7.00";
		
		String expadjustmentsAmount1Pick                =	"50.00";
		
		String expadjustmentsAmount2Pick                =	"50.00";
		
		String expadjustmentsAmount3Pick                =	"3.50";
		
		String expadjustmentsAmount4Pick                =	"50.00";
		
		String expexchangeGainLossForBaseCurrencyPick   =	"0.00";
		
		String expexchangeGainLossForLocalCurrencyPick  =	"0.00";
		
		System.out.println("*********************************************************************************************************");

		System.out.println("Bill reference Adjustment Bills  :"+ actAdjustbills+"                          "+ "expDate :"+expAdjustbills);
		System.out.println("actBillNewReference :             "+ actBillNewReference +"                    "+ "expBillNewReference :"+expBillNewReference);
		System.out.println("actBillTransactionCurrency       :"+ actBillTransactionCurrency + "            " + "expBillTransactionCurrency :"+expBillTransactionCurrency);
		System.out.println("actBillBaseCurrency :             "+ actBillBaseCurrency + "                   " +"expBillBaseCurrency :"+expBillBaseCurrency);
		System.out.println("actBillLocalCurrency :            "+ actBillLocalCurrency+ "                   " +"expBillLocalCurrency :"+expBillLocalCurrency);
		System.out.println("actBillBalanceNewRefAmount :      "+ actBillBalanceNewRefAmount + "            " + "expBillBalanceNewRefAmount :"+expBillBalanceNewRefAmount);

		System.out.println("actbillRefAdjustAmountInTransCurency :"+ actbillRefAdjustAmountInTransCurency+ "       " +"expbillRefAdjustAmountInTransCurency :"+expbillRefAdjustAmountInTransCurency);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrency :"+ actbillRefBalanceAmountAdjustInTrnasCurrency + "       " +"expbillRefBalanceAmountAdjustInTrnasCurrency :"+expbillRefBalanceAmountAdjustInTrnasCurrency);


		//////Pick

		System.out.println("actBillNewReferencePick :              "+ actBillNewReferencePick +"              "+ "expBillNewReferencePick :"+expBillNewReferencePick);
		System.out.println("actBillTransactionCurrencyPick :       "+ actBillTransactionCurrencyPick + "     " + "expBillTransactionCurrencyPick :"+expBillTransactionCurrencyPick);
		System.out.println("actBillBaseCurrencyPick :              "+ actBillBaseCurrencyPick + "            " + "expBillBaseCurrencyPick :"+expBillBaseCurrencyPick);
		System.out.println("actBillLocalCurrencyPick :             "+ actBillLocalCurrencyPick+ "                " +"expBillLocalCurrencyPick :"+expBillLocalCurrencyPick);
		System.out.println("actBillBalanceNewRefAmountPick :       "+ actBillBalanceNewRefAmountPick +        " " + "expBillBalanceNewRefAmountPick :"+expBillBalanceNewRefAmountPick);
		System.out.println("actbreakUpDetailsAccountPick :         "+ actbreakUpDetailsAccountPick+            " " +"expbreakUpDetailsAccountPick :"+expbreakUpDetailsAccountPick);
		System.out.println("actbreakUpDetailsDepartmentPick :      "+ actbreakUpDetailsDepartmentPick +        "  " +"expbreakUpDetailsDepartmentPick :"+expbreakUpDetailsDepartmentPick);
		//System.out.println("actconversationRateBaseCurrRatePick:   "+ actconversationRateBaseCurrencyRatePick+ "  " +"expconversationRateBaseCurrencyRatePick :"+expconversationRateBaseCurrencyRatePick);
		System.out.println("actconversationRateLocalCurRatePick :  "+ actconversationRateLocalCurrencyRatePick+ " " + "expconversationRateLocalCurrencyRatePick :"+expconversationRateLocalCurrencyRatePick);
		System.out.println("actasOnEntryDateTransAmtPick :         "+ actasOnEntryDateTransAmtPick  + "       " + "expasOnEntryDateTransAmtPick :"+expasOnEntryDateTransAmtPick);
		System.out.println("actOnEntryDateBaseConcersationRatePick :"+ actasOnEntryDateBaseConcersationRatePick + "       " +"expasOnEntryDateBaseConcersationRatePick :"+expasOnEntryDateBaseConcersationRatePick);
		System.out.println("actasOnEntryDateBaseAmountPick :       "+ actasOnEntryDateBaseAmountPick          + "       " +"expasOnEntryDateBaseAmountPick :"+expasOnEntryDateBaseAmountPick);
		System.out.println("actasOnEntryDateLocConverRatePick :    "+ actasOnEntryDateLocConversationRatePick + "       " +"expasOnEntryDateLocConversationRatePick :"+expasOnEntryDateLocConversationRatePick);
		System.out.println("actasOnEntryDateAmtPick :              "+ actasOnEntryDateAmtPick+ "       " + "expasOnEntryDateAmtPick :"+expasOnEntryDateAmtPick);

		System.out.println("actbalOnAdjstDateTransAmtPick :         "+ actbalOnAdjstDateTransAmtPick + "       " +"expbalOnAdjstDateTransAmtPick :"+expbalOnAdjstDateTransAmtPick);
		System.out.println("actbalOnAdjstDateBasrConversionRatePick :"+ actbalOnAdjstDateBasrConversionRatePick+ "       " +"expbalOnAdjstDateBasrConversionRatePick :"+expbalOnAdjstDateBasrConversionRatePick);
		System.out.println("actbalOnAdjstDateBaseAmountPick :        "+ actbalOnAdjstDateBaseAmountPick + "       " +"expbalOnAdjstDateBaseAmountPick :"+expbalOnAdjstDateBaseAmountPick);
		System.out.println("actbalOnAdjstDateLocalConversionRatePick:"+ actbalOnAdjstDateLocalConversionRatePick + "       " +"expbalOnAdjstDateLocalConversionRatePick :"+expbalOnAdjstDateLocalConversionRatePick);
		System.out.println("actbalOnAdjstDateAmtPick                 :"+ actbalOnAdjstDateAmtPick + "       " +"expbalOnAdjstDateAmtPick :"+expbalOnAdjstDateAmtPick);

		System.out.println("actadjustmentsAmount1Pick :   "+ actadjustmentsAmount1Pick+ "       " +"expadjustmentsAmount1Pick:"+expadjustmentsAmount1Pick);
		System.out.println("actadjustmentsAmount2PickPick :"+ actadjustmentsAmount2Pick + "       " +"expadjustmentsAmount2PickPick :"+expadjustmentsAmount2Pick);
		System.out.println("actadjustmentsAmount3PickPick :"+ actadjustmentsAmount3Pick+ "       " + "expadjustmentsAmount3Pick:"+expadjustmentsAmount3Pick);
		System.out.println("actadjustmentsAmount4PickPick :"+ actadjustmentsAmount4Pick + "       " +"expadjustmentsAmount4Pick :"+expadjustmentsAmount4Pick);

		System.out.println("actexchangeGainLossForBaseCurrencyPick : "+ actexchangeGainLossForBaseCurrencyPick + "       " +"expexchangeGainLossForBaseCurrencyPick :"+expexchangeGainLossForBaseCurrencyPick);
		System.out.println("actexchangeGainLossForLocalCurrencyPick :"+ actexchangeGainLossForLocalCurrencyPick+ "       " +"expexchangeGainLossForLocalCurrencyPick :"+expexchangeGainLossForLocalCurrencyPick);

		System.out.println("actbillRefAdjustAmountInTransCurencyPick :       "+ actbillRefAdjustAmountInTransCurencyPick+ "       " +"expbillRefAdjustAmountInTransCurencyPick :"+expbillRefAdjustAmountInTransCurencyPick);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrencyPick :"+ actbillRefBalanceAmountAdjustInTrnasCurrencyPick + "       " +"expbillRefBalanceAmountAdjustInTrnasCurrencyPick :"+expbillRefBalanceAmountAdjustInTrnasCurrencyPick);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));	
		voucherSaveBtn.click();

		

		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": 1";
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		
		
		if( actAdjustbills.equalsIgnoreCase(expAdjustbills)  && actBillNewReference.equalsIgnoreCase(expBillNewReference) && actBillTransactionCurrency.equalsIgnoreCase(expBillTransactionCurrency) && 
				actBillBaseCurrency.equalsIgnoreCase(expBillBaseCurrency) && actBillLocalCurrency.equalsIgnoreCase(expBillLocalCurrency) && 
				actBillBalanceNewRefAmount.equalsIgnoreCase(expBillBalanceNewRefAmount) && 
				actbillRefAdjustAmountInTransCurency.equalsIgnoreCase(expbillRefAdjustAmountInTransCurency) &&
				actbillRefBalanceAmountAdjustInTrnasCurrency.equalsIgnoreCase(expbillRefBalanceAmountAdjustInTrnasCurrency) &&   

				actBillNewReferencePick.equalsIgnoreCase(expBillNewReferencePick) && actBillTransactionCurrencyPick.equalsIgnoreCase(expBillTransactionCurrencyPick) && 
				actBillBaseCurrencyPick.equalsIgnoreCase(expBillBaseCurrencyPick) && actBillLocalCurrencyPick.equalsIgnoreCase(expBillLocalCurrencyPick) && 
				actBillBalanceNewRefAmountPick.equalsIgnoreCase(expBillBalanceNewRefAmountPick) && actbreakUpDetailsAccountPick.equalsIgnoreCase(expbreakUpDetailsAccountPick) && 
				actbreakUpDetailsDepartmentPick.equalsIgnoreCase(expbreakUpDetailsDepartmentPick)/* && 
				actconversationRateBaseCurrencyRatePick.equalsIgnoreCase(expconversationRateBaseCurrencyRatePick)*/ && actconversationRateLocalCurrencyRatePick.equalsIgnoreCase(expconversationRateLocalCurrencyRatePick) && 
				actasOnEntryDateTransAmtPick.equalsIgnoreCase(expasOnEntryDateTransAmtPick) && actasOnEntryDateBaseConcersationRatePick.equalsIgnoreCase(expasOnEntryDateBaseConcersationRatePick)&& 
				actasOnEntryDateBaseAmountPick.equalsIgnoreCase(expasOnEntryDateBaseAmountPick) && actasOnEntryDateLocConversationRatePick.equalsIgnoreCase(expasOnEntryDateLocConversationRatePick) && 
				actasOnEntryDateAmtPick.equalsIgnoreCase(expasOnEntryDateAmtPick) && actbalOnAdjstDateTransAmtPick.equalsIgnoreCase(expbalOnAdjstDateTransAmtPick) && 
				actbalOnAdjstDateBasrConversionRatePick.equalsIgnoreCase(expbalOnAdjstDateBasrConversionRatePick) && actbalOnAdjstDateBaseAmountPick.equalsIgnoreCase(expbalOnAdjstDateBaseAmountPick) && 
				actbalOnAdjstDateLocalConversionRatePick.equalsIgnoreCase(expbalOnAdjstDateLocalConversionRatePick) && actbalOnAdjstDateAmtPick.equalsIgnoreCase(expbalOnAdjstDateAmtPick) && 
				actbalOnAdjstDateAmtPick.equalsIgnoreCase(expbalOnAdjstDateAmtPick) && actadjustmentsAmount2Pick.equalsIgnoreCase(expadjustmentsAmount2Pick) && 
				actadjustmentsAmount1Pick.equalsIgnoreCase(expadjustmentsAmount1Pick) && actadjustmentsAmount3Pick.equalsIgnoreCase(expadjustmentsAmount3Pick) &&
				actadjustmentsAmount4Pick.equalsIgnoreCase(expadjustmentsAmount4Pick) && actexchangeGainLossForBaseCurrencyPick.equalsIgnoreCase(expexchangeGainLossForBaseCurrencyPick) &&
				actexchangeGainLossForLocalCurrencyPick.equalsIgnoreCase(expexchangeGainLossForLocalCurrencyPick) && actbillRefAdjustAmountInTransCurencyPick.equalsIgnoreCase(expbillRefAdjustAmountInTransCurencyPick) &&
				actbillRefBalanceAmountAdjustInTrnasCurrencyPick.equalsIgnoreCase(expbillRefBalanceAmountAdjustInTrnasCurrencyPick)

				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
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








	public boolean checkEditingRecepitsVATVocher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Amount));
		enter_Amount.sendKeys(excelReader.getCellData(xlSheetName, 48, 5));
		enter_Amount.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPartyName));

		String actPartyName = billRefPartyName.getText();
		String expPartyName ="Customer A (122-001)";

		System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);

		int Adjustbills=billRefAdjustBillsGridList.size();

		String actAdjustbills=Integer.toString(Adjustbills);

		String expAdjustbills="5";

		System.err.println("actAdjustbills : "+actAdjustbills +" Value Expected  : "+expAdjustbills);


		int billwiseAdjustBillsDocListcount=billwiseAdjustBillsDocList.size();

		HashSet<String> actbillwiseAdjustBillsDocList  = new HashSet<String>();

		for (int i = 0; i < billwiseAdjustBillsDocListcount; i++) 
		{
			String data		                              = billwiseAdjustBillsDocList.get(i).getText();
			actbillwiseAdjustBillsDocList.add(data);
		}

		String actDocumentNumberText=actbillwiseAdjustBillsDocList.toString();

		String expDocumentNumberText="[NDT50:1, NDT50:4, NDT50:5, NDT50:6, NDT50:7]";

		System.out.println("actDocumentNumberText   : "+actDocumentNumberText);
		System.out.println("expDocumentNumberText   : "+expDocumentNumberText);


		Thread.sleep(2000);			

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billrefAdjuBills1stChkbox));
		billrefAdjuBills1stChkbox.click();
		Thread.sleep(2000);


		//System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);


		//System.err.println("actAdjustbills : "+actAdjustbills +" Value Expected  : "+expAdjustbills);


		Thread.sleep(2000);

		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("arguments[0].scrollIntoView();", gridAdjustmentAmtRow1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		String actBillNewReference                           = billRefNewReferenceTxt.getAttribute("value");
		String actBillTransactionCurrency                    = billRefTransactionCurency.getText();
		String actBillBaseCurrency                           = billRefBaseCurrency.getText();
		String actBillLocalCurrency                          = localCurrencyDhs.getText();
		String actBillBalanceNewRefAmount                    = balanceNewReferenceAmt.getText();
		String actbillRefAdjustAmountInTransCurency          = billRefAdjustAmountInTransCurency.getText();
		String actbillRefBalanceAmountAdjustInTrnasCurrency  = billRefBalanceAmountAdjustInTrnasCurrency.getText();
		
		
		String expBillNewReference                      =	"0.00";
		
		String expBillTransactionCurrency               =	"100.00";
		
		String expBillBaseCurrency                      =	"100.00";
		
		String expBillLocalCurrency                     =	"7.00";
		
		String expBillBalanceNewRefAmount               =	"0.00";
		
		String expbillRefAdjustAmountInTransCurency     =	"0.00";
		
		String expbillRefBalanceAmountAdjustInTrnasCurrency =	"100.00";
		


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefGridFirstRowAdjustmentAmtTxt));
		billRefGridFirstRowAdjustmentAmtTxt.click();
		
		getAction().moveToElement(billRefPickIcon).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));

		String actBillNewReferencePick                           = billRefNewReferenceTxt.getAttribute("value");
		String actBillTransactionCurrencyPick                    = billRefTransactionCurency.getText();
		String actBillBaseCurrencyPick                           = billRefBaseCurrency.getText();
		String actBillLocalCurrencyPick                          = localCurrencyDhs.getText();
		String actBillBalanceNewRefAmountPick                    = balanceNewReferenceAmt.getText();
		String actbillRefAdjustAmountInTransCurencyPick          = billRefAdjustAmountInTransCurency.getText();
		String actbillRefBalanceAmountAdjustInTrnasCurrencyPick  = billRefBalanceAmountAdjustInTrnasCurrency.getText();


		String actbreakUpDetailsAccountPick              = breakUpDetailsAccount.getText();
		String actbreakUpDetailsItemPick                 = breakUpDetailsItem.getText(); 
		String actbreakUpDetailsDepartmentPick           = breakUpDetailsDepartment.getText();
		//String actconversationRateBaseCurrencyRatePick   = conversationRateBaseCurrencyRate.getText();
		String actconversationRateLocalCurrencyRatePick  = conversationRateLocalCurrencyRate.getText();

		String actasOnEntryDateTransAmtPick              = asOnEntryDateTransAmt.getText();
		String actasOnEntryDateBaseConcersationRatePick  = asOnEntryDateBaseConcersationRate.getText();
		String actasOnEntryDateBaseAmountPick            = asOnEntryDateBaseAmount.getText();
		String actasOnEntryDateLocConversationRatePick   = asOnEntryDateLocConversationRate.getText();
		String actasOnEntryDateAmtPick                   = asOnEntryDateAmt.getText();

		String actbalOnAdjstDateTransAmtPick             = balOnAdjstDateTransAmt.getText();
		String actbalOnAdjstDateBasrConversionRatePick   = balOnAdjstDateBasrConversionRate.getText();
		String actbalOnAdjstDateBaseAmountPick           = balOnAdjstDateBaseAmount.getText();
		String actbalOnAdjstDateLocalConversionRatePick  = balOnAdjstDateLocalConversionRate.getText();
		String actbalOnAdjstDateAmtPick                  = balOnAdjstDateAmt.getText();

		String actadjustmentsAmount1Pick                 = adjustmentsAmount1.getText();
		String actadjustmentsAmount2Pick                 = adjustmentsAmount2.getText();
		String actadjustmentsAmount3Pick                 = adjustmentsAmount3.getText();
		String actadjustmentsAmount4Pick                 = adjustmentsAmount4.getText();

		String actexchangeGainLossForBaseCurrencyPick    = exchangeGainLossForBaseCurrency.getText();	
		String actexchangeGainLossForLocalCurrencyPick   = exchangeGainLossForLocalCurrency.getText();

		
		String expBillNewReferencePick                          = 	"0.00";
		
		String expBillTransactionCurrencyPick                   = 	"100";
		
		String expBillBaseCurrencyPick                          = 	"100";
		
		String expBillLocalCurrencyPick                         = 	"7.00";
		
		String expBillBalanceNewRefAmountPick                   = 	"0.00";
		
		String expbillRefAdjustAmountInTransCurencyPick         = 	"100.00";
		
		String expbillRefBalanceAmountAdjustInTrnasCurrencyPick = 	"0.00";
		
		String expbreakUpDetailsAccountPick             =	"071-001";
	
		String expbreakUpDetailsDepartmentPick          =	"DUBAI";
		
		String expconversationRateLocalCurrencyRatePick =	"0.0700000000";
		
		String expasOnEntryDateTransAmtPick             =	"100.00";
		
		String expasOnEntryDateBaseConcersationRatePick =	"1.0000000000";
		
		String expasOnEntryDateBaseAmountPick           =	"100.00";
		
		String expasOnEntryDateLocConversationRatePick  =	"0.0700000000";
		
		String expasOnEntryDateAmtPick                  =	"7.00";
		
		String expbalOnAdjstDateTransAmtPick            =	"100.00";
		
		String expbalOnAdjstDateBasrConversionRatePick  =	"1";
		
		String expbalOnAdjstDateBaseAmountPick          =	"100.00";
		
		String expbalOnAdjstDateLocalConversionRatePick =	"0.0700000000";
		
		String expbalOnAdjstDateAmtPick                 =	"7.00";
		
		String expadjustmentsAmount1Pick                =	"0.00";
		
		String expadjustmentsAmount2Pick                =	"0.00";
		
		String expadjustmentsAmount3Pick                =	"0.00";
		
		String expadjustmentsAmount4Pick                =	"0.00";
		
		String expexchangeGainLossForBaseCurrencyPick   =	"0.00";
		
		String expexchangeGainLossForLocalCurrencyPick  =	"0.00";
		
		System.out.println("*********************************************************************************************************");

		System.out.println("Bill reference Adjustment Bills  :"+ actAdjustbills+"                          "+ "expDate :"+expAdjustbills);
		System.out.println("actBillNewReference :             "+ actBillNewReference +"                    "+ "expBillNewReference :"+expBillNewReference);
		System.out.println("actBillTransactionCurrency       :"+ actBillTransactionCurrency + "            " + "expBillTransactionCurrency :"+expBillTransactionCurrency);
		System.out.println("actBillBaseCurrency :             "+ actBillBaseCurrency + "                   " +"expBillBaseCurrency :"+expBillBaseCurrency);
		System.out.println("actBillLocalCurrency :            "+ actBillLocalCurrency+ "                   " +"expBillLocalCurrency :"+expBillLocalCurrency);
		System.out.println("actBillBalanceNewRefAmount :      "+ actBillBalanceNewRefAmount + "            " + "expBillBalanceNewRefAmount :"+expBillBalanceNewRefAmount);

		System.out.println("actbillRefAdjustAmountInTransCurency :"+ actbillRefAdjustAmountInTransCurency+ "       " +"expbillRefAdjustAmountInTransCurency :"+expbillRefAdjustAmountInTransCurency);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrency :"+ actbillRefBalanceAmountAdjustInTrnasCurrency + "       " +"expbillRefBalanceAmountAdjustInTrnasCurrency :"+expbillRefBalanceAmountAdjustInTrnasCurrency);




		//////Pick

		System.out.println("actBillNewReferencePick :              "+ actBillNewReferencePick +"              "+ "expBillNewReferencePick :"+expBillNewReferencePick);
		System.out.println("actBillTransactionCurrencyPick :       "+ actBillTransactionCurrencyPick + "     " + "expBillTransactionCurrencyPick :"+expBillTransactionCurrencyPick);
		System.out.println("actBillBaseCurrencyPick :              "+ actBillBaseCurrencyPick + "            " + "expBillBaseCurrencyPick :"+expBillBaseCurrencyPick);
		System.out.println("actBillLocalCurrencyPick :             "+ actBillLocalCurrencyPick+ "                " +"expBillLocalCurrencyPick :"+expBillLocalCurrencyPick);
		System.out.println("actBillBalanceNewRefAmountPick :       "+ actBillBalanceNewRefAmountPick +        " " + "expBillBalanceNewRefAmountPick :"+expBillBalanceNewRefAmountPick);
		System.out.println("actbreakUpDetailsAccountPick :         "+ actbreakUpDetailsAccountPick+            " " +"expbreakUpDetailsAccountPick :"+expbreakUpDetailsAccountPick);
		System.out.println("actbreakUpDetailsDepartmentPick :      "+ actbreakUpDetailsDepartmentPick +        "  " +"expbreakUpDetailsDepartmentPick :"+expbreakUpDetailsDepartmentPick);
		//System.out.println("actconversationRateBaseCurrRatePick:   "+ actconversationRateBaseCurrencyRatePick+ "  " +"expconversationRateBaseCurrencyRatePick :"+expconversationRateBaseCurrencyRatePick);
		System.out.println("actconversationRateLocalCurRatePick :  "+ actconversationRateLocalCurrencyRatePick+ " " + "expconversationRateLocalCurrencyRatePick :"+expconversationRateLocalCurrencyRatePick);
		System.out.println("actasOnEntryDateTransAmtPick :         "+ actasOnEntryDateTransAmtPick  + "       " + "expasOnEntryDateTransAmtPick :"+expasOnEntryDateTransAmtPick);
		System.out.println("actOnEntryDateBaseConcersationRatePick :"+ actasOnEntryDateBaseConcersationRatePick + "       " +"expasOnEntryDateBaseConcersationRatePick :"+expasOnEntryDateBaseConcersationRatePick);
		System.out.println("actasOnEntryDateBaseAmountPick :       "+ actasOnEntryDateBaseAmountPick          + "       " +"expasOnEntryDateBaseAmountPick :"+expasOnEntryDateBaseAmountPick);
		System.out.println("actasOnEntryDateLocConverRatePick :    "+ actasOnEntryDateLocConversationRatePick + "       " +"expasOnEntryDateLocConversationRatePick :"+expasOnEntryDateLocConversationRatePick);
		System.out.println("actasOnEntryDateAmtPick :              "+ actasOnEntryDateAmtPick+ "       " + "expasOnEntryDateAmtPick :"+expasOnEntryDateAmtPick);

		System.out.println("actbalOnAdjstDateTransAmtPick :         "+ actbalOnAdjstDateTransAmtPick + "       " +"expbalOnAdjstDateTransAmtPick :"+expbalOnAdjstDateTransAmtPick);
		System.out.println("actbalOnAdjstDateBasrConversionRatePick :"+ actbalOnAdjstDateBasrConversionRatePick+ "       " +"expbalOnAdjstDateBasrConversionRatePick :"+expbalOnAdjstDateBasrConversionRatePick);
		System.out.println("actbalOnAdjstDateBaseAmountPick :        "+ actbalOnAdjstDateBaseAmountPick + "       " +"expbalOnAdjstDateBaseAmountPick :"+expbalOnAdjstDateBaseAmountPick);
		System.out.println("actbalOnAdjstDateLocalConversionRatePick:"+ actbalOnAdjstDateLocalConversionRatePick + "       " +"expbalOnAdjstDateLocalConversionRatePick :"+expbalOnAdjstDateLocalConversionRatePick);
		System.out.println("actbalOnAdjstDateAmtPick                 :"+ actbalOnAdjstDateAmtPick + "       " +"expbalOnAdjstDateAmtPick :"+expbalOnAdjstDateAmtPick);

		System.out.println("actadjustmentsAmount1Pick :   "+ actadjustmentsAmount1Pick+ "       " +"expadjustmentsAmount1Pick:"+expadjustmentsAmount1Pick);
		System.out.println("actadjustmentsAmount2PickPick :"+ actadjustmentsAmount2Pick + "       " +"expadjustmentsAmount2PickPick :"+expadjustmentsAmount2Pick);
		System.out.println("actadjustmentsAmount3PickPick :"+ actadjustmentsAmount3Pick+ "       " + "expadjustmentsAmount3Pick:"+expadjustmentsAmount3Pick);
		System.out.println("actadjustmentsAmount4PickPick :"+ actadjustmentsAmount4Pick + "       " +"expadjustmentsAmount4Pick :"+expadjustmentsAmount4Pick);

		System.out.println("actexchangeGainLossForBaseCurrencyPick : "+ actexchangeGainLossForBaseCurrencyPick + "       " +"expexchangeGainLossForBaseCurrencyPick :"+expexchangeGainLossForBaseCurrencyPick);
		System.out.println("actexchangeGainLossForLocalCurrencyPick :"+ actexchangeGainLossForLocalCurrencyPick+ "       " +"expexchangeGainLossForLocalCurrencyPick :"+expexchangeGainLossForLocalCurrencyPick);

		System.out.println("actbillRefAdjustAmountInTransCurencyPick :       "+ actbillRefAdjustAmountInTransCurencyPick+ "       " +"expbillRefAdjustAmountInTransCurencyPick :"+expbillRefAdjustAmountInTransCurencyPick);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrencyPick :"+ actbillRefBalanceAmountAdjustInTrnasCurrencyPick + "       " +"expbillRefBalanceAmountAdjustInTrnasCurrencyPick :"+expbillRefBalanceAmountAdjustInTrnasCurrencyPick);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));	
		voucherSaveBtn.click();

		
		String expMessage1 ="Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": 1";
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		
		if( actAdjustbills.equalsIgnoreCase(expAdjustbills)  && actBillNewReference.equalsIgnoreCase(expBillNewReference) && actBillTransactionCurrency.equalsIgnoreCase(expBillTransactionCurrency) && 
				actBillBaseCurrency.equalsIgnoreCase(expBillBaseCurrency) && actBillLocalCurrency.equalsIgnoreCase(expBillLocalCurrency) && 
				actBillBalanceNewRefAmount.equalsIgnoreCase(expBillBalanceNewRefAmount) && 
				actbillRefAdjustAmountInTransCurency.equalsIgnoreCase(expbillRefAdjustAmountInTransCurency) &&
				actbillRefBalanceAmountAdjustInTrnasCurrency.equalsIgnoreCase(expbillRefBalanceAmountAdjustInTrnasCurrency) &&   

				actBillNewReferencePick.equalsIgnoreCase(expBillNewReferencePick) && actBillTransactionCurrencyPick.equalsIgnoreCase(expBillTransactionCurrencyPick) && 
				actBillBaseCurrencyPick.equalsIgnoreCase(expBillBaseCurrencyPick) && actBillLocalCurrencyPick.equalsIgnoreCase(expBillLocalCurrencyPick) && 
				actBillBalanceNewRefAmountPick.equalsIgnoreCase(expBillBalanceNewRefAmountPick) && actbreakUpDetailsAccountPick.equalsIgnoreCase(expbreakUpDetailsAccountPick) && 
				actbreakUpDetailsDepartmentPick.equalsIgnoreCase(expbreakUpDetailsDepartmentPick) /*&& 
				actconversationRateBaseCurrencyRatePick.equalsIgnoreCase(expconversationRateBaseCurrencyRatePick)*/ && actconversationRateLocalCurrencyRatePick.equalsIgnoreCase(expconversationRateLocalCurrencyRatePick) && 
				actasOnEntryDateTransAmtPick.equalsIgnoreCase(expasOnEntryDateTransAmtPick) && actasOnEntryDateBaseConcersationRatePick.equalsIgnoreCase(expasOnEntryDateBaseConcersationRatePick)&& 
				actasOnEntryDateBaseAmountPick.equalsIgnoreCase(expasOnEntryDateBaseAmountPick) && actasOnEntryDateLocConversationRatePick.equalsIgnoreCase(expasOnEntryDateLocConversationRatePick) && 
				actasOnEntryDateAmtPick.equalsIgnoreCase(expasOnEntryDateAmtPick) && actbalOnAdjstDateTransAmtPick.equalsIgnoreCase(expbalOnAdjstDateTransAmtPick) && 
				actbalOnAdjstDateBasrConversionRatePick.equalsIgnoreCase(expbalOnAdjstDateBasrConversionRatePick) && actbalOnAdjstDateBaseAmountPick.equalsIgnoreCase(expbalOnAdjstDateBaseAmountPick) && 
				actbalOnAdjstDateLocalConversionRatePick.equalsIgnoreCase(expbalOnAdjstDateLocalConversionRatePick) && actbalOnAdjstDateAmtPick.equalsIgnoreCase(expbalOnAdjstDateAmtPick) && 
				actbalOnAdjstDateAmtPick.equalsIgnoreCase(expbalOnAdjstDateAmtPick) && actadjustmentsAmount2Pick.equalsIgnoreCase(expadjustmentsAmount2Pick) && 
				actadjustmentsAmount1Pick.equalsIgnoreCase(expadjustmentsAmount1Pick) && actadjustmentsAmount3Pick.equalsIgnoreCase(expadjustmentsAmount3Pick) &&
				actadjustmentsAmount4Pick.equalsIgnoreCase(expadjustmentsAmount4Pick) && actexchangeGainLossForBaseCurrencyPick.equalsIgnoreCase(expexchangeGainLossForBaseCurrencyPick) &&
				actexchangeGainLossForLocalCurrencyPick.equalsIgnoreCase(expexchangeGainLossForLocalCurrencyPick) && actbillRefAdjustAmountInTransCurencyPick.equalsIgnoreCase(expbillRefAdjustAmountInTransCurencyPick) &&
				actbillRefBalanceAmountAdjustInTrnasCurrencyPick.equalsIgnoreCase(expbillRefBalanceAmountAdjustInTrnasCurrencyPick)

				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
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








	public boolean checkSavingRecepitsVATVocher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCashBankAccountTxt));
		newCashBankAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 88, 5));
		Thread.sleep(2000);
		newCashBankAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SHIFT,Keys.HOME,Keys.BACK_SPACE);
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 89, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeofSupplyTxt));
		placeofSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 90, 5));
		Thread.sleep(2000);
		placeofSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 91, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 92, 5));
		Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enterReceiptsVATTaxCode));	
		enterReceiptsVATTaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Amount));
		enter_Amount.sendKeys(excelReader.getCellData(xlSheetName, 93, 5));
		enter_Amount.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPartyName));

		String actPartyName = billRefPartyName.getText();
		String expPartyName = "Customer A (122-001)";

		System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);

		int Adjustbills=billRefAdjustBillsGridList.size();

		String actAdjustbills=Integer.toString(Adjustbills);

		String expAdjustbills="4";

		System.err.println("actAdjustbills : "+actAdjustbills +" Value Expected  : "+expAdjustbills);


		int billwiseAdjustBillsDocListcount=billwiseAdjustBillsDocList.size();

		HashSet<String> actbillwiseAdjustBillsDocList  = new HashSet<String>();

		for (int i = 0; i < billwiseAdjustBillsDocListcount; i++) 
		{
			String data		                              = billwiseAdjustBillsDocList.get(i).getText();
			actbillwiseAdjustBillsDocList.add(data);
		}

		String actDocumentNumberText=actbillwiseAdjustBillsDocList.toString();

		String expDocumentNumberText="[NDT50:4, NDT50:5, NDT50:6, NDT50:7]";

		System.out.println("actDocumentNumberText   : "+actDocumentNumberText);
		System.out.println("expDocumentNumberText   : "+expDocumentNumberText);

		Thread.sleep(2000);			

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billrefAdjuBills1stChkbox));
		billrefAdjuBills1stChkbox.click();
		
		Thread.sleep(2000);

		//System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);

		//System.err.println("actAdjustbills : "+actAdjustbills +" Value Expected  : "+expAdjustbills);


		Thread.sleep(2000);

		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("arguments[0].scrollIntoView();", gridAdjustmentAmtRow1);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		String actBillNewReference                           = billRefNewReferenceTxt.getAttribute("value");
		String actBillTransactionCurrency                    = billRefTransactionCurency.getText();
		String actBillBaseCurrency                           = billRefBaseCurrency.getText();
		String actBillLocalCurrency                          = localCurrencyDhs.getText();
		String actBillBalanceNewRefAmount                    = balanceNewReferenceAmt.getText();
		String actbillRefAdjustAmountInTransCurency          = billRefAdjustAmountInTransCurency.getText();
		String actbillRefBalanceAmountAdjustInTrnasCurrency  = billRefBalanceAmountAdjustInTrnasCurrency.getText();

		String expBillNewReference                      =	"0.00";
		
		String expBillTransactionCurrency               =	"50.00";
		
		String expBillBaseCurrency                      =	"50.00";
		
		String expBillLocalCurrency                     =	"3.50";
		
		String expBillBalanceNewRefAmount               =	"0.00";
		
		String expbillRefAdjustAmountInTransCurency         =	"50.00";
		
		String expbillRefBalanceAmountAdjustInTrnasCurrency =	"0.00";

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefGridFirstRowAdjustmentAmtTxt));
		billRefGridFirstRowAdjustmentAmtTxt.click();
		
		getAction().moveToElement(billRefPickIcon).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));

		String actBillNewReferencePick                           = billRefNewReferenceTxt.getAttribute("value");
		String actBillTransactionCurrencyPick                    = billRefTransactionCurency.getText();
		String actBillBaseCurrencyPick                           = billRefBaseCurrency.getText();
		String actBillLocalCurrencyPick                          = localCurrencyDhs.getText();
		String actBillBalanceNewRefAmountPick                    = balanceNewReferenceAmt.getText();
		String actbillRefAdjustAmountInTransCurencyPick          = billRefAdjustAmountInTransCurency.getText();
		String actbillRefBalanceAmountAdjustInTrnasCurrencyPick  = billRefBalanceAmountAdjustInTrnasCurrency.getText();


		String actbreakUpDetailsAccountPick              = breakUpDetailsAccount.getText();
		String actbreakUpDetailsItemPick                 = breakUpDetailsItem.getText(); 
		String actbreakUpDetailsDepartmentPick           = breakUpDetailsDepartment.getText();
		String actconversationRateBaseCurrencyRatePick   = conversationRateBaseCurrencyRate.getText();
		String actconversationRateLocalCurrencyRatePick  = conversationRateLocalCurrencyRate.getText();

		String actasOnEntryDateTransAmtPick              = asOnEntryDateTransAmt.getText();
		String actasOnEntryDateBaseConcersationRatePick  = asOnEntryDateBaseConcersationRate.getText();
		String actasOnEntryDateBaseAmountPick            = asOnEntryDateBaseAmount.getText();
		String actasOnEntryDateLocConversationRatePick   = asOnEntryDateLocConversationRate.getText();
		String actasOnEntryDateAmtPick                   = asOnEntryDateAmt.getText();

		String actbalOnAdjstDateTransAmtPick             = balOnAdjstDateTransAmt.getText();
		String actbalOnAdjstDateBasrConversionRatePick   = balOnAdjstDateBasrConversionRate.getText();
		String actbalOnAdjstDateBaseAmountPick           = balOnAdjstDateBaseAmount.getText();
		String actbalOnAdjstDateLocalConversionRatePick  = balOnAdjstDateLocalConversionRate.getText();
		String actbalOnAdjstDateAmtPick                  = balOnAdjstDateAmt.getText();

		String actadjustmentsAmount1Pick                 = adjustmentsAmount1.getText();
		String actadjustmentsAmount2Pick                 = adjustmentsAmount2.getText();
		String actadjustmentsAmount3Pick                 = adjustmentsAmount3.getText();
		String actadjustmentsAmount4Pick                 = adjustmentsAmount4.getText();

		String actexchangeGainLossForBaseCurrencyPick    = exchangeGainLossForBaseCurrency.getText();	
		String actexchangeGainLossForLocalCurrencyPick   = exchangeGainLossForLocalCurrency.getText();

		
		String expBillNewReferencePick                          = "0.00";
		
		String expBillTransactionCurrencyPick                   = 	"50";
		
		String expBillBaseCurrencyPick                          = 	"50";
		
		String expBillLocalCurrencyPick                         = 	"3.50";
		
		String expBillBalanceNewRefAmountPick                   = 	"0.00";
				
		String expbillRefAdjustAmountInTransCurencyPick         = 	"50.00";
		
		String expbillRefBalanceAmountAdjustInTrnasCurrencyPick = 	"0.00";
		
		String expbreakUpDetailsAccountPick             =	"071-001";
		
		
		String expbreakUpDetailsDepartmentPick          =	"DUBAI";
		
		String expconversationRateBaseCurrencyRatePick  =	"071-001";
		
		String expconversationRateLocalCurrencyRatePick =	"0.0700000000";
		
		String expasOnEntryDateTransAmtPick             =	"60.00";
		
		String expasOnEntryDateBaseConcersationRatePick =	"1.0000000000";
		
		String expasOnEntryDateBaseAmountPick           =	"60.00";
		
		String expasOnEntryDateLocConversationRatePick  =	"0.0700000000";
		
		String expasOnEntryDateAmtPick                  =	"4.20";
		
		String expbalOnAdjstDateTransAmtPick            =	"60.00";
		
		String expbalOnAdjstDateBasrConversionRatePick  =	"1";
		
		String expbalOnAdjstDateBaseAmountPick          =	"60.00";
		
		String expbalOnAdjstDateLocalConversionRatePick =	"0.0700000000";
		
		String expbalOnAdjstDateAmtPick                 =	"4.20";
			
		String expadjustmentsAmount1Pick                =	"50.00";
		
		String expadjustmentsAmount2Pick                =	"50.00";
		
		String expadjustmentsAmount3Pick                =	"3.50";
		
		String expadjustmentsAmount4Pick                =	"50.00";
		
		String expexchangeGainLossForBaseCurrencyPick   =	"0.00";
		
		String expexchangeGainLossForLocalCurrencyPick  =	"0.00";
		

		System.out.println("*********************************************************************************************************");

		System.out.println("Bill reference Adjustment Bills  :"+ actAdjustbills+"                          "+ "expDate :"+expAdjustbills);
		System.out.println("actBillNewReference :             "+ actBillNewReference +"                    "+ "expBillNewReference :"+expBillNewReference);
		System.out.println("actBillTransactionCurrency       :"+ actBillTransactionCurrency + "            " + "expBillTransactionCurrency :"+expBillTransactionCurrency);
		System.out.println("actBillBaseCurrency :             "+ actBillBaseCurrency + "                   " +"expBillBaseCurrency :"+expBillBaseCurrency);
		System.out.println("actBillLocalCurrency :            "+ actBillLocalCurrency+ "                   " +"expBillLocalCurrency :"+expBillLocalCurrency);
		System.out.println("actBillBalanceNewRefAmount :      "+ actBillBalanceNewRefAmount + "            " + "expBillBalanceNewRefAmount :"+expBillBalanceNewRefAmount);

		System.out.println("actbillRefAdjustAmountInTransCurency :"+ actbillRefAdjustAmountInTransCurency+ "       " +"expbillRefAdjustAmountInTransCurency :"+expbillRefAdjustAmountInTransCurency);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrency :"+ actbillRefBalanceAmountAdjustInTrnasCurrency + "       " +"expbillRefBalanceAmountAdjustInTrnasCurrency :"+expbillRefBalanceAmountAdjustInTrnasCurrency);


		// Pick

		System.out.println("actBillNewReferencePick :              "+ actBillNewReferencePick +"              "+ "expBillNewReferencePick :"+expBillNewReferencePick);
		System.out.println("actBillTransactionCurrencyPick :       "+ actBillTransactionCurrencyPick + "     " + "expBillTransactionCurrencyPick :"+expBillTransactionCurrencyPick);
		System.out.println("actBillBaseCurrencyPick :              "+ actBillBaseCurrencyPick + "            " + "expBillBaseCurrencyPick :"+expBillBaseCurrencyPick);
		System.out.println("actBillLocalCurrencyPick :             "+ actBillLocalCurrencyPick+ "                " +"expBillLocalCurrencyPick :"+expBillLocalCurrencyPick);
		System.out.println("actBillBalanceNewRefAmountPick :       "+ actBillBalanceNewRefAmountPick +        " " + "expBillBalanceNewRefAmountPick :"+expBillBalanceNewRefAmountPick);
		System.out.println("actbreakUpDetailsAccountPick :         "+ actbreakUpDetailsAccountPick+            " " +"expbreakUpDetailsAccountPick :"+expbreakUpDetailsAccountPick);
		System.out.println("actbreakUpDetailsDepartmentPick :      "+ actbreakUpDetailsDepartmentPick +        "  " +"expbreakUpDetailsDepartmentPick :"+expbreakUpDetailsDepartmentPick);
		System.out.println("actconversationRateBaseCurrRatePick:   "+ actconversationRateBaseCurrencyRatePick+ "  " +"expconversationRateBaseCurrencyRatePick :"+expconversationRateBaseCurrencyRatePick);
		System.out.println("actconversationRateLocalCurRatePick :  "+ actconversationRateLocalCurrencyRatePick+ " " + "expconversationRateLocalCurrencyRatePick :"+expconversationRateLocalCurrencyRatePick);
		System.out.println("actasOnEntryDateTransAmtPick :         "+ actasOnEntryDateTransAmtPick  + "       " + "expasOnEntryDateTransAmtPick :"+expasOnEntryDateTransAmtPick);
		System.out.println("actOnEntryDateBaseConcersationRatePick :"+ actasOnEntryDateBaseConcersationRatePick + "       " +"expasOnEntryDateBaseConcersationRatePick :"+expasOnEntryDateBaseConcersationRatePick);
		System.out.println("actasOnEntryDateBaseAmountPick :       "+ actasOnEntryDateBaseAmountPick          + "       " +"expasOnEntryDateBaseAmountPick :"+expasOnEntryDateBaseAmountPick);
		System.out.println("actasOnEntryDateLocConverRatePick :    "+ actasOnEntryDateLocConversationRatePick + "       " +"expasOnEntryDateLocConversationRatePick :"+expasOnEntryDateLocConversationRatePick);
		System.out.println("actasOnEntryDateAmtPick :              "+ actasOnEntryDateAmtPick+ "       " + "expasOnEntryDateAmtPick :"+expasOnEntryDateAmtPick);

		System.out.println("actbalOnAdjstDateTransAmtPick :         "+ actbalOnAdjstDateTransAmtPick + "       " +"expbalOnAdjstDateTransAmtPick :"+expbalOnAdjstDateTransAmtPick);
		System.out.println("actbalOnAdjstDateBasrConversionRatePick :"+ actbalOnAdjstDateBasrConversionRatePick+ "       " +"expbalOnAdjstDateBasrConversionRatePick :"+expbalOnAdjstDateBasrConversionRatePick);
		System.out.println("actbalOnAdjstDateBaseAmountPick :        "+ actbalOnAdjstDateBaseAmountPick + "       " +"expbalOnAdjstDateBaseAmountPick :"+expbalOnAdjstDateBaseAmountPick);
		System.out.println("actbalOnAdjstDateLocalConversionRatePick:"+ actbalOnAdjstDateLocalConversionRatePick + "       " +"expbalOnAdjstDateLocalConversionRatePick :"+expbalOnAdjstDateLocalConversionRatePick);
		System.out.println("actbalOnAdjstDateAmtPick                 :"+ actbalOnAdjstDateAmtPick + "       " +"expbalOnAdjstDateAmtPick :"+expbalOnAdjstDateAmtPick);

		System.out.println("actadjustmentsAmount1Pick :   "+ actadjustmentsAmount1Pick+ "       " +"expadjustmentsAmount1Pick:"+expadjustmentsAmount1Pick);
		System.out.println("actadjustmentsAmount2PickPick :"+ actadjustmentsAmount2Pick + "       " +"expadjustmentsAmount2PickPick :"+expadjustmentsAmount2Pick);
		System.out.println("actadjustmentsAmount3PickPick :"+ actadjustmentsAmount3Pick+ "       " + "expadjustmentsAmount3Pick:"+expadjustmentsAmount3Pick);
		System.out.println("actadjustmentsAmount4PickPick :"+ actadjustmentsAmount4Pick + "       " +"expadjustmentsAmount4Pick :"+expadjustmentsAmount4Pick);

		System.out.println("actexchangeGainLossForBaseCurrencyPick : "+ actexchangeGainLossForBaseCurrencyPick + "       " +"expexchangeGainLossForBaseCurrencyPick :"+expexchangeGainLossForBaseCurrencyPick);
		System.out.println("actexchangeGainLossForLocalCurrencyPick :"+ actexchangeGainLossForLocalCurrencyPick+ "       " +"expexchangeGainLossForLocalCurrencyPick :"+expexchangeGainLossForLocalCurrencyPick);

		System.out.println("actbillRefAdjustAmountInTransCurencyPick :       "+ actbillRefAdjustAmountInTransCurencyPick+ "       " +"expbillRefAdjustAmountInTransCurencyPick :"+expbillRefAdjustAmountInTransCurencyPick);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrencyPick :"+ actbillRefBalanceAmountAdjustInTrnasCurrencyPick + "       " +"expbillRefBalanceAmountAdjustInTrnasCurrencyPick :"+expbillRefBalanceAmountAdjustInTrnasCurrencyPick);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));	
		voucherSaveBtn.click();
		
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

		expMsg.add(excelReader.getCellData(xlSheetName, 126, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 127, 6)+docno);
		excelReader.setCellData(xlfile, xlSheetName, 126, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 126, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 127, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 126, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		if( actAdjustbills.equalsIgnoreCase(expAdjustbills)  && actBillNewReference.equalsIgnoreCase(expBillNewReference) && actBillTransactionCurrency.equalsIgnoreCase(expBillTransactionCurrency) && 
				actBillBaseCurrency.equalsIgnoreCase(expBillBaseCurrency) && actBillLocalCurrency.equalsIgnoreCase(expBillLocalCurrency) && 
				actBillBalanceNewRefAmount.equalsIgnoreCase(expBillBalanceNewRefAmount) && 
				actbillRefAdjustAmountInTransCurency.equalsIgnoreCase(expbillRefAdjustAmountInTransCurency) &&
				actbillRefBalanceAmountAdjustInTrnasCurrency.equalsIgnoreCase(expbillRefBalanceAmountAdjustInTrnasCurrency) &&   

				actBillNewReferencePick.equalsIgnoreCase(expBillNewReferencePick) && actBillTransactionCurrencyPick.equalsIgnoreCase(expBillTransactionCurrencyPick) && 
				actBillBaseCurrencyPick.equalsIgnoreCase(expBillBaseCurrencyPick) && actBillLocalCurrencyPick.equalsIgnoreCase(expBillLocalCurrencyPick) && 
				actBillBalanceNewRefAmountPick.equalsIgnoreCase(expBillBalanceNewRefAmountPick) && actbreakUpDetailsAccountPick.equalsIgnoreCase(expbreakUpDetailsAccountPick) && 
				actbreakUpDetailsDepartmentPick.equalsIgnoreCase(expbreakUpDetailsDepartmentPick) && 
				actconversationRateBaseCurrencyRatePick.equalsIgnoreCase(expconversationRateBaseCurrencyRatePick) && actconversationRateLocalCurrencyRatePick.equalsIgnoreCase(expconversationRateLocalCurrencyRatePick) && 
				actasOnEntryDateTransAmtPick.equalsIgnoreCase(expasOnEntryDateTransAmtPick) && actasOnEntryDateBaseConcersationRatePick.equalsIgnoreCase(expasOnEntryDateBaseConcersationRatePick)&& 
				actasOnEntryDateBaseAmountPick.equalsIgnoreCase(expasOnEntryDateBaseAmountPick) && actasOnEntryDateLocConversationRatePick.equalsIgnoreCase(expasOnEntryDateLocConversationRatePick) && 
				actasOnEntryDateAmtPick.equalsIgnoreCase(expasOnEntryDateAmtPick) && actbalOnAdjstDateTransAmtPick.equalsIgnoreCase(expbalOnAdjstDateTransAmtPick) && 
				actbalOnAdjstDateBasrConversionRatePick.equalsIgnoreCase(expbalOnAdjstDateBasrConversionRatePick) && actbalOnAdjstDateBaseAmountPick.equalsIgnoreCase(expbalOnAdjstDateBaseAmountPick) && 
				actbalOnAdjstDateLocalConversionRatePick.equalsIgnoreCase(expbalOnAdjstDateLocalConversionRatePick) && actbalOnAdjstDateAmtPick.equalsIgnoreCase(expbalOnAdjstDateAmtPick) && 
				actbalOnAdjstDateAmtPick.equalsIgnoreCase(expbalOnAdjstDateAmtPick) && actadjustmentsAmount2Pick.equalsIgnoreCase(expadjustmentsAmount2Pick) && 
				actadjustmentsAmount1Pick.equalsIgnoreCase(expadjustmentsAmount1Pick) && actadjustmentsAmount3Pick.equalsIgnoreCase(expadjustmentsAmount3Pick) &&
				actadjustmentsAmount4Pick.equalsIgnoreCase(expadjustmentsAmount4Pick) && actexchangeGainLossForBaseCurrencyPick.equalsIgnoreCase(expexchangeGainLossForBaseCurrencyPick) &&
				actexchangeGainLossForLocalCurrencyPick.equalsIgnoreCase(expexchangeGainLossForLocalCurrencyPick) && actbillRefAdjustAmountInTransCurencyPick.equalsIgnoreCase(expbillRefAdjustAmountInTransCurencyPick) &&
				actbillRefBalanceAmountAdjustInTrnasCurrencyPick.equalsIgnoreCase(expbillRefBalanceAmountAdjustInTrnasCurrencyPick)

				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 88, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 88, 8, resFail);
			return false;
		}	
	}






	public boolean checkDeletingRecepitsVATVocher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

	/*	getWaitForAlert();
		getAlert().accept();*/
		
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucher_ConfirmYesBtn)); 
		  voucher_ConfirmYesBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 128, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 128, 7, actMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();

		Thread.sleep(2000);

		if(actMessage.equalsIgnoreCase(expMessage))
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








	public boolean checkVerifingDetailsOfSavedReceiptsVATVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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

		expVouchersList.add(excelReader.getCellData(xlSheetName, 129, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 130, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 131, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 132, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 133, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 134, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 135, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 136, 6));
		
		excelReader.setCellData(xlfile, xlSheetName, 129, 7, actVouchersList.toString());

		System.out.println("**********************checkVerifingDetailsOfSavedReceiptsVATVoucherInHomepage*********************");

		System.out.println(actVouchersList);
		System.out.println(expVouchersList);

		if (actVouchersList.equals(expVouchersList)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 129, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 129, 8, resFail);
			return false;
		}
	}




	public static boolean CheckLogin() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	        Thread.sleep(3000);

	        getDriver().navigate().refresh();

	        LoginPage lp=new LoginPage(getDriver()); 

	        lp.checkLoginPageTitleByURLInputInBrowser();

	        String unamelt="su";

	        String pawslt="su";

	        lp.enterUserName(unamelt);

	        lp.enterPassword(pawslt);

	        Thread.sleep(2000);

	        String compname = "Automation Company : 08/10/2020";

	        Select oSelect = new Select(companyDropDownList);

	        List<WebElement> elementCount = oSelect.getOptions();

	        int cqSize = elementCount.size();

	        System.out.println("CompanyDropdownList Count :" + cqSize);

	        int i;

	        for (i = 0; i < elementCount.size(); i++) {

	                elementCount.get(i).getText();

	                String optionName = elementCount.get(i).getText();
	                if (optionName.toUpperCase().startsWith(compname.toUpperCase())) {
	                        System.out.println("q" + elementCount.get(i).getText());
	                        elementCount.get(i).click();
	                }

	        }

	        Thread.sleep(2000);

	        lp.clickOnSignInBtn();

	        Thread.sleep(2000);

	        String actUserInfo1=userNameDisplay.getText();

	        System.out.println("User Info  : "+actUserInfo1);

	        System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

	        //getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	        //companyLogo.click();

	        String getCompanyTxt1=Company_Name.getText();
	        String getLoginCompanyName1=getCompanyTxt1.substring(0, 31);
	        System.out.println("company name  :  "+ getLoginCompanyName1);
	        //companyLogo.click();

	        String expUserInfo1           ="SU";
	        String expLoginCompanyName1   ="Automation Company : 08/10/2020";

	        System.out.println("UserInfo1             : "+actUserInfo1            +" Value Expected : "+expUserInfo1);
	        System.out.println("LoginCompanyName1     : "+getLoginCompanyName1    +" Value Expected : "+expLoginCompanyName1);

	        if(actUserInfo1.equalsIgnoreCase(expUserInfo1) && getLoginCompanyName1.contains(expLoginCompanyName1))
	        {
	                return true;
	        }
	        else
	        {
	                return false;
	        }

	}


	public boolean checkLogoutReceiptsVATvoucherPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		 
		 try
			{
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
			  userNameDisplayLogo.click();
			  Thread.sleep(2000);
			 
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			  logoutOption.click();
			  
			  Thread.sleep(2000);
			  
			  boolean actUserLoginPage              = username.isDisplayed() && username.isEnabled()
	                                               && password.isDisplayed() && password.isEnabled();
	                                      
			  boolean expUserLoginPage              = true;
			  
			  if(actUserLoginPage==expUserLoginPage)  
		      {
				System.out.println("***Test Pass: Login Successfull***");
				
				return true;
			  }
		      else
		      {
		  	 
				System.out.println("***Test Fail: Login Not Successfull***");
				
				return false;
			  }
			}
			catch (Exception e)
			{
			 	String exception = e.getMessage();
			 		
				return false;
			}
		}









	public ReceiptsVATVoucherPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
