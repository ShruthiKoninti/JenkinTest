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

public class JVVATViewPage extends BaseEngine 
{
	private static String xlfile = getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlSheetName = "JVVATViewPage";
	
	public boolean checkSavingJVVATVocher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finTransJournalsMenu));
		finTransJournalsMenu.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(JVVATViewVoucher));
		JVVATViewVoucher.click();

		Thread.sleep(2000);

		checkDeleteLinkStatus();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDateCalenderBtn));
		voucherHeaderDueDateCalenderBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDateCalenderTodayBtn));
		voucherHeaderDueDateCalenderTodayBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SHIFT,Keys.HOME,Keys.BACK_SPACE);
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 8, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_CreditACTxt));
		enter_CreditACTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		Thread.sleep(2000);
		enter_CreditACTxt.sendKeys(Keys.TAB);



		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enterJVVATTaxCode));	
		enterJVVATTaxCode.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		Thread.sleep(2000);
		enterJVVATTaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Amount));
		enter_Amount.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		enter_Amount.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPartyName));

		String actPartyName = billRefPartyName.getText();
		String expPartyName = excelReader.getCellData(xlSheetName, 8, 6);
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actPartyName);

		System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);

		int Adjustbills=billRefAdjustBillsGridList.size();

		String actAdjustbills=Integer.toString(Adjustbills);

		String expAdjustbills=excelReader.getCellData(xlSheetName, 9, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, actAdjustbills);

		System.err.println("actAdjustbills : "+actAdjustbills +" Value Expected  : "+expAdjustbills);

		int billwiseAdjustBillsDocListcount=billwiseAdjustBillsDocList.size();

		HashSet<String> actbillwiseAdjustBillsDocList  = new HashSet<String>();

		for (int i = 0; i < billwiseAdjustBillsDocListcount; i++) 
		{
			String data		                              = billwiseAdjustBillsDocList.get(i).getText();
			actbillwiseAdjustBillsDocList.add(data);
		}

		String actDocumentNumberText=actbillwiseAdjustBillsDocList.toString();

		String expDocumentNumberText=excelReader.getCellData(xlSheetName, 10, 6);

		excelReader.setCellData(xlfile, xlSheetName, 10, 7, actDocumentNumberText);
		
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
		
		String expBillNewReference                      =excelReader.getCellData(xlSheetName, 11, 6);
		excelReader.setCellData(xlfile, xlSheetName, 11, 7, actBillNewReference);
		
		String expBillTransactionCurrency               =excelReader.getCellData(xlSheetName, 12, 6);
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actBillTransactionCurrency);
		
		String expBillBaseCurrency                      =excelReader.getCellData(xlSheetName, 13, 6);
		excelReader.setCellData(xlfile, xlSheetName, 13, 7, actBillBaseCurrency);
		
		String expBillLocalCurrency                     =excelReader.getCellData(xlSheetName, 14, 6);
		excelReader.setCellData(xlfile, xlSheetName, 14, 7, actBillLocalCurrency);
		
		String expBillBalanceNewRefAmount               =excelReader.getCellData(xlSheetName, 15, 6);
		excelReader.setCellData(xlfile, xlSheetName, 15, 7, actBillBalanceNewRefAmount);
		
		String expbillRefAdjustAmountInTransCurency         =excelReader.getCellData(xlSheetName, 16, 6);
		excelReader.setCellData(xlfile, xlSheetName, 16, 7, actbillRefAdjustAmountInTransCurency);
		
		String expbillRefBalanceAmountAdjustInTrnasCurrency =excelReader.getCellData(xlSheetName, 17, 6);
		excelReader.setCellData(xlfile, xlSheetName, 17, 7, actbillRefBalanceAmountAdjustInTrnasCurrency);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefGridFirstRowAdjustmentAmtTxt));
		billRefGridFirstRowAdjustmentAmtTxt.click();

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
	//	String actconversationRateBaseCurrencyRatePick   = conversationRateBaseCurrencyRate.getText();
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


		String expBillNewReferencePick                          = excelReader.getCellData(xlSheetName, 18, 6);
		excelReader.setCellData(xlfile, xlSheetName, 18, 7, actBillNewReferencePick);
		
		String expBillTransactionCurrencyPick                   = excelReader.getCellData(xlSheetName, 19, 6);
		excelReader.setCellData(xlfile, xlSheetName, 19, 7, actBillTransactionCurrencyPick);
		
		String expBillBaseCurrencyPick                          = excelReader.getCellData(xlSheetName, 20, 6);
		excelReader.setCellData(xlfile, xlSheetName, 20, 7, actBillBaseCurrencyPick);
		
		String expBillLocalCurrencyPick                         = excelReader.getCellData(xlSheetName, 21, 6);
		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actBillLocalCurrencyPick);
		
		String expBillBalanceNewRefAmountPick                   = excelReader.getCellData(xlSheetName, 22, 6);
		excelReader.setCellData(xlfile, xlSheetName, 22, 7, actBillBalanceNewRefAmountPick);
		
		String expbillRefAdjustAmountInTransCurencyPick         = excelReader.getCellData(xlSheetName, 23, 6);
		excelReader.setCellData(xlfile, xlSheetName, 23, 7, actbillRefAdjustAmountInTransCurencyPick);
		
		String expbillRefBalanceAmountAdjustInTrnasCurrencyPick = excelReader.getCellData(xlSheetName, 24, 6);
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actbillRefBalanceAmountAdjustInTrnasCurrencyPick);

		String expbreakUpDetailsAccountPick             =excelReader.getCellData(xlSheetName, 25, 6);
		excelReader.setCellData(xlfile, xlSheetName, 25, 7, actbreakUpDetailsAccountPick);
		
		String expbreakUpDetailsDepartmentPick          =excelReader.getCellData(xlSheetName, 26, 6);
		excelReader.setCellData(xlfile, xlSheetName, 26, 7, actbreakUpDetailsDepartmentPick);
		
	/*	String expconversationRateBaseCurrencyRatePick  =excelReader.getCellData(xlSheetName, 27, 6);
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actconversationRateBaseCurrencyRatePick);*/
		
		String expconversationRateLocalCurrencyRatePick =excelReader.getCellData(xlSheetName, 28, 6);
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actconversationRateLocalCurrencyRatePick);
		
		String expasOnEntryDateTransAmtPick             =excelReader.getCellData(xlSheetName, 29, 6);
		excelReader.setCellData(xlfile, xlSheetName, 29, 7, actasOnEntryDateTransAmtPick);
		
		String expasOnEntryDateBaseConcersationRatePick =excelReader.getCellData(xlSheetName, 30, 6);
		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actasOnEntryDateBaseConcersationRatePick);
		
		String expasOnEntryDateBaseAmountPick           =excelReader.getCellData(xlSheetName, 31, 6);
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actasOnEntryDateBaseAmountPick);
		
		String expasOnEntryDateLocConversationRatePick  =excelReader.getCellData(xlSheetName, 32, 6);
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actasOnEntryDateLocConversationRatePick);
		
		String expasOnEntryDateAmtPick                  =excelReader.getCellData(xlSheetName, 33, 6);
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actasOnEntryDateAmtPick);
		
		String expbalOnAdjstDateTransAmtPick            =excelReader.getCellData(xlSheetName, 34, 6);
		excelReader.setCellData(xlfile, xlSheetName, 34, 7, actbalOnAdjstDateTransAmtPick);
		
		String expbalOnAdjstDateBasrConversionRatePick  =excelReader.getCellData(xlSheetName, 35, 6);
		excelReader.setCellData(xlfile, xlSheetName, 35, 7, actbalOnAdjstDateBasrConversionRatePick);
		
		String expbalOnAdjstDateBaseAmountPick          =excelReader.getCellData(xlSheetName, 36, 6);
		excelReader.setCellData(xlfile, xlSheetName, 36, 7, actbalOnAdjstDateBaseAmountPick);
		
		String expbalOnAdjstDateLocalConversionRatePick =excelReader.getCellData(xlSheetName, 37, 6);
		excelReader.setCellData(xlfile, xlSheetName, 37, 7, actbalOnAdjstDateLocalConversionRatePick);
		
		String expbalOnAdjstDateAmtPick                 =excelReader.getCellData(xlSheetName, 38, 6);
		excelReader.setCellData(xlfile, xlSheetName, 38, 7, actbalOnAdjstDateAmtPick);
		
		String expadjustmentsAmount1Pick                =excelReader.getCellData(xlSheetName, 39, 6);
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actadjustmentsAmount1Pick);
		
		String expadjustmentsAmount2Pick                =excelReader.getCellData(xlSheetName, 40, 6);
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, actadjustmentsAmount2Pick);
		
		String expadjustmentsAmount3Pick                =excelReader.getCellData(xlSheetName, 41, 6);
		excelReader.setCellData(xlfile, xlSheetName, 41, 7, actadjustmentsAmount3Pick);
		
		String expadjustmentsAmount4Pick                =excelReader.getCellData(xlSheetName, 42, 6);
		excelReader.setCellData(xlfile, xlSheetName, 42, 7, actadjustmentsAmount4Pick);
		
		String expexchangeGainLossForBaseCurrencyPick   =excelReader.getCellData(xlSheetName, 43, 6);
		excelReader.setCellData(xlfile, xlSheetName, 43, 7, actexchangeGainLossForBaseCurrencyPick);
		
		String expexchangeGainLossForLocalCurrencyPick  =excelReader.getCellData(xlSheetName, 44, 6);
		excelReader.setCellData(xlfile, xlSheetName, 44, 7, actexchangeGainLossForLocalCurrencyPick);
		

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
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRsaveBtn));
		MRsaveBtn.click();

		/*HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : 1");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 45, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 46, 6));
		excelReader.setCellData(xlfile, xlSheetName, 45, 7, actMessage);
		
		String expMessage = expMsg.toString();

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
		
		
		if(actAdjustbills.equalsIgnoreCase(expAdjustbills)  && actBillNewReference.equalsIgnoreCase(expBillNewReference) && actBillTransactionCurrency.equalsIgnoreCase(expBillTransactionCurrency) && 
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
			excelReader.setCellData(xlfile, xlSheetName, 8, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 8, 8, resFail);
			return false;
		}
	}










	public boolean checkEditingJVVATVocher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded Successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		select1stRow_4thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Amount));
		enter_Amount.sendKeys(excelReader.getCellData(xlSheetName, 47, 5));
		enter_Amount.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPartyName));

		String actPartyName = billRefPartyName.getText();
		String expPartyName = excelReader.getCellData(xlSheetName, 47, 6);


		System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);

		int Adjustbills=billRefAdjustBillsGridList.size();

		String actAdjustbills=Integer.toString(Adjustbills);

		String expAdjustbills=excelReader.getCellData(xlSheetName, 48, 6);

		System.err.println("actAdjustbills : "+actAdjustbills +" Value Expected  : "+expAdjustbills);


		int billwiseAdjustBillsDocListcount=billwiseAdjustBillsDocList.size();

		HashSet<String> actbillwiseAdjustBillsDocList  = new HashSet<String>();

		for (int i = 0; i < billwiseAdjustBillsDocListcount; i++) 
		{
			String data		                              = billwiseAdjustBillsDocList.get(i).getText();
			actbillwiseAdjustBillsDocList.add(data);
		}

		String actDocumentNumberText=actbillwiseAdjustBillsDocList.toString();

		String expDocumentNumberText=excelReader.getCellData(xlSheetName, 49, 6);


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

		String expBillNewReference                      =excelReader.getCellData(xlSheetName, 50, 6);
		excelReader.setCellData(xlfile, xlSheetName, 50, 7, actBillNewReference);
		
		String expBillTransactionCurrency               =excelReader.getCellData(xlSheetName, 51, 6);
		excelReader.setCellData(xlfile, xlSheetName, 51, 7, actBillTransactionCurrency);
		
		String expBillBaseCurrency                      =excelReader.getCellData(xlSheetName, 52, 6);
		excelReader.setCellData(xlfile, xlSheetName, 52, 7, actBillBaseCurrency);
		
		String expBillLocalCurrency                     =excelReader.getCellData(xlSheetName, 53, 6);
		excelReader.setCellData(xlfile, xlSheetName, 53, 7, actBillLocalCurrency);
		
		String expBillBalanceNewRefAmount               =excelReader.getCellData(xlSheetName, 54, 6);
		excelReader.setCellData(xlfile, xlSheetName, 54, 7, actBillBalanceNewRefAmount);
		
		String expbillRefAdjustAmountInTransCurency         =excelReader.getCellData(xlSheetName, 55, 6);
		excelReader.setCellData(xlfile, xlSheetName, 55, 7, actbillRefAdjustAmountInTransCurency);
		
		String expbillRefBalanceAmountAdjustInTrnasCurrency =excelReader.getCellData(xlSheetName, 56, 6);
		excelReader.setCellData(xlfile, xlSheetName, 56, 7, actbillRefBalanceAmountAdjustInTrnasCurrency);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefGridFirstRowAdjustmentAmtTxt));
		billRefGridFirstRowAdjustmentAmtTxt.click();

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

		String expBillNewReferencePick                          = excelReader.getCellData(xlSheetName, 57, 6);
		excelReader.setCellData(xlfile, xlSheetName, 57, 7, actBillNewReferencePick);
		
		String expBillTransactionCurrencyPick                   = excelReader.getCellData(xlSheetName, 58, 6);
		excelReader.setCellData(xlfile, xlSheetName, 58, 7, actBillTransactionCurrencyPick);
		
		String expBillBaseCurrencyPick                          = excelReader.getCellData(xlSheetName, 59, 6);
		excelReader.setCellData(xlfile, xlSheetName, 59, 7, actBillBaseCurrencyPick);
		
		String expBillLocalCurrencyPick                         = excelReader.getCellData(xlSheetName, 60, 6);
		excelReader.setCellData(xlfile, xlSheetName, 60, 7, actBillLocalCurrencyPick);
		
		String expBillBalanceNewRefAmountPick                   = excelReader.getCellData(xlSheetName, 61, 6);
		excelReader.setCellData(xlfile, xlSheetName, 61, 7, actBillBalanceNewRefAmountPick);
		
		String expbillRefAdjustAmountInTransCurencyPick         = excelReader.getCellData(xlSheetName, 62, 6);
		excelReader.setCellData(xlfile, xlSheetName, 62, 7, actbillRefAdjustAmountInTransCurencyPick);
		
		String expbillRefBalanceAmountAdjustInTrnasCurrencyPick = excelReader.getCellData(xlSheetName, 63, 6);
		excelReader.setCellData(xlfile, xlSheetName, 63, 7, actbillRefBalanceAmountAdjustInTrnasCurrencyPick);

		String expbreakUpDetailsAccountPick             =excelReader.getCellData(xlSheetName, 64, 6);
		excelReader.setCellData(xlfile, xlSheetName, 64, 7, actbreakUpDetailsAccountPick);
		
		String expbreakUpDetailsDepartmentPick          =excelReader.getCellData(xlSheetName, 65, 6);
		excelReader.setCellData(xlfile, xlSheetName, 65, 7, actbreakUpDetailsDepartmentPick);
		
		/*String expconversationRateBaseCurrencyRatePick  =excelReader.getCellData(xlSheetName, 66, 6);
		excelReader.setCellData(xlfile, xlSheetName, 66, 7, actconversationRateBaseCurrencyRatePick);
		*/
		String expconversationRateLocalCurrencyRatePick =excelReader.getCellData(xlSheetName, 67, 6);
		excelReader.setCellData(xlfile, xlSheetName, 67, 7, actconversationRateLocalCurrencyRatePick);
		
		String expasOnEntryDateTransAmtPick             =excelReader.getCellData(xlSheetName, 68, 6);
		excelReader.setCellData(xlfile, xlSheetName, 68, 7, actasOnEntryDateTransAmtPick);
		
		String expasOnEntryDateBaseConcersationRatePick =excelReader.getCellData(xlSheetName, 69, 6);
		excelReader.setCellData(xlfile, xlSheetName, 69, 7, actasOnEntryDateBaseConcersationRatePick);
		
		String expasOnEntryDateBaseAmountPick           =excelReader.getCellData(xlSheetName, 70, 6);
		excelReader.setCellData(xlfile, xlSheetName, 70, 7, actasOnEntryDateBaseAmountPick);
		
		String expasOnEntryDateLocConversationRatePick  =excelReader.getCellData(xlSheetName, 71, 6);
		excelReader.setCellData(xlfile, xlSheetName, 71, 7, actasOnEntryDateLocConversationRatePick);
		
		String expasOnEntryDateAmtPick                  =excelReader.getCellData(xlSheetName, 72, 6);
		excelReader.setCellData(xlfile, xlSheetName, 72, 7, actasOnEntryDateAmtPick);
		
		String expbalOnAdjstDateTransAmtPick            =excelReader.getCellData(xlSheetName, 73, 6);
		excelReader.setCellData(xlfile, xlSheetName, 73, 7, actbalOnAdjstDateTransAmtPick);
		
		String expbalOnAdjstDateBasrConversionRatePick  =excelReader.getCellData(xlSheetName, 74, 6);
		excelReader.setCellData(xlfile, xlSheetName, 74, 7, actbalOnAdjstDateBasrConversionRatePick);
		
		String expbalOnAdjstDateBaseAmountPick          =excelReader.getCellData(xlSheetName, 75, 6);
		excelReader.setCellData(xlfile, xlSheetName, 75, 7, actbalOnAdjstDateBaseAmountPick);
		
		String expbalOnAdjstDateLocalConversionRatePick =excelReader.getCellData(xlSheetName, 76, 6);
		excelReader.setCellData(xlfile, xlSheetName, 76, 7, actbalOnAdjstDateLocalConversionRatePick);
		
		String expbalOnAdjstDateAmtPick                 =excelReader.getCellData(xlSheetName, 77, 6);
		excelReader.setCellData(xlfile, xlSheetName, 77, 7, actbalOnAdjstDateAmtPick);
		
		String expadjustmentsAmount1Pick                =excelReader.getCellData(xlSheetName, 78, 6);
		excelReader.setCellData(xlfile, xlSheetName, 78, 7, actadjustmentsAmount1Pick);
		
		String expadjustmentsAmount2Pick                =excelReader.getCellData(xlSheetName, 79, 6);
		excelReader.setCellData(xlfile, xlSheetName, 79, 7, actadjustmentsAmount2Pick);
		
		String expadjustmentsAmount3Pick                =excelReader.getCellData(xlSheetName, 80, 6);
		excelReader.setCellData(xlfile, xlSheetName, 80, 7, actadjustmentsAmount3Pick);
		
		String expadjustmentsAmount4Pick                =excelReader.getCellData(xlSheetName, 81, 6);
		excelReader.setCellData(xlfile, xlSheetName, 81, 7, actadjustmentsAmount4Pick);
		
		String expexchangeGainLossForBaseCurrencyPick   =excelReader.getCellData(xlSheetName, 82, 6);
		excelReader.setCellData(xlfile, xlSheetName, 82, 7, actexchangeGainLossForBaseCurrencyPick);
		
		String expexchangeGainLossForLocalCurrencyPick  =excelReader.getCellData(xlSheetName, 83, 6);
		excelReader.setCellData(xlfile, xlSheetName, 83, 7, actexchangeGainLossForLocalCurrencyPick);


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
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRsaveBtn));
		MRsaveBtn.click();

		/*HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : 1");
		expMessage.add("Saving in background.");*/
		
	/*	String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 84, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 85, 6));
		excelReader.setCellData(xlfile, xlSheetName, 84, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/

		
String expMessage1 = excelReader.getCellData(xlSheetName, 84, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 85, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 84, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		if(actAdjustbills.equalsIgnoreCase(expAdjustbills)  && actBillNewReference.equalsIgnoreCase(expBillNewReference) && actBillTransactionCurrency.equalsIgnoreCase(expBillTransactionCurrency) && 
				actBillBaseCurrency.equalsIgnoreCase(expBillBaseCurrency) && actBillLocalCurrency.equalsIgnoreCase(expBillLocalCurrency) && 
				actBillBalanceNewRefAmount.equalsIgnoreCase(expBillBalanceNewRefAmount) && 
				actbillRefAdjustAmountInTransCurency.equalsIgnoreCase(expbillRefAdjustAmountInTransCurency) &&
				actbillRefBalanceAmountAdjustInTrnasCurrency.equalsIgnoreCase(expbillRefBalanceAmountAdjustInTrnasCurrency) &&   

				actBillNewReferencePick.equalsIgnoreCase(expBillNewReferencePick) && actBillTransactionCurrencyPick.equalsIgnoreCase(expBillTransactionCurrencyPick) && 
				actBillBaseCurrencyPick.equalsIgnoreCase(expBillBaseCurrencyPick) && actBillLocalCurrencyPick.equalsIgnoreCase(expBillLocalCurrencyPick) && 
				actBillBalanceNewRefAmountPick.equalsIgnoreCase(expBillBalanceNewRefAmountPick) && actbreakUpDetailsAccountPick.equalsIgnoreCase(expbreakUpDetailsAccountPick) && 
				actbreakUpDetailsDepartmentPick.equalsIgnoreCase(expbreakUpDetailsDepartmentPick) /*&& 
				actconversationRateBaseCurrencyRatePick.equalsIgnoreCase(expconversationRateBaseCurrencyRatePick) */&& actconversationRateLocalCurrencyRatePick.equalsIgnoreCase(expconversationRateLocalCurrencyRatePick) && 
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
			excelReader.setCellData(xlfile, xlSheetName, 47, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 47, 8, resFail);
			return false;
		}
	}








	public boolean checkSavingJVVATVocher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDateCalenderBtn));
		voucherHeaderDueDateCalenderBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDateCalenderTodayBtn));
		voucherHeaderDueDateCalenderTodayBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SHIFT,Keys.HOME,Keys.BACK_SPACE);
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 86, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 87, 5));
		Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_CreditACTxt));
		enter_CreditACTxt.sendKeys(excelReader.getCellData(xlSheetName, 88, 5));
		Thread.sleep(2000);
		enter_CreditACTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enterJVVATTaxCode));	
		enterJVVATTaxCode.sendKeys(excelReader.getCellData(xlSheetName, 89, 5));
		Thread.sleep(2000);
		enterJVVATTaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Amount));
		enter_Amount.sendKeys(excelReader.getCellData(xlSheetName, 90, 5));
		enter_Amount.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPartyName));

		String actPartyName = billRefPartyName.getText();
		String expPartyName = excelReader.getCellData(xlSheetName, 86, 6);
		excelReader.setCellData(xlfile, xlSheetName, 86, 7, actPartyName);

		System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);

		int Adjustbills=billRefAdjustBillsGridList.size();

		String actAdjustbills=Integer.toString(Adjustbills);

		String expAdjustbills=excelReader.getCellData(xlSheetName, 87, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 87, 7, actAdjustbills);

		System.err.println("actAdjustbills : "+actAdjustbills +" Value Expected  : "+expAdjustbills);


		int billwiseAdjustBillsDocListcount=billwiseAdjustBillsDocList.size();

		HashSet<String> actbillwiseAdjustBillsDocList  = new HashSet<String>();

		for (int i = 0; i < billwiseAdjustBillsDocListcount; i++) 
		{
			String data		                              = billwiseAdjustBillsDocList.get(i).getText();
			actbillwiseAdjustBillsDocList.add(data);
		}

		String actDocumentNumberText=actbillwiseAdjustBillsDocList.toString();

		String expDocumentNumberText=excelReader.getCellData(xlSheetName, 88, 6);

		excelReader.setCellData(xlfile, xlSheetName, 88, 7, actDocumentNumberText);

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

		String expBillNewReference                      =excelReader.getCellData(xlSheetName, 89, 6);
		excelReader.setCellData(xlfile, xlSheetName, 89, 7, actBillNewReference);
		
		String expBillTransactionCurrency               =excelReader.getCellData(xlSheetName, 90, 6);
		excelReader.setCellData(xlfile, xlSheetName, 90, 7, actBillTransactionCurrency);
		
		String expBillBaseCurrency                      =excelReader.getCellData(xlSheetName, 91, 6);
		excelReader.setCellData(xlfile, xlSheetName, 91, 7, actBillBaseCurrency);
		
		String expBillLocalCurrency                     =excelReader.getCellData(xlSheetName, 92, 6);
		excelReader.setCellData(xlfile, xlSheetName, 92, 7, actBillLocalCurrency);
		
		String expBillBalanceNewRefAmount               =excelReader.getCellData(xlSheetName, 93, 6);
		excelReader.setCellData(xlfile, xlSheetName, 93, 7, actBillBalanceNewRefAmount);
		
		String expbillRefAdjustAmountInTransCurency         =excelReader.getCellData(xlSheetName, 94, 6);
		excelReader.setCellData(xlfile, xlSheetName, 94, 7, actbillRefAdjustAmountInTransCurency);
		
		String expbillRefBalanceAmountAdjustInTrnasCurrency =excelReader.getCellData(xlSheetName, 95, 6);
		excelReader.setCellData(xlfile, xlSheetName, 95, 7, actbillRefBalanceAmountAdjustInTrnasCurrency);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefGridFirstRowAdjustmentAmtTxt));
		billRefGridFirstRowAdjustmentAmtTxt.click();

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
	//	String actconversationRateBaseCurrencyRatePick   = conversationRateBaseCurrencyRate.getText();
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
		
		String expBillNewReferencePick                          = excelReader.getCellData(xlSheetName, 96, 6);
		excelReader.setCellData(xlfile, xlSheetName, 96, 7, actBillNewReferencePick);
		
		String expBillTransactionCurrencyPick                   = excelReader.getCellData(xlSheetName, 98, 6);
		excelReader.setCellData(xlfile, xlSheetName, 97, 7, actBillTransactionCurrencyPick);
		
		String expBillBaseCurrencyPick                          = excelReader.getCellData(xlSheetName, 98, 6);
		excelReader.setCellData(xlfile, xlSheetName, 98, 7, actBillBaseCurrencyPick);
		
		String expBillLocalCurrencyPick                         = excelReader.getCellData(xlSheetName, 99, 6);
		excelReader.setCellData(xlfile, xlSheetName, 99, 7, actBillLocalCurrencyPick);
		
		String expBillBalanceNewRefAmountPick                   = excelReader.getCellData(xlSheetName, 100, 6);
		excelReader.setCellData(xlfile, xlSheetName, 100, 7, actBillBalanceNewRefAmountPick);
		
		String expbillRefAdjustAmountInTransCurencyPick         = excelReader.getCellData(xlSheetName, 101, 6);
		excelReader.setCellData(xlfile, xlSheetName, 101, 7, actbillRefAdjustAmountInTransCurencyPick);
		
		String expbillRefBalanceAmountAdjustInTrnasCurrencyPick = excelReader.getCellData(xlSheetName, 102, 6);
		excelReader.setCellData(xlfile, xlSheetName, 102, 7, actbillRefBalanceAmountAdjustInTrnasCurrencyPick);

		String expbreakUpDetailsAccountPick             =excelReader.getCellData(xlSheetName, 103, 6);
		excelReader.setCellData(xlfile, xlSheetName, 103, 7, actbreakUpDetailsAccountPick);
		
		String expbreakUpDetailsDepartmentPick          =excelReader.getCellData(xlSheetName, 104, 6);
		excelReader.setCellData(xlfile, xlSheetName, 104, 7, actbreakUpDetailsDepartmentPick);
		
		/*String expconversationRateBaseCurrencyRatePick  =excelReader.getCellData(xlSheetName, 105, 6);
		excelReader.setCellData(xlfile, xlSheetName, 105, 7, actconversationRateBaseCurrencyRatePick);*/
		
		String expconversationRateLocalCurrencyRatePick =excelReader.getCellData(xlSheetName, 106, 6);
		excelReader.setCellData(xlfile, xlSheetName, 106, 7, actconversationRateLocalCurrencyRatePick);
		
		String expasOnEntryDateTransAmtPick             =excelReader.getCellData(xlSheetName, 107, 6);
		excelReader.setCellData(xlfile, xlSheetName, 107, 7, actasOnEntryDateTransAmtPick);
		
		String expasOnEntryDateBaseConcersationRatePick =excelReader.getCellData(xlSheetName, 108, 6);
		excelReader.setCellData(xlfile, xlSheetName, 108, 7, actasOnEntryDateBaseConcersationRatePick);
		
		String expasOnEntryDateBaseAmountPick           =excelReader.getCellData(xlSheetName, 109, 6);
		excelReader.setCellData(xlfile, xlSheetName, 109, 7, actasOnEntryDateBaseAmountPick);
		
		String expasOnEntryDateLocConversationRatePick  =excelReader.getCellData(xlSheetName, 110, 6);
		excelReader.setCellData(xlfile, xlSheetName, 110, 7, actasOnEntryDateLocConversationRatePick);
		
		String expasOnEntryDateAmtPick                  =excelReader.getCellData(xlSheetName, 111, 6);
		excelReader.setCellData(xlfile, xlSheetName, 111, 7, actasOnEntryDateAmtPick);
		
		String expbalOnAdjstDateTransAmtPick            =excelReader.getCellData(xlSheetName, 112, 6);
		excelReader.setCellData(xlfile, xlSheetName, 112, 7, actbalOnAdjstDateTransAmtPick);
		
		String expbalOnAdjstDateBasrConversionRatePick  =excelReader.getCellData(xlSheetName, 113, 6);
		excelReader.setCellData(xlfile, xlSheetName, 113, 7, actbalOnAdjstDateBasrConversionRatePick);
		
		String expbalOnAdjstDateBaseAmountPick          =excelReader.getCellData(xlSheetName, 114, 6);
		excelReader.setCellData(xlfile, xlSheetName, 114, 7, actbalOnAdjstDateBaseAmountPick);
		
		String expbalOnAdjstDateLocalConversionRatePick =excelReader.getCellData(xlSheetName, 115, 6);
		excelReader.setCellData(xlfile, xlSheetName, 115, 7, actbalOnAdjstDateLocalConversionRatePick);
		
		String expbalOnAdjstDateAmtPick                 =excelReader.getCellData(xlSheetName, 116, 6);
		excelReader.setCellData(xlfile, xlSheetName, 116, 7, actbalOnAdjstDateAmtPick);
		
		String expadjustmentsAmount1Pick                =excelReader.getCellData(xlSheetName, 117, 6);
		excelReader.setCellData(xlfile, xlSheetName, 117, 7, actadjustmentsAmount1Pick);
		
		String expadjustmentsAmount2Pick                =excelReader.getCellData(xlSheetName, 118, 6);
		excelReader.setCellData(xlfile, xlSheetName, 118, 7, actadjustmentsAmount2Pick);
		
		String expadjustmentsAmount3Pick                =excelReader.getCellData(xlSheetName, 119, 6);
		excelReader.setCellData(xlfile, xlSheetName, 119, 7, actadjustmentsAmount3Pick);
		
		String expadjustmentsAmount4Pick                =excelReader.getCellData(xlSheetName, 120, 6);
		excelReader.setCellData(xlfile, xlSheetName, 120, 7, actadjustmentsAmount4Pick);
		
		String expexchangeGainLossForBaseCurrencyPick   =excelReader.getCellData(xlSheetName, 121, 6);
		excelReader.setCellData(xlfile, xlSheetName, 121, 7, actexchangeGainLossForBaseCurrencyPick);
		
		String expexchangeGainLossForLocalCurrencyPick  =excelReader.getCellData(xlSheetName, 122, 6);
		excelReader.setCellData(xlfile, xlSheetName, 122, 7, actexchangeGainLossForLocalCurrencyPick);

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
	//	System.out.println("actconversationRateBaseCurrRatePick:   "+ actconversationRateBaseCurrencyRatePick+ "  " +"expconversationRateBaseCurrencyRatePick :"+expconversationRateBaseCurrencyRatePick);
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
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRsaveBtn));
		MRsaveBtn.click();

		/*HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet<String>();

		expMessage.add("Voucher saved successfully : 2");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 123, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 124, 6));
		excelReader.setCellData(xlfile, xlSheetName, 123, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		
String expMessage1 = excelReader.getCellData(xlSheetName, 123, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 124, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 123, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		if(actAdjustbills.equalsIgnoreCase(expAdjustbills)  && actBillNewReference.equalsIgnoreCase(expBillNewReference) && actBillTransactionCurrency.equalsIgnoreCase(expBillTransactionCurrency) && 
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
			excelReader.setCellData(xlfile, xlSheetName, 86, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 86, 8, resFail);
			return false;
		}
	}








	public boolean checkDeletingJVVATViewVoucher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded Successfully");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		String expMessage = excelReader.getCellData(xlSheetName, 125, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 125, 7, actMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();

		Thread.sleep(1000);

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 125, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 125, 8, resFail);
			return false;
		}	
	}









	public boolean checkVerifingDetailsOfSavedJVVATViewVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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

		expVouchersList.add(excelReader.getCellData(xlSheetName, 126, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 127, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 128, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 129, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 130, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 131, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 132, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 133, 6));
		
		excelReader.setCellData(xlfile, xlSheetName, 126, 7, actVouchersList.toString());

		System.out.println("**********************checkVerifingDetailsOfSavedJVVATViewVoucherInHomepage*********************");

		System.out.println(actVouchersList);
		System.out.println(expVouchersList);

		if (actVouchersList.equals(expVouchersList)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 126, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 126, 8, resFail);
			return false;
		}
	}





	public JVVATViewPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
