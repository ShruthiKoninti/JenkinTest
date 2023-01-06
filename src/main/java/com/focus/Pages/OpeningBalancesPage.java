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

public class OpeningBalancesPage extends BaseEngine 
{
	private static String xlfile = getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlSheetName = "OpeningBalancePage";
	
	public static boolean checkOpeningBalanceCustomizationRemovingLinkStatusColumn() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsJournalsMenu));
		financialsTransactionsJournalsMenu.click();

		Thread.sleep(1000);
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);",OpeningBalanceMenu);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(OpeningBalanceMenu));
		OpeningBalanceMenu.click();
		
		//JavascriptExecutor js=
		
		
	//	ClickSubMenu(JournalsSubMenusList, excelReader.getCellData(xlSheetName, 8, 5));

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int actBeforeDeletingCount = homePageBodyHeaderList.size();
		int expBeforeDeletingCount = Integer.valueOf(excelReader.getCellData(xlSheetName, 8, 6));
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, Integer.toString(actBeforeDeletingCount));

		System.err.println("Header List Count Before Deleting Link Status : "+actBeforeDeletingCount+"  Value Expected  "+expBeforeDeletingCount);

		ArrayList<String > beforeDelete = new ArrayList<String>();

		for (int i = 0; i < actBeforeDeletingCount; i++) 
		{
			String data = homePageBodyHeaderList.get(i).getText();
			beforeDelete.add(data);
		}

		String actHeaderListBeforeDelete = beforeDelete.toString();
		String expHeaderListBeforeDelete = excelReader.getCellData(xlSheetName, 9, 6);
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, actHeaderListBeforeDelete);

		System.err.println(actHeaderListBeforeDelete);
		System.err.println(expHeaderListBeforeDelete);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeBtn));
		customizeBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(custSaveBtn));

		int custBodyHeaderListCount = custBodyHeaderList.size();

		for (int i = 0; i < custBodyHeaderListCount; i++)
		{
			String data = custBodyHeaderList.get(i).getText();

			System.out.println(i+" "+data);

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 9, 5))) 
			{
				custBodyHeaderList.get(i).click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(custRemoveBtn));
				custRemoveBtn.click();

				Thread.sleep(1000);

				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(custSaveBtn));
		custSaveBtn.click();

		String expMessage  = excelReader.getCellData(xlSheetName, 10, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 10, 7, actMessage);
		
		Thread.sleep(2000);

		int actAfterDeletingCount = homePageBodyHeaderList.size();
		int expAfterDeletingCount = Integer.valueOf(excelReader.getCellData(xlSheetName, 11, 6));
		excelReader.setCellData(xlfile, xlSheetName, 11, 7, Integer.toString(actAfterDeletingCount));
		
		System.err.println("Header List Count After Deleting Link Status : "+actAfterDeletingCount+"  Value Expected  "+expAfterDeletingCount);

		ArrayList<String > afterDelete = new ArrayList<String>();

		for (int i = 0; i < actAfterDeletingCount; i++) 
		{
			String data = homePageBodyHeaderList.get(i).getText();
			afterDelete.add(data);
		}

		String actHeaderListAfterDelete = afterDelete.toString();
		String expHeaderListAfterDelete = excelReader.getCellData(xlSheetName, 12, 6);
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actHeaderListAfterDelete);

		System.err.println(actHeaderListAfterDelete);
		System.err.println(expHeaderListAfterDelete);

		if (actHeaderListBeforeDelete.equalsIgnoreCase(expHeaderListBeforeDelete) && actMessage.equalsIgnoreCase(expMessage) 
				&& actHeaderListAfterDelete.equalsIgnoreCase(expHeaderListAfterDelete)) 
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








	public static boolean checkOpeningBalanceCustomizationHideBalanceLinkValueColumn() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeBtn));
		customizeBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(custSaveBtn));

		int custBodyHeaderListCount = custBodyHeaderList.size();

		for (int i = 0; i < custBodyHeaderListCount; i++)
		{
			String data = custBodyHeaderList.get(i).getText();

			System.out.println(i+" "+data);

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 13, 5))) 
			{
				custBodyHeaderList.get(i).click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(hideThisColumnChkbox));
				
				if (hideThisColumnChkboxIsSelected.isSelected()==false) 
				{
					hideThisColumnChkbox.click();
				}

				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(custSaveBtn));
		custSaveBtn.click();

		String expHideMessage  = excelReader.getCellData(xlSheetName, 13, 6);
		String actHideMessage = checkValidationMessage(expHideMessage);
		excelReader.setCellData(xlfile, xlSheetName, 13, 7, actHideMessage);

		Thread.sleep(2000);

		int actAfterHidingCount = homePageBodyHeaderList.size();
		int expAfterHidingCount = Integer.valueOf(excelReader.getCellData(xlSheetName, 14, 6));
		excelReader.setCellData(xlfile, xlSheetName, 14, 7, Integer.toString(actAfterHidingCount));

		System.err.println("Header List Count After Hiding Balance link value : "+actAfterHidingCount+"  Value Expected  "+expAfterHidingCount);

		ArrayList<String > afterHiding = new ArrayList<String>();

		for (int i = 0; i < actAfterHidingCount; i++) 
		{
			String data = homePageBodyHeaderList.get(i).getText();
			afterHiding.add(data);
		}

		String actHeaderListAfterHiding = afterHiding.toString();
		String expHeaderListAfterHiding = excelReader.getCellData(xlSheetName, 15, 6);
		excelReader.setCellData(xlfile, xlSheetName, 15, 7, actHeaderListAfterHiding);

		System.err.println(actHeaderListAfterHiding);
		System.err.println(expHeaderListAfterHiding);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeBtn));
		customizeBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(custSaveBtn));

		for (int i = 0; i < custBodyHeaderListCount; i++)
		{
			String data = custBodyHeaderList.get(i).getText();

			System.out.println(i+" "+data);

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 14, 5))) 
			{
				custBodyHeaderList.get(i).click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(hideThisColumnChkbox));
				
				if (hideThisColumnChkboxIsSelected.isSelected()==true) 
				{
					hideThisColumnChkbox.click();
				}

				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(custSaveBtn));
		custSaveBtn.click();

		String expUnhideMessage  = excelReader.getCellData(xlSheetName, 16, 6);
		String actUnhideMessage = checkValidationMessage(expUnhideMessage);
		excelReader.setCellData(xlfile, xlSheetName, 16, 7, actUnhideMessage);

		Thread.sleep(2000);

		int actAfterUnHidingCount = homePageBodyHeaderList.size();
		int expAfterUnHidingCount = Integer.valueOf(excelReader.getCellData(xlSheetName, 17, 6));
		excelReader.setCellData(xlfile, xlSheetName, 17, 7, Integer.toString(actAfterUnHidingCount));

		System.err.println("Header List Count After UnHiding Balance link value : "+actAfterUnHidingCount+"  Value Expected  "+expAfterUnHidingCount);

		ArrayList<String > afterUnHiding = new ArrayList<String>();

		for (int i = 0; i < actAfterUnHidingCount; i++) 
		{
			String data = homePageBodyHeaderList.get(i).getText();
			afterUnHiding.add(data);
		}

		String actHeaderListAfterUnHiding = afterUnHiding.toString();
		String expHeaderListAfterUnHiding = excelReader.getCellData(xlSheetName, 18, 6);
		excelReader.setCellData(xlfile, xlSheetName, 18, 7, actHeaderListAfterUnHiding);

		System.err.println(actHeaderListAfterUnHiding);
		System.err.println(expHeaderListAfterUnHiding);

		if (actHideMessage.equalsIgnoreCase(expHideMessage) && actHeaderListAfterHiding.equalsIgnoreCase(expHeaderListAfterHiding) 
				&& actUnhideMessage.equalsIgnoreCase(expUnhideMessage) && actHeaderListAfterUnHiding.equalsIgnoreCase(expHeaderListAfterUnHiding)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 13, 7, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 13, 7, resFail);
			return false;
		}
	}





	



	public static boolean checkOpeningBalanceCustomizationAddingDepartmentColumn() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeBtn));
		customizeBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(custSaveBtn));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(custLeftPannelTransExpandBtn));
		custLeftPannelTransExpandBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(custTransExtraFieldExpandBtn));
		custTransExtraFieldExpandBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cusDepartmentExpandBtn));
		cusDepartmentExpandBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cusDepName));

		getAction().doubleClick(cusDepName).build().perform();

		int cusGridHeaderListCount = cusGridHeaderList.size();

		for(int i=1;i<cusGridHeaderListCount;i++)
		{
			String data = cusGridHeaderList.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 19, 5)))
			{
				cusGridHeaderList.get(i).click();
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(custSaveBtn));
		custSaveBtn.click();

		String expMessage  = excelReader.getCellData(xlSheetName, 19, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 19, 7, actMessage);

		Thread.sleep(2000);

		int actAfterAddingCount = homePageBodyHeaderList.size();
		int expAfterAddingCount = Integer.valueOf(excelReader.getCellData(xlSheetName, 20, 6));
		excelReader.setCellData(xlfile, xlSheetName, 20, 7, Integer.toString(actAfterAddingCount));

		System.err.println("Header List Count After Adding Balance link value : "+actAfterAddingCount+"  Value Expected  "+expAfterAddingCount);

		ArrayList<String > afterAdding = new ArrayList<String>();

		for (int i = 0; i < actAfterAddingCount; i++) 
		{
			String data = homePageBodyHeaderList.get(i).getText();
			afterAdding.add(data);
		}

		String actHeaderListAfterAdding = afterAdding.toString();
		String expHeaderListAfterAdding = excelReader.getCellData(xlSheetName, 21, 6);
		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actHeaderListAfterAdding);

		System.err.println(actHeaderListAfterAdding);
		System.err.println(expHeaderListAfterAdding);

		if (actMessage.equalsIgnoreCase(expMessage) && actHeaderListAfterAdding.equalsIgnoreCase(expHeaderListAfterAdding)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 19, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 19, 8, resFail);
			return false;
		}
	}



	



	public static boolean checkSavingOpeningBalanceVoucher1WithCurrencyAsINR() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsJournalsMenu));
		financialsTransactionsJournalsMenu.click();

		Thread.sleep(1000);
		
		//ClickSubMenu(JournalsSubMenusList, excelReader.getCellData(xlSheetName, 22, 5));

		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);",OpeningBalanceMenu);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(OpeningBalanceMenu));
		OpeningBalanceMenu.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 23, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 24, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 25, 5));
		Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_DebitTxt));
		enter_DebitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_CreditTxt));
		enter_CreditTxt.sendKeys(excelReader.getCellData(xlSheetName, 26, 5));
		enter_CreditTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		String actPartyName = billRefPartyName.getText();
		String expPartyName = excelReader.getCellData(xlSheetName, 22, 6);

		System.out.println("***********************************checkSavingOpeningBalanceVoucher1WithCurrencyAsINR*************************************");

		System.out.println("Bill wise Screen Cutomer Name "  + actPartyName + "  Value Expected  " + expPartyName);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefAdjustBillsGrid));

		
		String actAdjustbills= billRefAdjustBillsGrid.getAttribute("data-totalrows");
		String expAdjustbills=excelReader.getCellData(xlSheetName, 23, 6);
		excelReader.setCellData(xlfile, xlSheetName, 23, 7, actAdjustbills);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));

		String actBillNewReference                           = billRefNewReferenceTxt.getAttribute("value");
		String actBillTransactionCurrency                    = billRefTransactionCurency.getText();
		String actBillBaseCurrency                           = billRefBaseCurrency.getText();
		String actBillLocalCurrency                          = localCurrencyDhs.getText();
		String actBillBalanceNewRefAmount                    = balanceNewReferenceAmt.getText();
		String actbillRefAdjustAmountInTransCurency          = billRefAdjustAmountInTransCurency.getText();
		String actbillRefBalanceAmountAdjustInTrnasCurrency  = billRefBalanceAmountAdjustInTrnasCurrency.getText();

		String expBillNewReference                      =excelReader.getCellData(xlSheetName, 24, 6);
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actBillNewReference);
		
		String expBillTransactionCurrency               =excelReader.getCellData(xlSheetName, 25, 6);
		excelReader.setCellData(xlfile, xlSheetName, 25, 7, actBillTransactionCurrency);
		
		String expBillBaseCurrency                      =excelReader.getCellData(xlSheetName, 26, 6);
		excelReader.setCellData(xlfile, xlSheetName, 26, 7, actBillBaseCurrency);
		
		String expBillLocalCurrency                     =excelReader.getCellData(xlSheetName, 27, 6);
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actBillLocalCurrency);
		
		String expBillBalanceNewRefAmount               =excelReader.getCellData(xlSheetName, 28, 6);
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actBillBalanceNewRefAmount);
		
		String expbillRefAdjustAmountInTransCurency         =excelReader.getCellData(xlSheetName, 29, 6);
		excelReader.setCellData(xlfile, xlSheetName, 29, 7, actbillRefAdjustAmountInTransCurency);
		
		String expbillRefBalanceAmountAdjustInTrnasCurrency =excelReader.getCellData(xlSheetName, 30, 6);
		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actbillRefBalanceAmountAdjustInTrnasCurrency);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();

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

		//String actconversationRateBaseCurrencyRatePick   = conversationRateBaseCurrencyRate.getText();
		String actconversationRateLocalCurrencyRatePick  = conversationRateLocalCurrencyRate.getText();

		
		String expBillNewReferencePick                          = excelReader.getCellData(xlSheetName, 31, 6);
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actBillNewReferencePick);
		
		String expBillTransactionCurrencyPick                   = excelReader.getCellData(xlSheetName, 32, 6);
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actBillTransactionCurrencyPick);
		
		String expBillBaseCurrencyPick                          = excelReader.getCellData(xlSheetName, 33, 6);
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actBillBaseCurrencyPick);
		
		String expBillLocalCurrencyPick                         = excelReader.getCellData(xlSheetName, 34, 6);
		excelReader.setCellData(xlfile, xlSheetName, 34, 7, actBillLocalCurrencyPick);
		
		String expBillBalanceNewRefAmountPick                   = excelReader.getCellData(xlSheetName, 35, 6);
		excelReader.setCellData(xlfile, xlSheetName, 35, 7, actBillBalanceNewRefAmountPick);
		
		String expbillRefAdjustAmountInTransCurencyPick         = excelReader.getCellData(xlSheetName, 36, 6);
		excelReader.setCellData(xlfile, xlSheetName, 36, 7, actbillRefAdjustAmountInTransCurencyPick);
		
		String expbillRefBalanceAmountAdjustInTrnasCurrencyPick = excelReader.getCellData(xlSheetName, 37, 6);
		excelReader.setCellData(xlfile, xlSheetName, 37, 7, actbillRefBalanceAmountAdjustInTrnasCurrencyPick);
		
	/*	String expconversationRateBaseCurrencyRatePick = excelReader.getCellData(xlSheetName, 38, 6); 
		excelReader.setCellData(xlfile, xlSheetName, 38, 7, actconversationRateBaseCurrencyRatePick);
		*/
		String expconversationRateLocalCurrencyRatePick = excelReader.getCellData(xlSheetName, 39, 6);
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actconversationRateLocalCurrencyRatePick);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingBalancesSaveBtn));
		openingBalancesSaveBtn.click();

/*		HashSet<String> actMessage = new HashSet();

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

		//expMsg.add(excelReader.getCellData(xlSheetName, 40, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 41, 6));
		
		String expMessage = expMsg.toString();
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 40, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 41, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		

		System.out.println("Bill reference Adjustment Bills  :"+ actAdjustbills+"                          "+ "expDate :"+expAdjustbills);
		System.out.println("actBillNewReference :             "+ actBillNewReference +"                    "+ "expBillNewReference :"+expBillNewReference);
		System.out.println("actBillTransactionCurrency       :"+ actBillTransactionCurrency + "            " + "expBillTransactionCurrency :"+expBillTransactionCurrency);
		System.out.println("actBillBaseCurrency :             "+ actBillBaseCurrency + "                   " +"expBillBaseCurrency :"+expBillBaseCurrency);
		System.out.println("actBillLocalCurrency :            "+ actBillLocalCurrency+ "                   " +"expBillLocalCurrency :"+expBillLocalCurrency);
		System.out.println("actBillBalanceNewRefAmount :      "+ actBillBalanceNewRefAmount + "            " + "expBillBalanceNewRefAmount :"+expBillBalanceNewRefAmount);
		System.out.println("billRefAdjustAmountInTransCurencyPick  :"+actbillRefAdjustAmountInTransCurencyPick+ " expected "+expbillRefAdjustAmountInTransCurencyPick);
		System.out.println("billRefAdjustAmountInTransCurencyPick  :"+actbillRefAdjustAmountInTransCurencyPick+ " expected "+expbillRefAdjustAmountInTransCurencyPick);
		System.out.println("actbillRefAdjustAmountInTransCurency :"+ actbillRefAdjustAmountInTransCurency+ "       " +"expbillRefAdjustAmountInTransCurency :"+expbillRefAdjustAmountInTransCurency);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrency :"+ actbillRefBalanceAmountAdjustInTrnasCurrency + "       " +"expbillRefBalanceAmountAdjustInTrnasCurrency :"+expbillRefBalanceAmountAdjustInTrnasCurrency);

		// Pick

		System.out.println("actBillNewReferencePick :              "+ actBillNewReferencePick +"                " + "expBillNewReferencePick :"+expBillNewReferencePick);
		System.out.println("actBillTransactionCurrencyPick :       "+ actBillTransactionCurrencyPick + "        " + "expBillTransactionCurrencyPick :"+expBillTransactionCurrencyPick);
		System.out.println("actBillBaseCurrencyPick :              "+ actBillBaseCurrencyPick + "               " + "expBillBaseCurrencyPick :"+expBillBaseCurrencyPick);
		System.out.println("actBillLocalCurrencyPick :             "+ actBillLocalCurrency+ "                   " + "expBillLocalCurrencyPick :"+expBillLocalCurrency);
		System.out.println("actBillBalanceNewRefAmountPick :       "+ actBillBalanceNewRefAmountPick +          " " + "expBillBalanceNewRefAmountPick :"+expBillBalanceNewRefAmountPick);
		//System.out.println("actconversationRateBaseCurrRatePick:   "+ actconversationRateBaseCurrencyRatePick+  "  " +"expconversationRateBaseCurrencyRatePick :"+expconversationRateBaseCurrencyRatePick);
		System.out.println("actconversationRateLocalCurRatePick :  "+ actconversationRateLocalCurrencyRatePick+ " " + "expconversationRateLocalCurrencyRatePick :"+expconversationRateLocalCurrencyRatePick);

		if(actPartyName.equalsIgnoreCase(expPartyName) && actAdjustbills.equalsIgnoreCase(expAdjustbills) && actBillNewReference.equalsIgnoreCase(expBillNewReference) && actBillTransactionCurrency.equalsIgnoreCase(expBillTransactionCurrency) && 
				actBillBaseCurrency.equalsIgnoreCase(expBillBaseCurrency) && actBillLocalCurrency.equalsIgnoreCase(expBillLocalCurrency) && 
				actBillBalanceNewRefAmount.equalsIgnoreCase(expBillBalanceNewRefAmount) && 
				actbillRefAdjustAmountInTransCurency.equalsIgnoreCase(expbillRefAdjustAmountInTransCurency) &&
				actbillRefBalanceAmountAdjustInTrnasCurrency.equalsIgnoreCase(expbillRefBalanceAmountAdjustInTrnasCurrency) && 

				actBillNewReferencePick.equalsIgnoreCase(expBillNewReferencePick) && actBillTransactionCurrencyPick.equalsIgnoreCase(expBillTransactionCurrencyPick) && 
				actBillBaseCurrencyPick.equalsIgnoreCase(expBillBaseCurrencyPick) && actBillLocalCurrencyPick.equalsIgnoreCase(expBillLocalCurrencyPick) && 
				actBillBalanceNewRefAmountPick.equalsIgnoreCase(expBillBalanceNewRefAmountPick)/* &&
				actconversationRateBaseCurrencyRatePick.equalsIgnoreCase(expconversationRateBaseCurrencyRatePick)*/ && actconversationRateLocalCurrencyRatePick.equalsIgnoreCase(expconversationRateLocalCurrencyRatePick)

				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
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







	public static boolean checkVerifyAndDeleteOpeningBalanceVoucher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		String actDocno               = documentNumberTxt.getAttribute("value");
		String actCurrency            = voucherHeaderCurrency.getAttribute("value");
		String actDepartment          = departmentTxt.getAttribute("value");
		String actExchangeRate        = voucherHeaderExchangeRate.getAttribute("value");
		String actLocExchangeRate     = voucherHeaderLocalExchangeRate.getAttribute("value");
		String actR1Account           = select1stRow_1stColumn.getText();
		String actR1Credit            = select1stRow_3rdColumn.getText();
		String actR1Ref               = select1stRow_4thColumn.getText();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vocFooterExpandBtn));
		vocFooterExpandBtn.click();
		
		Thread.sleep(1000);
		
		String actFooterCreditAmt     = vocFooterCreditAmount.getText();
		String actVoucherDebitAmt     = vocFooterdebitAmount.getText();

		String expDocno               = excelReader.getCellData(xlSheetName, 42, 6);
		excelReader.setCellData(xlfile, xlSheetName, 42, 7, actDocno);
		
		String expCurrency            = excelReader.getCellData(xlSheetName, 43, 6);
		excelReader.setCellData(xlfile, xlSheetName, 43, 7, actCurrency);
		
		String expDepartment          = excelReader.getCellData(xlSheetName, 44, 6);
		excelReader.setCellData(xlfile, xlSheetName, 43, 7, actDepartment);
		
		String expExchangeRate        = excelReader.getCellData(xlSheetName, 45, 6);
		excelReader.setCellData(xlfile, xlSheetName, 44, 7, actExchangeRate);
		
		String expLocExchangeRate     = excelReader.getCellData(xlSheetName, 46, 6);
		excelReader.setCellData(xlfile, xlSheetName, 45, 7, actLocExchangeRate);
		
		String expR1Account           = excelReader.getCellData(xlSheetName, 47, 6);
		excelReader.setCellData(xlfile, xlSheetName, 46, 7, actR1Account);
		
		String expR1Credit            = excelReader.getCellData(xlSheetName, 48, 6);
		excelReader.setCellData(xlfile, xlSheetName, 47, 7, actR1Credit);
		
		String expR1Ref               = excelReader.getCellData(xlSheetName, 49, 6);
		excelReader.setCellData(xlfile, xlSheetName, 48, 7, actR1Ref);
		
		String expFooterCreditAmt     = excelReader.getCellData(xlSheetName, 50, 6);
		excelReader.setCellData(xlfile, xlSheetName, 49, 7, actFooterCreditAmt);
		
		String expVoucherDebitAmt     = excelReader.getCellData(xlSheetName, 51, 6);
		excelReader.setCellData(xlfile, xlSheetName, 50, 7, actVoucherDebitAmt);
		
		System.out.println("*******************************checkVerifyAndDeleteOpeningBalanceVoucher1**************************");

		System.out.println("Entry Page Document Number    "+actDocno           +"  value Expected  "+expDocno);
		System.out.println("Entry Page Currency           "+actCurrency        +"  value Expected  "+expCurrency);
		System.out.println("Entry Page Department         "+actDepartment      +"  value Expected  "+expDepartment);
		System.out.println("Entry Page Exchange Rate      "+actExchangeRate    +"  value Expected  "+expExchangeRate);
		System.out.println("Entry Page Department         "+actLocExchangeRate +"  value Expected  "+expLocExchangeRate);
		System.out.println("Entry Page R1Account          "+actR1Account       +"  value Expected  "+expR1Account);
		System.out.println("Entry Page R1Credit           "+actR1Credit        +"  value Expected  "+expR1Credit);
		System.out.println("Entry Page R1Ref              "+actR1Ref           +"  value Expected  "+expR1Ref);

		System.out.println("Entry Page Debit Amount       "+actVoucherDebitAmt +"  value Expected  "+expVoucherDebitAmt);
		System.out.println("Entry Page Credit Amount      "+actFooterCreditAmt +"  value Expected  "+expFooterCreditAmt);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		String expMessage = excelReader.getCellData(xlSheetName, 52, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 52, 7, actMessage);

		if(actDocno.equalsIgnoreCase(expDocno) && actCurrency.equalsIgnoreCase(expCurrency) &&actDepartment.equalsIgnoreCase(expDepartment) 
				&& actExchangeRate.equalsIgnoreCase(expExchangeRate) && actLocExchangeRate.equalsIgnoreCase(expLocExchangeRate) 
				&& actR1Account.equalsIgnoreCase(expR1Account) && actR1Credit.equalsIgnoreCase(expR1Credit) && actR1Ref.equalsIgnoreCase(expR1Ref)
				&& actVoucherDebitAmt.equalsIgnoreCase(expVoucherDebitAmt) && actFooterCreditAmt.equalsIgnoreCase(expFooterCreditAmt)
				&& actMessage.equalsIgnoreCase(expMessage))
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










	public static boolean checkSavingOpeningBalanceVoucher1WithCurrencyAsAED() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 53, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 54, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 55, 5));
		Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_DebitTxt));
		enter_DebitTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_CreditTxt));
		enter_CreditTxt.sendKeys(excelReader.getCellData(xlSheetName, 56, 5));
		enter_CreditTxt.sendKeys(Keys.TAB);
		
		System.out.println("***********************************checkSavingOpeningBalanceVoucher1WithCurrencyAsINR*************************************");
		
		String actR1PartyName = billRefPartyName.getText();
		String expR1PartyName = excelReader.getCellData(xlSheetName, 53, 6);
		excelReader.setCellData(xlfile, xlSheetName, 53, 7, actR1PartyName);
		
		System.out.println("Bill wise Screen Cutomer Name "  + actR1PartyName + "  Value Expected  " + expR1PartyName);
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefAdjustBillsGrid));
			
		String actR1Adjustbills=((WebElement) billRefAdjustBillsGrid).getAttribute("data-totalrows");
		String expR1Adjustbills=excelReader.getCellData(xlSheetName, 54, 6);
		excelReader.setCellData(xlfile, xlSheetName, 54, 7, actR1Adjustbills);
		
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
        
        String actR1BillNewReference                           = billRefNewReferenceTxt.getAttribute("value");
		String actR1BillTransactionCurrency                    = billRefTransactionCurency.getText();
		String actR1BillBaseCurrency                           = billRefBaseCurrency.getText();
		String actR1BillLocalCurrency                          = localCurrencyDhs.getText();
		String actR1BillBalanceNewRefAmount                    = balanceNewReferenceAmt.getText();
        String actR1billRefAdjustAmountInTransCurency          = billRefAdjustAmountInTransCurency.getText();
        String actR1billRefBalanceAmountAdjustInTrnasCurrency  = billRefBalanceAmountAdjustInTrnasCurrency.getText();
        
        String expR1BillNewReference                      =excelReader.getCellData(xlSheetName, 55, 6);
		excelReader.setCellData(xlfile, xlSheetName, 55, 7, actR1BillNewReference);
		
		String expR1BillTransactionCurrency               =excelReader.getCellData(xlSheetName, 56, 6);
		excelReader.setCellData(xlfile, xlSheetName, 56, 7, actR1BillTransactionCurrency);
		
		String expR1BillBaseCurrency                      =excelReader.getCellData(xlSheetName, 57, 6);
		excelReader.setCellData(xlfile, xlSheetName, 57, 7, actR1BillBaseCurrency);
		
		String expR1BillLocalCurrency                     =excelReader.getCellData(xlSheetName, 58, 6);
		excelReader.setCellData(xlfile, xlSheetName, 58, 7, actR1BillLocalCurrency);
		
		String expR1BillBalanceNewRefAmount               =excelReader.getCellData(xlSheetName, 59, 6);
		excelReader.setCellData(xlfile, xlSheetName, 59, 7, actR1BillBalanceNewRefAmount);
		
        String expR1billRefAdjustAmountInTransCurency         =excelReader.getCellData(xlSheetName, 60, 6);
        excelReader.setCellData(xlfile, xlSheetName, 60, 7, actR1billRefAdjustAmountInTransCurency);
        
        String expR1billRefBalanceAmountAdjustInTrnasCurrency =excelReader.getCellData(xlSheetName, 61, 6);
        excelReader.setCellData(xlfile, xlSheetName, 61, 7, actR1billRefBalanceAmountAdjustInTrnasCurrency);
        
 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
 		billRefNewReferenceTxt.click();
 		
 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
 		billRefPickIcon.click();
    	
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
      
        String actR1BillNewReferencePick                           = billRefNewReferenceTxt.getAttribute("value");
 		String actR1BillTransactionCurrencyPick                    = billRefTransactionCurency.getText();
 		String actR1BillBaseCurrencyPick                           = billRefBaseCurrency.getText();
 		String actR1BillLocalCurrencyPick                          = localCurrencyDhs.getText();
 		String actR1BillBalanceNewRefAmountPick                    = balanceNewReferenceAmt.getText();
 		String actR1billRefAdjustAmountInTransCurencyPick          = billRefAdjustAmountInTransCurency.getText();
	    String actR1billRefBalanceAmountAdjustInTrnasCurrencyPick  = billRefBalanceAmountAdjustInTrnasCurrency.getText();
 		
 	    //String actR1conversationRateBaseCurrencyRatePick   = conversationRateBaseCurrencyRate.getText();
        String actR1conversationRateLocalCurrencyRatePick  = conversationRateLocalCurrencyRate.getText();
        
        
        String expR1BillNewReferencePick                          = excelReader.getCellData(xlSheetName, 62, 6);
        excelReader.setCellData(xlfile, xlSheetName, 62, 7, actR1BillNewReferencePick);
        
 		String expR1BillTransactionCurrencyPick                   = excelReader.getCellData(xlSheetName, 63, 6);
 		excelReader.setCellData(xlfile, xlSheetName, 63, 7, actR1BillTransactionCurrencyPick);
 		
 		String expR1BillBaseCurrencyPick                          = excelReader.getCellData(xlSheetName, 64, 6);
 		excelReader.setCellData(xlfile, xlSheetName, 64, 7, actR1BillBaseCurrencyPick);
 		
 		String expR1BillLocalCurrencyPick                         = excelReader.getCellData(xlSheetName, 65, 6);
 		excelReader.setCellData(xlfile, xlSheetName, 65, 7, actR1BillLocalCurrencyPick);
 		
 		String expR1BillBalanceNewRefAmountPick                   = excelReader.getCellData(xlSheetName, 66, 6);
 		excelReader.setCellData(xlfile, xlSheetName, 66, 7, actR1BillBalanceNewRefAmountPick);
 		
 		String expR1billRefAdjustAmountInTransCurencyPick         = excelReader.getCellData(xlSheetName, 67, 6);
 		excelReader.setCellData(xlfile, xlSheetName, 67, 7, actR1billRefAdjustAmountInTransCurencyPick);
 		
        String expR1billRefBalanceAmountAdjustInTrnasCurrencyPick = excelReader.getCellData(xlSheetName, 68, 6);
        excelReader.setCellData(xlfile, xlSheetName, 68, 7, actR1billRefBalanceAmountAdjustInTrnasCurrencyPick);
   
      /*  String expR1conversationRateBaseCurrencyRatePick = excelReader.getCellData(xlSheetName, 69, 6); 
        excelReader.setCellData(xlfile, xlSheetName, 69, 7, actR1conversationRateBaseCurrencyRatePick);
        */
        String expR1conversationRateLocalCurrencyRatePick = excelReader.getCellData(xlSheetName, 70, 6);
        excelReader.setCellData(xlfile, xlSheetName, 70, 7, actR1conversationRateLocalCurrencyRatePick);
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
        billRefOkBtn.click();
		
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
        select2ndRow_1stColumn.click();
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
        enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 57, 5));
        Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_DebitTxt));
		enter_DebitTxt.sendKeys(excelReader.getCellData(xlSheetName, 58, 5));
		enter_DebitTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_CreditTxt));
		enter_CreditTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		String actR2PartyName = billRefPartyName.getText();
		String expR2PartyName = excelReader.getCellData(xlSheetName, 71, 6);
		excelReader.setCellData(xlfile, xlSheetName, 71, 7, actR2PartyName);
		
		System.out.println("Bill wise Screen Cutomer Name "  + actR2PartyName + "  Value Expected  " + expR2PartyName);
		
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefAdjustBillsGrid));
			
		String actR2Adjustbills =((WebElement) billRefAdjustBillsGrid).getAttribute("data-totalrows");
		String expR2Adjustbills =excelReader.getCellData(xlSheetName, 72, 6);
		excelReader.setCellData(xlfile, xlSheetName, 72, 7, actR2Adjustbills);
		
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
        
        String actR2BillNewReference                           = billRefNewReferenceTxt.getAttribute("value");
		String actR2BillTransactionCurrency                    = billRefTransactionCurency.getText();
		String actR2BillBaseCurrency                           = billRefBaseCurrency.getText();
		String actR2BillLocalCurrency                          = localCurrencyDhs.getText();
		String actR2BillBalanceNewRefAmount                    = balanceNewReferenceAmt.getText();
        String actR2billRefAdjustAmountInTransCurency          = billRefAdjustAmountInTransCurency.getText();
        String actR2billRefBalanceAmountAdjustInTrnasCurrency  = billRefBalanceAmountAdjustInTrnasCurrency.getText();
        
        
        String expR2BillNewReference                      =excelReader.getCellData(xlSheetName, 73, 6);
		excelReader.setCellData(xlfile, xlSheetName, 73, 7, actR2BillNewReference);
		
		String expR2BillTransactionCurrency               =excelReader.getCellData(xlSheetName, 74, 6);
		excelReader.setCellData(xlfile, xlSheetName, 74, 7, actR2BillTransactionCurrency);
		
		String expR2BillBaseCurrency                      =excelReader.getCellData(xlSheetName, 75, 6);
		excelReader.setCellData(xlfile, xlSheetName, 75, 7, actR2BillBaseCurrency);
		
		String expR2BillLocalCurrency                     =excelReader.getCellData(xlSheetName, 76, 6);
		excelReader.setCellData(xlfile, xlSheetName, 76, 7, actR2BillLocalCurrency);
		
		String expR2BillBalanceNewRefAmount               =excelReader.getCellData(xlSheetName, 77, 6);
		excelReader.setCellData(xlfile, xlSheetName, 77, 7, actR2BillBalanceNewRefAmount);
		
        String expR2billRefAdjustAmountInTransCurency         =excelReader.getCellData(xlSheetName, 78, 6);
        excelReader.setCellData(xlfile, xlSheetName, 78, 7, actR2billRefAdjustAmountInTransCurency);
        
        String expR2billRefBalanceAmountAdjustInTrnasCurrency =excelReader.getCellData(xlSheetName, 79, 6);
        excelReader.setCellData(xlfile, xlSheetName, 79, 7, actR2billRefBalanceAmountAdjustInTrnasCurrency);
        
 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
 		billRefNewReferenceTxt.click();
 		
 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
 		billRefPickIcon.click();
    	
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
      
        String actR2BillNewReferencePick                           = billRefNewReferenceTxt.getAttribute("value");
 		String actR2BillTransactionCurrencyPick                    = billRefTransactionCurency.getText();
 		String actR2BillBaseCurrencyPick                           = billRefBaseCurrency.getText();
 		String actR2BillLocalCurrencyPick                          = localCurrencyDhs.getText();
 		String actR2BillBalanceNewRefAmountPick                    = balanceNewReferenceAmt.getText();
 		String actR2billRefAdjustAmountInTransCurencyPick          = billRefAdjustAmountInTransCurency.getText();
	    String actR2billRefBalanceAmountAdjustInTrnasCurrencyPick  = billRefBalanceAmountAdjustInTrnasCurrency.getText();
 		
 	    String actR2conversationRateBaseCurrencyRatePick   = conversationRateBaseCurrencyRate.getText();
        String actR2conversationRateLocalCurrencyRatePick  = conversationRateLocalCurrencyRate.getText();
        

        String expR2BillNewReferencePick                          = excelReader.getCellData(xlSheetName, 80, 6);
        excelReader.setCellData(xlfile, xlSheetName, 80, 7, actR2BillNewReferencePick);
        
        String expR2BillTransactionCurrencyPick                   = excelReader.getCellData(xlSheetName, 81, 6);
        excelReader.setCellData(xlfile, xlSheetName, 81, 7, actR2BillTransactionCurrencyPick);
        
        String expR2BillBaseCurrencyPick                          = excelReader.getCellData(xlSheetName, 82, 6);
        excelReader.setCellData(xlfile, xlSheetName, 82, 7, actR2BillBaseCurrencyPick);
        
        String expR2BillLocalCurrencyPick                         = excelReader.getCellData(xlSheetName, 83, 6);
        excelReader.setCellData(xlfile, xlSheetName, 83, 7, actR2BillLocalCurrencyPick);
        
        String expR2BillBalanceNewRefAmountPick                   = excelReader.getCellData(xlSheetName, 84, 6);
        excelReader.setCellData(xlfile, xlSheetName, 84, 7, actR2BillBalanceNewRefAmountPick);
        
        String expR2billRefAdjustAmountInTransCurencyPick         = excelReader.getCellData(xlSheetName, 85, 6);
        excelReader.setCellData(xlfile, xlSheetName, 85, 7, actR2billRefAdjustAmountInTransCurencyPick);
        
        String expR2billRefBalanceAmountAdjustInTrnasCurrencyPick = excelReader.getCellData(xlSheetName, 86, 6);
        excelReader.setCellData(xlfile, xlSheetName, 86, 7, actR2billRefBalanceAmountAdjustInTrnasCurrencyPick);
        
        String expR2conversationRateBaseCurrencyRatePick = excelReader.getCellData(xlSheetName, 87, 6); 
        excelReader.setCellData(xlfile, xlSheetName, 87, 7, actR2conversationRateBaseCurrencyRatePick);
        
        String expR2conversationRateLocalCurrencyRatePick = excelReader.getCellData(xlSheetName, 88, 6);
        excelReader.setCellData(xlfile, xlSheetName, 88, 7, actR2conversationRateLocalCurrencyRatePick);
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
        billRefOkBtn.click();
        
        Thread.sleep(2000);
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingBalancesSaveBtn));
		openingBalancesSaveBtn.click();
		
		Thread.sleep(1000);

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

		//expMsg.add(excelReader.getCellData(xlSheetName, 89, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 90, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 89, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 89, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 90, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 89, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		

		System.out.println("Bill reference Adjustment Bills  :"+ actR1Adjustbills+"                          "+ "expR1Adjustbills :"+expR1Adjustbills);
		System.out.println("actR1BillNewReference :             "+ actR1BillNewReference +"                    "+ "expBillNewReference :"+expR1BillNewReference);
		System.out.println("actR1BillTransactionCurrency       :"+ actR1BillTransactionCurrency + "            " + "expBillTransactionCurrency :"+expR1BillTransactionCurrency);
		System.out.println("actBillBaseCurrency :             "+ actR1BillBaseCurrency + "                   " +"expBillBaseCurrency :"+expR1BillBaseCurrency);
		System.out.println("actBillLocalCurrency :            "+ actR1BillLocalCurrency+ "                   " +"expBillLocalCurrency :"+expR1BillLocalCurrency);
		System.out.println("actBillBalanceNewRefAmount :      "+ actR1BillBalanceNewRefAmount + "            " + "expBillBalanceNewRefAmount :"+expR1BillBalanceNewRefAmount);
		System.out.println("billRefAdjustAmountInTransCurency  :"+actR1billRefAdjustAmountInTransCurency+ " expected "+expR1billRefAdjustAmountInTransCurency);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrency :"+ actR1billRefBalanceAmountAdjustInTrnasCurrency + "expbillRefBalanceAmountAdjustInTrnasCurrency :"+expR1billRefBalanceAmountAdjustInTrnasCurrency);

		//////Pick

		System.out.println("actBillNewReferencePick :              "+ actR1BillNewReferencePick +"                "+ "expBillNewReferencePick :"+expR1BillNewReferencePick);
		System.out.println("actBillTransactionCurrencyPick :       "+ actR1BillTransactionCurrencyPick + "        " + "expBillTransactionCurrencyPick :"+expR1BillTransactionCurrencyPick);
		System.out.println("actBillBaseCurrencyPick :              "+ actR1BillBaseCurrencyPick + "               " + "expBillBaseCurrencyPick :"+expR1BillBaseCurrencyPick);
		System.out.println("actBillLocalCurrencyPick :             "+ actR1BillLocalCurrencyPick+ "                   " +"expBillLocalCurrencyPick :"+expR1BillLocalCurrencyPick);
		System.out.println("actBillBalanceNewRefAmountPick :       "+ actR1BillBalanceNewRefAmountPick +          " " + "expBillBalanceNewRefAmountPick :"+expR1BillBalanceNewRefAmountPick);
		//System.out.println("actconversationRateBaseCurrRatePick:   "+ actR1conversationRateBaseCurrencyRatePick+  "  " +"expconversationRateBaseCurrencyRatePick :"+expR1conversationRateBaseCurrencyRatePick);
		System.out.println("actconversationRateLocalCurRatePick :  "+ actR1conversationRateLocalCurrencyRatePick+ " " + "expconversationRateLocalCurrencyRatePick :"+expR1conversationRateLocalCurrencyRatePick);
		System.out.println("actbillRefAdjustAmountInTransCurency :"+ actR1billRefAdjustAmountInTransCurencyPick + "       " +"expbillRefAdjustAmountInTransCurency :"+expR1billRefAdjustAmountInTransCurencyPick);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrency :"+ actR1billRefBalanceAmountAdjustInTrnasCurrencyPick + "expbillRefBalanceAmountAdjustInTrnasCurrency :"+expR1billRefBalanceAmountAdjustInTrnasCurrencyPick);



		System.out.println("Bill reference Adjustment Bills  :"+ actR2Adjustbills+"                          "+ "expR2Adjustbills :"+expR2Adjustbills);
		System.out.println("actR2BillNewReference :             "+ actR2BillNewReference +"                    "+ "expBillNewReference :"+expR2BillNewReference);
		System.out.println("actR2BillTransactionCurrency       :"+ actR2BillTransactionCurrency + "            " + "expBillTransactionCurrency :"+expR2BillTransactionCurrency);
		System.out.println("actBillBaseCurrency :             "+ actR2BillBaseCurrency + "                   " +"expBillBaseCurrency :"+expR2BillBaseCurrency);
		System.out.println("actBillLocalCurrency :            "+ actR2BillLocalCurrency+ "                   " +"expBillLocalCurrency :"+expR2BillLocalCurrency);
		System.out.println("actBillBalanceNewRefAmount :      "+ actR2BillBalanceNewRefAmount + "            " + "expBillBalanceNewRefAmount :"+expR2BillBalanceNewRefAmount);
		System.out.println("actbillRefAdjustAmountInTransCurency :"+ actR2billRefAdjustAmountInTransCurency + "       " +"expbillRefAdjustAmountInTransCurency :"+expR2billRefAdjustAmountInTransCurency);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrency :"+ actR2billRefBalanceAmountAdjustInTrnasCurrency + "expbillRefBalanceAmountAdjustInTrnasCurrency :"+expR2billRefBalanceAmountAdjustInTrnasCurrency);

		//////Pick

		System.out.println("actBillNewReferencePick :              "+ actR2BillNewReferencePick +"                "+ "expBillNewReferencePick :"+expR2BillNewReferencePick);
		System.out.println("actBillTransactionCurrencyPick :       "+ actR2BillTransactionCurrencyPick + "        " + "expBillTransactionCurrencyPick :"+expR2BillTransactionCurrencyPick);
		System.out.println("actBillBaseCurrencyPick :              "+ actR2BillBaseCurrencyPick + "               " + "expBillBaseCurrencyPick :"+expR2BillBaseCurrencyPick);
		System.out.println("actBillLocalCurrencyPick :             "+ actR2BillLocalCurrencyPick+ "                   " +"expBillLocalCurrencyPick :"+expR2BillLocalCurrencyPick);
		System.out.println("actBillBalanceNewRefAmountPick :       "+ actR2BillBalanceNewRefAmountPick +          " " + "expBillBalanceNewRefAmountPick :"+expR2BillBalanceNewRefAmountPick);
		System.out.println("actconversationRateBaseCurrRatePick:   "+ actR2conversationRateBaseCurrencyRatePick+  "  " +"expconversationRateBaseCurrencyRatePick :"+expR2conversationRateBaseCurrencyRatePick);
		System.out.println("actconversationRateLocalCurRatePick :  "+ actR2conversationRateLocalCurrencyRatePick+ " " + "expconversationRateLocalCurrencyRatePick :"+expR2conversationRateLocalCurrencyRatePick);
		System.out.println("actbillRefAdjustAmountInTransCurency :"+ actR2billRefAdjustAmountInTransCurencyPick + "       " +"expbillRefAdjustAmountInTransCurency :"+expR2billRefAdjustAmountInTransCurencyPick);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrency :"+ actR2billRefBalanceAmountAdjustInTrnasCurrencyPick + "expbillRefBalanceAmountAdjustInTrnasCurrency :"+expR2billRefBalanceAmountAdjustInTrnasCurrencyPick);



		if(actR1PartyName.equalsIgnoreCase(expR1PartyName) && actR1Adjustbills.equalsIgnoreCase(expR1Adjustbills) && 
				actR1BillNewReference.equalsIgnoreCase(expR1BillNewReference) && actR1BillTransactionCurrency.equalsIgnoreCase(expR1BillTransactionCurrency) && 
				actR1BillBaseCurrency.equalsIgnoreCase(expR1BillBaseCurrency) && actR1BillLocalCurrency.equalsIgnoreCase(expR1BillLocalCurrency) && 
				actR1BillBalanceNewRefAmount.equalsIgnoreCase(expR1BillBalanceNewRefAmount) && 
				actR1billRefAdjustAmountInTransCurency.equalsIgnoreCase(expR1billRefAdjustAmountInTransCurency) &&
				actR1billRefBalanceAmountAdjustInTrnasCurrency.equalsIgnoreCase(expR1billRefBalanceAmountAdjustInTrnasCurrency) && 

				actR1BillNewReferencePick.equalsIgnoreCase(expR1BillNewReferencePick) && actR1BillTransactionCurrencyPick.equalsIgnoreCase(expR1BillTransactionCurrencyPick) && 
				actR1BillBaseCurrencyPick.equalsIgnoreCase(expR1BillBaseCurrencyPick) && actR1BillLocalCurrencyPick.equalsIgnoreCase(expR1BillLocalCurrencyPick) && 
				actR1BillBalanceNewRefAmountPick.equalsIgnoreCase(expR1BillBalanceNewRefAmountPick) /*&&
				actR1conversationRateBaseCurrencyRatePick.equalsIgnoreCase(expR1conversationRateBaseCurrencyRatePick)*/ && actR1conversationRateLocalCurrencyRatePick.equalsIgnoreCase(expR1conversationRateLocalCurrencyRatePick) &&



				actR2PartyName.equalsIgnoreCase(expR2PartyName) && actR2Adjustbills.equalsIgnoreCase(expR2Adjustbills) && actR2BillNewReference.equalsIgnoreCase(expR2BillNewReference) && 
				actR2BillTransactionCurrency.equalsIgnoreCase(expR2BillTransactionCurrency) && 
				actR2BillBaseCurrency.equalsIgnoreCase(expR2BillBaseCurrency) && actR2BillLocalCurrency.equalsIgnoreCase(expR2BillLocalCurrency) && 
				actR2BillBalanceNewRefAmount.equalsIgnoreCase(expR2BillBalanceNewRefAmount) && 
				actR2billRefAdjustAmountInTransCurency.equalsIgnoreCase(expR2billRefAdjustAmountInTransCurency) &&
				actR2billRefBalanceAmountAdjustInTrnasCurrency.equalsIgnoreCase(expR2billRefBalanceAmountAdjustInTrnasCurrency) && 

				actR2BillNewReferencePick.equalsIgnoreCase(expR2BillNewReferencePick) && actR2BillTransactionCurrencyPick.equalsIgnoreCase(expR2BillTransactionCurrencyPick) && 
				actR2BillBaseCurrencyPick.equalsIgnoreCase(expR2BillBaseCurrencyPick) && actR2BillLocalCurrencyPick.equalsIgnoreCase(expR2BillLocalCurrencyPick) && 
				actR2BillBalanceNewRefAmountPick.equalsIgnoreCase(expR2BillBalanceNewRefAmountPick) &&
				actR2conversationRateBaseCurrencyRatePick.equalsIgnoreCase(expR2conversationRateBaseCurrencyRatePick) && actR2conversationRateLocalCurrencyRatePick.equalsIgnoreCase(expR2conversationRateLocalCurrencyRatePick)

				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 53, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 53, 8, resFail);
			return false;
		}
	}








	public static boolean checkVerifyAndEditOpeningBalanceVoucher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();
		
		checkValidationMessage("Voucher loaded successfully");
		
		String actDocno               = documentNumberTxt.getAttribute("value");
        String actCurrency            = voucherHeaderCurrency.getAttribute("value");
		String actDepartment          = departmentTxt.getAttribute("value");
        String actExchangeRate        = voucherHeaderExchangeRate.getAttribute("value");
        String actLocExchangeRate     = voucherHeaderLocalExchangeRate.getAttribute("value");
		String actR1Account           = select1stRow_1stColumn.getText();
		String actR1Credit            = select1stRow_3rdColumn.getText();
        String actR1Ref               = select1stRow_4thColumn.getText();
        String actR2Account           = select2ndRow_1stColumn.getText();
		String actR2Debit             = select2ndRow_2ndColumn.getText();
        String actR2Ref               = select2ndRow_4thColumn.getText();
        
        /*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vocFooterExpandBtn));
		vocFooterExpandBtn.click();
		
		Thread.sleep(1000);
		
		ScrollIntoView(vocFooterdebitAmount);*/
		
    	String actFooterCreditAmt     = vocFooterCreditAmount.getText();
    	String actVoucherDebitAmt     = vocFooterdebitAmount.getText();
    	
		String expDocno               = excelReader.getCellData(xlSheetName, 91, 6);
		excelReader.setCellData(xlfile, xlSheetName, 91, 7, actDocno);
		
	    String expCurrency            = excelReader.getCellData(xlSheetName, 92, 6);
	    excelReader.setCellData(xlfile, xlSheetName, 92, 7, actCurrency);
	    
	    String expDepartment          = excelReader.getCellData(xlSheetName, 93, 6);
	    excelReader.setCellData(xlfile, xlSheetName, 93, 7, actDepartment);
	    
	    String expExchangeRate        = excelReader.getCellData(xlSheetName, 94, 6);
	    excelReader.setCellData(xlfile, xlSheetName, 94, 7, actExchangeRate);
	    
	    String expLocExchangeRate     = excelReader.getCellData(xlSheetName, 95, 6);
	    excelReader.setCellData(xlfile, xlSheetName, 95, 7, actLocExchangeRate);
	    
	    String expR1Account           = excelReader.getCellData(xlSheetName, 96, 6);
	    excelReader.setCellData(xlfile, xlSheetName, 96, 7, actR1Account);
	    
	    String expR1Credit            = excelReader.getCellData(xlSheetName, 97, 6);
	    excelReader.setCellData(xlfile, xlSheetName, 97, 7, actR1Credit);
	    
	    String expR1Ref               = excelReader.getCellData(xlSheetName, 98, 6);
	    excelReader.setCellData(xlfile, xlSheetName, 98, 7, actR1Ref);
	    
	    String expR2Account           = excelReader.getCellData(xlSheetName, 99, 6);
	    excelReader.setCellData(xlfile, xlSheetName, 99, 7, actR2Account);
	    
	    String expR2Debit             = excelReader.getCellData(xlSheetName, 100, 6);
	    excelReader.setCellData(xlfile, xlSheetName, 100, 7, actR2Debit);
	    
	    String expR2Ref               = excelReader.getCellData(xlSheetName, 101, 6);
	    excelReader.setCellData(xlfile, xlSheetName, 101, 7, actR2Ref);
	    
	    String expFooterCreditAmt     = excelReader.getCellData(xlSheetName, 102, 6);
	    excelReader.setCellData(xlfile, xlSheetName, 102, 7, actFooterCreditAmt);
	    
	    String expVoucherDebitAmt     = excelReader.getCellData(xlSheetName, 103, 6);
	    excelReader.setCellData(xlfile, xlSheetName, 103, 7, actVoucherDebitAmt);
	    
        System.out.println("*******************************checkVerifyAndEditOpeningBalanceVoucher1**************************");

		System.out.println("Entry Page Document Number    "+actDocno           +"  value Expected  "+expDocno);
		System.out.println("Entry Page Currency           "+actCurrency        +"  value Expected  "+expCurrency);
		System.out.println("Entry Page Department         "+actDepartment      +"  value Expected  "+expDepartment);
		System.out.println("Entry Page Exchange Rate      "+actExchangeRate    +"  value Expected  "+expExchangeRate);
		System.out.println("Entry Page Department         "+actLocExchangeRate +"  value Expected  "+expLocExchangeRate);
		System.out.println("Entry Page R1Account          "+actR1Account       +"  value Expected  "+expR1Account);
		System.out.println("Entry Page R1Credit           "+actR1Credit        +"  value Expected  "+expR1Credit);
		System.out.println("Entry Page R1Ref              "+actR1Ref           +"  value Expected  "+expR1Ref);
		System.out.println("Entry Page R2Account          "+actR2Account       +"  value Expected  "+expR2Account);
		System.out.println("Entry Page R2Credit           "+actR2Debit        +"  value Expected  "+expR2Debit);
		System.out.println("Entry Page R2Ref              "+actR2Ref           +"  value Expected  "+expR2Ref);
		
		System.out.println("Entry Page Debit Amount       "+actVoucherDebitAmt +"  value Expected  "+expVoucherDebitAmt);
		System.out.println("Entry Page Credit Amount      "+actFooterCreditAmt +"  value Expected  "+expFooterCreditAmt);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		 
		String docno=documentNumberTxt.getAttribute("value");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.click();
		
		voucherHeaderCurrency.sendKeys(Keys.SHIFT,Keys.HOME);
		
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 91, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_CreditTxt));
		enter_CreditTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		System.out.println("***********************************checkSavingOpeningBalanceVoucher1WithCurrencyAsINR*************************************");
		
		String actR1PartyName = billRefPartyName.getText();
		String expR1PartyName = excelReader.getCellData(xlSheetName, 104, 6);
		excelReader.setCellData(xlfile, xlSheetName, 104, 7, actR1PartyName);
		
		System.out.println("Bill wise Screen Cutomer Name "  + actR1PartyName + "  Value Expected  " + expR1PartyName);
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefAdjustBillsGrid));
			
		String actR1Adjustbills=((WebElement) billRefAdjustBillsGrid).getAttribute("data-totalrows");
		String expR1Adjustbills=excelReader.getCellData(xlSheetName, 105, 6);
		excelReader.setCellData(xlfile, xlSheetName, 105, 7, actR1Adjustbills);
		
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
        
        String actR1BillNewReference                           = billRefNewReferenceTxt.getAttribute("value");
		String actR1BillTransactionCurrency                    = billRefTransactionCurency.getText();
		String actR1BillBaseCurrency                           = billRefBaseCurrency.getText();
		String actR1BillLocalCurrency                          = localCurrencyDhs.getText();
		String actR1BillBalanceNewRefAmount                    = balanceNewReferenceAmt.getText();
        String actR1billRefAdjustAmountInTransCurency          = billRefAdjustAmountInTransCurency.getText();
        String actR1billRefBalanceAmountAdjustInTrnasCurrency  = billRefBalanceAmountAdjustInTrnasCurrency.getText();
        
        
        String expR1BillNewReference                      =excelReader.getCellData(xlSheetName, 106, 6);
        excelReader.setCellData(xlfile, xlSheetName, 106, 7, actR1BillNewReference);
		
		String expR1BillTransactionCurrency               =excelReader.getCellData(xlSheetName, 107, 6);
		excelReader.setCellData(xlfile, xlSheetName, 107, 7, actR1BillTransactionCurrency);
		
		String expR1BillBaseCurrency                      =excelReader.getCellData(xlSheetName, 108, 6);
		excelReader.setCellData(xlfile, xlSheetName, 108, 7, actR1BillBaseCurrency);
		
		String expR1BillLocalCurrency                     =excelReader.getCellData(xlSheetName, 109, 6);
		excelReader.setCellData(xlfile, xlSheetName, 109, 7, actR1BillLocalCurrency);
		
		String expR1BillBalanceNewRefAmount               =excelReader.getCellData(xlSheetName, 110, 6);
		excelReader.setCellData(xlfile, xlSheetName, 110, 7, actR1BillBalanceNewRefAmount);
		
        String expR1billRefAdjustAmountInTransCurency         =excelReader.getCellData(xlSheetName, 111, 6);
        excelReader.setCellData(xlfile, xlSheetName, 111, 7, actR1billRefAdjustAmountInTransCurency);
        
        String expR1billRefBalanceAmountAdjustInTrnasCurrency =excelReader.getCellData(xlSheetName, 112, 6);
        excelReader.setCellData(xlfile, xlSheetName, 112, 7, actR1billRefBalanceAmountAdjustInTrnasCurrency);
        
 		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
 		billRefNewReferenceTxt.click();
 		
 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
 		billRefPickIcon.click();*/
    	
 
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
      
        String actR1BillNewReferencePick                           = billRefNewReferenceTxt.getAttribute("value");
 		String actR1BillTransactionCurrencyPick                    = billRefTransactionCurency.getText();
 		String actR1BillBaseCurrencyPick                           = billRefBaseCurrency.getText();
 		String actR1BillLocalCurrencyPick                          = localCurrencyDhs.getText();
 		String actR1BillBalanceNewRefAmountPick                    = balanceNewReferenceAmt.getText();
 		String actR1billRefAdjustAmountInTransCurencyPick          = billRefAdjustAmountInTransCurency.getText();
	    String actR1billRefBalanceAmountAdjustInTrnasCurrencyPick  = billRefBalanceAmountAdjustInTrnasCurrency.getText();
 		
 	    String actR1conversationRateBaseCurrencyRatePick   = conversationRateBaseCurrencyRate.getText();
        String actR1conversationRateLocalCurrencyRatePick  = conversationRateLocalCurrencyRate.getText();
        
        
        String expR1BillNewReferencePick                          = excelReader.getCellData(xlSheetName, 113, 6);
        excelReader.setCellData(xlfile, xlSheetName, 113, 7, actR1BillNewReferencePick);
        
 		String expR1BillTransactionCurrencyPick                   = excelReader.getCellData(xlSheetName, 114, 6);
 		excelReader.setCellData(xlfile, xlSheetName, 114, 7, actR1BillTransactionCurrencyPick);
 		
 		String expR1BillBaseCurrencyPick                          = excelReader.getCellData(xlSheetName, 115, 6);
 		excelReader.setCellData(xlfile, xlSheetName, 115, 7, actR1BillBaseCurrencyPick);
 		
 		String expR1BillLocalCurrencyPick                         = excelReader.getCellData(xlSheetName, 116, 6);
 		excelReader.setCellData(xlfile, xlSheetName, 116, 7, actR1BillLocalCurrencyPick);
 		
 		String expR1BillBalanceNewRefAmountPick                   = excelReader.getCellData(xlSheetName, 117, 6);
 		excelReader.setCellData(xlfile, xlSheetName, 117, 7, actR1BillBalanceNewRefAmountPick);
 		
 		String expR1billRefAdjustAmountInTransCurencyPick         = excelReader.getCellData(xlSheetName, 118, 6);
 		excelReader.setCellData(xlfile, xlSheetName, 118, 7, actR1billRefAdjustAmountInTransCurencyPick);
 		
 		String expR1billRefBalanceAmountAdjustInTrnasCurrencyPick = excelReader.getCellData(xlSheetName, 119, 6);
 		excelReader.setCellData(xlfile, xlSheetName, 119, 7, actR1billRefBalanceAmountAdjustInTrnasCurrencyPick);
        
 		String expR1conversationRateBaseCurrencyRatePick = excelReader.getCellData(xlSheetName, 120, 6); 
 		excelReader.setCellData(xlfile, xlSheetName, 120, 7, actR1conversationRateBaseCurrencyRatePick);
 		
 		String expR1conversationRateLocalCurrencyRatePick = excelReader.getCellData(xlSheetName, 121, 6);
 		excelReader.setCellData(xlfile, xlSheetName, 121, 7, actR1conversationRateLocalCurrencyRatePick);
 		
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
        billRefOkBtn.click();
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
		select2ndRow_3rdColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_CreditTxt));
		enter_CreditTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		String actR2PartyName = billRefPartyName.getText();
		String expR2PartyName = excelReader.getCellData(xlSheetName, 122, 6);
		excelReader.setCellData(xlfile, xlSheetName, 122, 7, actR2PartyName);
		
		System.out.println("Bill wise Screen Cutomer Name "  + actR2PartyName + "  Value Expected  " + expR2PartyName);
		
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefAdjustBillsGrid));
			
		String actR2Adjustbills=((WebElement) billRefAdjustBillsGrid).getAttribute("data-totalrows");
		String expR2Adjustbills=excelReader.getCellData(xlSheetName, 123, 6);
		excelReader.setCellData(xlfile, xlSheetName, 123, 7, actR2Adjustbills);
		
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
        
        String actR2BillNewReference                           = billRefNewReferenceTxt.getAttribute("value");
		String actR2BillTransactionCurrency                    = billRefTransactionCurency.getText();
		String actR2BillBaseCurrency                           = billRefBaseCurrency.getText();
		String actR2BillLocalCurrency                          = localCurrencyDhs.getText();
		String actR2BillBalanceNewRefAmount                    = balanceNewReferenceAmt.getText();
        String actR2billRefAdjustAmountInTransCurency          = billRefAdjustAmountInTransCurency.getText();
        String actR2billRefBalanceAmountAdjustInTrnasCurrency  = billRefBalanceAmountAdjustInTrnasCurrency.getText();
        

        String expR2BillNewReference                      =excelReader.getCellData(xlSheetName, 124, 6);
        excelReader.setCellData(xlfile, xlSheetName, 124, 7, actR2BillNewReference);
		
		String expR2BillTransactionCurrency               =excelReader.getCellData(xlSheetName, 125, 6);
		excelReader.setCellData(xlfile, xlSheetName, 125, 7, actR2BillTransactionCurrency);
		
		String expR2BillBaseCurrency                      =excelReader.getCellData(xlSheetName, 126, 6);
		excelReader.setCellData(xlfile, xlSheetName, 126, 7, actR2BillBaseCurrency);
		
		String expR2BillLocalCurrency                     =excelReader.getCellData(xlSheetName, 127, 6);
		excelReader.setCellData(xlfile, xlSheetName, 127, 7, actR2BillLocalCurrency);
		
		String expR2BillBalanceNewRefAmount               =excelReader.getCellData(xlSheetName, 128, 6);
		excelReader.setCellData(xlfile, xlSheetName, 128, 7, actR2BillBalanceNewRefAmount);
		
        String expR2billRefAdjustAmountInTransCurency         =excelReader.getCellData(xlSheetName, 129, 6);
        excelReader.setCellData(xlfile, xlSheetName, 129, 7, actR2billRefAdjustAmountInTransCurency);
        
        String expR2billRefBalanceAmountAdjustInTrnasCurrency =excelReader.getCellData(xlSheetName, 130, 6);
        excelReader.setCellData(xlfile, xlSheetName, 130, 7, actR2billRefBalanceAmountAdjustInTrnasCurrency);
        
 		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
 		billRefNewReferenceTxt.click();
 		
 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
 		billRefPickIcon.click();*/
    	
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
      
        String actR2BillNewReferencePick                           = billRefNewReferenceTxt.getAttribute("value");
 		String actR2BillTransactionCurrencyPick                    = billRefTransactionCurency.getText();
 		String actR2BillBaseCurrencyPick                           = billRefBaseCurrency.getText();
 		String actR2BillLocalCurrencyPick                          = localCurrencyDhs.getText();
 		String actR2BillBalanceNewRefAmountPick                    = balanceNewReferenceAmt.getText();
 		String actR2billRefAdjustAmountInTransCurencyPick          = billRefAdjustAmountInTransCurency.getText();
	    String actR2billRefBalanceAmountAdjustInTrnasCurrencyPick  = billRefBalanceAmountAdjustInTrnasCurrency.getText();
 		
 	    String actR2conversationRateBaseCurrencyRatePick   = conversationRateBaseCurrencyRate.getText();
        String actR2conversationRateLocalCurrencyRatePick  = conversationRateLocalCurrencyRate.getText();
        
       
        String expR2BillNewReferencePick                          = excelReader.getCellData(xlSheetName, 131, 6);
        excelReader.setCellData(xlfile, xlSheetName, 131, 7, actR2BillNewReferencePick);
        
        String expR2BillTransactionCurrencyPick                   = excelReader.getCellData(xlSheetName, 132, 6);
        excelReader.setCellData(xlfile, xlSheetName, 132, 7, actR2BillTransactionCurrencyPick);
        
        String expR2BillBaseCurrencyPick                          = excelReader.getCellData(xlSheetName, 133, 6);
        excelReader.setCellData(xlfile, xlSheetName, 133, 7, actR2BillBaseCurrencyPick);
        
        String expR2BillLocalCurrencyPick                         = excelReader.getCellData(xlSheetName, 134, 6);
        excelReader.setCellData(xlfile, xlSheetName, 134, 7, actR2BillLocalCurrencyPick);
        
        String expR2BillBalanceNewRefAmountPick                   = excelReader.getCellData(xlSheetName, 135, 6);
        excelReader.setCellData(xlfile, xlSheetName, 135, 7, actR2BillBalanceNewRefAmountPick);
        
        String expR2billRefAdjustAmountInTransCurencyPick         = excelReader.getCellData(xlSheetName, 136, 6);
        excelReader.setCellData(xlfile, xlSheetName, 136, 7, actR2billRefAdjustAmountInTransCurencyPick);
        
        String expR2billRefBalanceAmountAdjustInTrnasCurrencyPick = excelReader.getCellData(xlSheetName, 137, 6);
        excelReader.setCellData(xlfile, xlSheetName, 137, 7, actR2billRefBalanceAmountAdjustInTrnasCurrencyPick);
        
        String expR2conversationRateBaseCurrencyRatePick = excelReader.getCellData(xlSheetName, 138, 6); 
        excelReader.setCellData(xlfile, xlSheetName, 138, 7, actR2conversationRateBaseCurrencyRatePick);
        
        String expR2conversationRateLocalCurrencyRatePick = excelReader.getCellData(xlSheetName, 139, 6);
        excelReader.setCellData(xlfile, xlSheetName, 139, 7, actR2conversationRateLocalCurrencyRatePick);
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
        billRefOkBtn.click();
        
        Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingBalancesSaveBtn));
		openingBalancesSaveBtn.click();
		
		Thread.sleep(1000);

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

		//expMsg.add(excelReader.getCellData(xlSheetName, 140, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 141, 6));
		
		String expMessage = expMsg.toString();
		excelReader.setCellData(xlfile, xlSheetName, 140, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 140, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 141, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 140, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		System.out.println("Bill reference Adjustment Bills  :"+ actR1Adjustbills+"                          "+ "expR1Adjustbills :"+expR1Adjustbills);
		System.out.println("actR1BillNewReference :             "+ actR1BillNewReference +"                    "+ "expBillNewReference :"+expR1BillNewReference);
		System.out.println("actR1BillTransactionCurrency       :"+ actR1BillTransactionCurrency + "            " + "expBillTransactionCurrency :"+expR1BillTransactionCurrency);
		System.out.println("actBillBaseCurrency :             "+ actR1BillBaseCurrency + "                   " +"expBillBaseCurrency :"+expR1BillBaseCurrency);
		System.out.println("actBillLocalCurrency :            "+ actR1BillLocalCurrency+ "                   " +"expBillLocalCurrency :"+expR1BillLocalCurrency);
		System.out.println("actBillBalanceNewRefAmount :      "+ actR1BillBalanceNewRefAmount + "            " + "expBillBalanceNewRefAmount :"+expR1BillBalanceNewRefAmount);
		System.out.println("billRefAdjustAmountInTransCurency  :"+actR1billRefAdjustAmountInTransCurency+ " expected "+expR1billRefAdjustAmountInTransCurency);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrency :"+ actR1billRefBalanceAmountAdjustInTrnasCurrency + "expbillRefBalanceAmountAdjustInTrnasCurrency :"+expR1billRefBalanceAmountAdjustInTrnasCurrency);

		// Pick

		System.out.println("actBillNewReferencePick :              "+ actR1BillNewReferencePick +"                "+ "expBillNewReferencePick :"+expR1BillNewReferencePick);
		System.out.println("actBillTransactionCurrencyPick :       "+ actR1BillTransactionCurrencyPick + "        " + "expBillTransactionCurrencyPick :"+expR1BillTransactionCurrencyPick);
		System.out.println("actBillBaseCurrencyPick :              "+ actR1BillBaseCurrencyPick + "               " + "expBillBaseCurrencyPick :"+expR1BillBaseCurrencyPick);
		System.out.println("actBillLocalCurrencyPick :             "+ actR1BillLocalCurrencyPick+ "                   " +"expBillLocalCurrencyPick :"+expR1BillLocalCurrencyPick);
		System.out.println("actBillBalanceNewRefAmountPick :       "+ actR1BillBalanceNewRefAmountPick +          " " + "expBillBalanceNewRefAmountPick :"+expR1BillBalanceNewRefAmountPick);
		System.out.println("actconversationRateBaseCurrRatePick:   "+ actR1conversationRateBaseCurrencyRatePick+  "  " +"expconversationRateBaseCurrencyRatePick :"+expR1conversationRateBaseCurrencyRatePick);
		System.out.println("actconversationRateLocalCurRatePick :  "+ actR1conversationRateLocalCurrencyRatePick+ " " + "expconversationRateLocalCurrencyRatePick :"+expR1conversationRateLocalCurrencyRatePick);
		System.out.println("actbillRefAdjustAmountInTransCurency :"+ actR1billRefAdjustAmountInTransCurencyPick + "       " +"expbillRefAdjustAmountInTransCurency :"+expR1billRefAdjustAmountInTransCurencyPick);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrency :"+ actR1billRefBalanceAmountAdjustInTrnasCurrencyPick + "expbillRefBalanceAmountAdjustInTrnasCurrency :"+expR1billRefBalanceAmountAdjustInTrnasCurrencyPick);



		System.out.println("Bill reference Adjustment Bills  :"+ actR2Adjustbills+"                          "+ "expR2Adjustbills :"+expR2Adjustbills);
		System.out.println("actR2BillNewReference :             "+ actR2BillNewReference +"                    "+ "expBillNewReference :"+expR2BillNewReference);
		System.out.println("actR2BillTransactionCurrency       :"+ actR2BillTransactionCurrency + "            " + "expBillTransactionCurrency :"+expR2BillTransactionCurrency);
		System.out.println("actBillBaseCurrency :             "+ actR2BillBaseCurrency + "                   " +"expBillBaseCurrency :"+expR2BillBaseCurrency);
		System.out.println("actBillLocalCurrency :            "+ actR2BillLocalCurrency+ "                   " +"expBillLocalCurrency :"+expR2BillLocalCurrency);
		System.out.println("actBillBalanceNewRefAmount :      "+ actR2BillBalanceNewRefAmount + "            " + "expBillBalanceNewRefAmount :"+expR2BillBalanceNewRefAmount);
		System.out.println("actbillRefAdjustAmountInTransCurency :"+ actR2billRefAdjustAmountInTransCurency + "       " +"expbillRefAdjustAmountInTransCurency :"+expR2billRefAdjustAmountInTransCurency);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrency :"+ actR2billRefBalanceAmountAdjustInTrnasCurrency + "expbillRefBalanceAmountAdjustInTrnasCurrency :"+expR2billRefBalanceAmountAdjustInTrnasCurrency);

		// Pick

		System.out.println("actBillNewReferencePick :              "+ actR2BillNewReferencePick +"                "+ "expBillNewReferencePick :"+expR2BillNewReferencePick);
		System.out.println("actBillTransactionCurrencyPick :       "+ actR2BillTransactionCurrencyPick + "        " + "expBillTransactionCurrencyPick :"+expR2BillTransactionCurrencyPick);
		System.out.println("actBillBaseCurrencyPick :              "+ actR2BillBaseCurrencyPick + "               " + "expBillBaseCurrencyPick :"+expR2BillBaseCurrencyPick);
		System.out.println("actBillLocalCurrencyPick :             "+ actR2BillLocalCurrencyPick+ "                   " +"expBillLocalCurrencyPick :"+expR2BillLocalCurrencyPick);
		System.out.println("actBillBalanceNewRefAmountPick :       "+ actR2BillBalanceNewRefAmountPick +          " " + "expBillBalanceNewRefAmountPick :"+expR2BillBalanceNewRefAmountPick);
		System.out.println("actconversationRateBaseCurrRatePick:   "+ actR2conversationRateBaseCurrencyRatePick+  "  " +"expconversationRateBaseCurrencyRatePick :"+expR2conversationRateBaseCurrencyRatePick);
		System.out.println("actconversationRateLocalCurRatePick :  "+ actR2conversationRateLocalCurrencyRatePick+ " " + "expconversationRateLocalCurrencyRatePick :"+expR2conversationRateLocalCurrencyRatePick);
		System.out.println("actbillRefAdjustAmountInTransCurency :"+ actR2billRefAdjustAmountInTransCurencyPick + "       " +"expbillRefAdjustAmountInTransCurency :"+expR2billRefAdjustAmountInTransCurencyPick);
		System.out.println("actbillRefBalanceAmountAdjustInTrnasCurrency :"+ actR2billRefBalanceAmountAdjustInTrnasCurrencyPick + "expbillRefBalanceAmountAdjustInTrnasCurrency :"+expR2billRefBalanceAmountAdjustInTrnasCurrencyPick);

		Thread.sleep(3000);

		if(actDocno.equalsIgnoreCase(expDocno) && actCurrency.equalsIgnoreCase(expCurrency) &&actDepartment.equalsIgnoreCase(expDepartment) && 
				actExchangeRate.equalsIgnoreCase(expExchangeRate) && actLocExchangeRate.equalsIgnoreCase(expLocExchangeRate) && 
				actR1Account.equalsIgnoreCase(expR1Account) && actR1Credit.equalsIgnoreCase(expR1Credit) && actR1Ref.equalsIgnoreCase(expR1Ref) &&
				actR2Account.equalsIgnoreCase(expR2Account) && actR2Debit.equalsIgnoreCase(expR2Debit) && actR2Ref.equalsIgnoreCase(expR2Ref) &&
				actVoucherDebitAmt.equalsIgnoreCase(expVoucherDebitAmt) && actFooterCreditAmt.equalsIgnoreCase(expFooterCreditAmt) &&


				actR1PartyName.equalsIgnoreCase(expR1PartyName) && actR1Adjustbills.equalsIgnoreCase(expR1Adjustbills) && 
				actR1BillNewReference.equalsIgnoreCase(expR1BillNewReference) && actR1BillTransactionCurrency.equalsIgnoreCase(expR1BillTransactionCurrency) && 
				actR1BillBaseCurrency.equalsIgnoreCase(expR1BillBaseCurrency) && actR1BillLocalCurrency.equalsIgnoreCase(expR1BillLocalCurrency) && 
				actR1BillBalanceNewRefAmount.equalsIgnoreCase(expR1BillBalanceNewRefAmount) && 
				actR1billRefAdjustAmountInTransCurency.equalsIgnoreCase(expR1billRefAdjustAmountInTransCurency) &&
				actR1billRefBalanceAmountAdjustInTrnasCurrency.equalsIgnoreCase(expR1billRefBalanceAmountAdjustInTrnasCurrency) && 

				actR1BillNewReferencePick.equalsIgnoreCase(expR1BillNewReferencePick) && actR1BillTransactionCurrencyPick.equalsIgnoreCase(expR1BillTransactionCurrencyPick) && 
				actR1BillBaseCurrencyPick.equalsIgnoreCase(expR1BillBaseCurrencyPick) && actR1BillLocalCurrencyPick.equalsIgnoreCase(expR1BillLocalCurrencyPick) && 
				actR1BillBalanceNewRefAmountPick.equalsIgnoreCase(expR1BillBalanceNewRefAmountPick) &&
				actR1conversationRateBaseCurrencyRatePick.equalsIgnoreCase(expR1conversationRateBaseCurrencyRatePick) && actR1conversationRateLocalCurrencyRatePick.equalsIgnoreCase(expR1conversationRateLocalCurrencyRatePick) &&


				actR2PartyName.equalsIgnoreCase(expR2PartyName) && actR2Adjustbills.equalsIgnoreCase(expR2Adjustbills) && actR2BillNewReference.equalsIgnoreCase(expR2BillNewReference) && 
				actR2BillTransactionCurrency.equalsIgnoreCase(expR2BillTransactionCurrency) && 
				actR2BillBaseCurrency.equalsIgnoreCase(expR2BillBaseCurrency) && actR2BillLocalCurrency.equalsIgnoreCase(expR2BillLocalCurrency) && 
				actR2BillBalanceNewRefAmount.equalsIgnoreCase(expR2BillBalanceNewRefAmount) && 
				actR2billRefAdjustAmountInTransCurency.equalsIgnoreCase(expR2billRefAdjustAmountInTransCurency) &&
				actR2billRefBalanceAmountAdjustInTrnasCurrency.equalsIgnoreCase(expR2billRefBalanceAmountAdjustInTrnasCurrency) && 

				actR2BillNewReferencePick.equalsIgnoreCase(expR2BillNewReferencePick) && actR2BillTransactionCurrencyPick.equalsIgnoreCase(expR2BillTransactionCurrencyPick) && 
				actR2BillBaseCurrencyPick.equalsIgnoreCase(expR2BillBaseCurrencyPick) && actR2BillLocalCurrencyPick.equalsIgnoreCase(expR2BillLocalCurrencyPick) && 
				actR2BillBalanceNewRefAmountPick.equalsIgnoreCase(expR2BillBalanceNewRefAmountPick) &&
				actR2conversationRateBaseCurrencyRatePick.equalsIgnoreCase(expR2conversationRateBaseCurrencyRatePick) && actR2conversationRateLocalCurrencyRatePick.equalsIgnoreCase(expR2conversationRateLocalCurrencyRatePick)

				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
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




	
	
	public static boolean checkSavingOpeningBalanceVoucher2AndSuspendingFromEntryPage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
		new_newBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.click();
		
		voucherHeaderCurrency.sendKeys(Keys.SHIFT,Keys.HOME);
		
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 142, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 143, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 144, 5));
		Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_DebitTxt));
		enter_DebitTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_CreditTxt));
		enter_CreditTxt.sendKeys(excelReader.getCellData(xlSheetName, 145, 5));
		enter_CreditTxt.sendKeys(Keys.TAB);
		
		System.out.println("***********************************checkSavingOpeningBalanceVoucher1WithCurrencyAsINR*************************************");
		
 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
 		billRefNewReferenceTxt.click();
 		
 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
 		billRefPickIcon.click();
    	
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
        billRefOkBtn.click();
		
        Thread.sleep(2000);
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingBalancesSaveBtn));
		openingBalancesSaveBtn.click();
		
		/*Thread.sleep(1000);*/

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

		//expMsg.add(excelReader.getCellData(xlSheetName, 142, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 143, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 142, 7, actMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
		*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 142, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 143, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 142, 7, actMessage);
	
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

		String expSuspendMessage1=excelReader.getCellData(xlSheetName, 144, 6);
		String expSuspendMessage2=excelReader.getCellData(xlSheetName, 145, 6);

		String actSuspendMessage=checkValidationMessage(expSuspendMessage1);
		
		excelReader.setCellData(xlfile, xlSheetName, 144, 7, actSuspendMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		
		String actSuspendStatus = voucherHomeRow1SuspendedStatus.getText();
		String expSuspendStatus = excelReader.getCellData(xlSheetName, 146, 6);
		excelReader.setCellData(xlfile, xlSheetName, 146, 7, actSuspendStatus);
		
		System.out.println("Suspended Status : "+actSuspendStatus+"  Value Expected  "+expSuspendStatus);
		
		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2) && actSuspendMessage.startsWith(expSuspendMessage1) && actSuspendMessage.endsWith(expSuspendMessage2)
				&& actSuspendStatus.equalsIgnoreCase(expSuspendStatus))
		{
			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			*/
			excelReader.setCellData(xlfile, xlSheetName, 142, 8, resPass);
			return true;
		}
		else
		{
		/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();*/
			excelReader.setCellData(xlfile, xlSheetName, 142, 8, resFail);
			return false;
		}
	}
	
	
	
	

	
	public boolean checkResavingOpeningBalanceVoucher2AndSuspendingFromHomepage() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 147, 5);

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
		expMessage.add("Voucher saved successfully : 2");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 147, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 148, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 147, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
		*/
		

		String expMessage1 = excelReader.getCellData(xlSheetName, 147, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 148, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 147, 7, actMessage);
	
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
			String data =excelReader.getCellData(xlSheetName, 148, 5);

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
		
		String expSuspendMessage = excelReader.getCellData(xlSheetName, 149, 6);
		String actSuspendMessage = checkValidationMessage(expSuspendMessage);
		excelReader.setCellData(xlfile, xlSheetName, 149, 7, actSuspendMessage);
		
		Thread.sleep(2000);
		
		String actSuspendStatus = voucherHomeRow1SuspendedStatus.getText();
		String expSuspendStatus = excelReader.getCellData(xlSheetName, 150, 6);
		excelReader.setCellData(xlfile, xlSheetName, 150, 7, actSuspendStatus);
		
		System.out.println("Suspended Status : "+actSuspendStatus+"  Value Expected  "+expSuspendStatus);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2) && actSuspendMessage.equalsIgnoreCase(expSuspendMessage)
				&& actSuspendStatus.equalsIgnoreCase(expSuspendStatus))
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
	
	
	
	
	

	public boolean checkDeleteOptionInOpeningBalanceHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 151, 5);

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

		String expAlertMessage=excelReader.getCellData(xlSheetName, 151, 6);
		excelReader.setCellData(xlfile, xlSheetName, 151, 7, actAlertMessage);


		System.out.println("Alert Message on Deleteing the Voucher :  "+actAlertMessage +"  "+expAlertMessage);

		Thread.sleep(3000);

		getAlert().accept();

		String expMessage = excelReader.getCellData(xlSheetName, 152, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 152, 7, actMessage);

		System.out.println("Message : "+actMessage+" Value Expected : "+expMessage);

		if(actAlertMessage.contains(expAlertMessage) && actMessage.equalsIgnoreCase(expMessage))
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

	
	
	



	public boolean checkVerifingDetailsOfSavedOpeningBalancesVouchersInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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

			if (data.isEmpty()==false && i!=9 && i!=10 && i!=22 && i!=21) 
			{
				System.out.println(data);

				actVouchersList.add(data);
			}
		}


		ArrayList<String>  expVouchersList= new ArrayList<String>(); 

		/*expVouchersList.add("01/01/2020");
			expVouchersList.add("1");
			expVouchersList.add("SU");
			expVouchersList.add("SU");
			expVouchersList.add(currentDate);
			expVouchersList.add(currentDate);
			expVouchersList.add("False");
			expVouchersList.add("Authorized");
			expVouchersList.add("INDIA");*/
		expVouchersList.add(excelReader.getCellData(xlSheetName, 153, 6));
		expVouchersList.add("01/01/2020");
		expVouchersList.add(excelReader.getCellData(xlSheetName, 154, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 155, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 156, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 157, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 158, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 159, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 160, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 161, 6));
		
		excelReader.setCellData(xlfile, xlSheetName, 153, 7, actVouchersList.toString());

		System.out.println("**********************checkVerifingDetailsOfSavedOpeningStocksVouchersInHomepage*********************");

		System.out.println(actVouchersList);
		System.out.println(expVouchersList);

		if (actVouchersList.equals(expVouchersList)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 153, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 153, 8, resFail);
			return false;
		}
	}







	public OpeningBalancesPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
