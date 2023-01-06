package com.focus.Pages;

import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
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

public class DocumentNumberingPage extends BaseEngine {
	
	@FindBy(xpath="//*[@id='4608']")
	public static WebElement docCustmReceiptsVoucher;
	
	
	@FindBy(xpath="//*[@id='4609']")
	public static WebElement ReceiptsVATVoucher;
	
	@FindBy(xpath="//*[@id='Accounts']")
	public static WebElement accountingTag;
	
	@FindBy(xpath="//*[@id='panelsStayOpen-headingOne']/button")
	private static WebElement generalBtn;
	
	@FindBy(xpath="//*[@id='dvTag']/div/div[1]/div")
	private static WebElement accounting;
	
	@FindBy(xpath="//*[@id='docNum_documentNumberAs']")
	private static WebElement docNumberTxt;
	
	@FindBy(xpath="//input[@id='id_header_268435470']")
	private static WebElement  placeofSupplyTxt;

	@FindBy(xpath="//tbody[@id='id_header_268435470_table_body']/tr/td[2]")
	private static List<WebElement>  placeofSupplyList;

	@FindBy(xpath="//tbody[@id='id_header_268435471_table_body']/tr/td[2]")
	private static List<WebElement>  jurisdictionList;

	
	public boolean checkNavigateToDocumentNumberingTabInReceiptsVATVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{	
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		
		Thread.sleep(2000);

		int count = SettingsSubMenusList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			ScrollIntoView(SettingsSubMenusList.get(i));
			
			if (data.equalsIgnoreCase("Configure Transactions")) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionPreferencesMinus));
		configureTransactionPreferencesMinus.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionDocumentCustomiztionMinus));
		configureTransactionDocumentCustomiztionMinus.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustmReceiptsVoucher));
		docCustmReceiptsVoucher.click();
		
		
		Thread.sleep(8000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ReceiptsVATVoucher));
		ReceiptsVATVoucher.click();
		Thread.sleep(4000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberingBtn));
		documentNumberingBtn.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow1));
		typeRow1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
		Select typeDropDownSelect = new Select(typeDropDown);

		typeDropDownSelect.selectByVisibleText("Login");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow1));
		fieldValueRow1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueDropDown));
		Select valueDropDownSelect= new Select(fieldValueDropDown);

		valueDropDownSelect.selectByVisibleText("Login Name");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingCharacterRow1));
		startingCharacterRow1.click();

		startingCharacterTxt.sendKeys(Keys.BACK_SPACE);
		
		startingCharacterTxt.sendKeys("0");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow1));
		noofCharactersRow1.click();

		noOfCharacterTxt.sendKeys("2");

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow2));
		typeRow2.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
		typeDropDownSelect.selectByVisibleText("Character Input");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow2));
		fieldValueRow2.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueDropDown));
		Select fieldValueDropDownSelect = new Select(fieldValueDropDown);

		fieldValueDropDownSelect.selectByVisibleText("/");


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow2));
		noofCharactersRow2.click();

		noOfCharacterTxt.sendKeys("1");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow3));
		typeRow3.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
		typeDropDownSelect.selectByVisibleText("Department");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow3));
		fieldValueRow3.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueDropDown));
		fieldValueDropDownSelect.selectByVisibleText("Name");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingCharacterRow3));
		startingCharacterRow3.click();

		startingCharacterTxt.sendKeys(Keys.BACK_SPACE);
		startingCharacterTxt.sendKeys("0");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow3));
		noofCharactersRow3.click();

		noOfCharacterTxt.sendKeys("3");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow4));
		typeRow4.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
		typeDropDownSelect.selectByVisibleText("Character Input");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow4));
		fieldValueRow4.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueDropDown));
		fieldValueDropDownSelect.selectByVisibleText("/");


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow4));
		noofCharactersRow4.click();

		noOfCharacterTxt.sendKeys("1");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow5));
		typeRow5.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
		typeDropDownSelect.selectByVisibleText("Input");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow5));
		fieldValueRow5.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueTxt));
		fieldValueTxt.sendKeys("Receipt");


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow5));
		noofCharactersRow5.click();

		noOfCharacterTxt.sendKeys("1");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow6));
		typeRow6.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
		typeDropDownSelect.selectByVisibleText("Increment");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow6));
		fieldValueRow6.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueTxt));
		fieldValueTxt.sendKeys("0");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingCharacterRow6));
		startingCharacterRow6.click();

		startingCharacterTxt.sendKeys("0");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow6));
		noofCharactersRow6.click();

		noOfCharacterTxt.sendKeys("1");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow7));
		typeRow7.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docNumberAppearAsValue));

		String actdocNumberAppearAsValue = docNumberAppearAsValue.getText();

		String expdocNumberAppearAsValue = "SU/Tag/Receipt0";
		
		

		
		


		System.out.println("docNumberAppearAsValue Display Value Actual      : " + actdocNumberAppearAsValue);
		System.out.println("docNumberAppearAsValue Display Value Expected    : " + expdocNumberAppearAsValue);

		if(actdocNumberAppearAsValue.equalsIgnoreCase(expdocNumberAppearAsValue)) 
		{
			
			return true;
		} 
		else 
		{
			
			return false;
		}
	}
	
	
	
	public boolean checkSavingRecepitsVATVocher1withDocumentNumber() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(2000);
		LogoutandLoginwithSU();
		
		getDriver().navigate().refresh();
		Thread.sleep(3000);
		
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
	finacinalsMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
	cashAndBankMenu.click();

	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATVoucher));
	recepitsVATVoucher.click();

	Thread.sleep(2000);

	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
	newBtn.click();

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCashBankAccountTxt));
	newCashBankAccountTxt.click();

	newCashBankAccountTxt.sendKeys(Keys.SPACE);


	int cashAndBAnkAccountListCount  = cashAndBAnkAccountList.size();

	System.err.println("cashAndBAnkAccountListCount   : "+cashAndBAnkAccountListCount);

	for (int i = 0; i < cashAndBAnkAccountListCount; i++) 
	{
		String data		  = cashAndBAnkAccountList.get(i).getText();

		if(data.equalsIgnoreCase("Bank"))
		{
			cashAndBAnkAccountList.get(i).click();

			break;
		}
	}

	newCashBankAccountTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));

	departmentTxt.click();
	departmentTxt.sendKeys(Keys.SHIFT,Keys.HOME,Keys.BACK_SPACE);
	departmentTxt.sendKeys(Keys.SPACE);
	Thread.sleep(2000);
	int departmentcount=departmentListCount.size();

	System.err.println(departmentcount);

	for(int i=0 ; i < departmentcount ;i++)
	{
		String data=departmentListCount.get(i).getText();

		if(data.equalsIgnoreCase("DUBAI"))
		{
			departmentListCount.get(i).click();
			break;
		}
	}

	Thread.sleep(1000);

	departmentTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeofSupplyTxt));
	placeofSupplyTxt.click();
	placeofSupplyTxt.sendKeys(Keys.SPACE);

	int placeOfSupplyListCount=placeofSupplyList.size();

	System.err.println(placeOfSupplyListCount);

	for(int i=0 ; i < placeOfSupplyListCount ;i++)
	{
		String data=placeofSupplyList.get(i).getText();

		if(data.equalsIgnoreCase("Abu Dhabi"))
		{
			placeofSupplyList.get(i).click();

			break;
		}
	}

	placeofSupplyTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
	jurisdictionTxt.click();
	jurisdictionTxt.sendKeys(Keys.SPACE);

	int jurisdictionListCount=jurisdictionList.size();

	System.err.println(jurisdictionListCount);

	for(int i=0 ; i < jurisdictionListCount ;i++)
	{
		String data=jurisdictionList.get(i).getText();

		if(data.equalsIgnoreCase("DUBAI"))
		{
			jurisdictionList.get(i).click();

			break;
		}
	}

	jurisdictionTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String docno=documentNumberTxt.getAttribute("value");

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	enter_AccountTxt.click();
	enter_AccountTxt.sendKeys("Cus");


	getFluentWebDriverWait().until(ExpectedConditions.visibilityOfAllElements(bodyAccountListInGrid));
	int accountCount=bodyAccountListInGrid.size();

	System.err.println(accountCount);

	for(int i=0 ; i < accountCount ;i++)
	{
		String data=bodyAccountListInGrid.get(i).getText();


		if(data.equalsIgnoreCase("Customer A"))
		{
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOfAllElements(bodyAccountListInGrid));
			bodyAccountListInGrid.get(i).click();

			break;
		}
	}

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));	
	enter_AccountTxt.sendKeys(Keys.TAB);


	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enterReceiptsVATTaxCode));	
	enterReceiptsVATTaxCode.sendKeys(Keys.TAB);


	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Amount));
	enter_Amount.sendKeys("50");
	enter_Amount.sendKeys(Keys.TAB);

	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
	newReferenceTxt.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
	billRefPickIcon.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
	billRefOkBtn.click();
	
	Thread.sleep(1000);
	
	
	System.err.println(docno);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));	
	voucherSaveBtn.click();
	
	if(docno.equalsIgnoreCase("SU/DUB/Receipt1"))
	{
		return true;
	}
	
	else
	{
	return false;
	}
	
	}
	
	@FindBy(xpath="//*[@id='id_transactionentry_settings']/label")
	public static WebElement receiptsSettingsBtn;
	
	@FindBy(xpath="//*[@id='docNum_Table']//tr[5]//td[1]")
	public static WebElement selectDocNum5thRow_1stColumn;
	
	@FindBy(xpath="//*[@id='DocumentNumbering']/div/div/div/div[2]/input[2]")
	public static WebElement insertRowBtn;
	
	
	@FindBy(xpath="//*[@id='DocumentNumbering']/div/div/div/div[2]/input[1]")
	public static WebElement deleteRowBtn;
	
	@FindBy(xpath="//*[@id='btnCustomizeClose']/i")
	public static WebElement receiptsSettingsClsBtn;
	
	
	@FindBy(xpath="//*[@id='docNum_Table_body']/tr[5]/td[2]")
	public static WebElement insertedRow_2ndColumn;
	
	@FindBy(xpath="//*[@id='docNum_Table_body']/tr[5]/td[3]")
	public static WebElement insertedRow_3rdColumn;
	
	@FindBy(xpath="//*[@id='docNum_Table_body']/tr[5]/td[4]")
	public static WebElement insertedRow_4thColumnTxt;
	
	@FindBy(xpath="//*[@id='docNum_Table_body']/tr[5]/td[5]")
	public static WebElement insertedRow_5thColumnTxt;
	
	@FindBy(xpath="//*[@id='docNum_noofCharacterTxtbox']")
	public static WebElement insertedRow_5thColumn;
	
	@FindBy(xpath="//*[@id='docNum_startCharacterTxtbox']")
	public static WebElement insertedRow_4thColumn;
	
	public void checkInsertRowinDocumentNumberofReceiptsVATVoucher() throws InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATVoucher));
		recepitsVATVoucher.click();

		Thread.sleep(2000);

		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));
		toggleBtn.click();
		Thread.sleep(1000);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(receiptsSettingsBtn));
		receiptsSettingsBtn.click();
		
		Thread.sleep(1000);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberingBtn));
		documentNumberingBtn.click();
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectDocNum5thRow_1stColumn));
		
		selectDocNum5thRow_1stColumn.click();
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(insertRowBtn));
		insertRowBtn.click();
		
		Thread.sleep(1000);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(insertedRow_2ndColumn));
		insertedRow_2ndColumn.click();
		
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow5));
		typeRow5.click();*/
		
		Select typeDropDownSelect = new Select(typeDropDown);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
		typeDropDownSelect.selectByVisibleText("Date");

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(insertedRow_3rdColumn));
		insertedRow_3rdColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueDropDown));
		Select valueDropDownSelect= new Select(fieldValueDropDown);

		valueDropDownSelect.selectByVisibleText("Month Name");
		Thread.sleep(1000);

		//insertedRow_3rdColumn.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(insertedRow_4thColumnTxt));
		insertedRow_4thColumnTxt.click();
		insertedRow_4thColumn.sendKeys("0");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(insertedRow_5thColumnTxt));
		insertedRow_5thColumnTxt.click();
		insertedRow_5thColumn.sendKeys("3");
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		Thread.sleep(4000);
		
			
		
	}
	
	
	public boolean checkSavingRecepitsVATVocher2withChangedDocumentNumber() throws InterruptedException
	{
		checkInsertRowinDocumentNumberofReceiptsVATVoucher();
		
		LogoutandLoginwithSU();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATVoucher));
		recepitsVATVoucher.click();

		Thread.sleep(2000);

		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCashBankAccountTxt));
		newCashBankAccountTxt.click();

		newCashBankAccountTxt.sendKeys(Keys.SPACE);


		int cashAndBAnkAccountListCount  = cashAndBAnkAccountList.size();

		System.err.println("cashAndBAnkAccountListCount   : "+cashAndBAnkAccountListCount);

		for (int i = 0; i < cashAndBAnkAccountListCount; i++) 
		{
			String data		  = cashAndBAnkAccountList.get(i).getText();

			if(data.equalsIgnoreCase("Bank"))
			{
				cashAndBAnkAccountList.get(i).click();

				break;
			}
		}

		newCashBankAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));

		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SHIFT,Keys.HOME,Keys.BACK_SPACE);
		departmentTxt.sendKeys(Keys.SPACE);
		Thread.sleep(2000);
		int departmentcount=departmentListCount.size();

		System.err.println(departmentcount);

		for(int i=0 ; i < departmentcount ;i++)
		{
			String data=departmentListCount.get(i).getText();

			if(data.equalsIgnoreCase("DUBAI"))
			{
				departmentListCount.get(i).click();
				break;
			}
		}

		Thread.sleep(1000);

		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeofSupplyTxt));
		placeofSupplyTxt.click();
		placeofSupplyTxt.sendKeys(Keys.SPACE);

		int placeOfSupplyListCount=placeofSupplyList.size();

		System.err.println(placeOfSupplyListCount);

		for(int i=0 ; i < placeOfSupplyListCount ;i++)
		{
			String data=placeofSupplyList.get(i).getText();

			if(data.equalsIgnoreCase("Abu Dhabi"))
			{
				placeofSupplyList.get(i).click();

				break;
			}
		}

		placeofSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(Keys.SPACE);

		int jurisdictionListCount=jurisdictionList.size();

		System.err.println(jurisdictionListCount);

		for(int i=0 ; i < jurisdictionListCount ;i++)
		{
			String data=jurisdictionList.get(i).getText();

			if(data.equalsIgnoreCase("DUBAI"))
			{
				jurisdictionList.get(i).click();

				break;
			}
		}

		jurisdictionTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
		enter_AccountTxt.click();
		enter_AccountTxt.sendKeys("Cus");


		getFluentWebDriverWait().until(ExpectedConditions.visibilityOfAllElements(bodyAccountListInGrid));
		int accountCount=bodyAccountListInGrid.size();

		System.err.println(accountCount);

		for(int i=0 ; i < accountCount ;i++)
		{
			String data=bodyAccountListInGrid.get(i).getText();


			if(data.equalsIgnoreCase("Customer A"))
			{
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOfAllElements(bodyAccountListInGrid));
				bodyAccountListInGrid.get(i).click();

				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));	
		enter_AccountTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enterReceiptsVATTaxCode));	
		enterReceiptsVATTaxCode.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Amount));
		enter_Amount.sendKeys("50");
		enter_Amount.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		Thread.sleep(1000);
		
		String actDocno=documentNumberTxt.getAttribute("value");;
		String expDocno="SU/DUB/MarReceipt1";
		
		System.err.println("Actual docno: " + actDocno +"Expected:   " + expDocno);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));	
		voucherSaveBtn.click();
		
		if(actDocno.equalsIgnoreCase(expDocno))
		{
			return true;
		}
		
		else
		{
		return false;
		}
		
		
		
		
		
	}
	
	
	public void checkDeleteRowinDocumentNumberofReceiptsVATVoucher() throws InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATVoucher));
		recepitsVATVoucher.click();

		Thread.sleep(2000);

		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));
		toggleBtn.click();
		Thread.sleep(1000);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(receiptsSettingsBtn));
		receiptsSettingsBtn.click();
		
		Thread.sleep(1000);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberingBtn));
		documentNumberingBtn.click();
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectDocNum5thRow_1stColumn));
		
		selectDocNum5thRow_1stColumn.click();
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteRowBtn));
		deleteRowBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		Thread.sleep(1000);
		
				
		
		
	}
	
	
	public boolean checkSavingRecepitsVATVocher3withChangedDocumentNumber() throws InterruptedException
	{
		checkDeleteRowinDocumentNumberofReceiptsVATVoucher();
		LogoutandLoginwithSU();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATVoucher));
		recepitsVATVoucher.click();

		Thread.sleep(2000);

		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCashBankAccountTxt));
		newCashBankAccountTxt.click();

		newCashBankAccountTxt.sendKeys(Keys.SPACE);


		int cashAndBAnkAccountListCount  = cashAndBAnkAccountList.size();

		System.err.println("cashAndBAnkAccountListCount   : "+cashAndBAnkAccountListCount);

		for (int i = 0; i < cashAndBAnkAccountListCount; i++) 
		{
			String data		  = cashAndBAnkAccountList.get(i).getText();

			if(data.equalsIgnoreCase("Bank"))
			{
				cashAndBAnkAccountList.get(i).click();

				break;
			}
		}

		newCashBankAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));

		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SHIFT,Keys.HOME,Keys.BACK_SPACE);
		departmentTxt.sendKeys(Keys.SPACE);
		Thread.sleep(2000);
		int departmentcount=departmentListCount.size();

		System.err.println(departmentcount);

		for(int i=0 ; i < departmentcount ;i++)
		{
			String data=departmentListCount.get(i).getText();

			if(data.equalsIgnoreCase("DUBAI"))
			{
				departmentListCount.get(i).click();
				break;
			}
		}

		Thread.sleep(1000);

		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeofSupplyTxt));
		placeofSupplyTxt.click();
		placeofSupplyTxt.sendKeys(Keys.SPACE);

		int placeOfSupplyListCount=placeofSupplyList.size();

		System.err.println(placeOfSupplyListCount);

		for(int i=0 ; i < placeOfSupplyListCount ;i++)
		{
			String data=placeofSupplyList.get(i).getText();

			if(data.equalsIgnoreCase("Abu Dhabi"))
			{
				placeofSupplyList.get(i).click();

				break;
			}
		}

		placeofSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(Keys.SPACE);

		int jurisdictionListCount=jurisdictionList.size();

		System.err.println(jurisdictionListCount);

		for(int i=0 ; i < jurisdictionListCount ;i++)
		{
			String data=jurisdictionList.get(i).getText();

			if(data.equalsIgnoreCase("DUBAI"))
			{
				jurisdictionList.get(i).click();

				break;
			}
		}

		jurisdictionTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
		enter_AccountTxt.click();
		enter_AccountTxt.sendKeys("Cus");


		getFluentWebDriverWait().until(ExpectedConditions.visibilityOfAllElements(bodyAccountListInGrid));
		int accountCount=bodyAccountListInGrid.size();

		System.err.println(accountCount);

		for(int i=0 ; i < accountCount ;i++)
		{
			String data=bodyAccountListInGrid.get(i).getText();


			if(data.equalsIgnoreCase("Customer A"))
			{
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOfAllElements(bodyAccountListInGrid));
				bodyAccountListInGrid.get(i).click();

				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));	
		enter_AccountTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enterReceiptsVATTaxCode));	
		enterReceiptsVATTaxCode.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Amount));
		enter_Amount.sendKeys("50");
		enter_Amount.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		Thread.sleep(1000);
		
		String actDocno=documentNumberTxt.getAttribute("value");;
		String expDocno="SU/DUB/Receipt2";
		
		System.err.println("Actual docno: " + actDocno +"Expected:   " + expDocno);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));	
		voucherSaveBtn.click();
		
		if(actDocno.equalsIgnoreCase(expDocno))
		{
			return true;
		}
		
		else
		{
		return false;
		}
		
		
		
		
		
	}
	
	public void LogoutandLoginwithSU() throws InterruptedException
	{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();
		
		Thread.sleep(3000);
		
		 getDriver().navigate().refresh();
		 Thread.sleep(1999);
		
		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt="su";
				      
		String pawslt="su";
				      
		LoginPage.enterUserName(unamelt);

		Thread.sleep(1000);
				
		LoginPage.enterPassword(pawslt);
	
		 Select oSelect = new Select(companyDropDownList);
		 
		 oSelect.selectByValue("36");
		
		 LoginPage.clickOnSignInBtn();
		 Thread.sleep(2000);
	}
		 
	
	
	public DocumentNumberingPage(WebDriver driver)
	{
		PageFactory.initElements(driver, this);

	}



}
