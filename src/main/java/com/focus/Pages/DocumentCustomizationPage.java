package com.focus.Pages;

import java.io.IOException;
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
import com.gargoylesoftware.htmlunit.javascript.host.dom.Document;

public class DocumentCustomizationPage extends BaseEngine 
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "SanityDocCustViewPage";
	
	public static String checkValidationMessage(String ExpMessage) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		try
		{
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(errorMessage));
			String actErrorMessage=errorMessage.getText();
			String expErrorMessage=ExpMessage;

			try
			{

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
				errorMessageCloseBtn.click();

				System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);

				return actErrorMessage;
			}
			catch(Exception ee)
			{

				System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);

				return actErrorMessage;
			}
		}
		catch(Exception e)
		{
			System.err.println("Error Message NOT Found or NOT Clickable");
			System.err.println(e.getMessage());

			String Exception=e.getMessage();

			return Exception;
		}
	}

@FindBy(xpath="//*[@id='panelsStayOpen-headingOne']/button")
private static WebElement generalBtn;

@FindBy(xpath="//*[@id='769']")
public static WebElement createdVoucher;

@FindBy(xpath="//*[@id='id_search_menu']/input")
public static WebElement searchTxt;


	public boolean checkSavingDocumentCustomizationPurchaseVoucherDocumentsTabAndVerify() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

	/*	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
		searchTxt.click();
		searchTxt.sendKeys("Configure Transactions");
		searchTxt.sendKeys(Keys.ENTER);
		Thread.sleep(3000);*/
		
		
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		
		Thread.sleep(2000);

		int count = SettingsSubMenusList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			//ScrollIntoView(SettingsSubMenusList.get(i));
			System.out.println(excelReader.getCellData(xlSheetName, 8, 5));
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 8, 5))) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}
		
		Thread.sleep(2000);
		
*/	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
		searchTxt.click();
		searchTxt.sendKeys("Configure Transactions");
		searchTxt.sendKeys(Keys.ENTER);
		Thread.sleep(30000);
		
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionPreferencesMinus));
		configureTransactionPreferencesMinus.click();
		
		Thread.sleep(2000);*/
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionDocumentCustomiztionMinus));
		configureTransactionDocumentCustomiztionMinus.click();*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustPurchaseVouchersBtn));
		docCustPurchaseVouchersBtn.click();

		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustCreateVoucherBtn));
		docCustCreateVoucherBtn.click();

		Thread.sleep(8000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdVoucher));
		createdVoucher.click();
		Thread.sleep(2000);
		
		
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(generalBtn));
		generalBtn.click();
		*/Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(titleTxt));
		titleTxt.click();
		titleTxt.clear();
		titleTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		titleTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masters1));
		masters1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDropDown));
		masterDropDown.click();
		masterDropDown.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		masterDropDown.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(position1));
		position1.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(positionDropdown));
		positionDropdown.click();
		positionDropdown.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		positionDropdown.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 8, 6);
		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actMessage);

		String acttitleTxt					 =titleTxt.getAttribute("value");
		String exptitleTxt					 =excelReader.getCellData(xlSheetName, 9, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, acttitleTxt);

		String actMastersR1C1                = masters1.getText();
		String actPositionR1C1               = position1.getText();

		String expMastersR1C1                = excelReader.getCellData(xlSheetName, 10, 6);		
		excelReader.setCellData(xlfile, xlSheetName, 10, 7, actMastersR1C1);
		
		String expPositionR1C1               = excelReader.getCellData(xlSheetName, 11, 6);
		excelReader.setCellData(xlfile, xlSheetName, 10, 7, actPositionR1C1);

		System.out.println("**************************checkSavingDocumentCustomizationPurchaseVoucherDocumentsTabAndVerify*****************************");

		System.out.println("Error Message                                 : " + actMessage       +     "     Value Expected : " + expMessage);
		System.out.println("Documents Tab Title                           : " + acttitleTxt      +     "	 Value Expected : " + exptitleTxt);
		System.out.println("Documents Tab Master First Row                : " + actMastersR1C1   +     "	 Value Expected : " + expMastersR1C1);
		System.out.println("Documents Tab Master Position First Row       : " + actPositionR1C1  +     "	 Value Expected : " + expPositionR1C1);

		if(actMessage.equalsIgnoreCase(expMessage) && acttitleTxt.equalsIgnoreCase(exptitleTxt) && actMastersR1C1.equalsIgnoreCase(expMastersR1C1) 
				&& actPositionR1C1.equalsIgnoreCase(expPositionR1C1))
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


	@FindBy(xpath = "//div[@id='id_focus_msgbox_title']/div[2]/span")
    private static WebElement ServerErrorPopupCloseBtn;

    @FindBy(xpath = "//*[@id='id_focus_msgbox_detail']/span/h2/i")
    private static WebElement ServerErrorPopupTxt;



	public boolean checkEditingPurchaseVoucherDocumentAndUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(titleTxt));
		titleTxt.click();
		titleTxt.clear();
		titleTxt.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		titleTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masters2));
		masters2.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDropDown));
		masterDropDown.click();
		masterDropDown.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
		
		masterDropDown.sendKeys(Keys.TAB);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(position2));
		position2.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(positionDropdown));
		positionDropdown.click();
		positionDropdown.sendKeys(excelReader.getCellData(xlSheetName, 14, 5));
		positionDropdown.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		Thread.sleep(2000);
		
		
		

		String expMessage = excelReader.getCellData(xlSheetName, 12, 6);
		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actMessage);


		String acttitleTxt					 =titleTxt.getAttribute("value");
		String exptitleTxt					 =excelReader.getCellData(xlSheetName, 13, 6);
		excelReader.setCellData(xlfile, xlSheetName, 13, 7, acttitleTxt);

		String actMastersR1C1                = masters2.getText();
		String actPositionR1C1               = position2.getText();

		String expMastersR1C1                = excelReader.getCellData(xlSheetName, 14, 6);
		excelReader.setCellData(xlfile, xlSheetName, 14, 7, actMastersR1C1);
		
		
		String expPositionR1C1               = excelReader.getCellData(xlSheetName, 15, 6);
		excelReader.setCellData(xlfile, xlSheetName, 15, 7, actPositionR1C1);

		System.out.println("*****************************checkEditingPurchaseVoucherDocumentAndUpdate****************************");

		System.out.println("Error Message                                 : " + actMessage       +     "     Value Expected : " + expMessage);
		System.out.println("Documents Tab Title                           : " + acttitleTxt      +     "	 Value Expected : " + exptitleTxt);
		System.out.println("Documents Tab Master First Row                : " + actMastersR1C1   +     "	 Value Expected : " + expMastersR1C1);
		System.out.println("Documents Tab Master Position First Row       : " + actPositionR1C1  +     "	 Value Expected : " + expPositionR1C1);

		if(actMessage.equalsIgnoreCase(expMessage) && acttitleTxt.equalsIgnoreCase(exptitleTxt) && actMastersR1C1.equalsIgnoreCase(expMastersR1C1) 
				&& actPositionR1C1.equalsIgnoreCase(expPositionR1C1))
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 8, resFail);
			return false;
		} 
	}





@FindBy(xpath="(//*[@id='panelsStayOpen-headingOne']/button)[2]")
public static WebElement miscVoucherSettingsBtn;



	public boolean checkUpdatingMiscellaneousTabInPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousTab));
		miscellaneousTab.click();
		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscVoucherSettingsBtn));
		miscVoucherSettingsBtn.click();
		*/Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(abbreviationTxt));
		abbreviationTxt.click();
		abbreviationTxt.clear();
		abbreviationTxt.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));
		abbreviationTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneusTabPostingDetailsMenu));
		ScrollIntoView(miscellaneusTabPostingDetailsMenu);
		miscellaneusTabPostingDetailsMenu.click();
		
		Thread.sleep(1000);
		
		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateFAChkBox));
		
		/*if(updateFAChkBoxIsSelected.isSelected()==false)
		{
			updateFAChkBox.click();
		}
		getAction().moveToElement(updateFAChkBox).build().perform();
		updateFAChkBox.click();*/
		
		((JavascriptExecutor)getDriver()).executeScript("document.getElementById('misc_PostingUpdateFA').checked=false;");
	
		/*	updateFAChkBox.click();
		misc_PostingUpdateFA
	*/	
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateStockChkBox));
		if(updateStockChkBoxIsSelected.isSelected()==false)
		{
			updateStockChkBox.click();
		}
		
		/*getAction().moveToElement(updateStockChkBox).click().build().perform();
		updateStockChkBox.click();*/
		
		//((JavascriptExecutor)getDriver()).executeScript("document.getElementById('misc_chkUpdateStock').checked=false;");
		Thread.sleep(1000);
		/*updateStockChkBox.click();
		Thread.sleep(2000);
		*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousTabCurrencyMenu));
		ScrollIntoView(miscellaneousTabCurrencyMenu);
		miscellaneousTabCurrencyMenu.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inputExchangeRateChkBox));
		
		/*if(inputExchangeRateChkBoxIsSelected.isSelected()==false)
		{
			inputExchangeRateChkBox.click();
		}
		*/
		
		/*inputExchangeRateChkBox.click();
		Thread.sleep(2000);*/

		//getAction().moveToElement(inputLocalExchangeRateChkBox).click().build().perform();
		((JavascriptExecutor)getDriver()).executeScript("document.getElementById('misc_currencyInputExchangeRate').checked=false;");
		Thread.sleep(2000);
		//inputExchangeRateChkBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inputLocalExchangeRateChkBox));
		
		/*if(inputLocalExchangeRateChkBoxIsSelected.isSelected()==false)
		{
			inputLocalExchangeRateChkBox.click();
		}
		*/
		
		/*inputLocalExchangeRateChkBox.click();
		getAction().moveToElement(inputLocalExchangeRateChkBox).click().build().perform();*/
		
		((JavascriptExecutor)getDriver()).executeScript("document.getElementById('misc_currencyInputLocalExchangeRate').checked=false;");
		/*inputLocalExchangeRateChkBox.click();
		*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		Thread.sleep(2000);

		String expMessage = excelReader.getCellData(xlSheetName, 16, 6);
		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 16, 7, actMessage);


		String actabbreviationTxt				= abbreviationTxt.getAttribute("value");
		String expabbreviationTxt				= excelReader.getCellData(xlSheetName, 17, 6);
		excelReader.setCellData(xlfile, xlSheetName, 17, 7, actabbreviationTxt);
		
		String actupdateFAChkBox				= Boolean.toString(updateFAChkBoxIsSelected.isSelected());
		String expupdateFAChkBox				= excelReader.getCellData(xlSheetName, 18, 6);
		excelReader.setCellData(xlfile, xlSheetName, 18, 7, actupdateFAChkBox);
		
		
		String actupdateStockChkBox			= Boolean.toString(updateStockChkBoxIsSelected.isSelected());
		String expupdateStockChkBox			= excelReader.getCellData(xlSheetName, 19, 6);
		excelReader.setCellData(xlfile, xlSheetName, 19, 7, actupdateStockChkBox);
		
		String actinputExchangeRateChkBox		= Boolean.toString(inputExchangeRateChkBoxIsSelected.isSelected());
		String expinputExchangeRateChkBox		= excelReader.getCellData(xlSheetName, 20, 6);
		excelReader.setCellData(xlfile, xlSheetName, 20, 7, actinputExchangeRateChkBox);
		
		String actinputLocalExchangeRateChkBox	= Boolean.toString(inputLocalExchangeRateChkBoxIsSelected.isSelected());		
		String expinputLocalExchangeRateChkBox	= excelReader.getCellData(xlSheetName, 21, 6);	
		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actinputLocalExchangeRateChkBox);

		System.out.println("********************************checkUpdatingMiscellaneousTabInPurchaseVoucher**************************************");

		System.out.println("ErrorMessage                  : " + actMessage                       + "  value expected  " + expMessage);
		System.out.println("AbbreviationTxt               : " + actabbreviationTxt               + "  value expected  " + expabbreviationTxt);
		System.out.println("UpdateFAChkBox                : " + actupdateFAChkBox                + "  value expected  " + expupdateFAChkBox);
		System.out.println("UpdateStockChkBox             : " + actupdateStockChkBox             + "  value expected  " + expupdateStockChkBox);
		System.out.println("InputExchangeRateChkBox       : " + actinputExchangeRateChkBox       + "  value expected  " + expinputExchangeRateChkBox);
		System.out.println("InputLocalExchangeRateChkBox  : " + actinputLocalExchangeRateChkBox  + "  value expected  " + expinputLocalExchangeRateChkBox);

		if(actMessage.equalsIgnoreCase(expMessage) && actabbreviationTxt.equalsIgnoreCase(actabbreviationTxt) && actupdateFAChkBox.equalsIgnoreCase(expupdateFAChkBox) 
				&& actupdateStockChkBox.equalsIgnoreCase(expupdateStockChkBox) && actinputExchangeRateChkBox.equalsIgnoreCase(expinputExchangeRateChkBox) && actinputLocalExchangeRateChkBox.equalsIgnoreCase(expinputLocalExchangeRateChkBox))
		{
			excelReader.setCellData(xlfile, xlSheetName, 16, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 16, 8, resFail);
			return false;
		}	
	}






	public boolean checkEditingMiscellaneousTabInPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(abbreviationTxt));
		abbreviationTxt.click();
		abbreviationTxt.clear();
		abbreviationTxt.sendKeys(excelReader.getCellData(xlSheetName, 22, 5));
		abbreviationTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneusTabPostingDetailsMenu));
		ScrollIntoView(miscellaneusTabPostingDetailsMenu);
		
		Thread.sleep(1000);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateFAChkBox));
		if(updateFAChkBoxIsSelected.isSelected()==false)
		{
			updateFAChkBox.click();
		}
		//updateFAChkBox.click();
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateStockChkBox));
		
		/*if(updateStockChkBoxIsSelected.isSelected()==false)
		{
			updateStockChkBox.click();
		}
		*///	updateStockChkBox.click();
		((JavascriptExecutor)getDriver()).executeScript("document.getElementById('misc_chkUpdateStock').checked=false;");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousTabCurrencyMenu));
		ScrollIntoView(miscellaneousTabCurrencyMenu);
		
		Thread.sleep(1000);
		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inputExchangeRateChkBox));
		if(inputExchangeRateChkBoxIsSelected.isSelected()==false)
		{
			inputExchangeRateChkBox.click();
		}
		
		
		//inputExchangeRateChkBox.click();
	
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inputLocalExchangeRateChkBox));
		
		if(inputLocalExchangeRateChkBoxIsSelected.isSelected()==false)
		{
			inputLocalExchangeRateChkBox.click();
		}
		//inputLocalExchangeRateChkBox.click();
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		Thread.sleep(2000);

		String expMessage = excelReader.getCellData(xlSheetName, 22, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 22, 7, actMessage);
		

		String actabbreviationTxt				= abbreviationTxt.getAttribute("value");
		String expabbreviationTxt				= excelReader.getCellData(xlSheetName, 23, 6);
		excelReader.setCellData(xlfile, xlSheetName, 23, 7, actabbreviationTxt);
		
		String actupdateFAChkBox				= Boolean.toString(updateFAChkBoxIsSelected.isSelected());
		String expupdateFAChkBox				= excelReader.getCellData(xlSheetName, 24, 6);
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actupdateFAChkBox);
		
		
		String actupdateStockChkBox			= Boolean.toString(updateStockChkBoxIsSelected.isSelected());
		String expupdateStockChkBox			= excelReader.getCellData(xlSheetName, 25, 6);
		excelReader.setCellData(xlfile, xlSheetName, 25, 7, actupdateStockChkBox);
		
		
		String actinputExchangeRateChkBox		= Boolean.toString(inputExchangeRateChkBoxIsSelected.isSelected());
		String expinputExchangeRateChkBox		= excelReader.getCellData(xlSheetName, 26, 6);
		excelReader.setCellData(xlfile, xlSheetName, 26, 7, actinputExchangeRateChkBox);
		
		
		String actinputLocalExchangeRateChkBox	= Boolean.toString(inputLocalExchangeRateChkBoxIsSelected.isSelected());		
		String expinputLocalExchangeRateChkBox	= excelReader.getCellData(xlSheetName, 27, 6);	
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actinputLocalExchangeRateChkBox);

		System.out.println("********************************checkEditingMiscellaneousTabInPurchaseVoucher**************************************");

		System.out.println("ErrorMessage                  : " + actMessage                       + "  value expected  " + expMessage);
		System.out.println("AbbreviationTxt               : " + actabbreviationTxt               + "  value expected  " + expabbreviationTxt);
		System.out.println("UpdateFAChkBox                : " + actupdateFAChkBox                + "  value expected  " + expupdateFAChkBox);
		System.out.println("UpdateStockChkBox             : " + actupdateStockChkBox             + "  value expected  " + expupdateStockChkBox);
		System.out.println("InputExchangeRateChkBox       : " + actinputExchangeRateChkBox       + "  value expected  " + expinputExchangeRateChkBox);
		System.out.println("InputLocalExchangeRateChkBox  : " + actinputLocalExchangeRateChkBox  + "  value expected  " + expinputLocalExchangeRateChkBox);

		if(actMessage.equalsIgnoreCase(expMessage) && actabbreviationTxt.equalsIgnoreCase(actabbreviationTxt) && actupdateFAChkBox.equalsIgnoreCase(expupdateFAChkBox) 
				&& actupdateStockChkBox.equalsIgnoreCase(expupdateStockChkBox) && actinputExchangeRateChkBox.equalsIgnoreCase(expinputExchangeRateChkBox) && actinputLocalExchangeRateChkBox.equalsIgnoreCase(expinputLocalExchangeRateChkBox))
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







	public boolean checkAddFieldInEditScreenBodyInTrialPurchaseUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
		editScreenTab.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBodyTab));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenaddBtn));
		editScreenaddBtn.click();

		Thread.sleep(5000);
	if(getIsAlertPresent())
	{
		getWaitForAlert();
		Thread.sleep(2000);
		getAlert().accept();
		Thread.sleep(2000);
	}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenCaptionTxt));
		editScreenCaptionTxt.click();
		editScreenCaptionTxt.sendKeys(excelReader.getCellData(xlSheetName, 28, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
		editScreenApplyBtn.click();

		System.out.println("*******************************************checkAddFieldInEditScreenBodyInTrialPurchaseUpdate*******************************************");

		String expMessage = excelReader.getCellData(xlSheetName, 28, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBodyRow3Caption));
		String actCaption = editScreenBodyRow3Caption.getText();
		String expCaption = excelReader.getCellData(xlSheetName, 29, 6);
		excelReader.setCellData(xlfile, xlSheetName, 29, 7, actCaption);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBodyRow3Formula));
		String actFormula = Boolean.toString(editScreenBodyRow3Formula.getText().isEmpty());
		String expFormula = excelReader.getCellData(xlSheetName, 30, 6);
		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actFormula);

		System.out.println("Caption        : "+actCaption+"  Value expected  "+expCaption);
		System.out.println("Formula        : "+actFormula+"  Value expected  "+expFormula);

		if(actMessage.equalsIgnoreCase(expMessage) && actCaption.equalsIgnoreCase(expCaption) && actFormula.equalsIgnoreCase(expFormula))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 28, 8, resPass);
			return  true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 28, 8, resFail);
			return  false;
		}
	}







	public boolean checkEditFieldInEditScreenBodyInTrialPurchaseUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBodyRow3Edit));
		editScreenBodyRow3Edit.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenCaptionTxt));
		editScreenCaptionTxt.click();
		editScreenCaptionTxt.clear();
		editScreenCaptionTxt.sendKeys(excelReader.getCellData(xlSheetName, 31, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenAddFieldsBehaviourDropdown));
		editScreenAddFieldsBehaviourDropdown.click();

		Select s=new Select(editScreenAddFieldsBehaviourDropdown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 32, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenAddBehaviourformulaTxt));
		editScreenAddBehaviourformulaTxt.click();
		editScreenAddBehaviourformulaTxt.sendKeys(excelReader.getCellData(xlSheetName, 33, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
		editScreenApplyBtn.click();

		System.out.println("*******************************************checkEditFieldInEditScreenBodyInTrialPurchaseUpdate*******************************************");

		String expMessage = excelReader.getCellData(xlSheetName, 31, 6);
		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBodyRow3Caption));
		String actCaption = editScreenBodyRow3Caption.getText();
		String expCaption = excelReader.getCellData(xlSheetName, 32, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actCaption);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBodyRow3Formula));
		String actFormula = editScreenBodyRow3Formula.getText();
		String expFormula = excelReader.getCellData(xlSheetName, 33, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actFormula);

		System.out.println("Caption        : "+actCaption+"  Value expected  "+expCaption);
		System.out.println("Formula        : "+actFormula+"  Value expected  "+expFormula);

		if(actMessage.equalsIgnoreCase(expMessage) && actCaption.equalsIgnoreCase(expCaption) && actFormula.equalsIgnoreCase(expFormula))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 31, 8, resPass);
			return  true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 31, 8, resFail);
			return  false;
		}
	}







	public boolean checkDeleteFieldInEditScreenBodyInTrialPurchaseUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBodyRow3Delete));
		editScreenBodyRow3Delete.click();

		getWaitForAlert();
		getAlert().accept();

		System.out.println("*******************************************checkDeleteFieldInEditScreenBodyInTrialPurchaseUpdate*******************************************");

		String expMessage = excelReader.getCellData(xlSheetName, 34, 6);
		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 34, 7, actMessage);

		String actRowCount = Integer.toString(editScreenBodyRowsCount.size());/*String.valueOf(editScreenBodyRowsCount.size());*/
		String expRowCount = excelReader.getCellData(xlSheetName, 35, 6); 

		excelReader.setCellData(xlfile, xlSheetName, 35, 7, actRowCount);
		Thread.sleep(2000);

		System.out.println("Row Count After Deleting : "+actRowCount+"  Value Expected  "+expRowCount);

		if(actMessage.equalsIgnoreCase(expMessage) && actRowCount.equalsIgnoreCase(expRowCount))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 34, 8, resPass);
			return  true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 34, 8, resFail);
			return  false;
		}
	}







	public boolean checkAddFieldInEditScreenFooterInTrialPurchaseUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFooterTab));
		editScreenFooterTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenaddBtn));
		editScreenaddBtn.click();

		Thread.sleep(2000);
		//getAlert().accept();
		//Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenCaptionTxt));
		editScreenCaptionTxt.click();
		editScreenCaptionTxt.sendKeys(excelReader.getCellData(xlSheetName, 36, 5));
		editScreenCaptionTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
		editScreenApplyBtn.click();
		Thread.sleep(2000);

		System.out.println("*******************************************checkAddFieldInEditScreenFooterInTrialPurchaseUpdate*******************************************");

		String expMessage = excelReader.getCellData(xlSheetName, 36, 6);
		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 36, 7, actMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFooterTab));
		editScreenFooterTab.click();
		Thread.sleep(1000);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFooterRow2Caption));
		String actCaption = editScreenFooterRow2Caption.getText();
		String expCaption = excelReader.getCellData(xlSheetName, 37, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 37, 7, actCaption);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFooterRow2Formula));
		String actFormula = Boolean.toString(editScreenFooterRow2Formula.getText().isEmpty());
		String expFormula = excelReader.getCellData(xlSheetName, 38, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 38, 7, actFormula);

		System.out.println("Caption        : "+actCaption+"  Value expected  "+expCaption);
		System.out.println("Formula        : "+actFormula+"  Value expected  "+expFormula);

		if(actMessage.equalsIgnoreCase(expMessage) && actCaption.equalsIgnoreCase(expCaption) && actFormula.equalsIgnoreCase(expFormula))
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







	public boolean checkEditFieldInEditScreenFooterInTrialPurchaseUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFooterTab));
		editScreenFooterTab.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFooterRow2Edit));
		editScreenFooterRow2Edit.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenCaptionTxt));
		editScreenCaptionTxt.click();
		editScreenCaptionTxt.clear();
		editScreenCaptionTxt.sendKeys(excelReader.getCellData(xlSheetName, 39, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenAddFieldsBehaviourDropdown));
		editScreenAddFieldsBehaviourDropdown.click();

		Select s=new Select(editScreenAddFieldsBehaviourDropdown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 40, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenAddBehaviourformulaTxt));
		editScreenAddBehaviourformulaTxt.click();
		editScreenAddBehaviourformulaTxt.sendKeys(excelReader.getCellData(xlSheetName, 41, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
		editScreenApplyBtn.click();

		System.out.println("*******************************************checkEditFieldInEditScreenFooterInTrialPurchaseUpdate*******************************************");

		String expMessage = excelReader.getCellData(xlSheetName, 39, 6);
		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFooterTab));
		editScreenFooterTab.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFooterRow2Caption));
		String actCaption = editScreenFooterRow2Caption.getText();
		String expCaption = excelReader.getCellData(xlSheetName, 40, 6);
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, actCaption);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFooterRow2Formula));
		String actFormula = editScreenFooterRow2Formula.getText();
		String expFormula = excelReader.getCellData(xlSheetName, 41, 6);
		excelReader.setCellData(xlfile, xlSheetName, 41, 7, actFormula);

		System.out.println("Caption        : "+actCaption+"  Value expected  "+expCaption);
		System.out.println("Formula        : "+actFormula+"  Value expected  "+expFormula);

		if(actMessage.equalsIgnoreCase(expMessage) && actCaption.equalsIgnoreCase(expCaption) && actFormula.equalsIgnoreCase(expFormula))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 39, 8, resPass);

			return  true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 39, 8, resFail);

			return  false;
		}
	}








	public boolean checkDeleteFieldInEditScreenFooterInTrialPurchaseUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFooterTab));
		editScreenFooterTab.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFooterRow2Delete));
		editScreenFooterRow2Delete.click();

		getWaitForAlert();
		getAlert().accept();

		System.out.println("*******************************************checkDeleteFieldInEditScreenFooterInTrialPurchaseUpdate*******************************************");

		String expMessage = excelReader.getCellData(xlSheetName, 42, 6);
		String actMessage = checkValidationMessage(expMessage);
		 excelReader.setCellData(xlfile, xlSheetName, 42, 7, actMessage);

		Thread.sleep(2000);

		String actRowCount = Integer.toString(editScreenFooterRowsCount.size());
		String expRowCount = excelReader.getCellData(xlSheetName, 43, 6); 
		 excelReader.setCellData(xlfile, xlSheetName, 43, 7, actRowCount);

		System.out.println("Row Count After Deleting : "+actRowCount+"  Value Expected  "+expRowCount);

		if(actMessage.equalsIgnoreCase(expMessage) && actRowCount.equalsIgnoreCase(expRowCount))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 42, 8, resPass);
			return  true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 42, 8, resFail);
			return  false;
		}
	}






@FindBy(xpath="(//*[@id='panelsStayOpen-headingOne']/button)[3]")
private static WebElement inventorSettingsBtn;


	public boolean checkSelectingApplyRateSchemesInInventoryOptionsTabInTrialPurchaseUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryoptionsTab));
		inventoryoptionsTab.click();
		Thread.sleep(2000);
		
		//getAction().moveToElement(inventorSettingsBtn).build().perform();
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventorSettingsBtn));
		inventorSettingsBtn.click();
		*/Thread.sleep(2000);
		
		
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(applyrateSchemesChkBox));
		applyrateSchemesChkBox.click();
		
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTab));
		editLayoutTab.click();	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryoptionsTab));
		inventoryoptionsTab.click();	


		String actapplyrateSchemesChkBox = Boolean.toString(applyrateSchemesChkBoxIsSelected.isSelected());
		String expapplyrateSchemesChkBox = excelReader.getCellData(xlSheetName, 44, 6);
		excelReader.setCellData(xlfile, xlSheetName, 44, 7, actapplyrateSchemesChkBox);

		System.out.println("*******************************************checkSelectingApplyRateSchemesInInventoryOptionsTabInTrialPurchaseUpdate*******************************************");

		System.out.println("ApplyrateSchemesChkBox : "+actapplyrateSchemesChkBox+"  Value Expected  "+expapplyrateSchemesChkBox);

		if(actapplyrateSchemesChkBox.equalsIgnoreCase(expapplyrateSchemesChkBox))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 44, 8, resPass);

			return  true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 44, 8, resFail);

			return  false;
		}
	}







	public boolean checkUnSelectingApplyRateSchemesInInventoryOptionsTabInTrialPurchaseUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(applyrateSchemesChkBox));
		applyrateSchemesChkBox.click();
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTab));
		editLayoutTab.click();	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryoptionsTab));
		inventoryoptionsTab.click();	

		String actapplyrateSchemesChkBox = Boolean.toString(applyrateSchemesChkBox.isSelected());
		String expapplyrateSchemesChkBox = excelReader.getCellData(xlSheetName, 45, 6);
		excelReader.setCellData(xlfile, xlSheetName, 45, 7, actapplyrateSchemesChkBox);

		System.out.println("*******************************************checkUnSelectingApplyRateSchemesInInventoryOptionsTabInTrialPurchaseUpdate*******************************************");

		System.out.println("ApplyrateSchemesChkBox : "+actapplyrateSchemesChkBox+"  Value Expected  "+expapplyrateSchemesChkBox);

		if(actapplyrateSchemesChkBox.equalsIgnoreCase(expapplyrateSchemesChkBox))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 45, 8, resPass);
			return  true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 45, 8, resFail);
			return  false;
		}
	}









	public boolean checkAddFieldInEditLayoutHeaderInTrialPurchaseUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTab));
		editLayoutTab.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutAddFieldsBtn));
		editLayoutAddFieldsBtn.click();
		//getAlert().accept();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCaptionTxt));
		editLayoutCaptionTxt.click();
		editLayoutCaptionTxt.sendKeys(excelReader.getCellData(xlSheetName, 46, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutApplyBtn));
		editLayoutApplyBtn.click();

		System.out.println("*******************************************checkAddFieldInEditLayoutHeaderInTrialPurchaseUpdate*******************************************");

		String expMessage = excelReader.getCellData(xlSheetName, 46, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 46, 7, actMessage);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutHeader2ndRowCaption));
		String actCaption = editLayoutHeader2ndRowCaption.getText();
		String expCaption = excelReader.getCellData(xlSheetName, 47, 6);
		excelReader.setCellData(xlfile, xlSheetName, 47, 7, actCaption);


		System.out.println("Caption        : "+actCaption+"  Value expected  "+expCaption);

		if(actMessage.equalsIgnoreCase(expMessage) && actCaption.equalsIgnoreCase(expCaption))
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







	public boolean checkEditFieldInEditLayoutHeaderInTrialPurchaseUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutHeader2ndRowEditBtn));
		editLayoutHeader2ndRowEditBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCaptionTxt));
		editLayoutCaptionTxt.click();
		editLayoutCaptionTxt.clear();
		editLayoutCaptionTxt.sendKeys(excelReader.getCellData(xlSheetName, 48, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutApplyBtn));
		editLayoutApplyBtn.click();

		Thread.sleep(2000);
		System.out.println("*******************************************checkEditFieldInEditLayoutHeaderInTrialPurchaseUpdate*******************************************");

		String expMessage = excelReader.getCellData(xlSheetName, 48, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 48, 7, actMessage);

		for (int i = 0; i < 4; i++) 
		{
			if (getIsAlertPresent()==true) 
			{
				getAlert().accept();
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutHeader2ndRowCaption));
		String actCaption = editLayoutHeader2ndRowCaption.getText();
		String expCaption = excelReader.getCellData(xlSheetName, 49, 6);
		excelReader.setCellData(xlfile, xlSheetName, 49, 7, actCaption);

		System.out.println("Caption        : "+actCaption+"  Value expected  "+expCaption);

		if(actMessage.equalsIgnoreCase(expMessage) && actCaption.equalsIgnoreCase(expCaption))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 48, 8, resPass);
			return  true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 48, 8, resFail);
			return  false;
		}
	}








	public boolean checkDeleteFieldInEditLayoutHeaderInTrialPurchaseUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutHeader2ndRowDeleteBtn));
		editLayoutHeader2ndRowDeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		System.out.println("*******************************************checkDeleteFieldInEditLayoutHeaderInTrialPurchaseUpdate*******************************************");

		String expMessage = excelReader.getCellData(xlSheetName, 50, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 50, 7, actMessage);

		Thread.sleep(2000);

		String actRowCount = Integer.toString(editLayoutHeaderRowsCount.size());
		String expRowCount = excelReader.getCellData(xlSheetName, 51, 6); 
		excelReader.setCellData(xlfile, xlSheetName, 51, 7, actRowCount);

		System.out.println("Row Count After Deleting : "+actRowCount+"  Value Expected  "+expRowCount);

		if(actMessage.equalsIgnoreCase(expMessage) && actRowCount.equalsIgnoreCase(expRowCount))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 50, 8, resPass);
			return  true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 50, 8, resFail);
			return  false;
		}
	}








	public boolean checkAddFieldInEditLayoutBodyInTrialPurchaseUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBodyTab));
		editLayoutBodyTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutAddFieldsBtn));
		editLayoutAddFieldsBtn.click();

		Thread.sleep(2000);
		
		//getAlert().accept();
		//Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCaptionTxt));
		editLayoutCaptionTxt.click();
		editLayoutCaptionTxt.sendKeys(excelReader.getCellData(xlSheetName, 52, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutApplyBtn));
		editLayoutApplyBtn.click();

		System.out.println("*******************************************checkAddFieldInEditLayoutBodyInTrialPurchaseUpdate*******************************************");

		String expMessage = excelReader.getCellData(xlSheetName, 52, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 52, 7, actMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBodyTab));
		editLayoutBodyTab.click();

		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutbody1stRowCaption));
		String actCaption = editLayoutbody1stRowCaption.getText();
		String expCaption = excelReader.getCellData(xlSheetName, 53, 6);
		excelReader.setCellData(xlfile, xlSheetName, 53, 7, actCaption);

		System.out.println("Caption        : "+actCaption+"  Value expected  "+expCaption);

		if(actMessage.equalsIgnoreCase(expMessage) && actCaption.equalsIgnoreCase(expCaption))
		{
			excelReader.setCellData(xlfile, xlSheetName, 52, 8, resPass);
			
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 52, 8, resFail);
			return false;
		}
	}







	public boolean checkEditFieldInEditLayoutBodyInTrialPurchaseUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBodyTab));
		editLayoutBodyTab.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBody1stRowEditBtn));
		editLayoutBody1stRowEditBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCaptionTxt));
		editLayoutCaptionTxt.click();
		editLayoutCaptionTxt.clear();
		editLayoutCaptionTxt.sendKeys(excelReader.getCellData(xlSheetName, 54, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutApplyBtn));
		editLayoutApplyBtn.click();

		System.out.println("*******************************************checkEditFieldInEditLayoutBodyInTrialPurchaseUpdate*******************************************");

		String expMessage = excelReader.getCellData(xlSheetName, 54, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 54, 7, actMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBodyTab));
		editLayoutBodyTab.click();

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutbody1stRowCaption));
		String actCaption = editLayoutbody1stRowCaption.getText();
		String expCaption = excelReader.getCellData(xlSheetName, 55, 6);
		excelReader.setCellData(xlfile, xlSheetName, 55, 7, actCaption);

		System.out.println("Caption        : "+actCaption+"  Value expected  "+expCaption);

		if(actMessage.equalsIgnoreCase(expMessage) && actCaption.equalsIgnoreCase(expCaption))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 54, 8, resPass);

			return  true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 54, 8, resFail);

			return  false;
		}
	}







	public boolean checkDeleteFieldInEditLayoutBodyInTrialPurchaseUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBodyTab));
		editLayoutBodyTab.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBody1stRowDeleteBtn));
		editLayoutBody1stRowDeleteBtn.click();

		Thread.sleep(2000);

		getWaitForAlert();
		getAlert().accept();

		System.out.println("*******************************************checkDeleteFieldInEditLayoutBodyInTrialPurchaseUpdate*******************************************");

		String expMessage = excelReader.getCellData(xlSheetName, 56, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 56, 7, actMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 56, 8, resPass);
			return  true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 56, 8, resFail);
			return  false;
		}
	}



@FindBy(xpath="(//*[@id='panelsStayOpen-headingOne'])[2]")
private static WebElement viewSettingsBtn;

@FindBy(xpath="//*[@id='panelsStayOpen-headingTwo']/button")
private static WebElement viewTabLoginBtn;




	public boolean checkSavingViewInTrialVoucherUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		
		
int count = SettingsSubMenusList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			//ScrollIntoView(SettingsSubMenusList.get(i));
			System.out.println(excelReader.getCellData(xlSheetName, 8, 5));
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 8, 5))) 
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
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustPurchaseVouchersBtn));
		docCustPurchaseVouchersBtn.click();

		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustCreateVoucherBtn));
		docCustCreateVoucherBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CreatedVoucher));
		CreatedVoucher.click();

		Thread.sleep(7000);
		
		*/
		
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
		searchTxt.click();
		searchTxt.sendKeys("Configure Transactions");
		searchTxt.sendKeys(Keys.ENTER);
		Thread.sleep(3000);
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionPreferencesMinus));
		configureTransactionPreferencesMinus.click();
		
		Thread.sleep(2000);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustPurchaseVouchersBtn));
		docCustPurchaseVouchersBtn.click();

		Thread.sleep(5000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CreatedVoucher));
		CreatedVoucher.click();

		
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", viewsBtn);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsBtn));
		viewsBtn.click();
		Thread.sleep(2000);
		
		

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewExistingViewTxt));
		ViewExistingViewTxt.click();
		//ViewExistingViewTxt.clear();
		ViewExistingViewTxt.sendKeys(excelReader.getCellData(xlSheetName, 57, 5));
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys(Keys.TAB);

		Thread.sleep(5000);
		
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,350)", "");
		
		Thread.sleep(3000);
	
		if(viewsUserAllOptionsSTChkboxSelected.isDisplayed()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
		
		viewTabLoginBtn.click();;
		}
		Thread.sleep(2000);
	
		if(viewsUserAllOptionsSTChkboxSelected.isSelected()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserAllOptionsSTChkbox));
		viewsUserAllOptionsSTChkbox.click();
		}
		Thread.sleep(2000);

	
		
	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSaveView));
		viewSaveView.click();
		Thread.sleep(3000);

		System.out.println("*********************************checkSavingViewInTrialVoucherUpdate*********************************");

		String expMessage = excelReader.getCellData(xlSheetName, 57, 6);

		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 57, 7, actMessage);
			
		Thread.sleep(2000);
		
		/*((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", viewTabLoginBtn);
		Thread.sleep(2000);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
		viewTabLoginBtn.click();
		
		Thread.sleep(3000);
		
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", viewsUserAllOptionsSTChkbox);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserAllOptionsSTChkbox));
		
		String actviewsUserAllOptionsSTChkbox = Boolean.toString(viewsUserAllOptionsSTChkboxSelected.isSelected());
		String expviewsUserAllOptionsSTChkbox = excelReader.getCellData(xlSheetName, 58, 6);
		excelReader.setCellData(xlfile, xlSheetName, 58, 7, actviewsUserAllOptionsSTChkbox);

		System.out.println("viewsUserAllOptionsSTChkbox  : "+actviewsUserAllOptionsSTChkbox+"  Value Expected  "+expviewsUserAllOptionsSTChkbox);

		if (actMessage.equalsIgnoreCase(expMessage) && actviewsUserAllOptionsSTChkbox.equalsIgnoreCase(expviewsUserAllOptionsSTChkbox)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 57, 8, resPass);

			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 57, 8, resFail);

			return false;
		}
	}









	public boolean checkEditingViewInTrialVoucherUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
	
		Thread.sleep(2000);
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsBtn));
		viewsBtn.click();*/
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewExistingViewTxt));
		ViewExistingViewTxt.click();
		ViewExistingViewTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		ViewExistingViewTxt.sendKeys(Keys.SPACE);
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		Thread.sleep(1000);
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsAuthorizeOnlyChkbox));
		
		
		if(viewsAuthorizeOnlyChkboxIsSelected.isSelected()==false)
		{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsAuthorizeOnlyChkbox));
		viewsAuthorizeOnlyChkbox.click();
		}
		
		Thread.sleep(1000);
	
		
		/*((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", viewTabLoginBtn);
		Thread.sleep(3000);*/

	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
		viewTabLoginBtn.click();
		Thread.sleep(4000);*/
		
		//((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", viewsUserAllOptionsSTChkbox);
		
		if(viewsUserAllOptionsSTChkboxSelected.isSelected()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserAllOptionsSTChkbox));
		viewsUserAllOptionsSTChkbox.click();
		}

		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserRestrictChkbox));
		
		if(viewsUserRestrictChkboxSelected.isSelected()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserRestrictChkbox));
		viewsUserRestrictChkbox.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSaveView));
		viewSaveView.click();
		Thread.sleep(2000);
		
		

		System.out.println("*********************************checkEditingViewInTrialVoucherUpdate*********************************");

		String expMessage = excelReader.getCellData(xlSheetName, 59, 6);

		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 59, 7, actMessage);
		
	
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsAuthorizeOnlyChkbox));
		String actviewsAuthorizeOnlyChkbox = Boolean.toString(viewsAuthorizeOnlyChkboxIsSelected.isSelected());
		String expviewsAuthorizeOnlyChkbox = excelReader.getCellData(xlSheetName, 60, 6);
		excelReader.setCellData(xlfile, xlSheetName, 60, 7, actviewsAuthorizeOnlyChkbox);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
		viewTabLoginBtn.click();
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserAllOptionsSTChkbox));
		String actviewsUserAllOptionsSTChkbox = Boolean.toString(viewsUserAllOptionsSTChkboxSelected.isSelected());
		String expviewsUserAllOptionsSTChkbox = excelReader.getCellData(xlSheetName, 61, 6);
		excelReader.setCellData(xlfile, xlSheetName, 61, 7, actviewsUserAllOptionsSTChkbox);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserRestrictChkbox));
		String actviewSuChkbox = Boolean.toString(viewsUserRestrictChkboxSelected.isSelected());
		String expviewSuChkbox = excelReader.getCellData(xlSheetName, 62, 6);
		excelReader.setCellData(xlfile, xlSheetName, 62, 7, actviewSuChkbox);


		System.out.println("viewsAuthorizeOnlyChkbox   : "+actviewsAuthorizeOnlyChkbox +"  Value Expected  "+expviewsAuthorizeOnlyChkbox);
		System.out.println("viewsUserAllOptionsSTChkbox : "+actviewsUserAllOptionsSTChkbox +"  Value Expected  "+expviewsUserAllOptionsSTChkbox);
		System.out.println("viewSuChkbox : "+actviewSuChkbox +"  Value Expected  "+expviewSuChkbox);

		if (actMessage.equalsIgnoreCase(expMessage) && actviewsAuthorizeOnlyChkbox.equalsIgnoreCase(expviewsAuthorizeOnlyChkbox) && actviewsUserAllOptionsSTChkbox.equalsIgnoreCase(expviewsUserAllOptionsSTChkbox)
				&& actviewSuChkbox.equalsIgnoreCase(expviewSuChkbox)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 59, 8, resPass);

			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 59, 8, resFail);

			return false;
		}
	}






	public boolean checkDeletingViewInTrialVoucherUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewExistingViewTxt));
		ViewExistingViewTxt.click();
		/*//ViewExistingViewTxt.clear();
		ViewExistingViewTxt.sendKeys(excelReader.getCellData(xlSheetName, 57, 5));
		Thread.sleep(2000);*/
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsDeleteViewIcon));
		viewsDeleteViewIcon.click();

		getWaitForAlert();

		getAlert().accept();
		Thread.sleep(2000);

		System.out.println("*********************************checkDeletingViewInTrialVoucherUpdate*********************************");

		String expMessage = excelReader.getCellData(xlSheetName, 63, 6);

		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 63, 7, actMessage);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSettingsBtn));
		viewSettingsBtn.click();*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsAuthorizeOnlyChkbox));
		String actviewsAuthorizeOnlyChkbox = Boolean.toString(viewsAuthorizeOnlyChkboxIsSelected.isSelected());
		String expviewsAuthorizeOnlyChkbox = excelReader.getCellData(xlSheetName, 64, 6);
		excelReader.setCellData(xlfile, xlSheetName, 64, 7, actviewsAuthorizeOnlyChkbox);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
		viewTabLoginBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserAllOptionsSTChkbox));
		String actviewsUserAllOptionsSTChkbox = Boolean.toString(viewsUserAllOptionsSTChkboxSelected.isSelected());
		String expviewsUserAllOptionsSTChkbox = excelReader.getCellData(xlSheetName, 65, 6);
		excelReader.setCellData(xlfile, xlSheetName, 65, 7, actviewsUserAllOptionsSTChkbox);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSuChkbox));
		String actviewSuChkbox = Boolean.toString(viewSuChkboxIsSelected.isSelected());
		String expviewSuChkbox = excelReader.getCellData(xlSheetName, 66, 6);
		excelReader.setCellData(xlfile, xlSheetName, 66, 7, actviewSuChkbox);

		System.out.println("viewsAuthorizeOnlyChkbox   : "+actviewsAuthorizeOnlyChkbox +"  Value Expected  "+expviewsAuthorizeOnlyChkbox);
		System.out.println("viewsUserAllOptionsSTChkbox : "+actviewsUserAllOptionsSTChkbox +"  Value Expected  "+expviewsUserAllOptionsSTChkbox);
		System.out.println("viewSuChkbox : "+actviewSuChkbox +"  Value Expected  "+expviewSuChkbox);

		if (actMessage.equalsIgnoreCase(expMessage) && actviewsAuthorizeOnlyChkbox.equalsIgnoreCase(expviewsAuthorizeOnlyChkbox) && actviewsUserAllOptionsSTChkbox.equalsIgnoreCase(expviewsUserAllOptionsSTChkbox)
				&& actviewSuChkbox.equalsIgnoreCase(expviewSuChkbox)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 63, 8, resPass);

			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 63, 8, resFail);

			return false;
		}
	}


@FindBy(xpath="(//*[@id='panelsStayOpen-headingTwo']/button)[2]")
private static WebElement triggersExpandBtn;

@FindBy(xpath="//*[@id='17_1_AdvanceFilter_']/table/tbody/tr/td[1]/select")
private static WebElement triggersConditionBtn;

@FindBy(xpath="//*[@id='17_1_AdvanceFilter_']/table/tbody/tr/td[2]/input")
private static WebElement triggersNameField;

@FindBy(xpath="//*[@id='17_1_AdvanceFilter_']/table/tbody/tr/td[3]/select")
private static WebElement triggersConjunctionBtn;


@FindBy(xpath="//*[@id='17_1_AdvanceFilter_']/table/tbody/tr/td[4]/select")
private static WebElement triggersValueBtn;


@FindBy(xpath="//*[@id='advancefilter_master_17_1_']")
private static WebElement triggersValueField;



	public boolean checkSavingTriggerInTrialVoucherUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		
		Thread.sleep(2000);

		int count = SettingsSubMenusList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			//ScrollIntoView(SettingsSubMenusList.get(i));
			System.out.println(excelReader.getCellData(xlSheetName, 8, 5));
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 8, 5))) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}
		
		Thread.sleep(2000);
*/	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
		searchTxt.click();
		searchTxt.sendKeys("Configure Transactions");
		searchTxt.sendKeys(Keys.ENTER);
		Thread.sleep(3000);
		
		
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionPreferencesMinus));
		configureTransactionPreferencesMinus.click();
		
		Thread.sleep(2000);*/
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionDocumentCustomiztionMinus));
		configureTransactionDocumentCustomiztionMinus.click();
		*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustPurchaseVouchersBtn));
		docCustPurchaseVouchersBtn.click();

		Thread.sleep(8000);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustCreateVoucherBtn));
		docCustCreateVoucherBtn.click();

		Thread.sleep(8000);
		*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdVoucher));
		createdVoucher.click();
		Thread.sleep(2000);
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		triggersBtn.click();

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersExpandBtn));
		triggersExpandBtn.click();
		*/Thread.sleep(1000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.clear();
		editLayoutTriggerName.sendKeys(excelReader.getCellData(xlSheetName, 67, 5));
		Thread.sleep(2000);
		editLayoutTriggerName.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		
		//condition for Trigger
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersConditionBtn));
		triggersConditionBtn.click();
		Select s1=new Select(triggersConditionBtn);
		s1.selectByValue("0");
		triggersConditionBtn.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersNameField));
		triggersNameField.click();
		triggersNameField.sendKeys("Quantity");
		triggersNameField.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersConjunctionBtn));
		triggersConjunctionBtn.click();
	Select s2=new Select(triggersConjunctionBtn);
	s2.selectByValue("0");
	triggersConjunctionBtn.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersValueBtn));
	triggersValueBtn.click();
	Select s3=new Select(triggersValueBtn);
	s3.selectByValue("0");
	triggersValueBtn.sendKeys(Keys.TAB);
		
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersValueField));
	triggersValueField.click();
	triggersValueField.sendKeys("10");
	triggersValueField.sendKeys(Keys.TAB);*/
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersConjuctionDrpdwn));
		triggersConjuctionDrpdwn.click();
		
		triggersConjuctionDrpdwn.sendKeys(Keys.ARROW_DOWN);
		
		triggersConjuctionDrpdwn.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
		triggersSelectFieldTxt.click();
		
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerVendorACExpansionBtn));
		triggerVendorACExpansionBtn.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerVendorAcName));
		triggerVendorAcName.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectOperdrpdwn));
		triggersSelectOperdrpdwn.click();
		triggersSelectOperdrpdwn.sendKeys(Keys.ARROW_DOWN);
		
		
		Thread.sleep(2000);
		triggersSelectOperdrpdwn.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersCompareWithdrpdwn));
		triggersCompareWithdrpdwn.click();
		triggersCompareWithdrpdwn.sendKeys(Keys.ARROW_DOWN);
		Thread.sleep(2000);
		triggersCompareWithdrpdwn.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		triggersValueToEnterTxt.sendKeys("Vendor A");	
	
		
		Thread.sleep(2000);
		
		triggersValueToEnterTxt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		editLayoutTriggersSaveBtn.click();
		//editLayoutTriggersSaveBtn.click();

		System.out.println("*************************************checkSavingTriggerInTrialVoucherUpdate*****************************************");

		String expSaveMessage = excelReader.getCellData(xlSheetName, 67, 6);

		String actSaveMessage = checkValidationMessage(expSaveMessage);
		excelReader.setCellData(xlfile, xlSheetName, 67, 7, actSaveMessage);

		Thread.sleep(2000);

		if(actSaveMessage.equalsIgnoreCase(expSaveMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 67, 8, resPass);

			return true;
		}
		else
		{excelReader.setCellData(xlfile, xlSheetName, 67, 8, resFail);

			return false;
		}
	}


@FindBy(xpath="(//*[@id='panelsStayOpen-headingTwo']/button)[3]")
private static WebElement actionsBtn;




	public boolean checkEditingTriggerInTrialVoucherUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.clear();
		editLayoutTriggerName.sendKeys(excelReader.getCellData(xlSheetName, 68, 5));
		Thread.sleep(2000);
		editLayoutTriggerName.sendKeys(Keys.TAB);

		Thread.sleep(2000);

	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(actionsBtn));
		actionsBtn.click();
		*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSuspended));
		if (triggersSuspendedChkboxIsSelected.isSelected()==false) 
		{
			triggersSuspended.click();
		}

		
		//validation
		 
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersConjuctionDrpdwn));
		Select ConjuctionSelect=new Select(triggersConjuctionDrpdwn);

		String  acttriggersTabSelectTabDrpdwn=ConjuctionSelect.getFirstSelectedOption().getText();
		String exptriggersTabSelectTabDrpdwn="";

		System.out.println("TriggersSelectTabDrpdwn : "+acttriggersTabSelectTabDrpdwn +" Value : "+exptriggersTabSelectTabDrpdwn);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		editLayoutTriggersSaveBtn.click();

		System.out.println("*************************************checkEditingTriggerInTrialVoucherUpdate*****************************************");

		String expSaveMessage = excelReader.getCellData(xlSheetName, 68, 6);

		String actSaveMessage = checkValidationMessage(expSaveMessage);
		excelReader.setCellData(xlfile, xlSheetName, 68, 7, actSaveMessage);

		if(actSaveMessage.equalsIgnoreCase(expSaveMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 68, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 68, 8, resFail);
			return false;
		}
	}









	public boolean checkDeletingTriggerInTrialVoucherUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.clear();
		editLayoutTriggerName.sendKeys(excelReader.getCellData(xlSheetName, 69, 5));
		Thread.sleep(3000);
		editLayoutTriggerName.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(actionsBtn));
		actionsBtn.click();
		Thread.sleep(2000);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSuspendedChkboxs));
		String acttriggersSuspendedChkboxs = Boolean.toString(triggersSuspendedChkboxIsSelected.isSelected());
		String exptriggersSuspendedChkboxs = excelReader.getCellData(xlSheetName, 69, 6);
		excelReader.setCellData(xlfile, xlSheetName, 69, 7, acttriggersSuspendedChkboxs);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerDeleteBtn));
		editLayoutTriggerDeleteBtn.click();

		if (getIsAlertPresent())
		{
			getWaitForAlert();

			getAlert().accept();
		}
		else
		{
			editLayoutTriggerDeleteBtn.click();

			getWaitForAlert();

			getAlert().accept();
		}

		System.out.println("*************************************checkDeletingTriggerInTrialVoucherUpdate*****************************************");

		System.out.println("triggersSuspendedChkboxs  : "+acttriggersSuspendedChkboxs+"  Value Expected  "+exptriggersSuspendedChkboxs);

		String expDeleteMessage = excelReader.getCellData(xlSheetName, 70, 6);

		String actDeleteMessage = checkValidationMessage(expDeleteMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 70, 7, actDeleteMessage);

		if(acttriggersSuspendedChkboxs.equalsIgnoreCase(exptriggersSuspendedChkboxs) && actDeleteMessage.equalsIgnoreCase(expDeleteMessage))
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







	public boolean checkSavingDocumentNumberingInTrialVoucherUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{	
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		
		Thread.sleep(2000);

		int count = SettingsSubMenusList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			//ScrollIntoView(SettingsSubMenusList.get(i));
			System.out.println(excelReader.getCellData(xlSheetName, 8, 5));
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 8, 5))) 
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
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustPurchaseVouchersBtn));
		docCustPurchaseVouchersBtn.click();

		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustCreateVoucherBtn));
		docCustCreateVoucherBtn.click();

		Thread.sleep(3000);
*/		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberingBtn));
		documentNumberingBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow1));
		typeRow1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));

		Select typeDropDownSelect= new Select(typeDropDown);

		typeDropDownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName, 71, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow1));
		fieldValueRow1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueDropDown));
		Select fieldValueDropDownSelect= new Select(fieldValueDropDown);

		fieldValueDropDownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName, 72, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingCharacterRow1));
		startingCharacterRow1.click();

		startingCharacterTxt.sendKeys(Keys.BACK_SPACE);
		startingCharacterTxt.sendKeys(excelReader.getCellData(xlSheetName, 73, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow1));
		noofCharactersRow1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noOfCharacterTxt));
		noOfCharacterTxt.sendKeys(excelReader.getCellData(xlSheetName, 74, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow2));
		typeRow2.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));

		typeDropDownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName, 75, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow2));
		fieldValueRow2.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueTxt));
		fieldValueTxt.sendKeys(excelReader.getCellData(xlSheetName, 76, 5));
		fieldValueTxt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingCharacterRow2));
		startingCharacterRow2.click();
		
		startingCharacterTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow2));
		noofCharactersRow2.click();

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noOfCharacterTxt));
		noOfCharacterTxt.click();
		noOfCharacterTxt.sendKeys("2");
		noOfCharacterTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		System.out.println("********************************checkSavingDocumentNumberingInTrialVoucherUpdate*********************************");

		String expSaveMessage = excelReader.getCellData(xlSheetName, 71, 6);

		String actSaveMessage = checkValidationMessage(expSaveMessage);
		excelReader.setCellData(xlfile, xlSheetName, 71, 7, actSaveMessage);

		String acttypeRow1 = typeRow1.getText();
		String actfieldValueRow1 = fieldValueRow1.getText();
		String actstartingCharacterRow1 = startingCharacterRow1.getText();
		String actnoofCharactersRow1 = noofCharactersRow1.getText();
		String acttypeRow2 = typeRow2.getText();
		String actfieldValueRow2 = fieldValueRow2.getText();
		String actstartingCharacterRow2 = startingCharacterRow2.getText();
		String actnoofCharactersRow2 = Boolean.toString(noofCharactersRow2.getText().isEmpty());
		System.out.println("**********************************"+noofCharactersRow2.getText());

		String exptypeRow1 = excelReader.getCellData(xlSheetName, 72, 6);
		excelReader.setCellData(xlfile, xlSheetName, 72, 7, acttypeRow1);
		
		String expfieldValueRow1 = excelReader.getCellData(xlSheetName, 73, 6);
		excelReader.setCellData(xlfile, xlSheetName, 73, 7, actfieldValueRow1);
		
		String expstartingCharacterRow1 = excelReader.getCellData(xlSheetName, 74, 6);
		excelReader.setCellData(xlfile, xlSheetName, 74, 7, actstartingCharacterRow1);
		
		String expnoofCharactersRow1 = excelReader.getCellData(xlSheetName, 75, 6);
		excelReader.setCellData(xlfile, xlSheetName, 75, 7, actnoofCharactersRow1);
		
		String exptypeRow2 = excelReader.getCellData(xlSheetName, 76, 6);
		excelReader.setCellData(xlfile, xlSheetName, 76, 7, acttypeRow2);
		
		String expfieldValueRow2 = excelReader.getCellData(xlSheetName, 77, 6);
		excelReader.setCellData(xlfile, xlSheetName, 77, 7, actfieldValueRow2);
		
		String expstartingCharacterRow2 = excelReader.getCellData(xlSheetName, 78, 6);
		excelReader.setCellData(xlfile, xlSheetName, 78, 7, actstartingCharacterRow2);
		
		String expnoofCharactersRow2 = excelReader.getCellData(xlSheetName, 79, 6);
		excelReader.setCellData(xlfile, xlSheetName, 79, 7, actnoofCharactersRow2);

		System.out.println("typeRow1    : "+acttypeRow1+"  Value Expected  "+exptypeRow1);
		System.out.println("fieldValueRow1    : "+actfieldValueRow1+"  Value Expected  "+expfieldValueRow1);
		System.out.println("startingCharacterRow1    : "+actstartingCharacterRow1+"  Value Expected  "+expstartingCharacterRow1);
		System.out.println("noofCharactersRow1    : "+actnoofCharactersRow1+"  Value Expected  "+expnoofCharactersRow1);
		System.out.println("typeRow2    : "+acttypeRow2+"  Value Expected  "+exptypeRow2);
		System.out.println("fieldValueRow2    : "+actfieldValueRow2+"  Value Expected  "+expfieldValueRow2);
		System.out.println("startingCharacterRow2    : "+actstartingCharacterRow2+"  Value Expected  "+expstartingCharacterRow2);
		System.out.println("noofCharactersRow2    : "+actnoofCharactersRow2+"  Value Expected  "+expnoofCharactersRow2);

		if (actSaveMessage.equalsIgnoreCase(expSaveMessage)
				&& acttypeRow1.equalsIgnoreCase(exptypeRow1) && actfieldValueRow1.equalsIgnoreCase(expfieldValueRow1) 
				&& actstartingCharacterRow1.equalsIgnoreCase(expstartingCharacterRow1) && actnoofCharactersRow1.equalsIgnoreCase(expnoofCharactersRow1)
				&& acttypeRow2.equalsIgnoreCase(exptypeRow2) && actfieldValueRow2.equalsIgnoreCase(expfieldValueRow2) 
				&& actstartingCharacterRow2.equalsIgnoreCase(expstartingCharacterRow2) && actnoofCharactersRow2.equalsIgnoreCase(expnoofCharactersRow2)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 71, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 71, 8, resFail);
			return false;
		}
	}		








	public boolean checkEditingDocumentNumberingInTrialVoucherUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{	
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docNumberingRow1Col1));
		docNumberingRow1Col1.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow2));
		noofCharactersRow2.click();

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noOfCharacterTxt));
		noOfCharacterTxt.click();
		noOfCharacterTxt.sendKeys("2");
		noOfCharacterTxt.sendKeys(Keys.TAB);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docNumberingRow1DeleteBtn));
		docNumberingRow1DeleteBtn.click();
*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		System.out.println("********************************checkEditingDocumentNumberingInTrialVoucherUpdate*********************************");

		String expSaveMessage = excelReader.getCellData(xlSheetName, 80, 6);

		String actSaveMessage = checkValidationMessage(expSaveMessage);
		excelReader.setCellData(xlfile, xlSheetName, 80, 7, actSaveMessage);

		String acttypeRow1 = typeRow2.getText();
		String actfieldValueRow1 = fieldValueRow2.getText();
		String actstartingCharacterRow1 = startingCharacterRow2.getText();
		String actnoofCharactersRow1 = Boolean.toString(noofCharactersRow2.getText().isEmpty());

		String exptypeRow1 = excelReader.getCellData(xlSheetName, 81, 6);
		excelReader.setCellData(xlfile, xlSheetName, 81, 7, acttypeRow1);
		
		String expfieldValueRow1 = excelReader.getCellData(xlSheetName, 82, 6);
		excelReader.setCellData(xlfile, xlSheetName, 82, 7, actfieldValueRow1);
		
		String expstartingCharacterRow1 = excelReader.getCellData(xlSheetName, 83, 6);
		excelReader.setCellData(xlfile, xlSheetName, 83, 7, actstartingCharacterRow1);
		
		String expnoofCharactersRow1 = excelReader.getCellData(xlSheetName, 84, 6);
		excelReader.setCellData(xlfile, xlSheetName, 84, 7, actnoofCharactersRow1);

		System.out.println("typeRow1    : "+acttypeRow1+"  Value Expected  "+exptypeRow1);
		System.out.println("fieldValueRow1    : "+actfieldValueRow1+"  Value Expected  "+expfieldValueRow1);
		System.out.println("startingCharacterRow1    : "+actstartingCharacterRow1+"  Value Expected  "+expstartingCharacterRow1);
		System.out.println("noofCharactersRow1    : "+actnoofCharactersRow1+"  Value Expected  "+expnoofCharactersRow1);

		if (actSaveMessage.equalsIgnoreCase(expSaveMessage)
				&& acttypeRow1.equalsIgnoreCase(exptypeRow1) && actfieldValueRow1.equalsIgnoreCase(expfieldValueRow1) 
				&& actstartingCharacterRow1.equalsIgnoreCase(expstartingCharacterRow1) && actnoofCharactersRow1.equalsIgnoreCase(expnoofCharactersRow1))
		{
			excelReader.setCellData(xlfile, xlSheetName, 80, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 80, 8, resFail);
			return false;
		}
	}		





	public boolean checkDeletingDocumentNumberingInTrialVoucherUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{	
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docNumberingRow2Col1));
		docNumberingRow2Col1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docNumberingDeleteRowBtn));
		docNumberingDeleteRowBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		System.out.println("********************************checkEditingDocumentNumberingInTrialVoucherUpdate*********************************");

		String expSaveMessage = excelReader.getCellData(xlSheetName, 85, 6);

		String actSaveMessage = checkValidationMessage(expSaveMessage);
		excelReader.setCellData(xlfile, xlSheetName, 85, 7, actSaveMessage);

		String acttypeRow1 = typeRow1.getText();
		String actfieldValueRow1 = fieldValueRow1.getText();
		String actstartingCharacterRow1 = startingCharacterRow1.getText();
		String actnoofCharactersRow1 = noofCharactersRow1.getText();

		/*String exptypeRow1 = excelReader.getCellData(xlSheetName, 86, 6);
		excelReader.setCellData(xlfile, xlSheetName, 86, 7, acttypeRow1);
*/		
		String exptypeRow1="Login";
		
		String expfieldValueRow1 = excelReader.getCellData(xlSheetName, 87, 6);
		excelReader.setCellData(xlfile, xlSheetName, 87, 7, actfieldValueRow1);
		
		String expstartingCharacterRow1 = excelReader.getCellData(xlSheetName, 88, 6);
		excelReader.setCellData(xlfile, xlSheetName, 88, 7, actstartingCharacterRow1);
		
		String expnoofCharactersRow1 = excelReader.getCellData(xlSheetName, 89, 6);
		excelReader.setCellData(xlfile, xlSheetName, 89, 7, actnoofCharactersRow1);

		System.out.println("typeRow1    : "+acttypeRow1+"  Value Expected  "+exptypeRow1);
		System.out.println("fieldValueRow1    : "+actfieldValueRow1+"  Value Expected  "+expfieldValueRow1);
		System.out.println("startingCharacterRow1    : "+actstartingCharacterRow1+"  Value Expected  "+expstartingCharacterRow1);
		System.out.println("noofCharactersRow1    : "+actnoofCharactersRow1+"  Value Expected  "+expnoofCharactersRow1);

		if (actSaveMessage.equalsIgnoreCase(expSaveMessage)
				&& acttypeRow1.equalsIgnoreCase(exptypeRow1) && actfieldValueRow1.equalsIgnoreCase(expfieldValueRow1) && actstartingCharacterRow1.equalsIgnoreCase(expstartingCharacterRow1) 
				&& actnoofCharactersRow1.equalsIgnoreCase(expnoofCharactersRow1)) 
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











	public boolean checkSavingRuleInTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		
		Thread.sleep(2000);

		int count = SettingsSubMenusList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			//ScrollIntoView(SettingsSubMenusList.get(i));
			System.out.println(excelReader.getCellData(xlSheetName, 8, 5));
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 8, 5))) 
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
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustPurchaseVouchersBtn));
		docCustPurchaseVouchersBtn.click();

		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustCreateVoucherBtn));
		docCustCreateVoucherBtn.click();
*/
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesBtn));
		rulesBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabAddRuleBtn));
		rulesTabAddRuleBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabRuleNameTxt));
		rulesTabRuleNameTxt.click();
		rulesTabRuleNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 90, 5));
		Thread.sleep(2000);
		rulesTabRuleNameTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		getAction().moveToElement(rulesNewRecordChekbox).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesNewRecordChekbox));
		rulesNewRecordChekbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesLoadChekbox));
		rulesLoadChekbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabNoConditionChkbbox));
		rulesTabNoConditionChkbbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessageTab));
		rulesIFMessageTab.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessagesGenralMessageTxt));
		rulesIFMessagesGenralMessageTxt.sendKeys(excelReader.getCellData(xlSheetName, 91, 5));

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTab_SaveRuleBtn));
		rulesTab_SaveRuleBtn.click();

		System.out.println("******************************checkSavingRuleInTrialVoucherUpdate*********************************");

		String expMessage = excelReader.getCellData(xlSheetName, 90, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 90, 7, actMessage);

		if (actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 90, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 90, 8, resFail);
			return false;
		}
	}









	public boolean checkEditingRuleInTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabRuleNameTxt));
		rulesTabRuleNameTxt.click();
		rulesTabRuleNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 92, 5));
		Thread.sleep(3000);
		rulesTabRuleNameTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesNewRecordChekbox));
		rulesNewRecordChekbox.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesEditChekbox));
		rulesEditChekbox.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesLoadChekbox));
		rulesLoadChekbox.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesOnEnterChekbox));
		rulesOnEnterChekbox.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTab_SaveRuleBtn));
		rulesTab_SaveRuleBtn.click();

		System.out.println("******************************checkEditingRuleInTrialVoucherUpdate*********************************");

		String expMessage = excelReader.getCellData(xlSheetName, 92, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 92, 7, actMessage);

		if (actMessage.equalsIgnoreCase(expMessage))
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








	public boolean checkDeletingRuleInTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabRuleNameTxt));
		rulesTabRuleNameTxt.click();
		rulesTabRuleNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 93, 5));
		Thread.sleep(3000);
		rulesTabRuleNameTxt.sendKeys(Keys.TAB);

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTab_DeleteRuleBtn));
		rulesTab_DeleteRuleBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteRuleYesBtn));
		deleteRuleYesBtn.click();

		System.out.println("******************************checkDeletingRuleInTrialVoucherUpdate*********************************");

		String expMessage = excelReader.getCellData(xlSheetName, 93, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 93, 7, actMessage);


		if (actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 93, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 93, 8, resFail);
			return false;
		}
	}







	public boolean checkDeletingTrialVoucherUpdate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
		searchTxt.click();
		searchTxt.sendKeys("Configure Transactions");
		searchTxt.sendKeys(Keys.ENTER);
		Thread.sleep(30000);
		
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionPreferencesMinus));
		configureTransactionPreferencesMinus.click();
		
		Thread.sleep(2000);*/
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionDocumentCustomiztionMinus));
		configureTransactionDocumentCustomiztionMinus.click();*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustPurchaseVouchersBtn));
		docCustPurchaseVouchersBtn.click();

		Thread.sleep(8000);
		
		
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdVoucher));
		createdVoucher.click();
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustCreateVoucherBtn));
		docCustCreateVoucherBtn.click();
		Thread.sleep(2000)*/;
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deletevoucherBtn));
		deletevoucherBtn.click();

		getWaitForAlert();
		Thread.sleep(2000);
		getAlert().accept();
		Thread.sleep(2000);

		String expMessage = excelReader.getCellData(xlSheetName, 94, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 94, 7, actMessage);


		Thread.sleep(3000);

		String acttitleTxt					 =titleTxt.getAttribute("value");
		String exptitleTxt					 =excelReader.getCellData(xlSheetName, 95, 6);
		excelReader.setCellData(xlfile, xlSheetName, 95, 7, acttitleTxt);

		String actMastersR1C1                = masters1.getText();
		String actPositionR1C1               = position1.getText();
		String actMastersR2C1                = masters2.getText();
		String actPositionR2C1               = position2.getText();

		String expMastersR1C1                = excelReader.getCellData(xlSheetName, 96, 6);
		excelReader.setCellData(xlfile, xlSheetName, 96, 7, actMastersR1C1);
		String expPositionR1C1               = excelReader.getCellData(xlSheetName, 97, 6);
		excelReader.setCellData(xlfile, xlSheetName, 97, 7, actPositionR1C1);
		String expMastersR2C1                = excelReader.getCellData(xlSheetName, 98, 6);
		excelReader.setCellData(xlfile, xlSheetName, 98, 7, actMastersR2C1);
		String expPositionR2C1               = excelReader.getCellData(xlSheetName, 99, 6);
		excelReader.setCellData(xlfile, xlSheetName, 99, 7, actPositionR2C1);

		System.out.println("*****************************checkDeletingPurchaseVoucherDocument****************************");

		System.out.println("Error Message                                 : " + actMessage       +     "     Value Expected : " + expMessage);
		System.out.println("Documents Tab Title                           : " + acttitleTxt      +     "	 Value Expected : " + exptitleTxt);
		System.out.println("Documents Tab Master First Row                : " + actMastersR1C1   +     "	 Value Expected : " + expMastersR1C1);
		System.out.println("Documents Tab Master Position First Row       : " + actPositionR1C1  +     "	 Value Expected : " + expPositionR1C1);
		System.out.println("Documents Tab Master Second Row               : " + actMastersR2C1   +     "	 Value Expected : " + expMastersR2C1);
		System.out.println("Documents Tab Master Position Second Row      : " + actPositionR2C1  +     "	 Value Expected : " + expPositionR2C1);



		if(actMessage.equalsIgnoreCase(expMessage) && acttitleTxt.equalsIgnoreCase(exptitleTxt) && actMastersR1C1.equalsIgnoreCase(expMastersR1C1) 
				&& actPositionR1C1.equalsIgnoreCase(expPositionR1C1) && actMastersR2C1.equalsIgnoreCase(expMastersR2C1) && actPositionR2C1.equalsIgnoreCase(expPositionR2C1))
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

	

	
	@FindBy(xpath="//*[@id='769']")
	public static WebElement CreatedVoucher;
	
	@FindBy(xpath="(//*[@id='panelsStayOpen-headingThree']/button)[2]")
	public static WebElement currencyBtn;
	
	@FindBy(xpath="//*[@id='panelsStayOpen-headingTwo']/button")
	public static WebElement PostingDetailsBtn;
	
	@FindBy(xpath="//*[@id='id_FooterClose']")
	public static WebElement editScreenCloseBtn;
	
	@FindBy(xpath="(//*[@id='id_FooterClose'])[1]")
	public static WebElement editLayoutCloseBtn;
	
	
	
	private static String xlSheetName1 = "SmokeDocumentCustomization";

	public boolean checkCreatingNewVoucherPurchaseVoucheVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		
		Thread.sleep(2000);

		int count = SettingsSubMenusList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			//ScrollIntoView(SettingsSubMenusList.get(i));
			System.out.println(excelReader.getCellData(xlSheetName, 8, 5));
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 8, 5))) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}
		
		Thread.sleep(2000);
		
*/	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
		searchTxt.click();
		searchTxt.sendKeys("Configure Transactions");
		searchTxt.sendKeys(Keys.ENTER);
		Thread.sleep(3000);
		
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionPreferencesMinus));
		configureTransactionPreferencesMinus.click();
		
		Thread.sleep(2000);*/
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionDocumentCustomiztionMinus));
		configureTransactionDocumentCustomiztionMinus.click();*/
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustPurchaseVouchersBtn));
		docCustPurchaseVouchersBtn.click();

		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustCreateVoucherBtn));
		docCustCreateVoucherBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CreatedVoucher));
		CreatedVoucher.click();

		Thread.sleep(7000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentsTab));
		documentsTab.click();
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(generalBtn));
		generalBtn.click();
*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(titleTxt));
		titleTxt.click();
		titleTxt.clear();
		titleTxt.sendKeys(excelReader.getCellData(xlSheetName1, 102, 6));
		titleTxt.sendKeys(Keys.TAB);

		String actTitle=titleTxt.getAttribute("value");
		String expTitle=excelReader.getCellData(xlSheetName1, 102, 7);
		excelReader.setCellData(xlfile, xlSheetName1, 102, 8, actTitle);

		System.err.println(" PurchaseVoucherVAt Title    : "+actTitle +" Value Expected : "+expTitle);

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masters1));
		masters1.click();

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDropDown));
		masterDropDown.click();
		Thread.sleep(2000);
		Select s1=new Select(masterDropDown);
		s1.selectByVisibleText(excelReader.getCellData(xlSheetName1, 103, 6));

		Select getValuesOfMasterC1=new Select(masterDropDown);
		String actMastersR1C1=getValuesOfMasterC1.getFirstSelectedOption().getText();
		String expMastersR1C1="Warehouse";
		//String expMastersR1C1=excelReader.getCellData(xlSheetName1, 103, 7);
		excelReader.setCellData(xlfile, xlSheetName1, 103, 8, actMastersR1C1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(position1));
		position1.click();

		positionDropdown.click();

		Select p1=new Select(positionDropdown);
		p1.selectByVisibleText(excelReader.getCellData(xlSheetName1, 104, 6));

		Select getValueOfPostionC1=new Select(positionDropdown);
		String actPositionR1C1=getValueOfPostionC1.getFirstSelectedOption().getText();
		String expPositionR1C1="Body";
		//String expPositionR1C1=excelReader.getCellData(xlSheetName1, 104, 7);
		excelReader.setCellData(xlfile, xlSheetName1, 104, 8, actPositionR1C1);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masters2));
		masters2.click();

		Select s2=new Select(masterDropDown);
		s2.selectByVisibleText(excelReader.getCellData(xlSheetName1, 105, 6));

		Select getValuesOfMasterC2=new Select(masterDropDown);
		String actMastersR2C1=getValuesOfMasterC1.getFirstSelectedOption().getText();
		String expMastersR2C1="Department";
		//String expMastersR2C1=excelReader.getCellData(xlSheetName1, 105, 7);
		excelReader.setCellData(xlfile, xlSheetName1, 105, 8, actMastersR2C1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(position2));
		position2.click();

		positionDropdown.click();

		Select p2=new Select(positionDropdown);
		p2.selectByVisibleText(excelReader.getCellData(xlSheetName1, 106, 6));

		Select getValueOfPostionC2=new Select(positionDropdown);
		String actPositionR2C2=getValueOfPostionC1.getFirstSelectedOption().getText();
		String expPositionR2C2="Header";
		//String expPositionR2C2=excelReader.getCellData(xlSheetName1, 106, 7);
		excelReader.setCellData(xlfile, xlSheetName1, 106, 8, actPositionR2C2);	

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousTab));
		miscellaneousTab.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyBtn));
		currencyBtn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addCurrencyInDropdown));
		addCurrencyInDropdown.click();
		Select s=new Select(addCurrencyInDropdown);
		s.selectByVisibleText("Header");

		if (inputExchangeRateChkBox.isSelected()==false && inputLocalExchangeRateChkBox.isSelected()==false) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inputExchangeRateChkBox));
			inputExchangeRateChkBox.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inputLocalExchangeRateChkBox));
			inputLocalExchangeRateChkBox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PostingDetailsBtn));
		PostingDetailsBtn.click();
		Thread.sleep(2000);

		if (selectAccountInEveryLineChkBox.isSelected()==false) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectAccountInEveryLineChkBox));
			selectAccountInEveryLineChkBox.click();
		}

		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
		editScreenTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenLoadFieldsBtn));
		editScreenLoadFieldsBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenLoadFieldsDropdown));
		Select editScreenLoadFieldsDropdownSelect =new Select(editScreenLoadFieldsDropdown);
		//editScreenLoadFieldsDropdownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName1, 100, 6));

		editScreenLoadFieldsDropdownSelect.selectByValue("3328");
		Thread.sleep(2000);



		int editScreenLoadFieldTxtListCount = editScreenLoadFieldTxtList.size();

		System.err.println("editScreenLoadFieldTxtListCount : "+editScreenLoadFieldTxtListCount);
		
		//System.err.println(editScreenLoadFieldTxtList.get(0)); 

		for(int i=0;i<editScreenLoadFieldTxtListCount;i++)
		{
			String data = editScreenLoadFieldTxtList.get(i).getText();
			System.err.println(data);

			if(data.equalsIgnoreCase("Avg Rate"))
			{
				editScreenLoadFieldchkboxList.get(i).click();

				break;
			}
		}	




		int editScreenLoadFieldTxtListCount1 = editScreenLoadFieldTxtList.size();

		System.err.println("editScreenLoadFieldTxtListCount1 : "+editScreenLoadFieldTxtListCount1);

		for(int i=0;i<editScreenLoadFieldTxtListCount1;i++)
		{
			String data = editScreenLoadFieldTxtList.get(i).getText();
			System.out.println(data);

			if(data.equalsIgnoreCase("Avg Rate(O)"))
			{
				editScreenLoadFieldchkboxList.get(i).click();

				break;
			}
		}	


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenokBtn));
		editScreenokBtn.click();

		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenCloseBtn));
		editScreenCloseBtn.click();
		Thread.sleep(1000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTab));
		editLayoutTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutHeaderTab));
		editLayoutHeaderTab.click();
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutLoadFieldsBtn));
		editLayoutLoadFieldsBtn.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutLoadFieldsDropdown));
		Select editLayoutLoadFieldsDropdownSelect =new Select(editLayoutLoadFieldsDropdown);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutLoadFieldsDropdown));
		//editLayoutLoadFieldsDropdownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName1, 107, 6));
		editLayoutLoadFieldsDropdownSelect.selectByValue("768");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutselectAllBtn));
		editLayoutselectAllBtn.click();
		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutokBtn));
		editLayoutokBtn.click();

		Thread.sleep(2000);
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCloseBtn));
		editLayoutCloseBtn.click();
		Thread.sleep(1000);
		*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTab));
		editLayoutTab.click();

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBodyTab));
		editLayoutBodyTab.click();

		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutLoadFieldsBtn));
		editLayoutLoadFieldsBtn.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutLoadFieldsDropdown));
		//editLayoutLoadFieldsDropdownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName1, 107, 6));
		editLayoutLoadFieldsDropdownSelect.selectByValue("768");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutselectAllBtn));
		editLayoutselectAllBtn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutokBtn));
		editLayoutokBtn.click();

		Thread.sleep(2000);
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCloseBtn));
		editLayoutCloseBtn.click();
		Thread.sleep(1000);
		*/


		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
*/
		/*String expMessage=excelReader.getCellData(xlSheetName1, 108, 7);

		String actMessage=checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName1, 108, 8, actMessage);

		Thread.sleep(2000);
*/		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CreatedVoucher));
		CreatedVoucher.click();

		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		Thread.sleep(2000);
		
		//String expMessage=excelReader.getCellData(xlSheetName1, 108, 7);
		String expMessage= "Data saved successfully";
		String actMessage=checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName1, 108, 8, actMessage);

		LogoutandLoginwithSU();
		Thread.sleep(2000);
		
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
		searchTxt.click();
		searchTxt.sendKeys("Settings Wizard");
		searchTxt.sendKeys(Keys.ENTER);
		Thread.sleep(3000);
		
		
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsWizardMenu));
		settingsWizardMenu.click();

		Thread.sleep(2000);
*/		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsWizardVATBtn));
		settingsWizardVATBtn.click();
		
		Thread.sleep(2000);
		
		int actVouchersCountAfterSelection = settingsWizardVATVouchersList.size();
		
		System.err.println(actVouchersCountAfterSelection);
		
		for (int i = 0; i < actVouchersCountAfterSelection; i++) 
		{
			String data=settingsWizardVATVouchersList.get(i).getText();
			System.err.println(data);

			//if (data.contains("VAT") && settingsWizardVATVoucherChkboxList.get(i).isSelected()==false ) 
			if(data.contains("Purchase Voucher VAT"))
			{
				System.out.println(settingsWizardVATVouchersList.get(i).getText());
				
				//settingsWizardVATVouchersList.get(i).click();
				( settingsWizardVATVoucherChkboxListSelected).get(i).click();
				
			}
			
			
			
		}


		
		int settingsWizardVATSelectedVouchersListCount = settingsWizardVATSelectedVouchersList.size();

		for(int i=0;i<settingsWizardVATSelectedVouchersListCount;i++)
		{
			settingsWizardVATSelectedVouchersList.get(i).click();

			Thread.sleep(1000);

			settingsWizardVATGrossChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsWizardUpdateBtn));
		settingsWizardUpdateBtn.click();


		String expSaveSettingWizardVATmessage = "Data saved Successfully"; 
		String actSaveSettingWizardVATmessage = checkValidationMessage(expSaveSettingWizardVATmessage);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsWizardVAtCloseBtn));
		settingsWizardVAtCloseBtn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		
		
int count1 = SettingsSubMenusList.size();
		
		for (int i = 0; i < count1; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			//ScrollIntoView(SettingsSubMenusList.get(i));
			System.out.println(excelReader.getCellData(xlSheetName, 8, 5));
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 8, 5))) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}
		
		Thread.sleep(3000);
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionPreferencesMinus));
		configureTransactionPreferencesMinus.click();
		
		Thread.sleep(2000);
*/	/*	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionDocumentCustomiztionMinus));
		configureTransactionDocumentCustomiztionMinus.click();*/
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustPurchaseVouchersBtn));
		docCustPurchaseVouchersBtn.click();

		Thread.sleep(5000);

	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustCreateVoucherBtn));
		docCustCreateVoucherBtn.click();
		
		Thread.sleep(3000);
		*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CreatedVoucher));
		CreatedVoucher.click();

		Thread.sleep(7000);
		
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(generalBtn));
		generalBtn.click();
		*/Thread.sleep(2000);
		

		System.out.println("****************************************checkCreatingNewVoucherInPostDatedPayments**************************************");

		System.out.println("MastersR1C1                   : "+actMastersR1C1         +"  value expected  "+expMastersR1C1);
		System.out.println("MastersR1C2                   : "+actMastersR2C1         +"  value expected  "+expMastersR2C1);
		System.out.println("PositionR2C2                  : "+actPositionR2C2        +"  value expected  "+expPositionR2C2);
		System.out.println("Message                       : "+actMessage             +"  value expected  "+expMessage );


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(3000);

		if(actMastersR2C1.equalsIgnoreCase(expMastersR2C1) && actPositionR2C2.equalsIgnoreCase(expPositionR2C2) &&
				actPositionR1C1 .equalsIgnoreCase(expPositionR1C1 ) && actMastersR1C1.equalsIgnoreCase(expMastersR1C1) && 
				actMessage.equalsIgnoreCase(expMessage) && actSaveSettingWizardVATmessage.equalsIgnoreCase(expSaveSettingWizardVATmessage))
		{
			excelReader.setCellData(xlfile, xlSheetName1, 101, 9, resPass);
			System.out.println("Test Pass : PV VAT Voucher is Created Successfully");
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName1, 101, 9, resFail);
			System.out.println("Test Fail :PV VAT  Voucher is Created Successfully");
			return false;
		}
	}


////
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
		 
		// getDriver().navigate().refresh();
		 //Thread.sleep(1999);

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt="su";
				      
		String pawslt="su";
				      
		LoginPage.enterUserName(unamelt);

		Thread.sleep(1000);
				
		LoginPage.enterPassword(pawslt);
	
		 Select oSelect = new Select(companyDropDownList);
		 //oSelect.selectByVisibleText(Compname); 
		 oSelect.selectByValue("36");
		
		 LoginPage.clickOnSignInBtn();
		 
		 Thread.sleep(2000);
		 LoginPage.reLogin("su", "su", "Automation Company");
		 
		
		 
		
		
	}





	public DocumentCustomizationPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
