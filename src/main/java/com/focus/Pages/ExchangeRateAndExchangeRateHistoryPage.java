package com.focus.Pages;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

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

public class ExchangeRateAndExchangeRateHistoryPage extends BaseEngine 
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "ExchangeRateAndHistoryPage";


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



	public static boolean checkOpenExchangeRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMenu));
		currencyMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateMenu));
		exchangeRateMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateDefinitionSaveIcon));

		String expected = excelReader.getCellData(xlSheetName, 8, 6);

		String actual = Boolean.toString(exchangeRateDefinitionSaveIcon.isDisplayed());

		System.out.println("Save Button is Displayed : "+actual+"  "+expected);

		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actual);

		if(actual.equalsIgnoreCase(expected))
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






	public static boolean checkDeleteOptionInExchangeRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyTxt));
		baseCurrencyTxt.click();
		baseCurrencyTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		Thread.sleep(2000);
		baseCurrencyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchnageRateDeleteOption));

		exchnageRateDeleteOption.click();

		getWaitForAlert();

		String actCompanySaveMsg=getAlert().getText();

		getAlert().accept();

		System.out.println("*******************************checkDeleteOptionInExchangeRate********************************");

		System.out.println("Company message is displaying as "+ actCompanySaveMsg);

		String expMessage=excelReader.getCellData(xlSheetName, 9, 6);

		String actMessage=checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 9, 7, actMessage);


		if(actMessage.equalsIgnoreCase(expMessage))
		{

			excelReader.setCellData(xlfile, xlSheetName, 9, 8, resPass);
			return true;
		}
		else
		{

			excelReader.setCellData(xlfile, xlSheetName, 9, 8, resFail);

			return false;
		}
	}






	public static boolean checkSavingCurrencyExchnageRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyTxt));
		baseCurrencyTxt.click();
		baseCurrencyTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));

		Thread.sleep(2000);

		baseCurrencyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(withEffectiveDateTxt));

		withEffectiveDateTxt.click();

		getAction().moveToElement(withEffectiveDateTxt).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridcurrencyName1Txt));

		gridcurrencyName1Txt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridDefineAsTxt));

		gridDefineAsTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateTxt));
		gridRateTxt.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		gridRateTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridcurrencyName1Txt));

		gridcurrencyName1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridDefineAsTxt));

		gridDefineAsTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateTxt));
		gridRateTxt.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		gridRateTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateDefinitionSaveIcon));
		exchangeRateDefinitionSaveIcon.click();

		System.out.println("***************checkSavingCurrencyExchnageRate***************");

		String expMessage=excelReader.getCellData(xlSheetName, 10, 6);

		String actMessage=checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 10, 7, actMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateDefinitionPageCloseIcon));
			exchangeRateDefinitionPageCloseIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 10, 8, resPass);


			return true;
		}	 
		else
		{    
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateDefinitionPageCloseIcon));
			exchangeRateDefinitionPageCloseIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 10, 8, resFail);
			return false;
		}	     
	}







	public static boolean checkOpenExchangeRateHistory() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMenu));
		currencyMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateHistoryMenu));
		exchangeRateHistoryMenu.click();

		Thread.sleep(2000);  
		String expected = excelReader.getCellData(xlSheetName, 13, 6);

		String actual = Boolean.toString(loadIcon.isDisplayed());

		System.out.println("Load Button is Displayed : "+actual+"  "+expected);

		excelReader.setCellData(xlfile, xlSheetName, 13, 7, actual);


		if(actual.equalsIgnoreCase(expected))
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








	public static boolean checkExchangeRateHistoryLoadButtonWithBaseCurrency() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loadIcon));
		loadIcon.click();

		String expMessage=excelReader.getCellData(xlSheetName, 14, 6);

		String actMessage=checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 14, 7, actMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyInExchangeRateHistory));

		baseCurrencyInExchangeRateHistory.sendKeys(excelReader.getCellData(xlSheetName, 14, 5));

		Thread.sleep(2000);

		baseCurrencyInExchangeRateHistory.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		System.out.println(defCurrecnyOne.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueUSD));

		String actValueUSD                  = valueUSD.getText();
		String actValueAED                  = valueAED.getText();
		String actselCurrencyOne            = selCurrencyOne.getText();
		String actselCurrecnyTwo            = selCurrecnyTwo.getText();
		String actdefCurrecnyOne            = defCurrecnyOne.getText();
		String actdefCurrecnyTwo            = defCurrecnyTwo.getText();

		String expValueUSD                  = "100.0000000000";
		String expValueAED                  = "50.0000000000";
		String expselCurrencyOne            = "USD";
		String expselCurrecnyTwo            = "AED";
		String expdefCurrecnyOne            = "INR";
		String expdefCurrecnyTwo            = "INR";

		System.out.println("********************************checkExchangeRateHistorySelectINR***********************************");

		System.out.println("valueUSD            "+actValueUSD       +"Value Excepted   :"+expValueUSD);
		System.out.println("ValueAED            "+actValueAED       +"Value Excepted   :"+expValueAED);
		System.out.println("selCurrencyOne      "+actselCurrencyOne +"Value Excepted   :"+expselCurrencyOne);
		System.out.println("selCurrecnyTwo      "+actselCurrecnyTwo +"Value Excepted   :"+expselCurrecnyTwo);
		System.out.println("defCurrecnyOne      "+actdefCurrecnyOne +"Value Excepted   :"+expdefCurrecnyOne);
		System.out.println("defCurrecnyTwo      "+actdefCurrecnyTwo +"Value Excepted   :"+expdefCurrecnyTwo);


		excelReader.setCellData(xlfile, xlSheetName, 15, 7, actValueUSD);
		excelReader.setCellData(xlfile, xlSheetName, 16, 7, actValueAED);
		excelReader.setCellData(xlfile, xlSheetName, 17, 7, actselCurrencyOne);
		excelReader.setCellData(xlfile, xlSheetName, 18, 7, actselCurrecnyTwo);
		excelReader.setCellData(xlfile, xlSheetName, 19, 7, actdefCurrecnyOne);
		excelReader.setCellData(xlfile, xlSheetName, 20, 7, actdefCurrecnyTwo);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateHistoryCloseIcon));

		exchangeRateHistoryCloseIcon.click();

		if (actValueUSD.equalsIgnoreCase(expValueUSD)&& actValueAED.equalsIgnoreCase(expValueAED) && 
				actselCurrencyOne.equalsIgnoreCase(expselCurrencyOne) && actselCurrecnyTwo.equalsIgnoreCase(expselCurrecnyTwo) && 
				actdefCurrecnyOne.equalsIgnoreCase(expdefCurrecnyOne) && actdefCurrecnyTwo.equalsIgnoreCase(expdefCurrecnyTwo))
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







	public static boolean checkUpdateExchangeRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		excelReader=new ExcelReader(POJOUtility.getExcelPath());


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMenu));
		currencyMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateMenu));
		exchangeRateMenu.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyTxt));
		baseCurrencyTxt.click();
		baseCurrencyTxt.sendKeys(excelReader.getCellData(xlSheetName, 21, 5));

		Thread.sleep(2000);

		baseCurrencyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(withEffectiveDateTxt));

		getAction().moveToElement(withEffectiveDateTxt).sendKeys(Keys.TAB).perform();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridcurrencyName1Txt));
		getAction().moveToElement(gridcurrencyName1Txt).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridDefineAsTxt));
		getAction().moveToElement(gridDefineAsTxt).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateTxt));
		gridRateTxt.sendKeys(excelReader.getCellData(xlSheetName, 22, 5));

		getAction().moveToElement(gridRateTxt).sendKeys(Keys.TAB).perform();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridcurrencyName1Txt));
		getAction().moveToElement(gridcurrencyName1Txt).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridDefineAsTxt));
		getAction().moveToElement(gridDefineAsTxt).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateTxt));
		gridRateTxt.sendKeys(excelReader.getCellData(xlSheetName, 23, 5));

		getAction().moveToElement(gridRateTxt).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateDefinitionSaveIcon));
		exchangeRateDefinitionSaveIcon.click();
		Thread.sleep(2000);

		System.out.println("*************************checkUpdateExchangeRate*************************");

		String expMessage=excelReader.getCellData(xlSheetName, 21, 6);

		String actMessage=checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateDefinitionPageCloseIcon));
		exchangeRateDefinitionPageCloseIcon.click();

		if(actMessage.equalsIgnoreCase(expMessage))
		{	

			excelReader.setCellData(xlfile, xlSheetName, 21, 8, resPass);
			return true;
		}	 
		else
		{    

			excelReader.setCellData(xlfile, xlSheetName, 21, 8, resFail);

			return false;
		}	     
	}







	public static boolean checkExchangeRateHistoryAfterUpdating() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMenu));
		currencyMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateHistoryMenu));
		exchangeRateHistoryMenu.click();

		Thread.sleep(2000);  

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyInExchangeRateHistory));

		baseCurrencyInExchangeRateHistory.sendKeys(excelReader.getCellData(xlSheetName, 24, 5));

		Thread.sleep(2000);

		baseCurrencyInExchangeRateHistory.sendKeys(Keys.TAB);

		System.out.println(defCurrecnyOne.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueUSD));

		String actValueUSD                  = valueUSD.getText();
		String actValueAED                  = valueAED.getText();
		String actselCurrencyOne            = selCurrencyOne.getText();
		String actselCurrecnyTwo            = selCurrecnyTwo.getText();
		String actdefCurrecnyOne            = defCurrecnyOne.getText();
		String actdefCurrecnyTwo            = defCurrecnyTwo.getText();

		String expValueUSD                  = "70.0000000000";
		String expValueAED                  = "14.0000000000";
		String expselCurrencyOne            = "USD";
		String expselCurrecnyTwo            = "AED";
		String expdefCurrecnyOne            = "INR";
		String expdefCurrecnyTwo            = "INR";

		System.out.println("********************************checkExchangeRateHistoryAfterUpdating***********************************");

		System.out.println("valueUSD            "+actValueUSD       +"Value Excepted   :"+expValueUSD);
		System.out.println("ValueAED            "+actValueAED       +"Value Excepted   :"+expValueAED);
		System.out.println("selCurrencyOne      "+actselCurrencyOne +"Value Excepted   :"+expselCurrencyOne);
		System.out.println("selCurrecnyTwo      "+actselCurrecnyTwo +"Value Excepted   :"+expselCurrecnyTwo);
		System.out.println("defCurrecnyOne      "+actdefCurrecnyOne +"Value Excepted   :"+expdefCurrecnyOne);
		System.out.println("defCurrecnyTwo      "+actdefCurrecnyTwo +"Value Excepted   :"+expdefCurrecnyTwo);

		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actValueUSD);
		excelReader.setCellData(xlfile, xlSheetName, 25, 7, actValueAED);
		excelReader.setCellData(xlfile, xlSheetName, 26, 7, actselCurrencyOne);
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actselCurrecnyTwo);
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actdefCurrecnyOne);
		excelReader.setCellData(xlfile, xlSheetName, 29, 7, actdefCurrecnyTwo);




		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateHistoryCloseIcon));
		exchangeRateHistoryCloseIcon.click();

		if (actValueUSD.equalsIgnoreCase(expValueUSD)&& actValueAED.equalsIgnoreCase(expValueAED) && 
				actselCurrencyOne.equalsIgnoreCase(expselCurrencyOne) && actselCurrecnyTwo.equalsIgnoreCase(expselCurrecnyTwo) && 
				actdefCurrecnyOne.equalsIgnoreCase(expdefCurrecnyOne) && actdefCurrecnyTwo.equalsIgnoreCase(expdefCurrecnyTwo))
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


	
	@FindBy(xpath="(//select[@id='DateOption'])[1]")
	public static WebElement selectDateRange;
	
	@FindBy(xpath="//*[@id='exc_FromDate']")
	public static WebElement FromDateTxt;
	
	@FindBy(xpath="//*[@id='exc_ToDate']")
	public static WebElement ToDateTxt;
	
	@FindBy(xpath="//tbody[@id='HistoryTableBody']/tr[2]/td[2]")
	private static WebElement valueEffectiveDate;
	
	
	public boolean checkDateRangeOptionAsOnDateinExchangeRateHistory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		    homeMenu.click();
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		    mastersMenu.click();
		     
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMenu));
	        currencyMenu.click();
	         
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateHistoryMenu));
	        exchangeRateHistoryMenu.click();
	        
	        Thread.sleep(2000);  
	        
	   
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyInExchangeRateHistory));
			
			baseCurrencyInExchangeRateHistory.sendKeys("INR");
			
			Thread.sleep(2000);
			
			baseCurrencyInExchangeRateHistory.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectDateRange));
			Select s=new Select(selectDateRange);
			s.selectByValue("1");
			Thread.sleep(1500);
			selectDateRange.sendKeys(Keys.TAB);
			
			String actFromDate=FromDateTxt.getAttribute("value");
			String expFromDate="02/01/2020";
					
			
			Calendar cal=Calendar.getInstance();
			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
			String currentDate = df.format(cal.getTime());
			
			DateFormat df1 = new SimpleDateFormat("dd-MM-yyyy");
			String currentDate1 = df1.format(cal.getTime());
			
			String actToDate=ToDateTxt.getAttribute("value");
			String expToDate=currentDate;
			
						
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loadIcon));
			loadIcon.click();
			
					
			System.out.println(defCurrecnyOne.getText());
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueEffectiveDate));
			
			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueUSD));*/
			String actEffectiveDate				= valueEffectiveDate.getText();
			String actValueUSD                  = valueUSD.getText();
			String actValueAED                  = valueAED.getText();
			String actselCurrencyOne            = selCurrencyOne.getText();
			String actselCurrecnyTwo            = selCurrecnyTwo.getText();
			String actdefCurrecnyOne            = defCurrecnyOne.getText();
			String actdefCurrecnyTwo            = defCurrecnyTwo.getText();
			
			String expEffectiveDate				= currentDate1;
			String expValueUSD                  = "70.0000000000";
			String expValueAED                  = "14.0000000000";
			String expselCurrencyOne            = "USD";
			String expselCurrecnyTwo            = "AED";
			String expdefCurrecnyOne            = "INR";
			String expdefCurrecnyTwo            = "INR";
			
			System.out.println("********************************checkExchangeRateHistorySelectINR***********************************");
			
			
			System.out.println("From Date		 "	+ actFromDate 	+"Expected		"	+ expFromDate);
			System.out.println("To Date			 "	+ actToDate 	+"Expected		"	+ expToDate);
			System.out.println("Effective Date	 "	+ actEffectiveDate 	+"Expected		"	+ expEffectiveDate);
			System.out.println("valueUSD            "+actValueUSD       +"Value Excepted   :"+expValueUSD);
			System.out.println("ValueAED            "+actValueAED       +"Value Excepted   :"+expValueAED);
			System.out.println("selCurrencyOne      "+actselCurrencyOne +"Value Excepted   :"+expselCurrencyOne);
			System.out.println("selCurrecnyTwo      "+actselCurrecnyTwo +"Value Excepted   :"+expselCurrecnyTwo);
			System.out.println("defCurrecnyOne      "+actdefCurrecnyOne +"Value Excepted   :"+expdefCurrecnyOne);
			System.out.println("defCurrecnyTwo      "+actdefCurrecnyTwo +"Value Excepted   :"+expdefCurrecnyTwo);
			
			
		/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateHistoryCloseIcon));
		     
			exchangeRateHistoryCloseIcon.click();*/
			
			if (actValueUSD.equalsIgnoreCase(expValueUSD)&& actValueAED.equalsIgnoreCase(expValueAED) && 
				actselCurrencyOne.equalsIgnoreCase(expselCurrencyOne) && actselCurrecnyTwo.equalsIgnoreCase(expselCurrecnyTwo) && 
				actdefCurrecnyOne.equalsIgnoreCase(expdefCurrecnyOne) && actdefCurrecnyTwo.equalsIgnoreCase(expdefCurrecnyTwo)
				&&actFromDate.equalsIgnoreCase(expFromDate) && actToDate.equalsIgnoreCase(expToDate) &&
				actEffectiveDate.equalsIgnoreCase(expEffectiveDate))
			{	
					return true;
			}	 
			else
			{
				   return false;
			}
	        
	}
	
	
	@FindBy(xpath="//tbody[@id='HistoryTableBody']/tr[1]/td[3]")
	private static WebElement valueAEDUSD;
	
	@FindBy(xpath="//tbody[@id='HistoryTableBody']/tr[1]/td[2]")
	private static WebElement valueAEDEffectiveDate;
	
	@FindBy(xpath="//tbody[@id='HistoryTableBody']/tr[1]/td[4]")
	private static WebElement valueAEDINR;
	
	
	
	public boolean checkDateOptionasCurrentMonthinExchangeRateHistory() throws InterruptedException
	{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyInExchangeRateHistory));
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.END , Keys.SHIFT ,Keys.HOME);
				
		baseCurrencyInExchangeRateHistory.sendKeys("AED");
		
		Thread.sleep(2000);
		
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectDateRange));
		Select s=new Select(selectDateRange);
		s.selectByValue("2");
		Thread.sleep(1500);
		selectDateRange.sendKeys(Keys.TAB);
		
		//current Month
		
		Calendar cal=Calendar.getInstance();
		cal.add(Calendar.MONTH, 0);
		
		
		cal.set(Calendar.DATE,cal.getActualMinimum(Calendar.DAY_OF_MONTH));
		Date fromDate=cal.getTime();
		cal.set(Calendar.DATE, cal.getActualMaximum(Calendar.DAY_OF_MONTH));
		Date toDate=cal.getTime();
			
		SimpleDateFormat f = new SimpleDateFormat("dd/MM/yyyy");
			
		
		String actFromDate=FromDateTxt.getAttribute("value");
		String expFromDate=f.format(fromDate);
		
		
		String actToDate=ToDateTxt.getAttribute("value");
		String expToDate=f.format(toDate);
		
					
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loadIcon));
		loadIcon.click();
		
				
		System.out.println(defCurrecnyOne.getText());
		
		SimpleDateFormat f1 = new SimpleDateFormat("dd-MM-yyyy");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueAEDEffectiveDate));
		String actEffectiveDate				= valueAEDEffectiveDate.getText();
		String actValueUSD                  = valueAEDUSD.getText();
		String actValueAED                  = valueAEDINR.getText();
		String actselCurrencyOne            = selCurrencyOne.getText();
		String actselCurrecnyTwo            = selCurrecnyTwo.getText();
		String actdefCurrecnyOne            = defCurrecnyOne.getText();
		String actdefCurrecnyTwo            = defCurrecnyTwo.getText();
		
		String expEffectiveDate				= f1.format(fromDate);
		String expValueUSD                  = "4.2800000000";
		String expValueAED                  = "0.0700000000";
		String expselCurrencyOne            = "USD";
		String expselCurrecnyTwo            = "INR";
		String expdefCurrecnyOne            = "AED";
		String expdefCurrecnyTwo            = "AED";
		
		System.out.println("********************************checkExchangeRateHistoryCurrentMonthSelectAED***********************************");
		
		
		System.out.println("From Date		:"	+ actFromDate 	+	"Expected		:"	+ expFromDate);
		System.out.println("To Date			:"	+ actToDate 	+	"Expected		:"	+ expToDate);
		System.out.println("Effective Date	 "	+ actEffectiveDate 	+"Expected		"	+ expEffectiveDate);
		System.out.println("valueUSD            "+actValueUSD       +"Value Excepted   :"+expValueUSD);
		System.out.println("ValueAED            "+actValueAED       +"Value Excepted   :"+expValueAED);
		System.out.println("selCurrencyOne      "+actselCurrencyOne +"Value Excepted   :"+expselCurrencyOne);
		System.out.println("selCurrecnyTwo      "+actselCurrecnyTwo +"Value Excepted   :"+expselCurrecnyTwo);
		System.out.println("defCurrecnyOne      "+actdefCurrecnyOne +"Value Excepted   :"+expdefCurrecnyOne);
		System.out.println("defCurrecnyTwo      "+actdefCurrecnyTwo +"Value Excepted   :"+expdefCurrecnyTwo);
		
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateHistoryCloseIcon));
	     
		exchangeRateHistoryCloseIcon.click();
		*/
		if (actValueUSD.equalsIgnoreCase(expValueUSD)&& actValueAED.equalsIgnoreCase(expValueAED) && 
			actselCurrencyOne.equalsIgnoreCase(expselCurrencyOne) && actselCurrecnyTwo.equalsIgnoreCase(expselCurrecnyTwo) && 
			actdefCurrecnyOne.equalsIgnoreCase(expdefCurrecnyOne) && actdefCurrecnyTwo.equalsIgnoreCase(expdefCurrecnyTwo)
			&&actFromDate.equalsIgnoreCase(expFromDate) && actToDate.equalsIgnoreCase(expToDate) &&
			actEffectiveDate.equalsIgnoreCase(expEffectiveDate))
		{	
				return true;
		}	 
		else
		{
			   return false;
		}
        
	}
	
	public boolean checkDateOptionasLastMonthinExchangeRateHistory() throws InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyInExchangeRateHistory));
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.END , Keys.SHIFT ,Keys.HOME);
		
		baseCurrencyInExchangeRateHistory.sendKeys("AED");
		
		Thread.sleep(2000);
		
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectDateRange));
		Select s=new Select(selectDateRange);
		s.selectByValue("3");
		Thread.sleep(1500);
		selectDateRange.sendKeys(Keys.TAB);
		
		//Last Month
		
		 Calendar cal=Calendar.getInstance();
			cal.add(Calendar.MONTH, -1);
			
			
			cal.set(Calendar.DATE,cal.getActualMinimum(Calendar.DAY_OF_MONTH));
			Date fromDate=cal.getTime();
			cal.set(Calendar.DATE, cal.getActualMaximum(Calendar.DAY_OF_MONTH));
			Date toDate=cal.getTime();
			
		SimpleDateFormat f = new SimpleDateFormat("dd/MM/yyyy");
	
		String actFromDate=FromDateTxt.getAttribute("value");
		String expFromDate=f.format(fromDate);
		
		
		String actToDate=ToDateTxt.getAttribute("value");
		String expToDate=f.format(toDate);
		
					
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loadIcon));
		loadIcon.click();
		
				
		System.out.println(defCurrecnyOne.getText());
		
		SimpleDateFormat f1 = new SimpleDateFormat("dd-MM-yyyy");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueAEDEffectiveDate));
		
		String actEffectiveDate				= valueAEDEffectiveDate.getText();
		String actValueUSD                  = valueAEDUSD.getText();
		String actValueAED                  = valueAEDINR.getText();
		String actselCurrencyOne            = selCurrencyOne.getText();
		String actselCurrecnyTwo            = selCurrecnyTwo.getText();
		String actdefCurrecnyOne            = defCurrecnyOne.getText();
		String actdefCurrecnyTwo            = defCurrecnyTwo.getText();
		
		
		String expEffectiveDate				= f1.format(fromDate);
		String expValueUSD                  = "4.2800000000";
		String expValueAED                  = "0.0700000000";
		String expselCurrencyOne            = "USD";
		String expselCurrecnyTwo            = "INR";
		String expdefCurrecnyOne            = "AED";
		String expdefCurrecnyTwo            = "AED";
		
		System.out.println("********************************checkExchangeRateHistoryLastMonthSelectAED***********************************");
		
		
		System.out.println("From Date		:"	+ actFromDate 	+"Expected		:"	+ expFromDate);
		System.out.println("To Date			:"	+ actToDate 	+"Expected		:"	+ expToDate);
		System.out.println("Effective Date	 "	+ actEffectiveDate 	+"Expected		"	+ expEffectiveDate);
		System.out.println("valueUSD            "+actValueUSD       +"Value Excepted   :"+expValueUSD);
		System.out.println("ValueAED            "+actValueAED       +"Value Excepted   :"+expValueAED);
		System.out.println("selCurrencyOne      "+actselCurrencyOne +"Value Excepted   :"+expselCurrencyOne);
		System.out.println("selCurrecnyTwo      "+actselCurrecnyTwo +"Value Excepted   :"+expselCurrecnyTwo);
		System.out.println("defCurrecnyOne      "+actdefCurrecnyOne +"Value Excepted   :"+expdefCurrecnyOne);
		System.out.println("defCurrecnyTwo      "+actdefCurrecnyTwo +"Value Excepted   :"+expdefCurrecnyTwo);
		
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateHistoryCloseIcon));
	     
		exchangeRateHistoryCloseIcon.click();*/
		
		if (actValueUSD.equalsIgnoreCase(expValueUSD)&& actValueAED.equalsIgnoreCase(expValueAED) && 
			actselCurrencyOne.equalsIgnoreCase(expselCurrencyOne) && actselCurrecnyTwo.equalsIgnoreCase(expselCurrecnyTwo) && 
			actdefCurrecnyOne.equalsIgnoreCase(expdefCurrecnyOne) && actdefCurrecnyTwo.equalsIgnoreCase(expdefCurrecnyTwo)
			&&actFromDate.equalsIgnoreCase(expFromDate) && actToDate.equalsIgnoreCase(expToDate) &&
			actEffectiveDate.equalsIgnoreCase(expEffectiveDate))
		{	
				return true;
		}	 
		else
		{
			   return false;
		}
        
	}
	
	
	
	public boolean checkDateOptionasCurrentWeekinExchangeRateHistory() throws InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyInExchangeRateHistory));
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.END , Keys.SHIFT ,Keys.HOME);
		baseCurrencyInExchangeRateHistory.sendKeys("USD");
		
		Thread.sleep(2000);
		
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectDateRange));
		Select s=new Select(selectDateRange);
		s.selectByValue("4");
		Thread.sleep(1500);
		selectDateRange.sendKeys(Keys.TAB);
		
		//Current Week
		
		Calendar cal=Calendar.getInstance();	
		cal.set(Calendar.DAY_OF_WEEK, Calendar.SUNDAY);
		Date fromDate=cal.getTime();
		cal.add(Calendar.DATE, 6);
		Date toDate=cal.getTime();
			
		SimpleDateFormat f = new SimpleDateFormat("dd/MM/yyyy");
		
		
		String actFromDate=FromDateTxt.getAttribute("value");
		String expFromDate=f.format(fromDate);
		
		
		String actToDate=ToDateTxt.getAttribute("value");
		String expToDate=f.format(toDate);
		
					
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loadIcon));
		loadIcon.click();
		
				
		System.out.println(defCurrecnyOne.getText());
		
		SimpleDateFormat f1 = new SimpleDateFormat("dd-MM-yyyy");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueAEDEffectiveDate));
		
		String actEffectiveDate				= valueAEDEffectiveDate.getText();
		
		String actValueUSD                  = valueAEDUSD.getText();
		String actValueAED                  = valueAEDINR.getText();
		String actselCurrencyOne            = selCurrencyOne.getText();
		String actselCurrecnyTwo            = selCurrecnyTwo.getText();
		String actdefCurrecnyOne            = defCurrecnyOne.getText();
		String actdefCurrecnyTwo            = defCurrecnyTwo.getText();
		
		String expEffectiveDate				= f1.format(fromDate);
		String expValueUSD                  = "0.2300000000";
		String expValueAED                  = "0.0100000000";
		String expselCurrencyOne            = "AED";
		String expselCurrecnyTwo            = "INR";
		String expdefCurrecnyOne            = "USD";
		String expdefCurrecnyTwo            = "USD";
		
		System.out.println("********************************checkExchangeRateHistoryCurrentWeekSelectUSD***********************************");
		
		
		System.out.println("From Date		:"	+ actFromDate 	+"Expected		:"	+ expFromDate);
		System.out.println("To Date			:"	+ actToDate 	+"Expected		:"	+ expToDate);
		System.out.println("Effective Date	 "	+ actEffectiveDate 	+"Expected		"	+ expEffectiveDate);
		System.out.println("valueUSD            "+actValueUSD       +"Value Excepted   :"+expValueUSD);
		System.out.println("ValueAED            "+actValueAED       +"Value Excepted   :"+expValueAED);
		System.out.println("selCurrencyOne      "+actselCurrencyOne +"Value Excepted   :"+expselCurrencyOne);
		System.out.println("selCurrecnyTwo      "+actselCurrecnyTwo +"Value Excepted   :"+expselCurrecnyTwo);
		System.out.println("defCurrecnyOne      "+actdefCurrecnyOne +"Value Excepted   :"+expdefCurrecnyOne);
		System.out.println("defCurrecnyTwo      "+actdefCurrecnyTwo +"Value Excepted   :"+expdefCurrecnyTwo);
		
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateHistoryCloseIcon));
	     
		exchangeRateHistoryCloseIcon.click();
		*/
		if (actValueUSD.equalsIgnoreCase(expValueUSD)&& actValueAED.equalsIgnoreCase(expValueAED) && 
			actselCurrencyOne.equalsIgnoreCase(expselCurrencyOne) && actselCurrecnyTwo.equalsIgnoreCase(expselCurrecnyTwo) && 
			actdefCurrecnyOne.equalsIgnoreCase(expdefCurrecnyOne) && actdefCurrecnyTwo.equalsIgnoreCase(expdefCurrecnyTwo)
			&&actFromDate.equalsIgnoreCase(expFromDate) && actToDate.equalsIgnoreCase(expToDate) &&
			actEffectiveDate.equalsIgnoreCase(expEffectiveDate))
		{	
				return true;
		}	 
		else
		{
			   return false;
		}
        
	}
	
	
	public boolean checkDateOptionasLastWeekinExchangeRateHistory() throws InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyInExchangeRateHistory));
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.END , Keys.SHIFT ,Keys.HOME);
		baseCurrencyInExchangeRateHistory.sendKeys("USD");
		
		Thread.sleep(2000);
		
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectDateRange));
		Select s=new Select(selectDateRange);
		s.selectByValue("5");
		Thread.sleep(1500);
		selectDateRange.sendKeys(Keys.TAB);
		
		//Last Week
		
		Calendar cal=Calendar.getInstance();	
		cal.set(Calendar.DAY_OF_WEEK, Calendar.SUNDAY);
		cal.add(Calendar.DATE, -7);
		Date fromDate=cal.getTime();
			
		cal.add(Calendar.DATE, 6);
		Date toDate=cal.getTime();
		 
		  
			
		SimpleDateFormat f = new SimpleDateFormat("dd/MM/yyyy");
		
		
		String actFromDate=FromDateTxt.getAttribute("value");
		String expFromDate=f.format(fromDate);
		
		
		String actToDate=ToDateTxt.getAttribute("value");
		String expToDate=f.format(toDate);
		
					
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loadIcon));
		loadIcon.click();
		
				
		System.out.println(defCurrecnyOne.getText());
		
		SimpleDateFormat f1 = new SimpleDateFormat("dd-MM-yyyy");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueAEDEffectiveDate));
		
		String actEffectiveDate				= valueAEDEffectiveDate.getText();
		String actValueUSD                  = valueAEDUSD.getText();
		String actValueAED                  = valueAEDINR.getText();
		String actselCurrencyOne            = selCurrencyOne.getText();
		String actselCurrecnyTwo            = selCurrecnyTwo.getText();
		String actdefCurrecnyOne            = defCurrecnyOne.getText();
		String actdefCurrecnyTwo            = defCurrecnyTwo.getText();
		
		String expEffectiveDate				= f1.format(fromDate);
		String expValueUSD                  = "0.2300000000";
		String expValueAED                  = "0.0100000000";
		String expselCurrencyOne            = "AED";
		String expselCurrecnyTwo            = "INR";
		String expdefCurrecnyOne            = "USD";
		String expdefCurrecnyTwo            = "USD";
		
		System.out.println("********************************checkExchangeRateHistoryLastWeekSelectUSD***********************************");
		
		
		System.out.println("From Date		:"	+ actFromDate 	+"Expected		:"	+ expFromDate);
		System.out.println("To Date			:"	+ actToDate 	+"Expected		:"	+ expToDate);
		System.out.println("Effective Date	 "	+ actEffectiveDate 	+"Expected		"	+ expEffectiveDate);
		System.out.println("valueUSD            "+actValueUSD       +"Value Excepted   :"+expValueUSD);
		System.out.println("ValueAED            "+actValueAED       +"Value Excepted   :"+expValueAED);
		System.out.println("selCurrencyOne      "+actselCurrencyOne +"Value Excepted   :"+expselCurrencyOne);
		System.out.println("selCurrecnyTwo      "+actselCurrecnyTwo +"Value Excepted   :"+expselCurrecnyTwo);
		System.out.println("defCurrecnyOne      "+actdefCurrecnyOne +"Value Excepted   :"+expdefCurrecnyOne);
		System.out.println("defCurrecnyTwo      "+actdefCurrecnyTwo +"Value Excepted   :"+expdefCurrecnyTwo);
		
		
		
		if (actValueUSD.equalsIgnoreCase(expValueUSD)&& actValueAED.equalsIgnoreCase(expValueAED) && 
			actselCurrencyOne.equalsIgnoreCase(expselCurrencyOne) && actselCurrecnyTwo.equalsIgnoreCase(expselCurrecnyTwo) && 
			actdefCurrecnyOne.equalsIgnoreCase(expdefCurrecnyOne) && actdefCurrecnyTwo.equalsIgnoreCase(expdefCurrecnyTwo)
			&&actFromDate.equalsIgnoreCase(expFromDate) && actToDate.equalsIgnoreCase(expToDate) &&
			actEffectiveDate.equalsIgnoreCase(expEffectiveDate))
		{	
				return true;
		}	 
		else
		{
			   return false;
		}
        
	}
	
	
	public boolean checkDateOptionasTodayinExchangeRateHistory() throws InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyInExchangeRateHistory));
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.END , Keys.SHIFT ,Keys.HOME);
		baseCurrencyInExchangeRateHistory.sendKeys("USD");
		
		Thread.sleep(2000);
		
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectDateRange));
		Select s=new Select(selectDateRange);
		s.selectByValue("6");
		Thread.sleep(1500);
		selectDateRange.sendKeys(Keys.TAB);
		
		//Today
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		String currentDate = df.format(cal.getTime());
		  
			
		//SimpleDateFormat f = new SimpleDateFormat("dd/MM/yyyy");
		
		
		String actFromDate=FromDateTxt.getAttribute("value");
		String expFromDate=currentDate;
		
		
		String actToDate=ToDateTxt.getAttribute("value");
		String expToDate=currentDate;
		
					
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loadIcon));
		loadIcon.click();
		
				
		System.out.println(defCurrecnyOne.getText());
		
		SimpleDateFormat f1 = new SimpleDateFormat("dd-MM-yyyy");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueAEDEffectiveDate));
		String actEffectiveDate				= valueAEDEffectiveDate.getText();
		
		String actValueUSD                  = valueAEDUSD.getText();
		String actValueAED                  = valueAEDINR.getText();
		String actselCurrencyOne            = selCurrencyOne.getText();
		String actselCurrecnyTwo            = selCurrecnyTwo.getText();
		String actdefCurrecnyOne            = defCurrecnyOne.getText();
		String actdefCurrecnyTwo            = defCurrecnyTwo.getText();
		
		String expEffectiveDate				= f1.format(cal.getTime());
		String expValueUSD                  = "0.2300000000";
		String expValueAED                  = "0.0100000000";
		String expselCurrencyOne            = "AED";
		String expselCurrecnyTwo            = "INR";
		String expdefCurrecnyOne            = "USD";
		String expdefCurrecnyTwo            = "USD";
		
		System.out.println("********************************checkExchangeRateHistoryTodaySelectUSD***********************************");
		
		
		System.out.println("From Date		:"	+ actFromDate 	+"Expected		:"	+ expFromDate);
		System.out.println("To Date			:"	+ actToDate 	+"Expected		:"	+ expToDate);
		System.out.println("Effective Date	 "	+ actEffectiveDate 	+"Expected		"	+ expEffectiveDate);
		System.out.println("valueUSD            "+actValueUSD       +"Value Excepted   :"+expValueUSD);
		System.out.println("ValueAED            "+actValueAED       +"Value Excepted   :"+expValueAED);
		System.out.println("selCurrencyOne      "+actselCurrencyOne +"Value Excepted   :"+expselCurrencyOne);
		System.out.println("selCurrecnyTwo      "+actselCurrecnyTwo +"Value Excepted   :"+expselCurrecnyTwo);
		System.out.println("defCurrecnyOne      "+actdefCurrecnyOne +"Value Excepted   :"+expdefCurrecnyOne);
		System.out.println("defCurrecnyTwo      "+actdefCurrecnyTwo +"Value Excepted   :"+expdefCurrecnyTwo);
		
		
		
		if (actValueUSD.equalsIgnoreCase(expValueUSD)&& actValueAED.equalsIgnoreCase(expValueAED) && 
			actselCurrencyOne.equalsIgnoreCase(expselCurrencyOne) && actselCurrecnyTwo.equalsIgnoreCase(expselCurrecnyTwo) && 
			actdefCurrecnyOne.equalsIgnoreCase(expdefCurrecnyOne) && actdefCurrecnyTwo.equalsIgnoreCase(expdefCurrecnyTwo)
			&&actFromDate.equalsIgnoreCase(expFromDate) && actToDate.equalsIgnoreCase(expToDate) &&
			actEffectiveDate.equalsIgnoreCase(expEffectiveDate))
		{	
				return true;
		}	 
		else
		{
			   return false;
		}
        
	}
		
	public boolean checkDateOptionasCurrentYearinExchangeRateHistory() throws InterruptedException
	{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyInExchangeRateHistory));
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.END , Keys.SHIFT ,Keys.HOME);
				
		baseCurrencyInExchangeRateHistory.sendKeys("AED");
		
		Thread.sleep(2000);
		
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectDateRange));
		Select s=new Select(selectDateRange);
		s.selectByValue("7");
		Thread.sleep(1500);
		selectDateRange.sendKeys(Keys.TAB);
		
		//current Year
		
		 Calendar cal=Calendar.getInstance();
		 cal.add(Calendar.YEAR,0);
		 cal.set(Calendar.DAY_OF_YEAR,1);
		 Date fromDate=cal.getTime();
		 Calendar cal2=Calendar.getInstance();
		 Date toDate=cal2.getTime();
		 
		 
		 
		SimpleDateFormat f = new SimpleDateFormat("dd/MM/yyyy");
			
		
		String actFromDate=FromDateTxt.getAttribute("value");
		String expFromDate=f.format(fromDate);
		
		
		String actToDate=ToDateTxt.getAttribute("value");
		String expToDate=f.format(toDate);
		
					
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loadIcon));
		loadIcon.click();
		
				
		System.out.println(defCurrecnyOne.getText());
		
		SimpleDateFormat f1 = new SimpleDateFormat("dd-MM-yyyy");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueAEDEffectiveDate));
		
		String actEffectiveDate				= valueAEDEffectiveDate.getText();
		
		String actValueUSD                  = valueAEDUSD.getText();
		String actValueAED                  = valueAEDINR.getText();
		String actselCurrencyOne            = selCurrencyOne.getText();
		String actselCurrecnyTwo            = selCurrecnyTwo.getText();
		String actdefCurrecnyOne            = defCurrecnyOne.getText();
		String actdefCurrecnyTwo            = defCurrecnyTwo.getText();
		
		String expEffectiveDate				= f1.format(fromDate);
		String expValueUSD                  = "4.2800000000";
		String expValueAED                  = "0.0700000000";
		String expselCurrencyOne            = "USD";
		String expselCurrecnyTwo            = "INR";
		String expdefCurrecnyOne            = "AED";
		String expdefCurrecnyTwo            = "AED";
		
		System.out.println("********************************checkExchangeRateHistoryCurrentYearSelectINR***********************************");
		
		
		System.out.println("From Date		:"	+ actFromDate 	+	"Expected		:"	+ expFromDate);
		System.out.println("To Date			:"	+ actToDate 	+	"Expected		:"	+ expToDate);
		System.out.println("Effective Date	 "	+ actEffectiveDate 	+"Expected		"	+ expEffectiveDate);
		System.out.println("valueUSD            "+actValueUSD       +"Value Excepted   :"+expValueUSD);
		System.out.println("ValueAED            "+actValueAED       +"Value Excepted   :"+expValueAED);
		System.out.println("selCurrencyOne      "+actselCurrencyOne +"Value Excepted   :"+expselCurrencyOne);
		System.out.println("selCurrecnyTwo      "+actselCurrecnyTwo +"Value Excepted   :"+expselCurrecnyTwo);
		System.out.println("defCurrecnyOne      "+actdefCurrecnyOne +"Value Excepted   :"+expdefCurrecnyOne);
		System.out.println("defCurrecnyTwo      "+actdefCurrecnyTwo +"Value Excepted   :"+expdefCurrecnyTwo);
		
		

		if (actValueUSD.equalsIgnoreCase(expValueUSD)&& actValueAED.equalsIgnoreCase(expValueAED) && 
			actselCurrencyOne.equalsIgnoreCase(expselCurrencyOne) && actselCurrecnyTwo.equalsIgnoreCase(expselCurrecnyTwo) && 
			actdefCurrecnyOne.equalsIgnoreCase(expdefCurrecnyOne) && actdefCurrecnyTwo.equalsIgnoreCase(expdefCurrecnyTwo)
			&&actFromDate.equalsIgnoreCase(expFromDate) && actToDate.equalsIgnoreCase(expToDate) &&
			actEffectiveDate.equalsIgnoreCase(expEffectiveDate))
		{	
				return true;
		}	 
		else
		{
			   return false;
		}
        
	}
	
	
	public boolean checkDateOptionasLastYearinExchangeRateHistory() throws InterruptedException
	{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyInExchangeRateHistory));
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.END , Keys.SHIFT ,Keys.HOME);
				
		baseCurrencyInExchangeRateHistory.sendKeys("AED");
		
		Thread.sleep(2000);
		
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectDateRange));
		Select s=new Select(selectDateRange);
		s.selectByValue("8");
		Thread.sleep(1500);
		selectDateRange.sendKeys(Keys.TAB);
		
		//Last Year
		
		 Calendar cal=Calendar.getInstance();
		 cal.add(Calendar.YEAR,-1);
		 cal.set(Calendar.DAY_OF_YEAR,1);
		 Date fromDate=cal.getTime();
		 cal.set(Calendar.DAY_OF_YEAR, cal.getActualMaximum(Calendar.DAY_OF_YEAR));
		 Date toDate=cal.getTime();
		 
		 
		 
		SimpleDateFormat f = new SimpleDateFormat("dd/MM/yyyy");
			
		
		String actFromDate=FromDateTxt.getAttribute("value");
		String expFromDate=f.format(fromDate);
		
		
		String actToDate=ToDateTxt.getAttribute("value");
		String expToDate=f.format(toDate);
		
					
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loadIcon));
		loadIcon.click();
		
				
		System.out.println(defCurrecnyOne.getText());
		
		SimpleDateFormat f1 = new SimpleDateFormat("dd-MM-yyyy");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueAEDEffectiveDate));
		
		String actEffectiveDate				= valueAEDEffectiveDate.getText();
		
		String actValueUSD                  = valueAEDUSD.getText();
		String actValueAED                  = valueAEDINR.getText();
		String actselCurrencyOne            = selCurrencyOne.getText();
		String actselCurrecnyTwo            = selCurrecnyTwo.getText();
		String actdefCurrecnyOne            = defCurrecnyOne.getText();
		String actdefCurrecnyTwo            = defCurrecnyTwo.getText();
		
		String expEffectiveDate				= f1.format(fromDate);
		String expValueUSD                  = "4.2800000000";
		String expValueAED                  = "0.0700000000";
		String expselCurrencyOne            = "USD";
		String expselCurrecnyTwo            = "INR";
		String expdefCurrecnyOne            = "AED";
		String expdefCurrecnyTwo            = "AED";
		
		System.out.println("********************************checkExchangeRateHistoryLastYearSelectINR***********************************");
		
		
		System.out.println("From Date		:"	+ actFromDate 	+	"Expected		:"	+ expFromDate);
		System.out.println("To Date			:"	+ actToDate 	+	"Expected		:"	+ expToDate);
		System.out.println("Effective Date	 "	+ actEffectiveDate 	+"Expected		"	+ expEffectiveDate);
		System.out.println("valueUSD            "+actValueUSD       +"Value Excepted   :"+expValueUSD);
		System.out.println("ValueAED            "+actValueAED       +"Value Excepted   :"+expValueAED);
		System.out.println("selCurrencyOne      "+actselCurrencyOne +"Value Excepted   :"+expselCurrencyOne);
		System.out.println("selCurrecnyTwo      "+actselCurrecnyTwo +"Value Excepted   :"+expselCurrecnyTwo);
		System.out.println("defCurrecnyOne      "+actdefCurrecnyOne +"Value Excepted   :"+expdefCurrecnyOne);
		System.out.println("defCurrecnyTwo      "+actdefCurrecnyTwo +"Value Excepted   :"+expdefCurrecnyTwo);
		
		

		if (actValueUSD.equalsIgnoreCase(expValueUSD)&& actValueAED.equalsIgnoreCase(expValueAED) && 
			actselCurrencyOne.equalsIgnoreCase(expselCurrencyOne) && actselCurrecnyTwo.equalsIgnoreCase(expselCurrecnyTwo) && 
			actdefCurrecnyOne.equalsIgnoreCase(expdefCurrecnyOne) && actdefCurrecnyTwo.equalsIgnoreCase(expdefCurrecnyTwo)
			&&actFromDate.equalsIgnoreCase(expFromDate) && actToDate.equalsIgnoreCase(expToDate) &&
			actEffectiveDate.equalsIgnoreCase(expEffectiveDate))
		{	
				return true;
		}	 
		else
		{
			   return false;
		}
        
	}
	







	public ExchangeRateAndExchangeRateHistoryPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
