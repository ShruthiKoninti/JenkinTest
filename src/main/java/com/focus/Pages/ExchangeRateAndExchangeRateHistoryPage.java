package com.focus.Pages;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
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

		System.out.println("Act Save Button is Displayed : "+actual);
		System.out.println("Exp Save Button is Displayed : "+expected);
		

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
		String expCompanySaveMsg="Are you sure that you want to delete all rows";

		getAlert().accept();

		System.out.println("*******************************checkDeleteOptionInExchangeRate********************************");

		System.out.println("Act Company message is displaying as "+ actCompanySaveMsg);
		System.out.println("Exp Company message is displaying as "+ expCompanySaveMsg);
		

		String expMessage=excelReader.getCellData(xlSheetName, 9, 6);
		String actMessage=checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, actMessage);

		System.out.println("actMessage: "+actMessage);
		System.out.println("expMessage: "+expMessage);

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
		withEffectiveDateTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridcurrencyName1Txt));
		//gridcurrencyName1Txt.click();
		//Thread.sleep(1000);
		clearValueFromTextBox(gridcurrencyName1Txt);
		Thread.sleep(1000);
		gridcurrencyName1Txt.sendKeys("USD");
		Thread.sleep(3000);
		gridcurrencyName1Txt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridDefineAsTxt));
		clearValueFromTextBox(gridDefineAsTxt);
		Thread.sleep(1000);
		gridDefineAsTxt.sendKeys("INR");
		Thread.sleep(2000);
		gridDefineAsTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateTxt));
		gridRateTxt.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		Thread.sleep(2000);
		gridRateTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridcurrencyName1Txt));
		gridcurrencyName1Txt.click();
		gridcurrencyName1Txt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		gridcurrencyName1Txt.sendKeys("AED");
		Thread.sleep(3000);
		gridcurrencyName1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridDefineAsTxt));
		clearValueFromTextBox(gridDefineAsTxt);
		Thread.sleep(1000);
		gridDefineAsTxt.sendKeys("INR");
		Thread.sleep(2000);
		gridDefineAsTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateTxt));
		gridRateTxt.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateDefinitionSaveIcon));
		exchangeRateDefinitionSaveIcon.click();
		Thread.sleep(2000);
		
		System.out.println("***************checkSavingCurrencyExchnageRate***************");

		String expMessage="Exchange Rate updated successfully.";
		String actMessage=checkValidationMessage(expMessage);

		System.out.println("actMessage: "+actMessage);
		System.out.println("expMessage: "+expMessage);

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

		System.out.println("Act Load Button is Displayed : "+actual);
		System.out.println("Exp Load Button is Displayed : "+expected);
		

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
		Thread.sleep(2000);

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

		System.out.println("valueUSD            "+actValueUSD       + "Value Excepted   :"+expValueUSD);
		System.out.println("ValueAED            "+actValueAED       + "Value Excepted   :"+expValueAED);
		System.out.println("selCurrencyOne      "+actselCurrencyOne + "Value Excepted   :"+expselCurrencyOne);
		System.out.println("selCurrecnyTwo      "+actselCurrecnyTwo + "Value Excepted   :"+expselCurrecnyTwo);
		System.out.println("defCurrecnyOne      "+actdefCurrecnyOne + "Value Excepted   :"+expdefCurrecnyOne);
		System.out.println("defCurrecnyTwo      "+actdefCurrecnyTwo + "Value Excepted   :"+expdefCurrecnyTwo);


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
		
		System.out.println("actMessage: "+actMessage);
		System.out.println("expMessage: "+expMessage);

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

		Thread.sleep(4000);  

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

		System.out.println("valueUSD            "+actValueUSD       + "Value Excepted   :"+expValueUSD);
		System.out.println("ValueAED            "+actValueAED       + "Value Excepted   :"+expValueAED);
		System.out.println("selCurrencyOne      "+actselCurrencyOne + "Value Excepted   :"+expselCurrencyOne);
		System.out.println("selCurrecnyTwo      "+actselCurrecnyTwo + "Value Excepted   :"+expselCurrecnyTwo);
		System.out.println("defCurrecnyOne      "+actdefCurrecnyOne + "Value Excepted   :"+expdefCurrecnyOne);
		System.out.println("defCurrecnyTwo      "+actdefCurrecnyTwo + "Value Excepted   :"+expdefCurrecnyTwo);

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
		  Thread.sleep(2000);  
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		    homeMenu.click();
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		    mastersMenu.click();
		     
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMenu));
	        currencyMenu.click();
	         
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateHistoryMenu));
	        exchangeRateHistoryMenu.click();
	        
	        Thread.sleep(4000);  
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyInExchangeRateHistory));
			baseCurrencyInExchangeRateHistory.sendKeys("INR");
			Thread.sleep(4000);
			baseCurrencyInExchangeRateHistory.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectDateRange));
			Select s=new Select(selectDateRange);
			s.selectByValue("1");
			Thread.sleep(1500);
			selectDateRange.sendKeys(Keys.TAB);
			
			String actFromDate=FromDateTxt.getAttribute("value");
			String expFromDate="02/01/2020";
					
			System.out.println("Act From Date "	+ actFromDate);
			System.out.println("Exp From Date "	+ expFromDate);
			
			Calendar cal=Calendar.getInstance();
			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
			String currentDate = df.format(cal.getTime());
			
			DateFormat df1 = new SimpleDateFormat("dd-MM-yyyy");
			String currentDate1 = df1.format(cal.getTime());
			
			String actToDate=ToDateTxt.getAttribute("value");
			String expToDate=currentDate;
			
			System.out.println("Act To Date "	+ actToDate);
			System.out.println("Exp To Date "	+ expToDate);
						
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loadIcon));
			loadIcon.click();
			
			Thread.sleep(2000);
					
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
			
			String expEffectiveDate				= currentDate;
			String expValueUSD                  = "70.0000000000";
			String expValueAED                  = "14.0000000000";
			String expselCurrencyOne            = "USD";
			String expselCurrecnyTwo            = "AED";
			String expdefCurrecnyOne            = "INR";
			String expdefCurrecnyTwo            = "INR";
			
			System.out.println("********************************checkExchangeRateHistorySelectINR***********************************");
			
			
			
			System.out.println("Effective Date	: "+ actEffectiveDate + "Expected		 : "+ expEffectiveDate);
			System.out.println("valueUSD        : "+actValueUSD       + "Value Excepted  : "+expValueUSD);
			System.out.println("ValueAED        : "+actValueAED       + "Value Excepted  : "+expValueAED);
			System.out.println("selCurrencyOne  : "+actselCurrencyOne + "Value Excepted  : "+expselCurrencyOne);
			System.out.println("selCurrecnyTwo  : "+actselCurrecnyTwo + "Value Excepted  : "+expselCurrecnyTwo);
			System.out.println("defCurrecnyOne  : "+actdefCurrecnyOne + "Value Excepted  : "+expdefCurrecnyOne);
			System.out.println("defCurrecnyTwo  : "+actdefCurrecnyTwo + "Value Excepted  : "+expdefCurrecnyTwo);
			
			
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
		
		Thread.sleep(5500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyInExchangeRateHistory));
		clearValueFromTextBox(baseCurrencyInExchangeRateHistory);
		baseCurrencyInExchangeRateHistory.sendKeys("AED");
		
		Thread.sleep(2000);
		
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectDateRange));
		Select s=new Select(selectDateRange);
		s.selectByValue("2");
		Thread.sleep(4000);
		selectDateRange.sendKeys(Keys.TAB);
		
		
		Thread.sleep(8000);
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
		
		System.out.println("Act From Date : "+ actFromDate);
		System.out.println("Exp From Date : "+ expFromDate);
		
		String actToDate=ToDateTxt.getAttribute("value");
		String expToDate=f.format(toDate);
		
		System.out.println("Act To Date	: "	+ actToDate);
		System.out.println("Exp To Date	: "	+ expToDate);
			
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
		
		String expEffectiveDate				= f.format(fromDate);
		String expValueUSD                  = "4.2800000000";
		String expValueAED                  = "0.0700000000";
		String expselCurrencyOne            = "USD";
		String expselCurrecnyTwo            = "INR";
		String expdefCurrecnyOne            = "AED";
		String expdefCurrecnyTwo            = "AED";
		
		System.out.println("********************************checkExchangeRateHistoryCurrentMonthSelectAED***********************************");
		
	
		System.out.println("Effective Date	 "	+ actEffectiveDate 	+"Expected		"	+ expEffectiveDate);
		System.out.println("valueUSD         "+actValueUSD       +"Value Excepted   :"+expValueUSD);
		System.out.println("ValueAED         "+actValueAED       +"Value Excepted   :"+expValueAED);
		System.out.println("selCurrencyOne   "+actselCurrencyOne +"Value Excepted   :"+expselCurrencyOne);
		System.out.println("selCurrecnyTwo   "+actselCurrecnyTwo +"Value Excepted   :"+expselCurrecnyTwo);
		System.out.println("defCurrecnyOne   "+actdefCurrecnyOne +"Value Excepted   :"+expdefCurrecnyOne);
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
	
	public boolean checkDateOptionasLastMonthinExchangeRateHistory() throws InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyInExchangeRateHistory));
		clearValueFromTextBox(baseCurrencyInExchangeRateHistory);
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
		
		
		String expEffectiveDate				= f.format(fromDate);
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
		
		Thread.sleep(4000);
		
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectDateRange));
		Select s=new Select(selectDateRange);
		s.selectByValue("4");
		Thread.sleep(4000);
		selectDateRange.sendKeys(Keys.TAB);
		Thread.sleep(3000);
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
		Thread.sleep(5000);
				
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
		
		String expEffectiveDate				= f.format(fromDate);
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
		
		Thread.sleep(6000);
		
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
		
		String expEffectiveDate				= f.format(fromDate);
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
		
		Thread.sleep(6000);
		
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
		
		String expEffectiveDate				= df.format(cal.getTime());
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
		
		String expEffectiveDate				= f.format(fromDate);
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
		
		String expEffectiveDate				= f.format(fromDate);
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
	


	public static boolean CheckLogin() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	        Thread.sleep(3000);

	        getDriver().navigate().refresh();

	        LoginPage lp=new LoginPage(getDriver()); 

	        lp.checkLoginPageTitleByURLInputInBrowser();

	        String unamelt="su";

	        String pawslt="su";

	        lp.enterUserName(unamelt);

	        Thread.sleep(2000);
	        
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

	        /*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	        companyLogo.click();*/

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


	public boolean checkLogoutExchangeRatePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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






	public ExchangeRateAndExchangeRateHistoryPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
