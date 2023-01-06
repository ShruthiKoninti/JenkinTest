package com.focus.Pages;

import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

public class CurrencyPage extends BaseEngine 
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "CurrencyPage";	
	
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
	
	
		
		
		
	
	public static boolean checkOpenCurrencyMenu() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		     
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMenu));
		currencyMenu.click();
	         
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMasterMenu));
		currencyMasterMenu.click();
		    	  
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveIcon));
		
        String expected = excelReader.getCellData(xlSheetName, 8, 6);
    	
    	String actual = Boolean.toString(saveIcon.isDisplayed());
    	
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
	
	
	
	
	
	
	
	
	public static boolean checkUpdatingISOCurrencyCodeSLLCoinName() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ISOCurrencyCodeTxt));
		ISOCurrencyCodeTxt.click();
		ISOCurrencyCodeTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		
		Thread.sleep(2000);
		
		ISOCurrencyCodeTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(coinsNameTxt));
		coinsNameTxt.click();
		coinsNameTxt.clear();
		coinsNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		
		getAction().moveToElement(coinsNameTxt).sendKeys(Keys.TAB).perform();	
		
		Thread.sleep(2000);
				    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveIcon));
		saveIcon.click();
		Thread.sleep(2000);
	  
		System.out.println("*********************checkUpdatingISOCurrencyCodeSLLCoinName******************************************");
		
		String expMessage=excelReader.getCellData(xlSheetName, 9, 6);
		
		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, actMessage);

		
		Thread.sleep(2000);
		
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
	
	
	
	
	
	
	
	
	public static boolean checkUpdatingAgainISOCurrencyCodeSLLCoinName() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ISOCurrencyCodeTxt));
		ISOCurrencyCodeTxt.click();
		ISOCurrencyCodeTxt.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		
		Thread.sleep(2000);
		
		ISOCurrencyCodeTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(coinsNameTxt));
		coinsNameTxt.click();
		coinsNameTxt.clear();
		coinsNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		
		getAction().moveToElement(coinsNameTxt).sendKeys(Keys.TAB).perform();	
		
		Thread.sleep(2000);
				    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveIcon));
		saveIcon.click();
		Thread.sleep(2000);
	  
		System.out.println("*********************checkUpdatingAgainISOCurrencyCodeSLLCoinName******************************************");
		
		String expMessage=excelReader.getCellData(xlSheetName, 11, 6);
		
		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 11, 7, actMessage);
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMasterCancelBtn));
			currencyMasterCancelBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 11, 8, resPass);
			return true;
		}	 
		else
		{    
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMasterCancelBtn));
			currencyMasterCancelBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 11, 8, resFail);

			return false;
		}
     }
	
	
	
	
	public CurrencyPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}
	
	
	
}
