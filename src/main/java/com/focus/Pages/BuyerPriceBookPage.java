package com.focus.Pages;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
import com.testautomationguru.utility.PDFUtil;

public class BuyerPriceBookPage extends BaseEngine 
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "BuyerPriceBookPage";
	
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


	
				 
				 

		public boolean checkSavingBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		  	homeMenu.click();
		  	
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		  	mastersMenu.click();
		  
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	      	homeMasterItemMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
		  	homeMasterItem_buyerPriceBookMenu.click();
		  	
		  	Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			buyerPriceBook_PriceBook.sendKeys(excelReader.getCellData(xlSheetName, 8, 5));
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
			itemFirstCellSelect.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
			itemValue.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
			
			Thread.sleep(2000);
			   
			itemValue.sendKeys(Keys.TAB);
			
			if (getIsAlertPresent()) 
			{
				getAlert().accept();
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
			rateValue.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
			rateValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
			buyerPriceBookSaveBtn.click();
			
			System.out.println("*********************************checkSavingBuyerPriceBook********************************");
			
			String expMessage=excelReader.getCellData(xlSheetName, 8, 6);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, xlSheetName, 8, 7, actMessage);

		
		   if (actMessage.equalsIgnoreCase(expMessage)) 
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
		
		
	
		
		
		
		
		public boolean checkUpdatingBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			buyerPriceBook_PriceBook.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
			Thread.sleep(2000);
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
			priceBookFilterAndLoadBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
			String actItem = itemFirstCellSelect.getText();
			String expItem = excelReader.getCellData(xlSheetName, 11, 6);
			
			excelReader.setCellData(xlfile, xlSheetName, 11, 7, actItem);

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateFirstCellSelect));
			String actRate = rateFirstCellSelect.getText();
			String expRate = excelReader.getCellData(xlSheetName, 12, 6);
			
			excelReader.setCellData(xlfile, xlSheetName, 12, 7, actRate);
			
			rateFirstCellSelect.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
			rateValue.clear();
			rateValue.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
			rateValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
			buyerPriceBookSaveBtn.click();
			Thread.sleep(2000);
			
			System.out.println("*********************************checkUpdatingBuyerPriceBook********************************");
			
			String expMessage = excelReader.getCellData(xlSheetName, 13, 6);
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, xlSheetName, 13, 7, actMessage);
			
			
			System.out.println("Item : "+actItem+"  Value Expected  "+expItem);
			System.out.println("Rate : "+actRate+"  Value Expected  "+expRate);
			
		   if (actItem.equalsIgnoreCase(expItem) && actRate.equalsIgnoreCase(expRate) && actMessage.equalsIgnoreCase(expMessage)) 
		   {
			   excelReader.setCellData(xlfile, xlSheetName, 11, 8, resPass);

			   
			   return true;
		   } 
		   else 
		   {
			   
			   excelReader.setCellData(xlfile, xlSheetName, 11, 8, resFail);

			   return false;
		   }
		}
		
		
		
		
		
		
		public boolean checkCopyPasteAndSavingBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			buyerPriceBook_PriceBook.sendKeys(excelReader.getCellData(xlSheetName, 14, 5));
			Thread.sleep(2000);
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
			priceBookFilterAndLoadBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCopyBtn));
			buyerPriceBookCopyBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
			deleteBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
			Thread.sleep(2000);
			
			String expDeleteMessage = excelReader.getCellData(xlSheetName, 14, 6);
			
			String actDeleteMessage=checkValidationMessage(expDeleteMessage);
			
			excelReader.setCellData(xlfile, xlSheetName, 14, 7, actDeleteMessage);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookClearBtn));
			buyerPriceBookClearBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			buyerPriceBook_PriceBook.sendKeys(excelReader.getCellData(xlSheetName, 15, 5));
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookPasteBtn));
			buyerPriceBookPasteBtn.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
			buyerPriceBookSaveBtn.click();
			Thread.sleep(2000);
			
			System.out.println("*********************************checkCopyPasteAndSavingBuyerPriceBook********************************");
			
			String expMessage = excelReader.getCellData(xlSheetName, 15, 6);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, xlSheetName, 15, 7, actMessage);
		
		   if (actDeleteMessage.equalsIgnoreCase(expDeleteMessage) && actMessage.equalsIgnoreCase(expMessage)) 
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
		
		
		
		
		
		
		
		public boolean checkDeletingBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			buyerPriceBook_PriceBook.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));
			Thread.sleep(2000);
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
			priceBookFilterAndLoadBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
			String actItem = itemFirstCellSelect.getText();
			String expItem = excelReader.getCellData(xlSheetName, 16, 6);
			
			excelReader.setCellData(xlfile, xlSheetName, 16, 7, actItem);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateFirstCellSelect));
			String actRate = rateFirstCellSelect.getText();
			String expRate = excelReader.getCellData(xlSheetName, 17, 6);
			
			excelReader.setCellData(xlfile, xlSheetName, 17, 7, actRate);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
			deleteBtn.click();
			Thread.sleep(1000);
			
			getWaitForAlert();
			Thread.sleep(2000);
			getAlert().accept();
			Thread.sleep(2000);
			
			System.out.println("*********************************checkDeletingBuyerPriceBook********************************");
			
			String expMessage = excelReader.getCellData(xlSheetName, 18, 6);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, xlSheetName, 18, 7, actMessage);
		
			System.out.println("Item : "+actItem+"  Value Expected  "+expItem);
			System.out.println("Rate : "+actRate+"  Value Expected  "+expRate);
			
		   if (actItem.equalsIgnoreCase(expItem) && actRate.equalsIgnoreCase(expRate) && actMessage.equalsIgnoreCase(expMessage)) 
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
		
		
		
		
		
		
		

		public boolean checkExportingBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			
			File Efile=new File(getBaseDir()+"\\autoIt\\ExportFiles\\TestSmokeBuyigRates.xlsx");
			
			if(Efile.exists())
			{
				Efile.delete();
			}
			
			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		  	homeMenu.click();
		  	
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		  	mastersMenu.click();
		  
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	      	homeMasterItemMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
		  	homeMasterItem_buyerPriceBookMenu.click();*/
		  	
		  	Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			buyerPriceBook_PriceBook.sendKeys(excelReader.getCellData(xlSheetName, 19, 5));
			Thread.sleep(1000);
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			JavascriptExecutor js = (JavascriptExecutor) getDriver(); 
			js.executeScript("arguments[0].scrollIntoView(true);", priceBookFilterAndLoadBtn);
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
			priceBookFilterAndLoadBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookExportToExcel));
			priceBookExportToExcel.click();
			Thread.sleep(2000);
			
			String expMessage=excelReader.getCellData(xlSheetName, 19, 6);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, xlSheetName, 19, 7, actMessage);
			
			
			
			Thread.sleep(2000);
			
			String FileName = checkDownloadedFileName(getDriver());
		 	
		 	ExcelReader excelReader = new ExcelReader(POJOUtility.getExcelPath());
		 	
		 	Workbook OrginalFile = new XSSFWorkbook(getBaseDir()+"\\autoIt\\ImportFiles\\TestSmokeBuyigRates.xlsx");
		 	Workbook Exportedfile= new XSSFWorkbook(getBaseDir()+"\\autoIt\\ExportFiles\\"+FileName);
		 	
		 	boolean result = excelReader.verifyDataInExcelBookAllSheets(OrginalFile, Exportedfile);
		 	
		 	System.out.println("*********************************checkExportingBuyerPriceBook********************************");
		 	
		 	System.out.println("Validation Message : "+actMessage+"  Value Expected  "+expMessage);
		 	
		 	System.out.println("Excel Comparing : "+result);
		
		   if (actMessage.equalsIgnoreCase(expMessage) && result) 
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
		
		
		
		

		
		
		public boolean checkImportingBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
		{
			//excelReader=new ExcelReader(POJOUtility.getExcelPath());
			
			ExcelReader excelReader = new ExcelReader(POJOUtility.getExcelPath());
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		  	homeMenu.click();
		  	
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		  	mastersMenu.click();
		  
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	      	homeMasterItemMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
		  	homeMasterItem_buyerPriceBookMenu.click();
		  	
		  	Thread.sleep(2000);
					
			
			String excelfile = getBaseDir()+"\\autoIt\\ImportFiles\\TestSmokeBuyigRatesImport.xlsx";
			String sheet = "PriceBook";
	  	   
			Calendar cal=Calendar.getInstance();
			SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
			String presentDate = currentDate.format(cal.getTime());
			
			System.err.println("Present date : "+presentDate);
			
			excelReader.setCellData(excelfile, sheet, 1, 5, presentDate);
			excelReader.setCellData(excelfile, sheet, 1, 6, presentDate);
			
			
			excelReader.setCellData(excelfile, sheet, 2, 5, presentDate);
			excelReader.setCellData(excelfile, sheet, 2, 6, presentDate);
			
						
			excelReader.setCellData(excelfile, sheet, 1, 7, "0");
			excelReader.setCellData(excelfile, sheet, 2, 7, "0");
			
			
			Thread.sleep(3000);
		  	
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			buyerPriceBook_PriceBook.click();
			buyerPriceBook_PriceBook.sendKeys(Keys.END);
			Thread.sleep(2000);
			buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
			Thread.sleep(2000);
			//buyerPriceBook_PriceBook.sendKeys(excelReader.getCellData(xlSheetName, 20, 5));
			buyerPriceBook_PriceBook.sendKeys("TestSmokeBuyigRatesImport");
			Thread.sleep(2000);
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
			//buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
			
			Thread.sleep(1000);
		  	
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_ImportFromExcel));
		  	buyerPriceBook_ImportFromExcel.click();
		  	
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(browseButton));
			browseButton.click();
			
			Thread.sleep(2000);
			
			Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\FocusBuyerImportPrice.exe");
			
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetName));
			sheetName.click();
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetOkButton));
			sheetOkButton.click();
			Thread.sleep(2000);
			
			
			getWaitForAlert();
			Thread.sleep(2000);
			
			//String expAlertMessage =excelReader.getCellData(xlSheetName, 20, 6);
			String expAlertMessage="Do you want to load fields mapping?";
			String actAlertMessage =getAlert().getText();
			
			excelReader.setCellData(xlfile, xlSheetName, 20, 7, actAlertMessage);
			
			getAlert().accept();
			
			Thread.sleep(3000);
			
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetItemHeading));
			sheetItemHeading.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_MappingDrpdwn));
			Select pricebook = new Select(buyerPriceBook_MappingDrpdwn);
			pricebook.selectByVisibleText("TestSmokeBuyigRatesImport");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetTagHeading));
			sheetTagHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_MappingDrpdwn));
			pricebook.selectByVisibleText("Vendor");
			Thread.sleep(2000);
			
			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetTagHeading));
			sheetTagHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetTag));
			sheetTag.click();
			sheetTag.sendKeys("Department");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetCurrencyHeading));
			sheetCurrencyHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetCurrency));
			sheetCurrency.click();
			sheetCurrency.sendKeys("Currency");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetStartDateHeading));
			sheetStartDateHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetStartDate));
			sheetStartDate.click();
			sheetStartDate.sendKeys("StartDate");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetEndDateHeading));
			sheetEndDateHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetEndDate));
			sheetEndDate.click();
			sheetEndDate.sendKeys("EndDate");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetMinQtyHeading));
			sheetMinQtyHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetMinQty));
			sheetMinQty.click();
			sheetMinQty.sendKeys("MinQty");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetMaxQtyHeading));
			sheetMaxQtyHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetMaxQty));
			sheetMaxQty.click();
			sheetMaxQty.sendKeys("MaxQty");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetUnitHeading));
			sheetUnitHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetUnitQty));
			sheetUnitQty.click();
			sheetUnitQty.sendKeys("UnitName");
			*/
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetRateHeading));
			sheetRateHeading.click();
			
			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetRate));
			sheetRate.click();
			sheetRate.sendKeys("Rate");*/
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_MappingDrpdwn));
			pricebook.selectByVisibleText("Rate");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetPOHeading));
			sheetPOHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_MappingDrpdwn));
			pricebook.selectByVisibleText("Po");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetPTHeading));
			sheetPTHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_MappingDrpdwn));
			pricebook.selectByVisibleText("Pt");
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importButton));
		  	importButton.click();
		  	Thread.sleep(3000);
		  	
		  	String expMessage=excelReader.getCellData(xlSheetName, 21, 6);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, xlSheetName, 21, 7, actMessage);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			buyerPriceBook_PriceBook.click();
			buyerPriceBook_PriceBook.sendKeys(Keys.END);
			buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
			buyerPriceBook_PriceBook.sendKeys(excelReader.getCellData(xlSheetName, 21, 5));
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
			Thread.sleep(3000);
			
			JavascriptExecutor js = (JavascriptExecutor) getDriver(); 
			js.executeScript("arguments[0].scrollIntoView(true);", priceBookFilterAndLoadBtn);
			Thread.sleep(2000);

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
			priceBookFilterAndLoadBtn.click();
			
			Thread.sleep(3000);
			
			js.executeScript("arguments[0].scrollIntoView(true);", buyerPriceBook_CustamizeSettingBtn);
			Thread.sleep(2000);
			
			getAction().moveToElement(buyerPriceBook_CustamizeSettingBtn).build().perform();

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
			buyerPriceBook_CustamizeSettingBtn.click();
			
		 	Thread.sleep(1000);
		 	
		 	int actCount = buyerPriceBook_CustomizeCheckBoxList.size();
			int expCount = 6;
			
			for (int i = 0; i < actCount; i++) 
			{
				if (buyerPriceBook_CustomizeCheckBoxList.get(i).isSelected()==false) 
				{
					buyerPriceBook_CustomizeCheckBoxListSelected.get(i).click();
				}
			}
			
			Thread.sleep(3000);
		 	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
			buyerPriceBook_CustamizeSettingBtn.click();
			
			int count = buyerPriceBook_TableRow1List.size();
			
			ArrayList<String> pricebookRow1List = new ArrayList<String>();
			
			for (int i = 0; i < count; i++) 
			{
				String data = buyerPriceBook_TableRow1List.get(i).getText();
				
				pricebookRow1List.add(data);
			}
			
			String actPriceBookRow1 = pricebookRow1List.toString();
			String expPriceBookRow1 ="[1, ITEMS GROUP, , AMERICA, "+presentDate+", "+presentDate+", , Indian Rupees, 0, 0, , 200.000000, 100.000000, 50.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";
			
			System.out.println(actPriceBookRow1);
			System.out.println(expPriceBookRow1);
			
			
			ArrayList<String> pricebookRow2List = new ArrayList<String>();
			
			for (int i = 0; i < count; i++) 
			{
				String data = buyerPriceBook_TableRow2List.get(i).getText();
				
				pricebookRow2List.add(data);
			}
			
			String actPriceBookRow2 = pricebookRow2List.toString();
			String expPriceBookRow2 ="[2, BR COGS ITEM, , AMERICA, "+presentDate+", "+presentDate+", , Indian Rupees, 0, 0, Dozs, 300.000000, 200.000000, 100.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";
			
			System.out.println(actPriceBookRow2);
			System.out.println(expPriceBookRow2);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
			buyerPriceBookSaveBtn.click();
			
			String expUpdatingMessage = "Price Book updated successfully";
			String actUpdatingMessage=checkValidationMessage(expMessage);
			
		 	
		 	System.out.println("*********************************checkImportingBuyerPriceBook********************************");
		 	
		 	System.out.println("AlertMessage  : "+actAlertMessage+" Value Expected : "+expAlertMessage);
		 	System.out.println("Validation Message : "+actMessage+"  Value Expected  "+expMessage);
		 	System.out.println("Customization Checkbox Count : "+actCount+" Value Expected "+expCount);
		 	System.out.println("Updating Validation Message : "+actUpdatingMessage+"  Value Expected  "+expUpdatingMessage);
		 	
		
		   if (actAlertMessage.equalsIgnoreCase(expAlertMessage) && actMessage.equalsIgnoreCase(expMessage) && actCount==expCount
				  && actPriceBookRow1.equalsIgnoreCase(expPriceBookRow1) && actPriceBookRow2.equalsIgnoreCase(expPriceBookRow2)
				  && actUpdatingMessage.equalsIgnoreCase(expUpdatingMessage)) 
		   {
			   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCloseBtn));
			   buyerPriceBookCloseBtn.click();
			   Thread.sleep(2000);
			   return true;
		   } 
		   else 
		   {
			   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCloseBtn));
			   buyerPriceBookCloseBtn.click();
			   Thread.sleep(2000); 
			   return false;
		   }
		}
		
		
	////////////////////////////////////////////////////////
		
		
		@FindBy(xpath="//*[@id='btnSheetOkAID']")
		private static WebElement sheetAdvOkButton;
		
		
		@FindBy(xpath="//*[@id='AIDImportGrid_col_1-2']")
		private static WebElement sheetAdvItemHeading;
		
		@FindBy(xpath="//select[@id='AIDImportGrid_ExternalFields']")
		private static WebElement AdvbuyerPriceBook_MappingDrpdwn;
				
		@FindBy(xpath="//*[@id='AIDImportGrid_col_3-2']")
		private static WebElement sheetAdvAccountHeading;
		
		@FindBy(xpath="//*[@id='AIDImportGrid_col_12-2']")
		private static WebElement sheetAdvPOHeading;
		
		
		@FindBy(xpath="//*[@id='AIDImportGrid_col_13-2']")
		private static WebElement sheetAdvPTHeading;
		
		
		@FindBy(xpath="//*[@id='AdvanceImportModalDiv']//div[3]/div/input[1]")
		private static WebElement AdvImportDataBtn;
		
		
		@FindBy(xpath="//*[@id='btnPBPrintRateHistory']")
		private static WebElement buyerPriceBook_PrintRateHistory;
		
		@FindBy(xpath="//*[@id='PRHProduct']")
		private static WebElement PRHItemTxt;
		
		
		@FindBy(xpath="//*[@title='Create layout']/i")
		private static WebElement PRHNewLayoutBtn;
		
		@FindBy(xpath="(//*[@id='drpdownSlider']/a/i)[1]")
		private static WebElement PRHNewLayout_NewBtn;
		
		@FindBy(xpath="(//*[@id='SliderMenu']/a[1])[1]")
		private static WebElement PRHNewLayout_NewLayout;
		
		@FindBy(xpath="//*[@id='PRHLayout']")
		private static WebElement PRHNewLayoutTxt;
		
		@FindBy(xpath="//*[@id='PreviewMenu']")
		private static WebElement PRHPreviewTab;
		
		@FindBy(xpath="//*[@id='DesignMenu']")
		private static WebElement PRHDesignTab;
		
		
		@FindBy(xpath="(//*[@class='icon-save hiconright2'])[2]")
		private static WebElement PRHSaveIcon;
		
		@FindBy(xpath="//*[@onclick='INVOICEDESIGN.Save()']")
		private static WebElement PRHSave;
		
		@FindBy(xpath="//*[@id='id_PopuplayoutName']")
		private static WebElement PRHSaveTxt;
		
		@FindBy(xpath="//*[@id='btn_PopUpOk']")
		private static WebElement PRHSaveBtn;
		
		@FindBy(xpath="//*[@id='id_InvoiceDesingCancel']")
		private static WebElement PRHExitBtn;
		
		@FindBy(xpath="//*[@id='btnPrintPreview']")
		private static WebElement PRHPreviewBtn;
		
		@FindBy(xpath="//*[@id='btnPrintPrh']")
		private static WebElement PRHPrintBtn;
		
		
		@FindBy(xpath="//*[@id='btnClearPrh']")
		private static WebElement PRHClearBtn;
		
		
		@FindBy(xpath="//*[@id='btnClosePrh']")
		private static WebElement PRHCloseBtn;
		
		
		@FindBy(xpath="//*[@id='btnPBExportToXML']")
		private static WebElement PricebookExporttoXML;
		
		@FindBy(xpath="//*[@id='formLoadDataXML']/span")
		private static WebElement PricebookImportFromXML;
		
		
		@FindBy(xpath="//*[@id='btnAdvExcelImport']")
		private static WebElement buyerPriceBook_AdvExcelImport;
		
		@FindBy(xpath="//*[@id='btnImportAID']")
		private static WebElement AdvbrowseButton;
		
		
		public boolean checkAdvExcelImportingBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
		{
			ExcelReader excelReader = new ExcelReader(POJOUtility.getExcelPath());
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		  	homeMenu.click();
		  	
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		  	mastersMenu.click();
		  
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	      	homeMasterItemMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
		  	homeMasterItem_buyerPriceBookMenu.click();
		  	
		  	Thread.sleep(2000);
			
			
			String excelfile = getBaseDir()+"\\autoIt\\ImportFiles\\TestSmokeBuyigRatesImport.xlsx";
			String sheet = "PriceBook";
	  	   
			Calendar cal=Calendar.getInstance();
			SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
			String presentDate = currentDate.format(cal.getTime());
			
			cal.add(Calendar.DATE, 5); 
			
			String endDate = currentDate.format(cal.getTime());
			
			System.err.println("Present date : "+presentDate);
			System.err.println("End date     : "+endDate);
			
			excelReader.setCellData(excelfile, sheet, 1, 5, presentDate);
			excelReader.setCellData(excelfile, sheet, 1, 7, endDate);
			
			excelReader.setCellData(excelfile, sheet, 2, 5, presentDate);
			excelReader.setCellData(excelfile, sheet, 2, 7, endDate);
			
			Thread.sleep(3000);
		  	
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			buyerPriceBook_PriceBook.click();
			buyerPriceBook_PriceBook.sendKeys(Keys.END);
			Thread.sleep(2000);
			buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
			Thread.sleep(2000);
			buyerPriceBook_PriceBook.sendKeys("TestSmokeBuyigRatesAdvExcelImport");
			Thread.sleep(2000);
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
			//buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
			
			Thread.sleep(3000);
		  	
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AdvExcelImport));
		  	buyerPriceBook_AdvExcelImport.click();
		  	
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvbrowseButton));
			AdvbrowseButton.click();
			
			Thread.sleep(2000);
			
			Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\FocusBuyerImportPrice.exe");
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetName));
			sheetName.click();
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAdvOkButton));
			sheetAdvOkButton.click();
			
			getWaitForAlert();
			
			String expAlertMessage ="Do you want to load fields mapping?";
			
			String actAlertMessage =getAlert().getText();
			
			getAlert().accept();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAdvItemHeading));
			sheetAdvItemHeading.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvbuyerPriceBook_MappingDrpdwn));
			AdvbuyerPriceBook_MappingDrpdwn.click();
			Thread.sleep(2000);
			Select pricebook = new Select(AdvbuyerPriceBook_MappingDrpdwn);
			pricebook.selectByVisibleText("TestSmokeBuyigRatesAdvExcelImport");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAdvAccountHeading));
			sheetAdvAccountHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvbuyerPriceBook_MappingDrpdwn));
			pricebook.selectByVisibleText("Vendor");
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAdvPOHeading));
			sheetAdvPOHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvbuyerPriceBook_MappingDrpdwn));
			pricebook.selectByVisibleText("Po");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAdvPTHeading));
			sheetAdvPTHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvbuyerPriceBook_MappingDrpdwn));
			pricebook.selectByVisibleText("Pt");
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvImportDataBtn));
			AdvImportDataBtn.click();
		  	
		  	String expMessage="Price book imported successfully";
			
			String actMessage=checkValidationMessage(expMessage);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			buyerPriceBook_PriceBook.click();
			buyerPriceBook_PriceBook.sendKeys(Keys.END);
			buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
			buyerPriceBook_PriceBook.sendKeys("TestSmokeBuyigRatesAdvExcelImport");
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
			priceBookFilterAndLoadBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
			buyerPriceBook_CustamizeSettingBtn.click();
			
		 	Thread.sleep(1000);
		 	
		 	int actCount = buyerPriceBook_CustomizeCheckBoxListSelected.size();
			int expCount = 6;
			
			for (int i = 0; i < actCount; i++) 
			{
				if (buyerPriceBook_CustomizeCheckBoxListSelected.get(i).isSelected()==false) 
				{
					buyerPriceBook_CustomizeCheckBoxListSelected.get(i).click();
				}
			}
			
			Thread.sleep(2000);
		 	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
			buyerPriceBook_CustamizeSettingBtn.click();
			
			int count = buyerPriceBook_TableRow1List.size();
			
			ArrayList<String> pricebookRow1List = new ArrayList<String>();
			
			for (int i = 0; i < count; i++) 
			{
				String data = buyerPriceBook_TableRow1List.get(i).getText();
				
				pricebookRow1List.add(data);
			}
			
			String actPriceBookRow1 = pricebookRow1List.toString();
			String expPriceBookRow1 ="[1, ITEMS GROUP, , AMERICA, "+presentDate+", "+presentDate+", , Indian Rupees, 0, 0, , 200.000000, 100.000000, 50.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";
			
			System.out.println(actPriceBookRow1);
			System.out.println(expPriceBookRow1);
			
			
			ArrayList<String> pricebookRow2List = new ArrayList<String>();
			
			for (int i = 0; i < count; i++) 
			{
				String data = buyerPriceBook_TableRow2List.get(i).getText();
				
				pricebookRow2List.add(data);
			}
			
			String actPriceBookRow2 = pricebookRow2List.toString();
			String expPriceBookRow2 ="[2, BR COGS ITEM, , AMERICA, "+presentDate+", "+presentDate+", , Indian Rupees, 0, 0, Dozs, 300.000000, 200.000000, 100.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";
			
			System.out.println(actPriceBookRow2);
			System.out.println(expPriceBookRow2);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
			buyerPriceBookSaveBtn.click();
			
			String expUpdatingMessage = "Price Book updated successfully";
			String actUpdatingMessage=checkValidationMessage(expMessage);
			
		 	
		 	System.out.println("*********************************checkImportingBuyerPriceBook********************************");
		 	
		 	System.out.println("AlertMessage  : "+actAlertMessage+" Value Expected : "+expAlertMessage);
		 	System.out.println("Validation Message : "+actMessage+"  Value Expected  "+expMessage);
		 	System.out.println("Customization Checkbox Count : "+actCount+" Value Expected "+expCount);
		 	System.out.println("Updating Validation Message : "+actUpdatingMessage+"  Value Expected  "+expUpdatingMessage);
		 	
		
		   if (actAlertMessage.equalsIgnoreCase(expAlertMessage) && actMessage.equalsIgnoreCase(expMessage) && actCount==expCount
				  && actPriceBookRow1.equalsIgnoreCase(expPriceBookRow1) && actPriceBookRow2.equalsIgnoreCase(expPriceBookRow2)
				  && actUpdatingMessage.equalsIgnoreCase(expUpdatingMessage)) 
		   {
			   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCloseBtn));
			   buyerPriceBookCloseBtn.click();
			   Thread.sleep(2000);
			   return true;
		   } 
		   else 
		   {
			   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCloseBtn));
			   buyerPriceBookCloseBtn.click();
			   Thread.sleep(2000); 
			   return false;
		   }
		}
		
		
		public boolean checkExportingXMLBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
		{
			File Efile=new File(getBaseDir()+"\\autoIt\\ExportFiles\\TestSmokeBuyigRates.xml");
			
			if(Efile.exists())
			{
				Efile.delete();
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		  	homeMenu.click();
		  	
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		  	mastersMenu.click();
		  
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	      	homeMasterItemMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
		  	homeMasterItem_buyerPriceBookMenu.click();
		  	
		  	Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			buyerPriceBook_PriceBook.sendKeys("TestSmokeBuyigRates");
			Thread.sleep(1000);
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
			priceBookFilterAndLoadBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PricebookExporttoXML));
			PricebookExporttoXML.click();
			
			String expMessage="Pricebook exported successfully";
			
			String actMessage=checkValidationMessage(expMessage);
			
			
			
			Thread.sleep(6000);
			
			Robot robot = new Robot();
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			
			Thread.sleep(3000);
			
			ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
			
			int actOpenWindowsCount = getDriver().getWindowHandles().size();
			
			System.err.println("No of tabs : "+actOpenWindowsCount);
			
			getDriver().switchTo().window(openTabs.get(1)).close();
			Thread.sleep(2000);
			
			getDriver().switchTo().window(openTabs.get(0));
			Thread.sleep(1000);
			
		 
		 
		 	
		 	
		 	
		 
		 	
		 	System.out.println("*********************************checkExportingXMLBuyerPriceBook********************************");
		 	
		 	System.out.println("Validation Message : "+actMessage+"  Value Expected  "+expMessage);
		 	
		 	
		
		   if (actMessage.equalsIgnoreCase(expMessage)) 
		   {
			   return true;
		   } 
		   else 
		   {
			   return false;
		   }
		}
		
			
		
	////
		public boolean checkImportingXMLBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
		{
			ExcelReader excelReader = new ExcelReader(POJOUtility.getExcelPath());
			
			File Efile=new File(getBaseDir()+"\\autoIt\\ExportFiles\\TestSmokeBuyingRatesImport.xml");
			
			if(Efile.exists())
			{
				Efile.delete();
			}
			
			Thread.sleep(2000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		  	homeMenu.click();
		  	
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		  	mastersMenu.click();
		  
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	      	homeMasterItemMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
		  	homeMasterItem_buyerPriceBookMenu.click();
		  	
		  	Thread.sleep(2000);
			
			
			/*String excelfile = getBaseDir()+"\\autoIt\\ImportFiles\\TestSmokeBuyigRatesImport.xlsx";
			String sheet = "PriceBook";
	  	   
			Calendar cal=Calendar.getInstance();
			SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
			String presentDate = currentDate.format(cal.getTime());
			
			cal.add(Calendar.DATE, 5); 
			
			String endDate = currentDate.format(cal.getTime());
			
			System.err.println("Present date : "+presentDate);
			System.err.println("End date     : "+endDate);
			
			excelReader.setCellData(excelfile, sheet, 1, 5, presentDate);
			excelReader.setCellData(excelfile, sheet, 1, 6, endDate);
			
			excelReader.setCellData(excelfile, sheet, 2, 5, presentDate);
			excelReader.setCellData(excelfile, sheet, 2, 6, endDate);
			
			Thread.sleep(3000);*/
		  	
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			buyerPriceBook_PriceBook.click();
			buyerPriceBook_PriceBook.sendKeys(Keys.END);
			Thread.sleep(2000);
			buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
			Thread.sleep(2000);
			buyerPriceBook_PriceBook.sendKeys("TestSmokeBuyigRatesXMLImport");
			Thread.sleep(2000);
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
			//buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
			
			Thread.sleep(1000);
		  	
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PricebookImportFromXML));
		  	PricebookImportFromXML.click();
		  	
		  /*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvbrowseButton));
			AdvbrowseButton.click();
			
			*/Thread.sleep(2000);
			
			Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\FocusBuyerImportPriceXML.exe");
			
			Thread.sleep(3000);
			
			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetName));
			sheetName.click();
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAdvOkButton));
			sheetAdvOkButton.click();
			
			getWaitForAlert();
			
			String expAlertMessage ="Do you want to load fields mapping?";
			
			String actAlertMessage =getAlert().getText();
			
			getAlert().accept();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAdvItemHeading));
			sheetAdvItemHeading.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvbuyerPriceBook_MappingDrpdwn));
			AdvbuyerPriceBook_MappingDrpdwn.click();
			Thread.sleep(2000);
			Select pricebook = new Select(AdvbuyerPriceBook_MappingDrpdwn);
			pricebook.selectByVisibleText("TestSmokeBuyigRatesAdvExcelImport");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAdvAccountHeading));
			sheetAdvAccountHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvbuyerPriceBook_MappingDrpdwn));
			pricebook.selectByVisibleText("Vendor");
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAdvPOHeading));
			sheetAdvPOHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvbuyerPriceBook_MappingDrpdwn));
			pricebook.selectByVisibleText("Po");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAdvPTHeading));
			sheetAdvPTHeading.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvbuyerPriceBook_MappingDrpdwn));
			pricebook.selectByVisibleText("Pt");
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvImportDataBtn));
			AdvImportDataBtn.click();*/
		  	
		  	String expMessage="Price book imported successfully";
			
			String actMessage=checkValidationMessage(expMessage);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			buyerPriceBook_PriceBook.click();
			buyerPriceBook_PriceBook.sendKeys(Keys.END);
			buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
			buyerPriceBook_PriceBook.sendKeys("TestSmokeBuyigRatesXMLImport");
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
			Thread.sleep(3000);
			
			getAction().moveToElement(priceBookFilterAndLoadBtn).build().perform();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
			priceBookFilterAndLoadBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
			buyerPriceBook_CustamizeSettingBtn.click();
			
		 	Thread.sleep(1000);
		 	
		 	int actCount = buyerPriceBook_CustomizeCheckBoxList.size();
			int expCount = 6;
			
			for (int i = 0; i < actCount; i++) 
			{
				if (buyerPriceBook_CustomizeCheckBoxList.get(i).isSelected()==false) 
				{
					buyerPriceBook_CustomizeCheckBoxList.get(i).click();
				}
			}
			
			Thread.sleep(2000);
		 	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
			buyerPriceBook_CustamizeSettingBtn.click();
			
			int count = buyerPriceBook_TableRow1List.size();
			
			ArrayList<String> pricebookRow1List = new ArrayList<String>();
			
			for (int i = 0; i < count; i++) 
			{
				String data = buyerPriceBook_TableRow1List.get(i).getText();
				
				pricebookRow1List.add(data);
			}
			
			String actPriceBookRow1 = pricebookRow1List.toString();
			String expPriceBookRow1 ="[1, BR COGS ITEM, , AMERICA, 09/10/2020, , , Indian Rupees, 0, 0, Dozs, 200.000000, 100.000000, 50.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";
			
			System.out.println(actPriceBookRow1);
			System.out.println(expPriceBookRow1);
			
			
		/*	ArrayList<String> pricebookRow2List = new ArrayList<String>();
			
			for (int i = 0; i < count; i++) 
			{
				String data = buyerPriceBook_TableRow2List.get(i).getText();
				
				pricebookRow2List.add(data);
			}
			
			String actPriceBookRow2 = pricebookRow2List.toString();
			String expPriceBookRow2 ="[2, BR COGS ITEM, , AMERICA, "+presentDate+", "+endDate+", , Indian Rupees, 0, 0, Dozs, 300.000000, 200.000000, 100.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";
			
			System.out.println(actPriceBookRow2);
			System.out.println(expPriceBookRow2);
			*/
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
			buyerPriceBookSaveBtn.click();
			
			String expUpdatingMessage = "Price Book updated successfully";
			String actUpdatingMessage=checkValidationMessage(expMessage);
			
		 	
		 	System.out.println("*********************************checkImportingBuyerPriceBook********************************");
		 	
		 	//System.out.println("AlertMessage  : "+actAlertMessage+" Value Expected : "+expAlertMessage);
		 	System.out.println("Validation Message : "+actMessage+"  Value Expected  "+expMessage);
		 	System.out.println("Customization Checkbox Count : "+actCount+" Value Expected "+expCount);
		 	System.out.println("Updating Validation Message : "+actUpdatingMessage+"  Value Expected  "+expUpdatingMessage);
		 	
		
		   if (/*actAlertMessage.equalsIgnoreCase(expAlertMessage) &&*/ actMessage.equalsIgnoreCase(expMessage) && actCount==expCount
				  && actPriceBookRow1.equalsIgnoreCase(expPriceBookRow1) /*&& actPriceBookRow2.equalsIgnoreCase(expPriceBookRow2)*/
				  && actUpdatingMessage.equalsIgnoreCase(expUpdatingMessage)) 
		   {
			   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCloseBtn));
			   buyerPriceBookCloseBtn.click();
			   Thread.sleep(2000);
			   return true;
		   } 
		   else 
		   {
			   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCloseBtn));
			   buyerPriceBookCloseBtn.click();
			   Thread.sleep(2000); 
			   return false;
		   }
		}
		
		
		String FileName2 = null;
		
		public boolean checkPrintRateHistoryBuyerPriceBook() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException, AWTException
		{
		
			
			File Efile=new File(getBaseDir()+"\\autoIt\\ExportFiles\\BuyerPriceBookToPDF.pdf");
			
			if(Efile.exists())
			{
				Efile.delete();
			}

			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	  	homeMenu.click();
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	  	mastersMenu.click();
	  
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
      	homeMasterItemMenu.click();
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
	  	homeMasterItem_buyerPriceBookMenu.click();
	  	
	  	Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.sendKeys("TestSmokeBuyigRates");
		Thread.sleep(1000);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PrintRateHistory));
		buyerPriceBook_PrintRateHistory.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHItemTxt));
		PRHItemTxt.click();
		PRHItemTxt.sendKeys("BR COGS ITEM" );
		Thread.sleep(2000);		
		PRHItemTxt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHNewLayoutBtn));
		PRHNewLayoutBtn.click();
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHNewLayout_NewBtn));
		PRHNewLayout_NewBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHNewLayout_NewLayout));
		PRHNewLayout_NewLayout.click();
		Thread.sleep(2000);
		
		
		
		
		Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\PrintRateHistoryNew.exe");
		
		Thread.sleep(95000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHPreviewTab));
		PRHPreviewTab.click();
		Thread.sleep(3000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHDesignTab));
		PRHDesignTab.click();
		Thread.sleep(3000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHSaveIcon));
		PRHSaveIcon.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHSave));
		PRHSave.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHSaveTxt));
		PRHSaveTxt.sendKeys("BuyerLayout");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHSaveBtn));
		PRHSaveBtn.click();
		Thread.sleep(2000);
		
		String expMessage="Layout Saved successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHExitBtn));
		PRHExitBtn.click();
		Thread.sleep(5000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHNewLayoutTxt));
		PRHNewLayoutTxt.click();
		PRHNewLayoutTxt.sendKeys("BuyerLayout");
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHPrintBtn));
		PRHPrintBtn.click();
		Thread.sleep(2000);
		
		/*
	 	
	 	
	 	
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
		
		Thread.sleep(8000);*/
		Robot robot = new Robot();
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_S);
		robot.keyRelease(KeyEvent.VK_S);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		
		Thread.sleep(5000);
			
		Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\SavingBuyerPriceBookPdf.exe");
		
		Thread.sleep(5000);
		
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_J);
		robot.keyRelease(KeyEvent.VK_J);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		
		Thread.sleep(2000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
			
		int actOpenWindowsCount = getDriver().getWindowHandles().size();
		int expOpenWindowsCount = 3;
		
		System.out.println("Number of Windows  : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
		
		Thread.sleep(1000);

	 	getDriver().switchTo().window(openTabs.get(2)).close();
	 	Thread.sleep(1000);
	 	getDriver().switchTo().window(openTabs.get(1)).close();
	 	Thread.sleep(1000);
	 	getDriver().switchTo().window(openTabs.get(0));

	 	Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("dd-MM-yyyy");
		String currentDate = df.format(cal.getTime());
		
		
		
	 	String actPDF = getBaseDir()+"\\autoIt\\ExportFiles\\BuyerPriceBookToPDF.pdf";
		String expPDF = getBaseDir()+"\\autoIt\\ImportFiles\\BuyerPriceBookToPDF.pdf";
		
		PDFUtil pdfutil = new PDFUtil();
		
		boolean result = pdfutil.compare(actPDF, expPDF);
		
		String actData = pdfutil.getText(actPDF);
		String expData = pdfutil.getText(expPDF).replaceAll("26-10-2022", currentDate);;
		
		System.err.println(actData);
		System.err.println(expData);
		
		System.out.println("Compared Result  : "+result);
		
		
	 	
	 	
		
	 	if(actMessage.equalsIgnoreCase(expMessage) && actData.equalsIgnoreCase(expData))
	 	{
		
		return true;
	 	}
	 	else
	 	{
	 		return false;
	 	}
		
		}
		
		
		@FindBy(xpath="(//*[@id='btnCancel'])[6]")
		private static WebElement previewTabClsBtn;
		
		public boolean checkPreviewandClearBtninPriceRateHistory() throws InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHPreviewBtn));
			PRHPreviewBtn.click();
			Thread.sleep(2000);
			
					
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previewTabClsBtn));
		 	previewTabClsBtn.click();
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHClearBtn));
		 	PRHClearBtn.click();
		 	Thread.sleep(1000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHItemTxt));
		 	boolean actTextAfterClear=PRHItemTxt.getText().isEmpty();
		 	boolean expTextAfterClear=true;
		 	
		 	
		 	if(actTextAfterClear==expTextAfterClear)
		 	{
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHCloseBtn));
		 		PRHCloseBtn.click();
		 		return true;
		 		
		 	}
		 	else
		 	{
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PRHCloseBtn));
		 		PRHCloseBtn.click();
		 		return false;
		 		
		 	}
		 	
			
		}
		
		
	@FindBy(xpath="//*[@id='btnViewPrice']")			 
	private static WebElement buyerPriceBook_ViewPrice;
	
	@FindBy(xpath="//*[@id='ctrlOptionDeptUPH2']")			 
	private static WebElement ViewPriceDepartmentTxt;
	
	@FindBy(xpath="//*[@id='ctrlOptionCustUPH2']")			 
	private static WebElement ViewPriceVendorTxt;
	
	@FindBy(xpath="//*[@id='ctrlOptionProUPH2']")			 
	private static WebElement ViewPriceItemTxt;
	
	@FindBy(xpath="//*[@id='ctrlOptionCurrUPH2']")			 
	private static WebElement ViewPriceCurrencyTxt;
	
	@FindBy(xpath="//*[@id='ctrlOptionUnitUPH2']")			 
	private static WebElement ViewPriceUnitsTxt;
	
	
	@FindBy(xpath="//*[@id='myModal2']//div[2]/div/input")			 
	private static WebElement ViewPriceLoadBtn;
	
	
	@FindBy(xpath="//*[@id='myModal2']//ul/li[1]/span")			 
	private static WebElement ViewPriceCustomizeBtn;
	
	
	@FindBy(xpath="//table[@id='WebGridUPT3']//th")			 
	private static List<WebElement> ViewPrice_HeaderList;
	
	@FindBy(xpath="//*[@id='chk_14']/following-sibling::span")			 
	private static WebElement ViewPrice_CustDepartmentCodeChkBox;
	
	@FindBy(xpath="//*[@id='chk_13']/following-sibling::span")			 
	private static WebElement ViewPrice_CustDepartmentCodeDeselectChkBox;
	
	@FindBy(xpath="//*[@id='WidthOfS']")			 
	private static WebElement ViewPrice_CustDepartmentCodeColWidth;
	
	@FindBy(xpath="//*[@id='btnOk']")			 
	private static WebElement ViewPrice_CustOKBtn;
	
	@FindBy(xpath="//*[@id='btnClearForView']")			 
	private static WebElement ViewPriceClearBtn;
	
	@FindBy(xpath="//*[@id='btnCloseForView']")			 
	private static WebElement ViewPriceCloseBtn;
	
	@FindBy(xpath="//tr[@id='WebGridUPT3_row_1']/td")			 
	private static List<WebElement> ViewPrice_FirstRowList;
	
	
	@FindBy(xpath="//tr[@id='WebGridUPT3_row_2']/td")			 
	private static List<WebElement> ViewPrice_SecondRowList;
	
	@FindBy(xpath="//tr[@id='WebGridUPT3_row_3']/td")			 
	private static List<WebElement> ViewPrice_ThirdRowList;
	
	
	public boolean checkViewPriceCustomizationBtnBuyerPriceBook() throws InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_ViewPrice));
		buyerPriceBook_ViewPrice.click();
		Thread.sleep(5000);
		
		ArrayList<String> BeforeHeaderList = new ArrayList<String>();
		
		for (int i = 2; i <15; i++) 
		{
			String data = ViewPrice_HeaderList.get(i).getText();
			
			BeforeHeaderList.add(data);
		}
		
		String actHeaderListBeforeCust = BeforeHeaderList.toString();
		String expHeaderListBeforeCust ="[Pricebook Name, Item Code, Item Name, Currency, Min Qty, Max Qty, Unit, Rate, Start Date, End Date, , , ]";
		
		System.out.println(actHeaderListBeforeCust);
		System.out.println(expHeaderListBeforeCust);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewPriceCustomizeBtn));
		ViewPriceCustomizeBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewPrice_CustDepartmentCodeChkBox));
		ViewPrice_CustDepartmentCodeChkBox.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewPrice_CustDepartmentCodeColWidth));
		ViewPrice_CustDepartmentCodeColWidth.click();
		ViewPrice_CustDepartmentCodeColWidth.sendKeys("150");
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewPrice_CustOKBtn));
		ViewPrice_CustOKBtn.click();
		Thread.sleep(2000);
		
		ArrayList<String> AfterHeaderList = new ArrayList<String>();
		
		for (int i = 2; i <16; i++) 
		{
			String data = ViewPrice_HeaderList.get(i).getText();
			
			AfterHeaderList.add(data);
		}
		
		String actHeaderListAfterCust = AfterHeaderList.toString();
		String expHeaderListAfterCust ="[Pricebook Name, Item Code, Item Name, Currency, Min Qty, Max Qty, Unit, Rate, Start Date, End Date, ModifiedBy, Created By, , Department]";
		
		System.out.println(actHeaderListAfterCust);
		System.out.println(expHeaderListAfterCust);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewPriceCustomizeBtn));
		ViewPriceCustomizeBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewPrice_CustDepartmentCodeDeselectChkBox));
		ViewPrice_CustDepartmentCodeDeselectChkBox.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewPrice_CustOKBtn));
		ViewPrice_CustOKBtn.click();
		Thread.sleep(2000);*/
		
		
		
		
		if(actHeaderListBeforeCust.equalsIgnoreCase(expHeaderListBeforeCust)&&
				actHeaderListAfterCust.equalsIgnoreCase(expHeaderListAfterCust))
		{
			return true;
		}
		else
		{
			return false;
		}
		
		
	}
	
	
	
	
	
	
	public boolean checkViewPriceBuyerPriceBook() throws InterruptedException 
	{
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_ViewPrice));
		buyerPriceBook_ViewPrice.click();*/
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewPriceDepartmentTxt));
		ViewPriceDepartmentTxt.click();
		ViewPriceDepartmentTxt.sendKeys("AMERICA");
		Thread.sleep(2000);
		ViewPriceDepartmentTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewPriceItemTxt));
		ViewPriceItemTxt.click();
		ViewPriceItemTxt.sendKeys("BR COGS ITEM");
		Thread.sleep(2000);
		ViewPriceItemTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewPriceCurrencyTxt));
		ViewPriceCurrencyTxt.click();
		ViewPriceCurrencyTxt.sendKeys("Indian Rupees");
		Thread.sleep(2000);
		ViewPriceCurrencyTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewPriceUnitsTxt));
		ViewPriceUnitsTxt.click();
		ViewPriceUnitsTxt.sendKeys("Dozs");
		Thread.sleep(2000);
		ViewPriceUnitsTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewPriceLoadBtn));
		ViewPriceLoadBtn.click();
		Thread.sleep(1000);
		
		
		
		int count = ViewPrice_FirstRowList.size();
		
		
		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		String presentDate = currentDate.format(cal.getTime());
		
		cal.add(Calendar.DATE, 5); 
		
		String endDate = currentDate.format(cal.getTime());
		
		ArrayList<String> pricebookRow1List = new ArrayList<String>();
		
		for (int i = 2; i < 12; i++) 
		{
			String data = ViewPrice_FirstRowList.get(i).getText();
			
			pricebookRow1List.add(data);
		}
		
		String actPriceBookRow1 = pricebookRow1List.toString();
		String expPriceBookRow1 ="[TestSmokeBuyigRatesImport, BR COGS ITEM, BR COGS ITEM, Indian Rupees, 0, 0, Dozs, 300.000000, "+presentDate+", "+presentDate+"]";
		
		System.out.println(actPriceBookRow1);
		System.out.println(expPriceBookRow1);
		
		
		ArrayList<String> pricebookRow2List = new ArrayList<String>();
		
		for (int i = 2; i < 11; i++) 
		{
			String data = ViewPrice_SecondRowList.get(i).getText();
			
			pricebookRow2List.add(data);
		}
		
		String actPriceBookRow2 = pricebookRow2List.toString();
		String expPriceBookRow2 ="[TestSmokeBuyigRatesXMLImport, BR COGS ITEM, BR COGS ITEM, Indian Rupees, 0, 0, Dozs, 200.000000, 09/10/2020]";
		
		System.out.println(actPriceBookRow2);
		System.out.println(expPriceBookRow2);
		
		ArrayList<String> pricebookRow3List = new ArrayList<String>();
		
		for (int i = 2; i < 11; i++) 
		{
			String data = ViewPrice_ThirdRowList.get(i).getText();
			
			pricebookRow3List.add(data);
		}
		
		String actPriceBookRow3 = pricebookRow3List.toString();
		String expPriceBookRow3 ="[TestSmokeBuyigRates, BR COGS ITEM, BR COGS ITEM, Indian Rupees, 0, 0, Dozs, 200.000000, 09/10/2020]";
		
		System.out.println(actPriceBookRow3);
		System.out.println(expPriceBookRow3);
		
		if(actPriceBookRow1.equalsIgnoreCase(expPriceBookRow1)  && actPriceBookRow2.equalsIgnoreCase(expPriceBookRow2)
			&&actPriceBookRow3.equalsIgnoreCase(expPriceBookRow3))
		{
		
		return true;
		}
		else
		{
			return false;
		}
		
		
		
	}
	
	
	@FindBy(xpath="(//*[@id='btnExporttoExcel'])[2]")
	private static WebElement ViewPrice_ExporttoExcel;
	
	
	public boolean checkExportingtoExcelinViewPriceBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
	{
		
		File Efile=new File(getBaseDir()+"\\autoIt\\ExportFiles\\PriceBook.xlsx");
		
	/*	if(Efile.exists())
		{
			Efile.delete();
		}
		
		Thread.sleep(3000);
		*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewPrice_ExporttoExcel));
		ViewPrice_ExporttoExcel.click();
		Thread.sleep(3000);
		
		String expMessage="Pricebook exported successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		
		
		Thread.sleep(2000);
		
		/*Robot robot = new Robot();
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_J);
		robot.keyRelease(KeyEvent.VK_J);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		
		Thread.sleep(2000);
		*/
		
		if(Efile.exists())
		{
			Efile.delete();
		}
		
		Thread.sleep(3000);
		
		

	   	
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
		

		
		Thread.sleep(2000);
		
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		
		Thread.sleep(5000);
		
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		
		robot.keyPress(KeyEvent.VK_F12);
		robot.keyRelease(KeyEvent.VK_F12);
		
		
		Thread.sleep(2000);
		
		
		/*robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		
		Thread.sleep(2000);
		
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);*/
		
		Thread.sleep(2000);
		Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\PriceBook.exe");
		
		Thread.sleep(15000);
		
		robot.keyPress(KeyEvent.VK_ENTER);
		robot.keyRelease(KeyEvent.VK_ENTER);
		Thread.sleep(3000);
		
		
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_F4);
		
		robot.keyRelease(KeyEvent.VK_ALT);
		robot.keyRelease(KeyEvent.VK_F4);
		Thread.sleep(2000);
		
		
		
		
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		int actOpenWindowsCount = getDriver().getWindowHandles().size();
		
		System.err.println("No of tabs : "+actOpenWindowsCount);
		
	 	getDriver().switchTo().window(openTabs.get(1)).close();
	 	
	 	Thread.sleep(1000);
	 	
	 	getDriver().switchTo().window(openTabs.get(0));
	 	
	 	/*ExcelReader excelReader = new ExcelReader(POJOUtility.getExcelPath());
	 	
	 	Workbook OrginalFile = new XSSFWorkbook(getBaseDir()+"\\autoIt\\ImportFiles\\PriceBook.xlsx");
	 	Workbook Exportedfile=new XSSFWorkbook(getBaseDir()+"\\autoIt\\ExportFiles\\PriceBook.xlsx");
	 	
	 	boolean result = excelReader.verifyDataInExcelBookAllSheets(OrginalFile, Exportedfile);
	 	*/
	 	 ExcelReader excelReader = new ExcelReader(POJOUtility.getExcelPath());
	     

	   	String actExcelfile = getBaseDir()+"\\autoIt\\ExportFiles\\PriceBook.xlsx";
	   	String expExcelfile = getBaseDir()+"\\autoIt\\ImportFiles\\PriceBook.xlsx";
	   	String sheet = "PriceBook";
	   	   
	   	
	   	
	   	FileInputStream fip1 = new FileInputStream(actExcelfile);
	   	Workbook workbook1  = WorkbookFactory.create(fip1);
	   	
	   	FileInputStream fip2 = new FileInputStream(expExcelfile);
	   	Workbook workbook2  = WorkbookFactory.create(fip2);
	   	
	   	boolean result = excelReader.verifyDataInExcelBookAllSheetsNew(workbook1, workbook2);
	   	
	   	fip1.close();
	   	fip2.close();
	   	
	   	
	   	System.err.println(result);
	   	
	 
	 	
	 	System.out.println("*********************************checkExportingBuyerPriceBook********************************");
	 	
	 	System.out.println("Validation Message : "+actMessage+"  Value Expected  "+expMessage);
	 	
	 	System.out.println("Excel Comparing : "+result);
	
	   if (actMessage.equalsIgnoreCase(expMessage) && result) 
	   {
		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewPriceCloseBtn));
		   ViewPriceCloseBtn.click();
		   return true;
	   } 
	   else 
	   {
		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewPriceCloseBtn));
		   ViewPriceCloseBtn.click();
		   return false;
	   }
	}
		
		
		@FindBy(xpath="//*[@id='btnUpdatePriceList']")
		private static WebElement buyerPriceBookUpdatePriceList;
		
		@FindBy(xpath="//select[@id='BasisOn']")
		private static WebElement UpdatePriceList_BasisOn;
		
		
		
		@FindBy(xpath="//select[@id='ddlIncreaseDecrease']")
		private static WebElement UpdatePriceList_Action;
		
		@FindBy(xpath="//*[@id='txtIncreaseDecrease']")
		private static WebElement UpdatePriceList_ActionTxt;
		
			
		@FindBy(xpath="//*[@id='ddlIncreaseDecreaseControl']")
		private static WebElement UpdatePriceList_ActionControl;
		
		@FindBy(xpath="//*[@id='rdoAppend1']")
		private static WebElement UpdatePriceList_AppendRadioBtn;
		
		@FindBy(xpath="//*[@id='rdoOverride1']")
		private static WebElement UpdatePriceList_OverrideRadioBtn;
		
		@FindBy(xpath="//*[@id='ctrlOptionProUP']")
		private static WebElement UpdatePriceList_ItemTxt;
		
		@FindBy(xpath="//*[@id='ctrlOptionDeptUP']")
		private static WebElement UpdatePriceList_DepatmentTxt;
		
		@FindBy(xpath="//*[@id='ctrlOptionCustUP']")
		private static WebElement UpdatePriceList_VendorTxt;
		
		@FindBy(xpath="//*[@id='ctrlOptionCurrUP']")
		private static WebElement UpdatePriceList_CurrencyTxt;
		
		@FindBy(xpath="//*[@id='ctrlOptionUnitUP']")
		private static WebElement UpdatePriceList_UnitTxt;
		
		@FindBy(xpath="//*[@id='btnOkForUpdate']")
		private static WebElement UpdatePriceList_OKBtn;
		
		@FindBy(xpath="//*[@id='btnCloseForUpdate']")
		private static WebElement UpdatePriceList_CloseBtn;
		
		
		
		
		
		
		
		
		
		
		
		
				 
			
	
	
	
	
	
	
	public BuyerPriceBookPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}
	
	
	

}
