package com.focus.Pages;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
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

public class SellerPriceBookPage extends BaseEngine 
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "SellerPriceBookPage";
	
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








	public boolean checkSavingSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		homeMasterItemMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
		homeMasterItem_sellerPriceBookMenu.click();

		Thread.sleep(3000);

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
		Thread.sleep(2000);

		System.out.println("*********************************checkSavingSellerPriceBook********************************");

		String expMessage=excelReader.getCellData(xlSheetName, 8, 6);

		String actMessage=checkValidationMessage(expMessage);
		
		 excelReader.setCellData(xlfile, xlSheetName, 8, 7, actMessage);


		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			System.out.println("Pass: Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");
			excelReader.setCellData(xlfile, xlSheetName, 8, 8, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail:  Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");
			excelReader.setCellData(xlfile, xlSheetName, 8, 8, resFail);
			return false;
		}
	}







	public boolean checkUpdatingSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		Thread.sleep(2000);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

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

		System.out.println("*********************************checkUpdatingSellerPriceBook********************************");

		String expMessage = excelReader.getCellData(xlSheetName, 13, 6);
		String actMessage = checkValidationMessage(expMessage);

		System.out.println("Item : "+actItem+"  Value Expected  "+expItem);
		System.out.println("Rate : "+actRate+"  Value Expected  "+expRate);
		
		excelReader.setCellData(xlfile, xlSheetName, 13, 7, actMessage);

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






	public boolean checkCopyPasteAndSavingSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
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

		System.out.println("*********************************checkCopyPasteAndSavingSellerPriceBook********************************");

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







	public boolean checkDeletingSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));
		Thread.sleep(2000);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		Thread.sleep(2000);

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

		getWaitForAlert();
		getAlert().accept();
		
		Thread.sleep(2000);

		System.out.println("*********************************checkDeletingSellerPriceBook********************************");

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







	public boolean checkExportingSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		File Efile=new File(getBaseDir()+"\\autoIt\\ExportFiles\\TestSmokeSellingRates.xlsx");

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

      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
      	homeMasterItem_sellerPriceBookMenu.click();*/

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.sendKeys(excelReader.getCellData(xlSheetName, 19, 5));
		Thread.sleep(1000);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookExportToExcel));
		priceBookExportToExcel.click();
		Thread.sleep(3000);

		String expMessage=excelReader.getCellData(xlSheetName, 19, 6);

		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 19, 7, actMessage);

		Thread.sleep(4000);

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

		String filename = checkDownloadedFileName(getDriver());
		
		ExcelReader excelReader = new ExcelReader(POJOUtility.getExcelPath());

		Workbook OrginalFile = new XSSFWorkbook(getBaseDir()+"\\autoIt\\ImportFiles\\TestSmokeSellingRates.xlsx");
		Workbook Exportedfile= new XSSFWorkbook(getBaseDir()+"\\autoIt\\ExportFiles\\"+filename);

		boolean result = excelReader.verifyDataInExcelBookAllSheets(OrginalFile, Exportedfile);

		System.out.println("*********************************checkExportingSellerPriceBook********************************");

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




	


	public boolean checkImportingSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader = new ExcelReader(POJOUtility.getExcelPath());

		String excelfile = getBaseDir()+"\\autoIt\\ImportFiles\\TestSmokeSellingRatesImport.xlsx";
		String sheet = "PriceBook";

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		String presentDate = currentDate.format(cal.getTime());

		System.err.println("Present date : "+presentDate);

		excelReader.setCellData(excelfile, sheet, 1, 5, presentDate);
		excelReader.setCellData(excelfile, sheet, 2, 5, presentDate);

		excelReader.setCellData(excelfile, sheet, 1, 6, presentDate);
		excelReader.setCellData(excelfile, sheet, 2, 6, presentDate);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.END);
		Thread.sleep(2000);
		buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		//buyerPriceBook_PriceBook.sendKeys(excelReader.getCellData(xlSheetName, 20, 5));
		buyerPriceBook_PriceBook.sendKeys("TestSmokeSellerRatesImport");
		Thread.sleep(2000);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		//buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_ImportFromExcel));
		buyerPriceBook_ImportFromExcel.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(browseButton));
		browseButton.click();

		Thread.sleep(2000);

		Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\FocusSellerImportPrice.exe");

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetName));
		sheetName.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetOkButton));
		sheetOkButton.click();

		getWaitForAlert();
		
		Thread.sleep(2000);

		String expAlertMessage = "Do you want to load fields mapping?";/*excelReader.getCellData(xlSheetName, 20, 6);*/

		String actAlertMessage =getAlert().getText();
		
		excelReader.setCellData(xlfile, xlSheetName, 20, 7, actAlertMessage);

		getAlert().accept();
		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetItemHeading));
		sheetItemHeading.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_MappingDrpdwn));
		Select pricebook = new Select(buyerPriceBook_MappingDrpdwn);
		pricebook.selectByVisibleText("TestSmokeSellerRatesImport");

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAccountHeading));
			sheetAccountHeading.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAccount));
			sheetAccount.click();
			sheetAccount.sendKeys("ItemName");*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetTagHeading));
		sheetTagHeading.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_MappingDrpdwn));
		pricebook.selectByVisibleText("Customer");

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetCurrencyHeading));
			sheetCurrencyHeading.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetCurrency));
			sheetCurrency.click();
			sheetCurrency.sendKeys("Department");


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetStartDateHeading));
			sheetStartDateHeading.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetStartDate));
			sheetStartDate.click();
			sheetStartDate.sendKeys("Currency");


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetEndDateHeading));
			sheetEndDateHeading.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetEndDate));
			sheetEndDate.click();
			sheetEndDate.sendKeys("StartDate");


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetMinQtyHeading));
			sheetMinQtyHeading.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetMinQty));
			sheetMinQty.click();
			sheetMinQty.sendKeys("EndDate");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetMaxQtyHeading));
			sheetMaxQtyHeading.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetMaxQty));
			sheetMaxQty.click();
			sheetMaxQty.sendKeys("MinQty");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetUnitHeading));
			sheetUnitHeading.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetUnitQty));
			sheetUnitQty.click();
			sheetUnitQty.sendKeys("MaxQty");


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetRateHeading));
			sheetRateHeading.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetRate));
			sheetRate.click();
			sheetRate.sendKeys("UnitName");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetPOHeading));
			sheetPOHeading.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetPO));
			sheetPO.click();
			sheetPO.sendKeys("Rate");


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetPTHeading));
			sheetPTHeading.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetPT));
			sheetPT.click();
			sheetPT.sendKeys("SO");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetHeading));
			sheetHeading.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetP));
			Select s = new Select(sheetP);
			s.selectByValue("12");*/

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importButton));
		importButton.click();
		Thread.sleep(2000);

		String expMessage=excelReader.getCellData(xlSheetName, 21, 6);

		String actMessage=checkValidationMessage(expMessage);
        
		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actMessage);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.END);
		buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
		//buyerPriceBook_PriceBook.sendKeys(excelReader.getCellData(xlSheetName, 22, 5));
		buyerPriceBook_PriceBook.sendKeys("TestSmokeSellerRatesImport");
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		Thread.sleep(2000);
	
		JavascriptExecutor js = (JavascriptExecutor) getDriver(); 
		js.executeScript("arguments[0].scrollIntoView(true);", priceBookFilterAndLoadBtn);
		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();

		Thread.sleep(2000);
		
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
		String expPriceBookRow1 ="[1, ITEMS GROUP, , AMERICA, "+presentDate+", "+presentDate+", , Indian Rupees, 0, 0, , 300.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

		System.out.println(actPriceBookRow1);
		System.out.println(expPriceBookRow1);


		ArrayList<String> pricebookRow2List = new ArrayList<String>();

		for (int i = 0; i < count; i++) 
		{
			String data = buyerPriceBook_TableRow2List.get(i).getText();

			pricebookRow2List.add(data);
		}

		String actPriceBookRow2 = pricebookRow2List.toString();
		String expPriceBookRow2 ="[2, BR COGS ITEM, , AMERICA, "+presentDate+", "+presentDate+", , Indian Rupees, 0, 0, Pcs, 500.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

		System.out.println(actPriceBookRow2);
		System.out.println(expPriceBookRow2);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();

		String expUpdatingMessage = "Price Book updated successfully";
		String actUpdatingMessage=checkValidationMessage(expMessage);


		System.out.println("*********************************checkImportingSellerPriceBook********************************");

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


	
	

	@FindBy(xpath="//select[@id='iProductType']")
	public static WebElement itemTypeDropdown;
	
	/*@FindBy(xpath="//*[@id='lblnchecked']/input")
	public static List<WebElement> accChkBoxList;
	
	@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[12]")
	private static List<WebElement> accNameList;
	*/

	@FindBy(xpath="//*[@id='PriceBookWebGrid_val1']")
	private static WebElement val1Value;
	
	@FindBy(xpath="//*[@id='PriceBookWebGrid_option_1_input_container']//tr/td[2]")
	private static List<WebElement> customerComboBoxList;

	
////////////////////////////////////////////////////////////////////////////////////////////////////
	
	
public boolean checkSavingSellerPriceBookforItemGroupCustomerWisePricebook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
homeMasterItem_sellerPriceBookMenu.click();

Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
buyerPriceBook_PriceBook.sendKeys("Seller1");
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();

Thread.sleep(1000);
if(buyerPriceBook_CustomizeVendorCheckBox.isDisplayed()==false)
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustomizeVendorCheckBoxSelected));
buyerPriceBook_CustomizeVendorCheckBoxSelected.click();
}

Thread.sleep(1500);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
itemFirstCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.SPACE);

int count = itemComboBoxList.size();

for(int i=0;i<count;i++)
{
String data = itemComboBoxList.get(i).getText();

if(data.equalsIgnoreCase("Grp1"))
{
itemComboBoxList.get(i).click();
break;
}
}

itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
vendorValue.click();
vendorValue.sendKeys(Keys.SPACE);

int count1 = customerComboBoxList.size();

for(int i=0;i<count1;i++)
{
String data = customerComboBoxList.get(i).getText();

if(data.equalsIgnoreCase("Customer A"))
{
customerComboBoxList.get(i).click();
break;
}
}

vendorValue.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("33");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("44");
val1Value.sendKeys(Keys.TAB);

//second cell


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
itemSecondCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.SPACE);

int count2 = itemComboBoxList.size();

for(int i=0;i<count2;i++)
{
String data = itemComboBoxList.get(i).getText();

if(data.equalsIgnoreCase("Grp1"))
{
itemComboBoxList.get(i).click();
break;
}
}

itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
vendorValue.click();
vendorValue.sendKeys(Keys.END);
vendorValue.sendKeys(Keys.SHIFT,Keys.HOME);
vendorValue.sendKeys(Keys.SPACE);

int count3 = customerComboBoxList.size();

for(int i=0;i<count3;i++)
{
String data = customerComboBoxList.get(i).getText();

if(data.equalsIgnoreCase("Customer B"))
{
customerComboBoxList.get(i).click();
break;
}
}

vendorValue.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("300");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
val1Value.click();
val1Value.sendKeys("400");
val1Value.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book created successfully";

String actMessage=checkValidationMessage(expMessage);

if (actMessage.equalsIgnoreCase(expMessage)) 
{
System.out.println("Pass: Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");
// excelReader.setCellData(xlfile, "Sheet1", 803, 9, resPass);
return true;
} 
else 
{
System.out.println("Fail:  Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");
//excelReader.setCellData(xlfile, "Sheet1", 803, 9, resFail);
return false;
}

}



@FindBy(xpath="//span[@class='icon-financial icon-font1']")
private static WebElement  finacinalsMenu;

@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[12]/a")
private static WebElement  settingsBtnSalesInvoice;

@FindBy(xpath="//*[contains(text(),'Edit Screen')]")
private static WebElement  editScreenSalesInvoice;


@FindBy(xpath="//select[@id='cmbBehaviour']")
private static WebElement  editScreenBehaviourComboBox;


@FindBy(xpath="//*[@id='editScreen_qtyFormula_textbox']")
private static WebElement  editScreenFormulaTxtBox;


@FindBy(xpath="//*[@id='editScreen_qtyFormula_Ok']")
private static WebElement  editScreenFormulaOkBtn;


@FindBy(xpath="//*[@id='updateButton']")
private static WebElement  settingsUpdateBtn;



@FindBy(xpath="//*[@id='btnCustomizeClose']")
private static WebElement  settingsCloseBtn;
@FindBy(xpath="//*[@id='callbackbtn']/i")
private static WebElement  PV_settingsCloseBtn;


@FindBy(xpath = "//tbody[@id='id_header_4_table_body']/tr/td[2]")
public static List<WebElement> customerAccountListCount;

@FindBy(xpath = "//tbody[@id='id_body_23_table_body']/tr/td[2]")
public static List<WebElement> itemListCount;

public boolean checkSavedSellerPriceBookatVoucherSalesInvoice() throws InterruptedException
{
	getDriver().navigate().refresh();
	Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
financialsMenu.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
financialsTransactionMenu.click();
Thread.sleep(2000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
financialTransactionSalesMenu.click();
Thread.sleep(2000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVoucher));
salesInvoiceVoucher.click();

Thread.sleep(3000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsBtnSalesInvoice));
settingsBtnSalesInvoice.click();
Thread.sleep(1500);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenSalesInvoice));
editScreenSalesInvoice.click();
Thread.sleep(1500);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBehaviourComboBox));
Select s=new Select(editScreenBehaviourComboBox);
s.selectByVisibleText("Formula");
Thread.sleep(1500);
editScreenBehaviourComboBox.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFormulaTxtBox));
editScreenFormulaTxtBox.click();
editScreenFormulaTxtBox.clear();
editScreenFormulaTxtBox.sendKeys("&salrtgrp(0)");
Thread.sleep(1000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFormulaOkBtn));
editScreenFormulaOkBtn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsUpdateBtn));
settingsUpdateBtn.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsCloseBtn));
settingsCloseBtn.click();

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
newBtn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
customerAccountTxt.click();
customerAccountTxt.sendKeys(Keys.END);
customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
customerAccountTxt.sendKeys(Keys.SPACE);

int customercount=customerAccountListCount.size();

System.err.println(customercount);

for(int i=0 ; i < customercount ;i++)
{
String data=customerAccountListCount.get(i).getText();

if(data.equalsIgnoreCase("Customer A"))
{
customerAccountListCount.get(i).click();
Thread.sleep(2000);

break;
}
}

customerAccountTxt.sendKeys(Keys.TAB);

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
select1stRow_1stColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
enter_ItemTxt.click();
enter_ItemTxt.sendKeys(Keys.END);
enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
enter_ItemTxt.sendKeys(Keys.SPACE);

int itemcount1=itemListCount.size();

System.err.println(itemcount1);

for(int i=0 ; i < itemcount1 ;i++)
{
String data=itemListCount.get(i).getText();

if(data.equalsIgnoreCase("item1"))
{
itemListCount.get(i).click();

break;
}
}

enter_ItemTxt.sendKeys(Keys.TAB);

Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
select1stRow_11thColumn.click();
Thread.sleep(2000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
enter_Rate.click();
Thread.sleep(1000);

String actRate=enter_Rate.getAttribute("value");
String expRate="33.00";

System.out.println("Actual Rate :		" +	actRate 		+ "Value Expected:		"	+expRate);

if(actRate.equalsIgnoreCase(expRate))

{

return true;
}
else
{
return false;
}
}



public boolean checkDeleteSellerPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
homeMasterItem_sellerPriceBookMenu.click();

Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
buyerPriceBook_PriceBook.sendKeys("Seller1");
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();

Thread.sleep(3000);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
deleteBtn.click();
Thread.sleep(2000);

getWaitForAlert();
getAlert().accept();

System.out.println("*********************************checkDeletingSellerPriceBook********************************");

String expMessage = "Pricebook details deleted successfully.";

String actMessage=checkValidationMessage(expMessage);

System.out.println(actMessage);
System.out.println(expMessage);

if ( actMessage.equalsIgnoreCase(expMessage)) 
{
return true;
} 
else 
{
return false;
}

}


//*[@id='StartingDateH']
@FindBy(xpath="//*[@id='StartingDateH']")
private static WebElement buyerPriceBook_StartDate;


@FindBy(xpath="//*[@id='EndingDateH']")
private static WebElement buyerPriceBook_EndDate;

@FindBy(xpath="//*[@id='EndingDateH_checkbox']")
private static WebElement buyerPriceBook_EndDateChkBox;


public boolean checkCreatingTwodiffBuyerPricebookforSameProductatDiffDates() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
homeMasterItem_buyerPriceBookMenu.click();

Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
buyerPriceBook_PriceBook.sendKeys("Buyer1");
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_StartDate));
DateFormat format = new SimpleDateFormat("dd/MM/yyyy");
Calendar cal = Calendar.getInstance();
cal.add(Calendar.DATE, -10);
Date first=cal.getTime();
buyerPriceBook_StartDate.sendKeys(format.format(first));
Thread.sleep(1000);
buyerPriceBook_StartDate.sendKeys(Keys.TAB);

System.out.println(format.format(first));

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_EndDateChkBox));
buyerPriceBook_EndDateChkBox.sendKeys(Keys.TAB);
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_EndDate));

Calendar cal1 = Calendar.getInstance();
cal1.add(Calendar.DATE, -5);
Date d1=cal1.getTime();
buyerPriceBook_EndDate.sendKeys(format.format(d1));
Thread.sleep(2000);
buyerPriceBook_EndDate.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
itemFirstCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.SPACE);

int count = itemComboBoxList.size();

for(int i=0;i<count;i++)
{
String data = itemComboBoxList.get(i).getText();

if(data.equalsIgnoreCase("item1"))
{
itemComboBoxList.get(i).click();
break;
}
}

itemValue.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("10");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("100");
val1Value.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book created successfully";

String actMessage=checkValidationMessage(expMessage);


///Saving second PriceBook

Thread.sleep(2500);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
buyerPriceBook_PriceBook.click();
buyerPriceBook_PriceBook.sendKeys("Buyer2");
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
itemFirstCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.SPACE);

int count1 = itemComboBoxList.size();

for(int i=0;i<count1;i++)
{
String data = itemComboBoxList.get(i).getText();

if(data.equalsIgnoreCase("item1"))
{
itemComboBoxList.get(i).click();
break;
}
}

itemValue.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("15");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("150");
val1Value.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingBuyerPriceBook********************************");

String expMessage1="Price Book created successfully";

String actMessage1=checkValidationMessage(expMessage1);




if (actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1) ) 
{
System.out.println("Pass: Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");

return true;
} 
else 
{
System.out.println("Fail:  Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");

return false;
}

}

@FindBy(xpath="//*[@id='2008']/span")
private static WebElement  purchaseVouchers;

public boolean checkUpdatedBuyerPriceBooksRateValueatVoucherPurchaseVoucher() throws InterruptedException
{


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
financialsMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
financialsTransactionMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
financialsTransactionsPurchaseMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchers));
purchaseVouchers.click();

Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsBtnSalesInvoice));
settingsBtnSalesInvoice.click();
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenSalesInvoice));
editScreenSalesInvoice.click();
Thread.sleep(1500);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBehaviourComboBox));
Select s=new Select(editScreenBehaviourComboBox);
s.selectByVisibleText("Formula");
Thread.sleep(1500);
editScreenBehaviourComboBox.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFormulaTxtBox));
editScreenFormulaTxtBox.click();
editScreenFormulaTxtBox.clear();
editScreenFormulaTxtBox.sendKeys("&purrt(0)");
Thread.sleep(1000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFormulaOkBtn));
editScreenFormulaOkBtn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsUpdateBtn));
settingsUpdateBtn.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsCloseBtn));
settingsCloseBtn.click();

Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
newBtn.click();



Thread.sleep(2000);





getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
select1stRow_1stColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
enter_ItemTxt.click();
enter_ItemTxt.sendKeys(Keys.END);
enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
enter_ItemTxt.sendKeys(Keys.SPACE);

int itemcount1=itemListCount.size();

System.err.println(itemcount1);

for(int i=0 ; i < itemcount1 ;i++)
{
String data=itemListCount.get(i).getText();

if(data.equalsIgnoreCase("item1"))
{
itemListCount.get(i).click();

break;
}
}

enter_ItemTxt.sendKeys(Keys.TAB);
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
select1stRow_6thColumn.click();
Thread.sleep(1500);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
String actRate=enter_Rate.getAttribute("value");
String expRate="15.00";

System.out.println("Actual Rate :		" +	actRate 		+ "Value Expected:		"	+expRate);

if(actRate.equalsIgnoreCase(expRate))

{

return true;
}
else
{
return false;
}

}



public boolean checkDeleteBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
homeMasterItem_buyerPriceBookMenu.click();

Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
buyerPriceBook_PriceBook.sendKeys("Buyer1");
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();

Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
deleteBtn.click();
Thread.sleep(2000);

getWaitForAlert();
getAlert().accept();



String expMessage = "Pricebook details deleted successfully.";

String actMessage=checkValidationMessage(expMessage);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
buyerPriceBook_PriceBook.sendKeys("Buyer2");
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();

Thread.sleep(3000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
deleteBtn.click();
Thread.sleep(2000);

getWaitForAlert();
getAlert().accept();

System.out.println("*********************************checkDeletingBuyerPriceBook********************************");

String expMessage1 = "Pricebook details deleted successfully.";

String actMessage1=checkValidationMessage(expMessage1);

System.out.println(actMessage);
System.out.println(expMessage);


System.out.println(actMessage1);
System.out.println(expMessage1);



if ( actMessage.equalsIgnoreCase(expMessage)) 
{
return true;
} 
else 
{
return false;
}

}



public boolean checkCreatingBuyerPriceBookwithMaxandMinQuantity() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
checkSettingsofBuyerPriceBook("Warehouse");
Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
homeMasterItem_buyerPriceBookMenu.click();

Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
buyerPriceBook_PriceBook.sendKeys("Buyer1");
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();

Thread.sleep(1000);
if(buyerPriceBook_WarehouseCheckBox.isSelected()==false)
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable( buyerPriceBook_WarehouseCheckBoxSelected));
buyerPriceBook_WarehouseCheckBoxSelected.click();
}
Thread.sleep(1500);
if(buyerPriceBook_QualityRangeCheckBox.isSelected()==false)
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable( buyerPriceBook_QualityRangeCheckBoxSelected));
buyerPriceBook_QualityRangeCheckBoxSelected.click();
}
Thread.sleep(1500);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();

Thread.sleep(1000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
itemFirstCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.SPACE);

int count = itemComboBoxList.size();

for(int i=0;i<count;i++)
{
String data = itemComboBoxList.get(i).getText();

if(data.equalsIgnoreCase("Grp1"))
{
itemComboBoxList.get(i).click();
break;
}
}

itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
wareHouseValue.click();
wareHouseValue.sendKeys(Keys.SPACE);

int count1 = warehouseComboBoxList.size();

for(int i=0;i<count1;i++)
{
String data = warehouseComboBoxList.get(i).getText();

if(data.equalsIgnoreCase("Hyderabad"))
{
warehouseComboBoxList.get(i).click();
Thread.sleep(1000);
break;
}
}

wareHouseValue.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
minQtyValue.sendKeys("10.5");
minQtyValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
maxQtyValue.sendKeys("20.5");
maxQtyValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("10");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("100");
val1Value.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book created successfully";

String actMessage=checkValidationMessage(expMessage);

if (actMessage.equalsIgnoreCase(expMessage)) 
{
System.out.println("Pass: Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");

return true;
} 
else 
{
System.out.println("Fail:  Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");

return false;
}

}



public boolean checkPurchaseVoucherRateFieldgivingMinorMaxQuantitygiveninBuyerPriceBook() throws InterruptedException
{

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
financialsMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
financialsTransactionMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
financialsTransactionsPurchaseMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchers));
purchaseVouchers.click();

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
newBtn.click();



Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
warehouseTxt.click();
warehouseTxt.sendKeys("Hyderabad");
Thread.sleep(1500);
warehouseTxt.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
select1stRow_1stColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
enter_ItemTxt.click();
enter_ItemTxt.sendKeys(Keys.END);
enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
enter_ItemTxt.sendKeys(Keys.SPACE);

int itemcount1=itemListCount.size();

System.err.println(itemcount1);

for(int i=0 ; i < itemcount1 ;i++)
{
String data=itemListCount.get(i).getText();

if(data.equalsIgnoreCase("item1"))
{
itemListCount.get(i).click();

break;
}
}

enter_ItemTxt.sendKeys(Keys.TAB);

Thread.sleep(2000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
select1stRow_6thColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
String actRate=enter_Rate.getAttribute("value");
String expRate="0.00";


System.out.println("Rate value before giving Quantity");
System.out.println("Actual Rate :		" +	actRate 		+ "Value Expected:		"	+expRate);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
select1stRow_4thColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

//enter_Quantity.click();
enter_Quantity.sendKeys("12");
enter_Quantity.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
String actRate1=enter_Rate.getAttribute("value");
String expRate1="10.00";


System.out.println("Rate value after giving Quantity");
System.out.println("Actual Rate :		" +	actRate1 		+ "Value Expected:		"	+expRate1);



if(actRate.equalsIgnoreCase(expRate) && actRate1.equalsIgnoreCase(expRate1))

{

return true;
}
else
{
return false;
}
}

@FindBy(xpath="//select[@id='BuyingRates']")
public static WebElement buyingRatesSelect;


@FindBy(xpath="//*[@id='dvTag']/div/label")
private static List<WebElement>  configureTransactionPageList;


@FindBy(xpath="//a[@id='0']")
private static WebElement  TagsBtn;
@FindBy(xpath = "//*[@id='16']")
public static WebElement settingsMenu;


public void checkSettingsofBuyerPriceBook(String buyerName) throws InterruptedException
{

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsMenu));
settingsMenu.click();

Thread.sleep(1500);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(preferencesBtn));
preferencesBtn.click();
Thread.sleep(2000);

/*
ArrayList<String>actconfigureTransactionPageList=new ArrayList<String>();

int configureTransactionPageCount=configureTransactionPageList.size();

System.out.println(" configureTransactionPageCount  :  "+configureTransactionPageCount);

for (int i = 0; i < configureTransactionPageCount; i++)
{
String data=configureTransactionPageList.get(i).getText();
actconfigureTransactionPageList.add(data);
}*/


/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(TagsBtn));
TagsBtn.click();

getAction().moveToElement(buyingRatesSelect).build().perform();
*/

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyingRatesSelect));
buyingRatesSelect.click();
Thread.sleep(2000);

Select s = new Select(buyingRatesSelect);
s.selectByVisibleText(buyerName);
//s.selectByValue("9");
Thread.sleep(2000);

//buyingRatesSelect.sendKeys(Keys.TAB);

//Thread.sleep(1500);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
updateBtn.click();

Thread.sleep(1000);

getWaitForAlert();

getAlert().accept();
Thread.sleep(2000);

/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsCloseBtn));

settingsCloseBtn.click();
Thread.sleep(2000);*/

}


public boolean checkDeleteCreatedBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
homeMasterItem_buyerPriceBookMenu.click();

Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
buyerPriceBook_PriceBook.sendKeys("Buyer1");
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();

Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
deleteBtn.click();
Thread.sleep(2000);

getWaitForAlert();
getAlert().accept();



String expMessage = "Pricebook details deleted successfully.";

String actMessage=checkValidationMessage(expMessage);



System.out.println("*********************************checkDeletingBuyerPriceBook********************************");


System.out.println(actMessage);
System.out.println(expMessage);





if ( actMessage.equalsIgnoreCase(expMessage)) 
{
return true;
} 
else 
{
return false;
}


}




public boolean checkCreateTwoDiffBuyerPriceBooksatDiffDateswithSameItem() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
homeMasterItem_buyerPriceBookMenu.click();

Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
buyerPriceBook_PriceBook.sendKeys("Buyer1");
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);




DateFormat format = new SimpleDateFormat("dd/MM/yyyy");

Calendar cal1=Calendar.getInstance();
cal1.add(Calendar.YEAR,0);
cal1.set(Calendar.DAY_OF_YEAR,1);
Date start=cal1.getTime();

buyerPriceBook_StartDate.sendKeys(format.format(start));

Thread.sleep(1000);
buyerPriceBook_StartDate.sendKeys(Keys.TAB);

System.out.println(format.format(start));

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_EndDateChkBox));
buyerPriceBook_EndDateChkBox.sendKeys(Keys.TAB);
Thread.sleep(2000);


//cal1.set(Calendar.DAY_OF_YEAR, cal1.getActualMaximum(Calendar.DAY_OF_YEAR));
Calendar cal11 = Calendar.getInstance();
cal11.add(Calendar.DATE, 5);
Date d1=cal11.getTime();

buyerPriceBook_EndDate.sendKeys(format.format(d1));
Thread.sleep(2000);
buyerPriceBook_EndDate.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
itemFirstCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.SPACE);

int count = itemComboBoxList.size();

for(int i=0;i<count;i++)
{
String data = itemComboBoxList.get(i).getText();

if(data.equalsIgnoreCase("item1"))
{
itemComboBoxList.get(i).click();
break;
}
}

itemValue.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("10");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("100");
val1Value.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book created successfully";

String actMessage=checkValidationMessage(expMessage);


///Saving second PriceBook

Thread.sleep(2500);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
buyerPriceBook_PriceBook.click();
buyerPriceBook_PriceBook.sendKeys("Buyer2");
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_StartDate));
buyerPriceBook_StartDate.click();
buyerPriceBook_StartDate.sendKeys(Keys.TAB);
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_EndDateChkBox));
buyerPriceBook_EndDateChkBox.click();
buyerPriceBook_EndDateChkBox.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
itemFirstCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.SPACE);

int count1 = itemComboBoxList.size();

for(int i=0;i<count1;i++)
{
String data = itemComboBoxList.get(i).getText();

if(data.equalsIgnoreCase("item1"))
{
itemComboBoxList.get(i).click();
break;
}
}

itemValue.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("15");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("150");
val1Value.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingBuyerPriceBook********************************");

String expMessage1="Price Book created successfully";

String actMessage1=checkValidationMessage(expMessage1);




if (actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1) ) 
{
System.out.println("Pass: Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");

return true;
} 
else 
{
System.out.println("Fail:  Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");

return false;
}


}




public boolean checkValidatingtheRateFieldatPurchaseVoucher() throws InterruptedException
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
financialsMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
financialsTransactionMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
financialsTransactionsPurchaseMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchers));
purchaseVouchers.click();

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
newBtn.click();

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
documentNumberTxt.click();
documentNumberTxt.sendKeys(Keys.TAB);
Thread.sleep(1500);

DateFormat format = new SimpleDateFormat("dd/MM/yyyy");
Calendar cal11 = Calendar.getInstance();
cal11.add(Calendar.DATE, 5);
Date d1=cal11.getTime();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dateTxt));
dateTxt.sendKeys(format.format(d1));
Thread.sleep(1500);
dateTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
warehouseTxt.click();
warehouseTxt.sendKeys("Hyderabad");
Thread.sleep(1500);
warehouseTxt.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
select1stRow_1stColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
enter_ItemTxt.click();
enter_ItemTxt.sendKeys(Keys.END);
enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
enter_ItemTxt.sendKeys(Keys.SPACE);

int itemcount1=itemListCount.size();

System.err.println(itemcount1);

for(int i=0 ; i < itemcount1 ;i++)
{
String data=itemListCount.get(i).getText();

if(data.equalsIgnoreCase("item1"))
{
itemListCount.get(i).click();

break;
}
}

enter_ItemTxt.sendKeys(Keys.TAB);

Thread.sleep(2000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
select1stRow_6thColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
String actRate=enter_Rate.getAttribute("value");
String expRate="10.00";


System.out.println("Rate value before current Date");
System.out.println("Actual Rate :		" +	actRate 		+ "Value Expected:		"	+expRate);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
documentNumberTxt.click();
documentNumberTxt.sendKeys(Keys.TAB);
Thread.sleep(1500);

Calendar cal1=Calendar.getInstance();
DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
String currentDate = df.format(cal1.getTime());

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dateTxt));
dateTxt.sendKeys(currentDate);
Thread.sleep(1500);
dateTxt.sendKeys(Keys.TAB);


getWaitForAlert();
getAlert().accept();
Thread.sleep(1500);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
warehouseTxt.click();
warehouseTxt.sendKeys(Keys.END);
warehouseTxt.sendKeys(Keys.SHIFT,Keys.HOME);
warehouseTxt.sendKeys(Keys.BACK_SPACE);
warehouseTxt.sendKeys("Hyderabad");
Thread.sleep(1500);
warehouseTxt.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
select1stRow_1stColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
enter_ItemTxt.click();
enter_ItemTxt.sendKeys(Keys.END);
enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
enter_ItemTxt.sendKeys(Keys.SPACE);

int itemcount11=itemListCount.size();

System.err.println(itemcount11);

for(int i=0 ; i < itemcount11 ;i++)
{
String data=itemListCount.get(i).getText();

if(data.equalsIgnoreCase("item1"))
{
itemListCount.get(i).click();

break;
}
}

enter_ItemTxt.sendKeys(Keys.TAB);

Thread.sleep(2000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
select1stRow_6thColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
String actRate1=enter_Rate.getAttribute("value");
String expRate1="15.00";

System.out.println("Rate value after giving CurrentDate");
System.out.println("Actual Rate :		" +	actRate1 		+ "Value Expected:		"	+expRate1);



if(actRate.equalsIgnoreCase(expRate) && actRate1.equalsIgnoreCase(expRate1))

{

return true;
}
else
{
return false;
}


}




public boolean checkDeletingtheBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
boolean result=checkDeletingBuyerPriceBook("Buyer1");
boolean result1= checkDeletingBuyerPriceBook("Buyer2");
if(result==true && result1==true)
{
return true;
}
else
{
return false;
}
}


public boolean checkDeletingBuyerPriceBook(String bname) throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
homeMasterItem_buyerPriceBookMenu.click();

Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
buyerPriceBook_PriceBook.sendKeys(bname);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();

Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
deleteBtn.click();
Thread.sleep(2000);

getWaitForAlert();
getAlert().accept();



String expMessage = "Pricebook details deleted successfully.";

String actMessage=checkValidationMessage(expMessage);



System.out.println("*********************************checkDeletingBuyerPriceBook********************************");


System.out.println(actMessage);
System.out.println(expMessage);

if ( actMessage.equalsIgnoreCase(expMessage)) 
{
return true;
} 
else 
{
return false;
}


}



public boolean checkCreatingBuyerPriceBookwithVendorandRate() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
homeMasterItem_buyerPriceBookMenu.click();

Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
buyerPriceBook_PriceBook.sendKeys("BuyerVendor1");
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_StartDate));
//buyerPriceBook_StartDate.click();

DateFormat format = new SimpleDateFormat("dd/MM/yyyy");

Calendar cal=Calendar.getInstance();
cal.add(Calendar.YEAR,0);
cal.set(Calendar.DAY_OF_YEAR,1);
Date start=cal.getTime();

buyerPriceBook_StartDate.sendKeys(format.format(start));

Thread.sleep(1000);
buyerPriceBook_StartDate.sendKeys(Keys.TAB);

System.out.println(format.format(start));

Thread.sleep(2000);

Calendar cal1=Calendar.getInstance();
cal1.add(Calendar.MONTH, 0);

//cal1.set(Calendar.DATE,cal1.getActualMinimum(Calendar.DAY_OF_MONTH));
//Date first=cal1.getTime();
cal1.set(Calendar.DATE, cal1.getActualMaximum(Calendar.DAY_OF_MONTH));
Date last=cal1.getTime();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_EndDateChkBox));
buyerPriceBook_EndDateChkBox.sendKeys(Keys.TAB);
Thread.sleep(2000);


buyerPriceBook_EndDate.sendKeys(format.format(last));
Thread.sleep(2000);
buyerPriceBook_EndDate.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();

Thread.sleep(1000);

if(buyerPriceBook_CustomizeVendorCheckBox.isSelected()==false)
{

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable( buyerPriceBook_CustomizeVendorCheckBoxSelected));
buyerPriceBook_CustomizeVendorCheckBoxSelected.click();
}
Thread.sleep(1500);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();

Thread.sleep(1000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
itemFirstCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.SPACE);

int count = itemComboBoxList.size();

for(int i=0;i<count;i++)
{
String data = itemComboBoxList.get(i).getText();

if(data.equalsIgnoreCase("item2"))
{
itemComboBoxList.get(i).click();
break;
}
}

itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
vendorValue.click();
vendorValue.sendKeys(Keys.SPACE);

int count1 = vendorComboBoxList.size();

for(int i=0;i<count1;i++)
{
String data = vendorComboBoxList.get(i).getText();

if(data.equalsIgnoreCase("Vendor A"))
{
vendorComboBoxList.get(i).click();
Thread.sleep(1000);
break;
}
}

vendorValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("10");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("100");
val1Value.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book created successfully";

String actMessage=checkValidationMessage(expMessage);

if (actMessage.equalsIgnoreCase(expMessage)) 
{
System.out.println("Pass: Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");

return true;
} 
else 
{
System.out.println("Fail:  Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");

return false;
}
}



public boolean checkPurchaseVoucherRateFieldforVendorAccount() throws InterruptedException
{

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
financialsMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
financialsTransactionMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
financialsTransactionsPurchaseMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchers));
purchaseVouchers.click();

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
newBtn.click();

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
documentNumberTxt.click();
documentNumberTxt.sendKeys(Keys.TAB);
Thread.sleep(1500);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vcustomerAccountTxt));
vcustomerAccountTxt.click();
vcustomerAccountTxt.sendKeys("Vendor A");
Thread.sleep(1500);
vcustomerAccountTxt.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
select1stRow_1stColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
enter_ItemTxt.click();
enter_ItemTxt.sendKeys(Keys.END);
enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
enter_ItemTxt.sendKeys(Keys.SPACE);

int itemcount1=itemListCount.size();

System.err.println(itemcount1);

for(int i=0 ; i < itemcount1 ;i++)
{
String data=itemListCount.get(i).getText();

if(data.equalsIgnoreCase("item2"))
{
itemListCount.get(i).click();

break;
}
}

enter_ItemTxt.sendKeys(Keys.TAB);

Thread.sleep(2000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
select1stRow_6thColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
String actRate=enter_Rate.getAttribute("value");
String expRate="10.00";


System.out.println("Rate value for Vendor A");
System.out.println("Actual Rate :		" +	actRate 		+ "Value Expected:		"	+expRate);


if(actRate.equalsIgnoreCase(expRate) )

{

return true;
}
else
{
return false;
}

}



public boolean checkDeletetheBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
boolean result=checkDeletingBuyerPriceBook("BuyerVendor1");

if(result==true )
{
return true;
}
else
{
return false;
}
}

public boolean checkCreatingBuyerPriceBookforMinandMaxQuantity() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
homeMasterItem_buyerPriceBookMenu.click();

Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
buyerPriceBook_PriceBook.sendKeys("BuyerQuantity1");
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();

Thread.sleep(1000);

/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable( buyerPriceBook_WarehouseCheckBox));
buyerPriceBook_WarehouseCheckBox.click();
Thread.sleep(1500);*/

if(buyerPriceBook_QualityRangeCheckBox.isSelected()==false)
{

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable( buyerPriceBook_QualityRangeCheckBoxSelected));
buyerPriceBook_QualityRangeCheckBoxSelected.click();
}
Thread.sleep(1500);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();

Thread.sleep(1000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
itemFirstCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.SPACE);

int count = itemComboBoxList.size();

for(int i=0;i<count;i++)
{
String data = itemComboBoxList.get(i).getText();

if(data.equalsIgnoreCase("item2"))
{
itemComboBoxList.get(i).click();
break;
}
}

itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
minQtyValue.sendKeys("10");
minQtyValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
maxQtyValue.sendKeys("50");
maxQtyValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("100");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("10");
val1Value.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book created successfully";

String actMessage=checkValidationMessage(expMessage);

if (actMessage.equalsIgnoreCase(expMessage)) 
{
System.out.println("Pass: Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");

return true;
} 
else 
{
System.out.println("Fail:  Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");

return false;
}


}


public boolean checkVerifyingPurchaseVoucherforRateQuantityApplied() throws InterruptedException
{



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
financialsMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
financialsTransactionMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
financialsTransactionsPurchaseMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchers));
purchaseVouchers.click();

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
newBtn.click();

Thread.sleep(2000);

/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
documentNumberTxt.click();
documentNumberTxt.sendKeys(Keys.TAB);
Thread.sleep(1500);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vcustomerAccountTxt));
vcustomerAccountTxt.click();
vcustomerAccountTxt.sendKeys("Vendor A");
Thread.sleep(1500);
vcustomerAccountTxt.sendKeys(Keys.TAB);*/



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
select1stRow_1stColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
enter_ItemTxt.click();
enter_ItemTxt.sendKeys(Keys.END);
enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
enter_ItemTxt.sendKeys(Keys.SPACE);

int itemcount1=itemListCount.size();

System.err.println(itemcount1);

for(int i=0 ; i < itemcount1 ;i++)
{
String data=itemListCount.get(i).getText();

if(data.equalsIgnoreCase("item2"))
{
itemListCount.get(i).click();

break;
}
}

enter_ItemTxt.sendKeys(Keys.TAB);

//Minimum Quantity
Thread.sleep(2000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
select1stRow_4thColumn.click();


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
enter_Quantity.sendKeys("10");
enter_Quantity.sendKeys(Keys.TAB);

/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
select1stRow_6thColumn.click();*/

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
String actRate=enter_Rate.getAttribute("value");
String expRate="10.00";


System.out.println("Rate value for Minimum Quantity");
System.out.println("Actual Rate :		" +	actRate 		+ "Value Expected:		"	+expRate);

//Between Quantity	
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
select1stRow_4thColumn.click();


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
enter_Quantity.sendKeys("25");
enter_Quantity.sendKeys(Keys.TAB);

/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
select1stRow_6thColumn.click();*/

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
String actRate1=enter_Rate.getAttribute("value");
String expRate1="10.00";


System.out.println("Rate value for between Quantity");
System.out.println("Actual Rate :		" +	actRate1 		+ "Value Expected:		"	+expRate1);

//Maximum Quantity	

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
select1stRow_4thColumn.click();


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
enter_Quantity.sendKeys("50");
enter_Quantity.sendKeys(Keys.TAB);

/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
select1stRow_6thColumn.click();*/

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
String actRate2=enter_Rate.getAttribute("value");
String expRate2="10.00";


System.out.println("Rate value for Maximum Quantity");
System.out.println("Actual Rate :		" +	actRate2 		+ "Value Expected:		"	+expRate2);

/*//Less than Minimum Quantity

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
select1stRow_4thColumn.click();


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
enter_Quantity.sendKeys("5");
enter_Quantity.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
select1stRow_6thColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
String actRate3=enter_Rate.getAttribute("value");
String expRate3="0.00";


System.out.println("Rate value for Lessthan Minimum Quantity");
System.out.println("Actual Rate :		" +	actRate3 		+ "Value Expected:		"	+expRate3);


//Grater than Maximum Quantity

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
select1stRow_4thColumn.click();


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
enter_Quantity.sendKeys("75");
enter_Quantity.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
select1stRow_6thColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
String actRate4=enter_Rate.getAttribute("value");
String expRate4="0.00";


System.out.println("Rate value for Grater than Quantity");
System.out.println("Actual Rate :		" +	actRate4 		+ "Value Expected:		"	+expRate4);

*/

if(actRate.equalsIgnoreCase(expRate) && actRate1.equalsIgnoreCase(expRate1) 
&& actRate2.equalsIgnoreCase(expRate2) /* && actRate3.equalsIgnoreCase(expRate3)*/ 
/*&& actRate4.equalsIgnoreCase(expRate4)*/ )

{

return true;
}
else
{
return false;
}


}

public boolean checkDeletetheQuantityBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

boolean result=checkDeletingBuyerPriceBook("BuyerQuantity1");

if(result==true )
{
return true;
}
else
{
return false;
}

}




@FindBy(xpath="//*[@id='PriceBookWebGrid_option_9_table_body']/tr/td[2]")
private static List<WebElement> statesComboBoxList;

@FindBy(xpath="//input[@id='PriceBookWebGrid_option_9']")
private static WebElement stateValue; 



//Creating Extra Master

public boolean checkCreatingBuyerPriceBookforState() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
checkSettingsofBuyerPriceBook("State");
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
homeMasterItem_buyerPriceBookMenu.click();

Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
buyerPriceBook_PriceBook.sendKeys("BuyerState1");
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();

Thread.sleep(1000);

if(buyerPriceBook_WarehouseCheckBox.isSelected()==false)
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable( buyerPriceBook_WarehouseCheckBoxSelected));
buyerPriceBook_WarehouseCheckBoxSelected.click();
}
Thread.sleep(1500);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();

Thread.sleep(1000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
itemFirstCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.SPACE);

int count = itemComboBoxList.size();

for(int i=0;i<count;i++)
{
String data = itemComboBoxList.get(i).getText();

if(data.equalsIgnoreCase("item1"))
{
itemComboBoxList.get(i).click();
break;
}
}

itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stateValue));
stateValue.click();
stateValue.sendKeys(Keys.SPACE);

int count1 = statesComboBoxList.size();

for(int i=0;i<count1;i++)
{
String data = statesComboBoxList.get(i).getText();

if(data.equalsIgnoreCase("AP"))
{
statesComboBoxList.get(i).click();
Thread.sleep(1000);
break;
}
}

stateValue.sendKeys(Keys.TAB);


/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
minQtyValue.sendKeys("10.5");
minQtyValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
maxQtyValue.sendKeys("20.5");
maxQtyValue.sendKeys(Keys.TAB);


*/

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("1000");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("100");
val1Value.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book created successfully";

String actMessage=checkValidationMessage(expMessage);

if (actMessage.equalsIgnoreCase(expMessage)) 
{
System.out.println("Pass: Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");

return true;
} 
else 
{
System.out.println("Fail:  Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");

return false;
}


}

@FindBy(xpath="//*[@id='id_header_268435465']")
private static WebElement StateTxt;

public boolean checkPurchaseVoucherforLoadingRateforStateApplied() throws InterruptedException
{


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
financialsMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
financialsTransactionMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
financialsTransactionsPurchaseMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchers));
purchaseVouchers.click();

Thread.sleep(2000);


checkAddingExtraFieldinPurchaseVoucher();
Thread.sleep(3000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
newBtn.click();



Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
warehouseTxt.click();
warehouseTxt.sendKeys("Hyderabad");
Thread.sleep(1500);
warehouseTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StateTxt));
StateTxt.click();
StateTxt.sendKeys("AP");
Thread.sleep(1000);
StateTxt.sendKeys(Keys.TAB);

Thread.sleep(1500);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
select1stRow_1stColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
enter_ItemTxt.click();
enter_ItemTxt.sendKeys(Keys.END);
enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
enter_ItemTxt.sendKeys(Keys.SPACE);

int itemcount1=itemListCount.size();

System.err.println(itemcount1);

for(int i=0 ; i < itemcount1 ;i++)
{
String data=itemListCount.get(i).getText();

if(data.equalsIgnoreCase("item1"))
{
itemListCount.get(i).click();

break;
}
}

enter_ItemTxt.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
select1stRow_4thColumn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

//enter_Quantity.click();
enter_Quantity.sendKeys("12");
enter_Quantity.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
String actRate=enter_Rate.getAttribute("value");
String expRate="1000.00";


System.out.println("Rate value ");
System.out.println("Actual Rate :		" +	actRate 		+ "Value Expected:		"	+expRate);



if(actRate.equalsIgnoreCase(expRate) )

{

return true;
}
else
{
return false;
}

}

@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[12]/a")
public static WebElement settingsBtn;

@FindBy(xpath="//*[@id='doc_TagsTableMasterDropDown']")
public static WebElement stateMasterField;

@FindBy(xpath="//*[@id='doc_TagsPositionDropDown']")
public static WebElement statePositionField;



@FindBy(xpath="(//*[@id='doc_TagsTable']//tr[3])[2]")
public static WebElement thirdRowinMaster;


@FindBy(xpath="//td[@id='doc_TagsTable_col_3-1']")
private static WebElement  mastersR3C1;

@FindBy(xpath="//td[@id='doc_TagsTable_col_3-2']")
private static WebElement  mastersPositionR3C2;

public void checkAddingExtraFieldinPurchaseVoucher() throws InterruptedException
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsBtn));
settingsBtn.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(thirdRowinMaster));
//thirdRowinMaster.click();
Thread.sleep(1500);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersR3C1)); 
mastersR3C1.click();


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stateMasterField));
stateMasterField.click();
Select s=new Select(stateMasterField);
s.selectByVisibleText("State");
Thread.sleep(1500);
stateMasterField.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersPositionR3C2));
mastersPositionR3C2.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(statePositionField));
statePositionField.click();
Select s1=new Select(statePositionField);
s1.selectByVisibleText("Header");
Thread.sleep(1500);
statePositionField.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
updateBtn.click();

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsCloseBtn));
settingsCloseBtn.click();
Thread.sleep(2000);


}


public boolean checkDeletetheStateBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

boolean result=checkDeletingBuyerPriceBook("BuyerState1");

if(result==true )
{
return true;
}
else
{
return false;
}

}


///Adding Item Group

@FindBy(xpath="//*[@id='lblnchecked']/input")
public static List<WebElement> accChkBoxList;


@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[12]")
private static List<WebElement> accNameList;


public boolean checkSavingItemGroupalongwithItemsinItemMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
homeMasterItem_ItemMenu.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAddGroupBtn));
masterAddGroupBtn.click();

System.out.println("********************************checkSavingItemGroupItemMaster*****************************");

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
nameTxt.click();
nameTxt.clear();
nameTxt.sendKeys("Grp1");
nameTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
codeTxt.click();
codeTxt.clear();
codeTxt.sendKeys("Grp1");
codeTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemTypeDropdown));
itemTypeDropdown.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
saveBtn.click();

String expMessage="Saved Successfully";

String actMessage=checkValidationMessage(expMessage);

System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);

Thread.sleep(1000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
masterCloseBtn.click();


int accNameListCount = accNameList.size();

for(int i=0;i<accNameListCount;i++)
{
String data = accNameList.get(i).getText();
Thread.sleep(2000);
if(data.equals("Grp1"))
{
Thread.sleep(2000);
getAction().doubleClick(accChkBoxList.get(i)).build().perform();
Thread.sleep(2000);
break;
}
}


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
masterNewBtn.click();

System.out.println("********************************checkSavingItemSubGroupItemMaster*****************************");

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
nameTxt.click();
nameTxt.clear();
nameTxt.sendKeys("item1");
nameTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
codeTxt.click();
codeTxt.clear();
codeTxt.sendKeys("item1");
codeTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemTypeDropdown));
itemTypeDropdown.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
saveBtn.click();

Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
nameTxt.click();
nameTxt.clear();
nameTxt.sendKeys("item2");
nameTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
codeTxt.click();
codeTxt.clear();
codeTxt.sendKeys("item2");
codeTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemTypeDropdown));
itemTypeDropdown.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
saveBtn.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
masterCloseBtn.click();


if(actMessage.equalsIgnoreCase(expMessage))
{
return true;
}
else
{
return false;
}

}


@FindBy(xpath="//*[@id='1112']/span")
public static WebElement StateMaster;

public boolean checkSavingStatesinStateMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();
Thread.sleep(2000);


((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,300)","");
Thread.sleep(2000);

getAction().moveToElement(StateMaster).build().perform();
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StateMaster));
StateMaster.click();

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
masterNewBtn.click();

System.out.println("********************************checkSavingItemSubGroupItemMaster*****************************");

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
nameTxt.click();
nameTxt.clear();
nameTxt.sendKeys("AP");
nameTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
codeTxt.click();
codeTxt.clear();
codeTxt.sendKeys("AP");
codeTxt.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
saveBtn.click();

Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
nameTxt.click();
nameTxt.clear();
nameTxt.sendKeys("Telangana");
nameTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
codeTxt.click();
codeTxt.clear();
codeTxt.sendKeys("Telangana");
codeTxt.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
saveBtn.click();
Thread.sleep(2000);

String expMessage="Saved Successfully";

String actMessage=checkValidationMessage(expMessage);

System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);

Thread.sleep(1000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
masterCloseBtn.click();


if(actMessage.equalsIgnoreCase(expMessage))
{
return true;
}

else
{
return false;
}


}

public boolean checkCreatingTwoPriceBooksforSameItemwithDiffRates() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
  	homeMenu.click();
  	
  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
  	mastersMenu.click();
  
  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
  	homeMasterItemMenu.click();
  	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
	homeMasterItem_buyerPriceBookMenu.click();
	
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	buyerPriceBook_PriceBook.sendKeys("Buyer1");
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	
	
	
	
	/* DateFormat format = new SimpleDateFormat("dd/MM/yyyy");
	
	 Calendar cal1=Calendar.getInstance();
	 cal1.add(Calendar.YEAR,0);
	 cal1.set(Calendar.DAY_OF_YEAR,1);
	 Date start=cal1.getTime();
	 
	 buyerPriceBook_StartDate.sendKeys(format.format(start));

	 Thread.sleep(1000);
	buyerPriceBook_StartDate.sendKeys(Keys.TAB);
		
	System.out.println(format.format(start));
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_EndDateChkBox));
		buyerPriceBook_EndDateChkBox.sendKeys(Keys.TAB);
		Thread.sleep(2000);
	 
	 
	 //cal1.set(Calendar.DAY_OF_YEAR, cal1.getActualMaximum(Calendar.DAY_OF_YEAR));
		Calendar cal11 = Calendar.getInstance();
		  cal11.add(Calendar.DATE, -5);
		  Date d1=cal11.getTime();

	buyerPriceBook_EndDate.sendKeys(format.format(d1));
	Thread.sleep(2000);
	buyerPriceBook_EndDate.sendKeys(Keys.TAB);
*/		
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.SPACE);
	
	int count = itemComboBoxList.size();
	   
	for(int i=0;i<count;i++)
	{
		String data = itemComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("item1"))
		{
			itemComboBoxList.get(i).click();
			break;
		}
	}
	   
	itemValue.sendKeys(Keys.TAB);

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("10");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("100");
	val1Value.sendKeys(Keys.TAB);*/
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	
	System.out.println("*********************************checkSavingSellerPriceBook********************************");
	
	String expMessage="Price Book created successfully";
	
	String actMessage=checkValidationMessage(expMessage);
	
	
	///Saving second PriceBook
	
	Thread.sleep(2500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys("Buyer2");
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_StartDate));
	buyerPriceBook_StartDate.click();
	buyerPriceBook_StartDate.sendKeys(Keys.TAB);
	Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_EndDateChkBox));
		buyerPriceBook_EndDateChkBox.click();
		buyerPriceBook_EndDateChkBox.sendKeys(Keys.TAB);
		Thread.sleep(2000);*/
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.SPACE);
	
	int count1 = itemComboBoxList.size();
	   
	for(int i=0;i<count1;i++)
	{
		String data = itemComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("item1"))
		{
			itemComboBoxList.get(i).click();
			break;
		}
	}
	   
	itemValue.sendKeys(Keys.TAB);

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("15");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("150");
	val1Value.sendKeys(Keys.TAB);
	*/
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	
	System.out.println("*********************************checkSavingBuyerPriceBook********************************");
	
	String expMessage1="Price Book created successfully";
	
	String actMessage1=checkValidationMessage(expMessage1);

	
	

   if (actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1) ) 
   {
	   System.out.println("Pass: Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");
	
	   return true;
   } 
   else 
   {
	   System.out.println("Fail:  Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");
	   
	   return false;
   }
}



public boolean checkRateValueatPurchaseVoucher() throws InterruptedException
{

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	financialsMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchers));
	purchaseVouchers.click();

	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
	newBtn.click();
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	documentNumberTxt.click();
	documentNumberTxt.sendKeys(Keys.TAB);
	Thread.sleep(1500);
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dateTxt));
	dateTxt.sendKeys("11/06/2022");
	Thread.sleep(1500);
	dateTxt.sendKeys(Keys.TAB);
	*/
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
	warehouseTxt.click();
	warehouseTxt.sendKeys("Hyderabad");
	Thread.sleep(1500);
	warehouseTxt.sendKeys(Keys.TAB);
	*/
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.click();
	enter_ItemTxt.sendKeys(Keys.END);
	enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
	enter_ItemTxt.sendKeys(Keys.SPACE);

	int itemcount1=itemListCount.size();

	System.err.println(itemcount1);

	for(int i=0 ; i < itemcount1 ;i++)
	{
		String data=itemListCount.get(i).getText();

		if(data.equalsIgnoreCase("item1"))
		{
			itemListCount.get(i).click();

			break;
		}
	}

	enter_ItemTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
	select1stRow_6thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	String actRate=enter_Rate.getAttribute("value");
	String expRate="15.00";
	
	
	System.out.println("Rate value before current Date");
	System.out.println("Actual Rate :		" +	actRate 		+ "Value Expected:		"	+expRate);
	

	if(actRate.equalsIgnoreCase(expRate) )
	
	{
	
	return true;
	}
	else
	{
		return false;
	}

}

public boolean checkUpdateCreatedPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{


	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
  	homeMenu.click();
  	
  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
  	mastersMenu.click();
  
  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
  	homeMasterItemMenu.click();
  	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
	homeMasterItem_buyerPriceBookMenu.click();
	
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	buyerPriceBook_PriceBook.sendKeys("Buyer1");
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);*/
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	
	Thread.sleep(3000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.SPACE);
	
	int count = itemComboBoxList.size();
	   
	for(int i=0;i<count;i++)
	{
		String data = itemComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("item2"))
		{
			itemComboBoxList.get(i).click();
			break;
		}
	}
	   
	itemValue.sendKeys(Keys.TAB);

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("100");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	
	System.out.println("*********************************checkSavingSellerPriceBook********************************");
	
	String expMessage="Price Book updated successfully";
	
	String actMessage=checkValidationMessage(expMessage);
	


   if (actMessage.equalsIgnoreCase(expMessage)  ) 
   {
	   System.out.println("Pass: Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");
	
	   return true;
   } 
   else 
   {
	   System.out.println("Fail:  Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");
	   
	   return false;
   }

}


public boolean checkRateFieldatPurchaseVoucherafterUpdatingPricebook() throws InterruptedException
{

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	financialsMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchers));
	purchaseVouchers.click();

	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
	newBtn.click();
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	documentNumberTxt.click();
	documentNumberTxt.sendKeys(Keys.TAB);
	Thread.sleep(1500);
	

	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.click();
	enter_ItemTxt.sendKeys(Keys.END);
	enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
	enter_ItemTxt.sendKeys(Keys.SPACE);

	int itemcount1=itemListCount.size();

	System.err.println(itemcount1);

	for(int i=0 ; i < itemcount1 ;i++)
	{
		String data=itemListCount.get(i).getText();

		if(data.equalsIgnoreCase("item1"))
		{
			itemListCount.get(i).click();

			break;
		}
	}

	enter_ItemTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
	select1stRow_6thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	String actRate=enter_Rate.getAttribute("value");
	String expRate="10.00";
	
	
	System.out.println("Rate value before current Date");
	System.out.println("Actual Rate :		" +	actRate 		+ "Value Expected:		"	+expRate);

	if(actRate.equalsIgnoreCase(expRate) )
	
	{
	
	return true;
	}
	else
	{
		return false;
	}

}



public boolean checkDeleteTwoBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	
		boolean result=checkDeletingBuyerPriceBook("Buyer1");
		
		
		boolean result1=checkDeletingBuyerPriceBook("Buyer2");
	
		if(result==true  && result1==true)
		{
			return true;
		}
		else
		{
			return false;
		}
	
}






	public SellerPriceBookPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
