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

public class UnitConversionPage extends BaseEngine 
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "UnitConversionPage";

	public static String checkValidationMessage(String ExpMessage) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		try
		{
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(errorMessage));
			String actErrorMessage=errorMessage.getText();
			String expErrorMessage=ExpMessage;

			/* try
	                {*/
	                	if (validationConfirmationMessage.getText().isEmpty()==false) 
	                	{
	                		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
	                		errorMessageCloseBtn.click();
	                	}

	                	System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);

	                	return actErrorMessage;
	                	/* }
	                catch(Exception ee)
	                {

	                	System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);

	                	return actErrorMessage;
	                }*/
		}
		catch(Exception e)
		{
			System.err.println("Error Message NOT Found or NOT Clickable");
			System.err.println(e.getMessage());

			String Exception=e.getMessage();

			return Exception;
		}
	}






	public boolean checkDeletingUnitConversion() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		homeMasterItemMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_UnitConversionMenu));
		homeMasterItem_UnitConversionMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionBaseUnitTxt));
		unitConversionBaseUnitTxt.sendKeys(excelReader.getCellData(xlSheetName, 8, 5));
		Thread.sleep(2000);
		unitConversionBaseUnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionItemTxt));
		unitConversionItemTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversion_EnterUnitNameTxt));
		String actUnitName = unitConversion_EnterUnitNameTxt.getAttribute("value");
		String expUnitName = excelReader.getCellData(xlSheetName, 8, 6);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col2));
		String actXFactor = unitConversionTableRow1Col2.getText();
		String expXFactor = excelReader.getCellData(xlSheetName, 9, 6);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col5));
		String actDescription = unitConversionTableRow1Col5.getText();
		String expDescription = excelReader.getCellData(xlSheetName, 10, 6);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionDeleteBtn));
		unitConversionDeleteBtn.click();

		Thread.sleep(2000);

		getWaitForAlert();
		getAlert().accept();

		String expMessage=excelReader.getCellData(xlSheetName, 11, 6);

		String actMessage=checkValidationMessage(expMessage);

		System.out.println("***************************************checkDeletingUnitConversion*************************************************");

		System.out.println("Unit Name   : "+actUnitName    +"  Value Expected  "+expUnitName);
		System.out.println("XFactor     : "+actXFactor     +"  Value Expected  "+expXFactor);
		System.out.println("Description : "+actDescription +"  Value Expected  "+expDescription);
		System.out.println("Message     : "+actMessage     +"  Value Expected  "+expMessage);
		
		 excelReader.setCellData(xlfile, xlSheetName, 8, 7, actUnitName);
		 excelReader.setCellData(xlfile, xlSheetName, 9, 7, actXFactor);
		 excelReader.setCellData(xlfile, xlSheetName, 10, 7, actDescription);
		 excelReader.setCellData(xlfile, xlSheetName, 11, 7, actMessage);


		if(actUnitName.equalsIgnoreCase(expUnitName) && actXFactor.equalsIgnoreCase(expXFactor) && actDescription.equalsIgnoreCase(expDescription) 
				&& actMessage.equalsIgnoreCase(expMessage))
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




	public boolean checkSavingUnitConversion() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionBaseUnitTxt));
		unitConversionBaseUnitTxt.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		Thread.sleep(2000);
		unitConversionBaseUnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionItemTxt));
		unitConversionItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col1));
		unitConversionTableRow1Col1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversion_EnterUnitNameTxt));
		unitConversion_EnterUnitNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
		Thread.sleep(2000);
		unitConversion_EnterUnitNameTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversion_EnterXFactorTxt));
		unitConversion_EnterXFactorTxt.sendKeys(excelReader.getCellData(xlSheetName, 14, 5));
		unitConversion_EnterXFactorTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveIcon));
		saveIcon.click();
		Thread.sleep(2000);

		String expMessage=excelReader.getCellData(xlSheetName, 12, 6);

		String actMessage=checkValidationMessage(expMessage);

		System.out.println("***************************************checkSavingUnitConversion*************************************************");

		System.out.println("Message     : "+actMessage     +"  Value Expected  "+expMessage);
		
		 excelReader.setCellData(xlfile, xlSheetName, 12, 7, actMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
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







	public boolean checkUpdatingUnitConversion() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionBaseUnitTxt));
		unitConversionBaseUnitTxt.sendKeys(excelReader.getCellData(xlSheetName, 15, 5));
		Thread.sleep(2000);
		unitConversionBaseUnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionItemTxt));
		unitConversionItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversion_EnterUnitNameTxt));
		String actUnitName = unitConversion_EnterUnitNameTxt.getAttribute("value");
		String expUnitName = excelReader.getCellData(xlSheetName, 15, 6);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col2));
		String actXFactor = unitConversionTableRow1Col2.getText();
		String expXFactor = excelReader.getCellData(xlSheetName, 16, 6);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col5));
		String actDescription = unitConversionTableRow1Col5.getText();
		String expDescription = excelReader.getCellData(xlSheetName, 17, 6);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col2));
		unitConversionTableRow1Col2.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversion_EnterXFactorTxt));
		unitConversion_EnterXFactorTxt.click();
		unitConversion_EnterXFactorTxt.clear();
		unitConversion_EnterXFactorTxt.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));
		unitConversion_EnterXFactorTxt.sendKeys(Keys.TAB);

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveIcon));
		saveIcon.click();
		Thread.sleep(2000);

		String expMessage=excelReader.getCellData(xlSheetName, 18, 6);

		String actMessage=checkValidationMessage(expMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionBaseUnitTxt));
		unitConversionBaseUnitTxt.sendKeys(excelReader.getCellData(xlSheetName, 17, 5));
		Thread.sleep(2000);
		unitConversionBaseUnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionItemTxt));
		unitConversionItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversion_EnterUnitNameTxt));
		String actAfterUpdateUnitName = unitConversion_EnterUnitNameTxt.getAttribute("value");
		String expAfterUpdateUnitName = excelReader.getCellData(xlSheetName, 19, 6);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col2));
		String actAfterUpdateXFactor = unitConversionTableRow1Col2.getText();
		String expAfterUpdateXFactor = excelReader.getCellData(xlSheetName, 20, 6);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col5));
		String actAfterUpdateDescription = unitConversionTableRow1Col5.getText();
		String expAfterUpdateDescription = excelReader.getCellData(xlSheetName, 21, 6);

		System.out.println("***************************************checkUpdatingUnitConversion*************************************************");

		System.out.println("Unit Name   : "+actUnitName    +"  Value Expected  "+expUnitName);
		System.out.println("XFactor     : "+actXFactor     +"  Value Expected  "+expXFactor);
		System.out.println("Description : "+actDescription +"  Value Expected  "+expDescription);
		System.out.println("Message     : "+actMessage     +"  Value Expected  "+expMessage);

		System.out.println("Unit Name AfterUpdate  : "+actAfterUpdateUnitName    +"  Value Expected  "+expAfterUpdateUnitName);
		System.out.println("XFactor AfterUpdate    : "+actAfterUpdateXFactor     +"  Value Expected  "+expAfterUpdateXFactor);
		System.out.println("Description AfterUpdate: "+actAfterUpdateDescription +"  Value Expected  "+expAfterUpdateDescription);
		
		 excelReader.setCellData(xlfile, xlSheetName, 15, 7, actUnitName);
		 excelReader.setCellData(xlfile, xlSheetName, 16, 7, actXFactor);
		 excelReader.setCellData(xlfile, xlSheetName, 17, 7, actDescription);
		 excelReader.setCellData(xlfile, xlSheetName, 18, 7, actMessage);
		 excelReader.setCellData(xlfile, xlSheetName, 19, 7, actAfterUpdateUnitName);
		 excelReader.setCellData(xlfile, xlSheetName, 20, 7, actAfterUpdateXFactor);
		 excelReader.setCellData(xlfile, xlSheetName, 21, 7, actAfterUpdateDescription);
		 

		if(actUnitName.equalsIgnoreCase(expUnitName) && actXFactor.equalsIgnoreCase(expXFactor) && actDescription.equalsIgnoreCase(expDescription) 
				&& actMessage.equalsIgnoreCase(expMessage)
				&& actAfterUpdateUnitName.equalsIgnoreCase(expAfterUpdateUnitName) && actAfterUpdateXFactor.equalsIgnoreCase(expAfterUpdateXFactor) 
				&& actAfterUpdateDescription.equalsIgnoreCase(expAfterUpdateDescription))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionCancelBtn));
			unitConversionCancelBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 15, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionCancelBtn));
			unitConversionCancelBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 15, 8, resFail);
			return false;
		}
	}















	public UnitConversionPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
