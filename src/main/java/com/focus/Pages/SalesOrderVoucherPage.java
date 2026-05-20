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

public class SalesOrderVoucherPage extends BaseEngine 
{
	private static String xlfile = getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlSheetName = "SalesOrderVoucherPage";
	
	@FindBy(xpath="//*[@id='id_FooterClose']")
	public static WebElement editScreenLoadClsBtn;
	
	@FindBy(xpath="//*[@id='editScreen_qtyPreLoaded']")
	public static WebElement clickeditScreen;
	
	
	@FindBy(xpath="//*[@class='icon-settings hiconright2']")
	public static WebElement SettingsBtnSO;
	
	
	
	
	@FindBy(xpath="//*[@id='btnCustomizeClose']")
	public static WebElement closeBtnSO;
	
	
	public boolean checkSalesOrderVoucherCustomization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsSalesMenu));
		inventoryTransactionsSalesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesOrdersVoucher));
		salesOrdersVoucher.click();

		Thread.sleep(4000);

		checkDeleteLinkStatus();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));
		toggleBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsOption));
		settingsOption.click();


		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masters1)); 
		masters1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDropDown));
		masterDropDown.click();

		Select s1=new Select(masterDropDown);
		s1.selectByVisibleText("Warehouse");

		Select getValuesOfMasterC1=new Select(masterDropDown);
		String actMastersR1C1=getValuesOfMasterC1.getFirstSelectedOption().getText();
		String expMastersR1C1="Warehouse";


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(position1));
		position1.click();

		positionDropdown.click();

		Select p1=new Select(positionDropdown);
		p1.selectByVisibleText("Body");

		Select getValueOfPostionC1=new Select(positionDropdown);
		String actPositionR1C1=getValueOfPostionC1.getFirstSelectedOption().getText();
		String expPositionR1C1="Body";*/

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
		editScreenTab.click();

		


		int editScreenbodyCaptionCount = editScreenbodyCaption.size();

		System.err.println("editScreenbodyCaption : "+editScreenbodyCaptionCount);

		for(int i=0;i<editScreenbodyCaptionCount;i++)
		{
			String data = editScreenbodyCaption.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 8, 5)))
			{
				editScreenbodyDeleteBtn.get(i).click();

				break;
			}
		}	

		Thread.sleep(4000);

		getWaitForAlert();
		Thread.sleep(2000);
		getAlert().accept();
		Thread.sleep(6000);

		//checkValidationMessage("");
	
		Thread.sleep(2000);
		int editScreenbodyCaptionCount1 = editScreenbodyCaption.size();

		System.err.println("editScreenbodyCaption : "+editScreenbodyCaptionCount);

		for(int i=0;i<editScreenbodyCaptionCount1;i++)
		{
			String data = editScreenbodyCaption.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 9, 5)))
			{
				editScreenbodyDeleteBtn.get(i).click();

				break;
			}
		}	

		Thread.sleep(4000);

		getWaitForAlert();

		Thread.sleep(2000);
		
		getAlert().accept();
		Thread.sleep(6000);

		//checkValidationMessage("");

		Thread.sleep(2000);

		int editScreenbodyCaptionCount2 = editScreenbodyCaption.size();

		System.err.println("editScreenbodyCaption2 : "+editScreenbodyCaptionCount2);

		for(int i=0;i<editScreenbodyCaptionCount2;i++)
		{
			String data = editScreenbodyCaption.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 10, 5)))
			{
				editScreenbodyDeleteBtn.get(i).click();

				break;
			}
		}	

		Thread.sleep(4000);
		if(getIsAlertPresent())
		{
		getWaitForAlert();
		Thread.sleep(2000);
		getAlert().accept();
		}
		Thread.sleep(6000);

		//checkValidationMessage("");
		Thread.sleep(4000);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenLoadFieldsBtn));
		editScreenLoadFieldsBtn.click();

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenLoadFieldsDropdown));
		editScreenLoadFieldsDropdown.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenLoadFieldsDropdown));
		Select editLayoutLoadFieldsDropdownSelect =new Select(editScreenLoadFieldsDropdown);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenLoadFieldsDropdown));
		editLayoutLoadFieldsDropdownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName, 11, 5));

		Thread.sleep(2000);

		int editScreenLoadFieldTxtListCount = editScreenLoadFieldTxtList.size();

		System.err.println("editScreenLoadFieldTxtListCount : "+editScreenLoadFieldTxtListCount);

		for(int i=0;i<editScreenLoadFieldTxtListCount;i++)
		{
			String data = editScreenLoadFieldTxtList.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 12, 5)))
			{
				editScreenLoadFieldchkboxList.get(i).click();

				break;
			}
		}	

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenokBtn));
		editScreenokBtn.click();

		Thread.sleep(8000);
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenLoadClsBtn));
		editScreenLoadClsBtn.click();
		Thread.sleep(2000);
*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenLoadFieldsBtn));
		editScreenLoadFieldsBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenLoadFieldsDropdown));
		editScreenLoadFieldsDropdown.click();


		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenLoadFieldsDropdown));
		/*Select editLayoutLoadFieldsDropdownSelect =new Select(editScreenLoadFieldsDropdown);*/


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenLoadFieldsDropdown));
		editLayoutLoadFieldsDropdownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName, 13, 5));

		Thread.sleep(2000);


		int editScreenLoadFieldTxtListCount1 = editScreenLoadFieldTxtList.size();

		//System.err.println("editScreenLoadFieldTxtListCount1 : "+editScreenLoadFieldTxtListCount1);

		for(int i=0;i<editScreenLoadFieldTxtListCount1;i++)
		{
			String data = editScreenLoadFieldTxtList.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 14, 5)))
			{
				editScreenLoadFieldchkboxList.get(i).click();

				break;
			}
		}	


		Thread.sleep(4000);

		int editScreenLoadFieldTxtListCount2 = editScreenLoadFieldTxtList.size();

		//System.err.println("editScreenLoadFieldTxtListCount2 : "+editScreenLoadFieldTxtListCount2);

		for(int i=0;i<editScreenLoadFieldTxtListCount2;i++)
		{
			String data = editScreenLoadFieldTxtList.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 15, 5)))
			{
				editScreenLoadFieldchkboxList.get(i).click();

				break;
			}
		}	


		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenokBtn));
		editScreenokBtn.click();

		Thread.sleep(8000);
		
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(behaviourSelect));
		Select s1=new Select(behaviourSelect);
		s1.selectByValue("16");
		Thread.sleep(2000);
		behaviourSelect.sendKeys(Keys.TAB);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtnSO));
		closeBtnSO.click();
		
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsBtnSO));
		SettingsBtnSO.click();
		Thread.sleep(1500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
		editScreenTab.click();
		Thread.sleep(2000);*/
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenQuantityPreloadedTxt));
		//editScreenFormulaTxt.sendKeys("&(SBO3+SBO4)");		
		editScreenQuantityPreloadedTxt.click();
		editScreenQuantityPreloadedTxt.clear();
		
		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		
		/*editScreenFormulaTxt.sendKeys(Keys.TAB);
		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("arguments[0].scrollIntoView(true);", editScreenQuantityPreloadedTxt);
		Thread.sleep(2000);
		
		js.executeScript("arguments[0].click();", editScreenQuantityPreloadedTxt);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenQuantityPreloadedTxt));
		editScreenQuantityPreloadedTxt.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFormulaTxt));
		//editScreenFormulaTxt.sendKeys("&(SBO3+SBO4)");		
		editScreenFormulaTxt.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenQuantityPreloadedTxt));
		editScreenQuantityPreloadedTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(qtyPreloadedformulaTxt));
		qtyPreloadedformulaTxt.click();
*/
		
		//js.executeScript("arguments[0].scrollIntoView(true);", qtyPreloadedformulaTxt);
		
		Thread.sleep(2000);

		String AQValue=null,FQValue=null ;

		int editScreeAtyPreLoadFormulaListCount = editScreeAtyPreLoadFormulaList.size();

		//System.err.println("editScreeAtyPreLoadFormulaListCount : "+editScreeAtyPreLoadFormulaListCount);

		for(int i=0;i<editScreeAtyPreLoadFormulaListCount;i++)
		{
			String data = editScreeAtyPreLoadFormulaList.get(i).getText();
			//System.err.println(data);

			if(data.equalsIgnoreCase("AQ Calculated Value")) /*excelReader.getCellData(xlSheetName, 16, 5)*/
			{
				AQValue=editScreeAtyPreLoadFormulaVariableList.get(i-1).getText();
			//	AQValue=editScreeAtyPreLoadFormulaList.get(i).getText();

				System.out.println("AQValue  : "+AQValue);

				break;
			}
		}

		Thread.sleep(5000);

		for(int i=0;i<editScreeAtyPreLoadFormulaListCount;i++)
		{
			String data = editScreeAtyPreLoadFormulaList.get(i).getText();
			//System.err.println(data);

			if(data.equalsIgnoreCase("FQ Calculated Value"))/*excelReader.getCellData(xlSheetName, 17, 5)*/
			{
				FQValue=editScreeAtyPreLoadFormulaVariableList.get(i-1).getText();

				System.out.println("FQValue  : "+FQValue);

				break;

			}
		}	

	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(qtyPreloadedformulaTxt));
		qtyPreloadedformulaTxt.click();*/
		
		js.executeScript("arguments[0].click();", editScreenQuantityPreloadedTxt);
		editScreenQuantityPreloadedTxt.sendKeys("&("+AQValue+"+"+FQValue+")");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(qtyPreloadedFormulaokBtn));
		qtyPreloadedFormulaokBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryoptionsTab));
		inventoryoptionsTab.click();

		Thread.sleep(3000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryoptionsTabInventorySettingBtn));
		inventoryoptionsTabInventorySettingBtn.click();
		
		Thread.sleep(2000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reservationTypeDropdown));
		reservationTypeDropdown.click();

		Select rev=new Select(reservationTypeDropdown);
		rev.selectByValue(excelReader.getCellData(xlSheetName, 18, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		String expMessage=excelReader.getCellData(xlSheetName, 8, 6);

		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actMessage);
		
		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
		settings_closeBtn.click();

		Thread.sleep(3000);

		
		if(actMessage.equalsIgnoreCase(expMessage))
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







	public boolean checkSavingSalesOrderVoucher1WithBatchItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsSalesMenu));
		inventoryTransactionsSalesMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesOrdersVoucher));
		salesOrdersVoucher.click();

		Thread.sleep(10000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");
		Thread.sleep(6000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 19, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 20, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 21, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
		so_enter_AQTxt.sendKeys(/*excelReader.getCellData(xlSheetName, 22, 5)*/"6");
		so_enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_FQTxt));
		so_enter_FQTxt.sendKeys(/*excelReader.getCellData(xlSheetName, 23, 5)*/"6");
		so_enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
		
		//select1stRow_7thColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		String actQtyInVoucher	=enter_Quantity.getAttribute("value");
		String expQtyInVoucher	="12.00";
	
		enter_Quantity.sendKeys(Keys.TAB);
		

		System.out.println("Qty Preload is displaying as Value In Sales Order Vouche  : "+actQtyInVoucher  +" Value Expected : "+expQtyInVoucher);

		Thread.sleep(2000);
		
		int batchReservationPopGridBodyCount=batchReservationPopGridBody.size();

		ArrayList<String> batchReservationPopGridBodyList = new ArrayList<String>();

		for (int i = 0; i < batchReservationPopGridBodyCount; i++) 
		{
			String data=batchReservationPopGridBody.get(i).getText();

			if(i==8 || i==9 || i==19 || i==20 || i==30 ||i==31 || i==41 || i==42 || i==52 || i==53)
			{
				data = "Date Field";
			}
			//System.err.println("data : " +data);
			batchReservationPopGridBodyList.add(data);
		}

		String actbatchReservationPopGridBody = batchReservationPopGridBodyList.toString();

		String expbatchReservationPopGridBody = "[1,  , 0.00, EX BATCH UPDATE, HYDERABAD, 12.00, 0.00, 0.00, Date Field, Date Field, 12.00, 2,  , 0.00, PVATR2#V1, HYDERABAD, 24.00, 0.00, 0.00, Date Field, Date Field, 24.00, 3,  , 0.00, PVATR1#V1, HYDERABAD, 12.00, 0.00, 0.00, Date Field, Date Field, 12.00, 4,  , 0.00, BATCH#OSVR2, HYDERABAD, 12.00, 0.00, 0.00, Date Field, Date Field, 12.00, 5,  , 0.00, BATCH#OSVR1, HYDERABAD, 12.00, 0.00, 0.00, Date Field, Date Field, 12.00]";
		
		System.out.println(actbatchReservationPopGridBody);
		System.out.println(expbatchReservationPopGridBody);

		int batchReservationPopGridBatchCount = batchReservationPopGridBatch.size();

		ArrayList<String> batchReservationPopGridBatchArray = new ArrayList<String> ();

		for(int i=0;i<batchReservationPopGridBatchCount;i++)
		{
			String data = batchReservationPopGridBatch.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 24, 5)))
			{
				batchReservationPopGridQtyToRelease.get(i).click();

				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_QtyToReserve));
		enter_QtyToReserve.sendKeys(excelReader.getCellData(xlSheetName, 25, 5));
		enter_QtyToReserve.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_OkBtn));
		res_OkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 26, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));

		String actReserveColumn	=select1stRow_8thColumn.getText();
		String expReserveColumn	="Reserved";

		System.out.println("Reservation Column  : "+actReserveColumn+"  Value Expected  "+expReserveColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRsaveBtn));
		MRsaveBtn.click();
		Thread.sleep(2000);
		
		

		
		
		
		String expMessage1="Voucher saved successfully";
		String actMessage = checkValidationMessage(expMessage1);
		
		
		String expMessage2 =": 1";
		
	
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

	
		
		
		if (actQtyInVoucher.equalsIgnoreCase(expQtyInVoucher) && actbatchReservationPopGridBody.equalsIgnoreCase(expbatchReservationPopGridBody) &&
				actReserveColumn.equalsIgnoreCase(expReserveColumn) && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
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






	


	public static boolean checkEditingSalesOrderVoucher1AndAdding2ndRow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocumentNo = documentNumberTxt.getAttribute("value");
		String expDocumentNo = excelReader.getCellData(xlSheetName, 27, 6);
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actDocumentNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomerAccount = newCashBankAccountTxt.getAttribute("value");
		String expCustomerAccount = excelReader.getCellData(xlSheetName, 28, 6);
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actCustomerAccount);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actWarehouse = select1stRow_1stColumn.getText();
		String expWarehouse = excelReader.getCellData(xlSheetName, 29, 6);
		excelReader.setCellData(xlfile, xlSheetName, 29, 7, actWarehouse);

		String actItemName = select1stRow_2ndColumn.getText();
		String expItemName = excelReader.getCellData(xlSheetName, 30, 6);
		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actItemName);

		String actUnits = select1stRow_3rdColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 31, 6);
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actUnits);

		String actAQ = select1stRow_5thColumn.getText();
		String expAQ = excelReader.getCellData(xlSheetName, 32, 6);
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actAQ);

		String actFQ = select1stRow_6thColumn.getText();
		String expFQ = excelReader.getCellData(xlSheetName, 33, 6);
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actFQ);

		String actQuantity = select1stRow_7thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 34, 6);
		excelReader.setCellData(xlfile, xlSheetName, 34, 7, actQuantity);
		
		String actReserve = select1stRow_8thColumn.getText();
		String expReserve = excelReader.getCellData(xlSheetName, 35, 6);
		excelReader.setCellData(xlfile, xlSheetName, 35, 7, actReserve);

		String actRate = select1stRow_10thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 36, 6);
		excelReader.setCellData(xlfile, xlSheetName, 36, 7, actRate);

		String actGross = select1stRow_11thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 37, 6);
		excelReader.setCellData(xlfile, xlSheetName, 37, 7, actGross);
 
		System.out.println("Document Number   : "+actDocumentNo       +" Value Expected "+expDocumentNo);
		System.out.println("Customer Account  : "+actCustomerAccount  +" Value Expected "+expCustomerAccount);
		System.out.println("Warehouse         : "+actWarehouse        +" Value Expected "+expWarehouse);
		System.out.println("Item Name         : "+actItemName         +" Value Expected "+expItemName);
		System.out.println("Units             : "+actUnits            +" Value Expected "+expUnits);
		System.out.println("AQ                : "+actAQ               +" Value Expected "+expAQ);
		System.out.println("FQ                : "+actFQ               +" Value Expected "+expFQ);
		System.out.println("Quantity          : "+actQuantity         +" Value Expected "+expQuantity);
		System.out.println("Reserve           : "+actReserve          +" Value Expected "+expReserve);
		System.out.println("Rate              : "+actRate             +" Value Expected "+expRate);
		System.out.println("Gross             : "+actGross            +" Value Expected "+expGross);


		click(newCashBankAccountTxt);
		newCashBankAccountTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		newCashBankAccountTxt.sendKeys("Customer A");
		Thread.sleep(4000);
		newCashBankAccountTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 27, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 28, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
		select2ndRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
		so_enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 29, 5));
		so_enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_FQTxt));
		so_enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 30, 5));
		so_enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		String actQtyInVoucher	=enter_Quantity.getAttribute("value");
		String expQtyInVoucher	=excelReader.getCellData(xlSheetName, 38, 6);
		
		enter_Quantity.sendKeys(Keys.TAB);
		excelReader.setCellData(xlfile, xlSheetName, 38, 7, actQtyInVoucher);

		System.out.println("Qty Preload is displaying as Value In Sales Order Vouche  : "+actQtyInVoucher  +" Value Expected : "+expQtyInVoucher);

		int batchReservationPopGridBodyCount=batchReservationPopGridBody.size();

		ArrayList<String> batchReservationPopGridBodyList = new ArrayList<String>();

		for (int i = 0; i < batchReservationPopGridBodyCount; i++) 
		{
			String data=batchReservationPopGridBody.get(i).getText();

			if(i==8 || i==9 || i==19 || i==20 || i==30 ||i==31 || i==41 || i==42)
			{
				data = "Date Field";
			}
			//System.err.println("data : " +data);
			batchReservationPopGridBodyList.add(data);
		}

		String actbatchReservationPopGridBody = batchReservationPopGridBodyList.toString();

		String expbatchReservationPopGridBody = "[1,  , 0.00, PVATR2#V1, HYDERABAD, 24.00, 0.00, 0.00, Date Field, Date Field, 24.00, 2,  , 0.00, PVATR1#V1, HYDERABAD, 12.00, 0.00, 0.00, Date Field, Date Field, 12.00, 3,  , 0.00, BATCH#OSVR2, HYDERABAD, 12.00, 0.00, 0.00, Date Field, Date Field, 12.00, 4,  , 0.00, BATCH#OSVR1, HYDERABAD, 12.00, 0.00, 0.00, Date Field, Date Field, 12.00, 5,  ,  ,  ,  ,  ,  ,  ,  ,  ,  ]";
		
		System.out.println(actbatchReservationPopGridBody);
		System.out.println(expbatchReservationPopGridBody);

		int batchReservationPopGridBatchCount = batchReservationPopGridBatch.size();

		ArrayList<String> batchReservationPopGridBatchArray = new ArrayList<String> ();

		for(int i=0;i<50;i++)
		{
			String data = batchReservationPopGridBatch.get(i).getText();

			if(data.equalsIgnoreCase("PVATR2#V1"))
			{
				batchReservationPopGridQtyToRelease.get(i).click();

				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_QtyToReserve));
		enter_QtyToReserve.sendKeys(excelReader.getCellData(xlSheetName, 32, 5));
		enter_QtyToReserve.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_OkBtn));
		res_OkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 33, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));

		String actReserveColumn	=select1stRow_8thColumn.getText();
		String expReserveColumn	="Reserved";

		System.out.println("Reservation Column  : "+actReserveColumn+"  Value Expected  "+expReserveColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRsaveBtn));
		MRsaveBtn.click();
		Thread.sleep(2000);

		
		
		
		
		
		String expMessage1="Voucher saved successfully";
		String actMessage = checkValidationMessage(expMessage1);
			
		String expMessage2 =": 1";
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		
		
		if (actDocumentNo.equalsIgnoreCase(expDocumentNo) && actCustomerAccount.equalsIgnoreCase(expCustomerAccount) &&
				actWarehouse.equalsIgnoreCase(expWarehouse) && actItemName.equalsIgnoreCase(expItemName) && actUnits.equalsIgnoreCase(expUnits) &&
				actAQ.equalsIgnoreCase(expAQ) && actFQ.equalsIgnoreCase(expFQ) && actQuantity.equalsIgnoreCase(expQuantity) && 
				actReserve.equalsIgnoreCase(expReserve) && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross) &&

				actQtyInVoucher.equalsIgnoreCase(expQtyInVoucher) && actbatchReservationPopGridBody.equalsIgnoreCase(expbatchReservationPopGridBody) &&
				actReserveColumn.equalsIgnoreCase(expReserveColumn) && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 27, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 27, 8, resFail);
			return false;
		}
	}








	public boolean checkSavingSalesOrderVoucher2WithBatchItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 43, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 44, 5));
		Thread.sleep(4000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 45, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
		so_enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 46, 5));
		so_enter_AQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_FQTxt));
		so_enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 47, 5));
		so_enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		String actQtyInVoucher	=enter_Quantity.getAttribute("value");
		String expQtyInVoucher	=excelReader.getCellData(xlSheetName, 43, 6);
		enter_Quantity.sendKeys(Keys.TAB);
		excelReader.setCellData(xlfile, xlSheetName, 43, 7, actQtyInVoucher);

		System.out.println("Qty Preload is displaying as Value In Sales Order Vouche  : "+actQtyInVoucher  +" Value Expected : "+expQtyInVoucher);


		int batchReservationPopGridBodyCount=batchReservationPopGridBody.size();

		ArrayList<String> batchReservationPopGridBodyList = new ArrayList<String>();

		for (int i = 0; i < batchReservationPopGridBodyCount; i++) 
		{
			String data=batchReservationPopGridBody.get(i).getText();

			if(i==8 || i==9 || i==19 || i==20 || i==30 ||i==31 || i==41 || i==42)
			{
				data = "Date Field";
			}
			//System.err.println("data : " +data);
			batchReservationPopGridBodyList.add(data);
		}

		String actbatchReservationPopGridBody = batchReservationPopGridBodyList.toString();

		String expbatchReservationPopGridBody = "[1,  , 0.00, PVATR2#V1, HYDERABAD, 14.00, 0.00, 0.00, Date Field, Date Field, 24.00, 2,  , 0.00, PVATR1#V1, HYDERABAD, 12.00, 0.00, 0.00, Date Field, Date Field, 12.00, 3,  , 0.00, BATCH#OSVR2, HYDERABAD, 12.00, 0.00, 0.00, Date Field, Date Field, 12.00, 4,  , 0.00, BATCH#OSVR1, HYDERABAD, 12.00, 0.00, 0.00, Date Field, Date Field, 12.00, 5,  ,  ,  ,  ,  ,  ,  ,  ,  ,  ]";
		
		System.out.println(actbatchReservationPopGridBody);
		System.out.println(expbatchReservationPopGridBody);

		int batchReservationPopGridBatchCount = batchReservationPopGridBatch.size();

		ArrayList<String> batchReservationPopGridBatchArray = new ArrayList<String> ();

		for(int i=0;i<50;i++)
		{
			String data = batchReservationPopGridBatch.get(i).getText();

			if(data.equalsIgnoreCase("PVATR2#V1"))
			{
				batchReservationPopGridQtyToRelease.get(i).click();

				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_QtyToReserve));
		enter_QtyToReserve.sendKeys(Keys.TAB);
		enter_QtyToReserve.sendKeys(excelReader.getCellData(xlSheetName, 49, 5));
		enter_QtyToReserve.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_OkBtn));
		res_OkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 50, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));

		String actReserveColumn	=select1stRow_8thColumn.getText();
		String expReserveColumn	=excelReader.getCellData(xlSheetName, 48, 6);
		excelReader.setCellData(xlfile, xlSheetName, 48, 7, actReserveColumn);

		System.out.println("Reservation Column  : "+actReserveColumn+"  Value Expected  "+expReserveColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRsaveBtn));
		MRsaveBtn.click();

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 2");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 49, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 50, 6));
		excelReader.setCellData(xlfile, xlSheetName, 49, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		/*String expMessage1 = excelReader.getCellData(xlSheetName, 49, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 50, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 49, 7, actMessage);*/
		
		String expMessage1="Voucher saved successfully";
		String actMessage = checkValidationMessage(expMessage1);
		
		//String expMessage2 = excelReader.getCellData(xlSheetName, 26, 6);
		String expMessage2 =": 2";
		
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		if (actQtyInVoucher.equalsIgnoreCase(expQtyInVoucher) && actbatchReservationPopGridBody.equalsIgnoreCase(expbatchReservationPopGridBody) &&
				actReserveColumn.equalsIgnoreCase(expReserveColumn) && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 43, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 43, 8, resFail);
			return false;
		}
	}







	public static boolean checkVerifyingSalesOrderVoucher2AndDeleting() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocumentNo = documentNumberTxt.getAttribute("value");
		String expDocumentNo = excelReader.getCellData(xlSheetName, 51, 6);
		excelReader.setCellData(xlfile, xlSheetName, 51, 7, actDocumentNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomerAccount = newCashBankAccountTxt.getAttribute("value");
		String expCustomerAccount = excelReader.getCellData(xlSheetName, 52, 6);
		excelReader.setCellData(xlfile, xlSheetName, 52, 7, actCustomerAccount);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actWarehouse = select1stRow_1stColumn.getText();
		String expWarehouse = excelReader.getCellData(xlSheetName, 53, 6);
		excelReader.setCellData(xlfile, xlSheetName, 53, 7, actWarehouse);

		String actItemName = select1stRow_2ndColumn.getText();
		String expItemName = excelReader.getCellData(xlSheetName, 54, 6);
		excelReader.setCellData(xlfile, xlSheetName, 54, 7, actItemName);

		String actUnits = select1stRow_3rdColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 55, 6);
		excelReader.setCellData(xlfile, xlSheetName, 55, 7, actUnits);

		String actAQ = select1stRow_5thColumn.getText();
		String expAQ = excelReader.getCellData(xlSheetName, 56, 6);
		excelReader.setCellData(xlfile, xlSheetName, 56, 7, actAQ);

		String actFQ = select1stRow_6thColumn.getText();
		String expFQ = excelReader.getCellData(xlSheetName, 57, 6);
		excelReader.setCellData(xlfile, xlSheetName, 57, 7, actFQ);

		String actQuantity = select1stRow_7thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 58, 6);
		excelReader.setCellData(xlfile, xlSheetName, 58, 7, actQuantity);

		String actReserve = select1stRow_8thColumn.getText();
		String expReserve = excelReader.getCellData(xlSheetName, 59, 6);
		excelReader.setCellData(xlfile, xlSheetName, 59, 7, actReserve);

		String actRate = select1stRow_10thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 60, 6);
		excelReader.setCellData(xlfile, xlSheetName, 60, 7, actRate);

		String actGross = select1stRow_11thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 61, 6);
		excelReader.setCellData(xlfile, xlSheetName, 61, 7, actGross);

		System.out.println("Document Number   : "+actDocumentNo       +" Value Expected "+expDocumentNo);
		System.out.println("Customer Account  : "+actCustomerAccount  +" Value Expected "+expCustomerAccount);
		System.out.println("Warehouse         : "+actWarehouse        +" Value Expected "+expWarehouse);
		System.out.println("Item Name         : "+actItemName         +" Value Expected "+expItemName);
		System.out.println("Units             : "+actUnits            +" Value Expected "+expUnits);
		System.out.println("AQ                : "+actAQ               +" Value Expected "+expAQ);
		System.out.println("FQ                : "+actFQ               +" Value Expected "+expFQ);
		System.out.println("Quantity          : "+actQuantity         +" Value Expected "+expQuantity);
		System.out.println("Reserve           : "+actReserve          +" Value Expected "+expReserve);
		System.out.println("Rate              : "+actRate             +" Value Expected "+expRate);
		System.out.println("Gross             : "+actGross            +" Value Expected "+expGross);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

	/*	getWaitForAlert();
		getAlert().accept();*/
		
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucher_ConfirmYesBtn)); 
		  voucher_ConfirmYesBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 62, 6);
		String actMessage  = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 62, 7, actMessage);

		System.out.println("Deleting Message  : "+actMessage+"  Value Expected  "+expMessage);

		if (actDocumentNo.equalsIgnoreCase(expDocumentNo) && actCustomerAccount.equalsIgnoreCase(expCustomerAccount) &&
				actWarehouse.equalsIgnoreCase(expWarehouse) && actItemName.equalsIgnoreCase(expItemName) && actUnits.equalsIgnoreCase(expUnits) &&
				actAQ.equalsIgnoreCase(expAQ) && actFQ.equalsIgnoreCase(expFQ) && actQuantity.equalsIgnoreCase(expQuantity) && 
				actReserve.equalsIgnoreCase(expReserve) && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross) &&

				actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 51, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 51, 8, resFail);
			return false;
		}
	}







	public boolean checkSavingSalesOrderVoucher2WithStockItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(5500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 63, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 64, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 65, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
		so_enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 66, 5));
		so_enter_AQTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_FQTxt));
		so_enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 67, 5));
		so_enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		String actQtyInVoucher	=enter_Quantity.getAttribute("value");
		String expQtyInVoucher	="12.00";
		enter_Quantity.sendKeys(Keys.TAB);
	

		System.out.println("Qty Preload is displaying as Value In Sales Order Vouche  : "+actQtyInVoucher  +" Value Expected : "+expQtyInVoucher);


		int stockReservationPopGridBodyCount=stockReservationPopGridBody.size();

		ArrayList<String> stockReservationPopGridBodyList = new ArrayList<String>();

		for (int i = 0; i < stockReservationPopGridBodyCount; i++) 
		{
			String data=stockReservationPopGridBody.get(i).getText();
			stockReservationPopGridBodyList.add(data);
		}

		String actStockReservationPopGridBody = stockReservationPopGridBodyList.toString();

		String expStockReservationPopGridBody = "[,  , HYDERABAD, Stock, 0.00, 96.00, 0.00, 96.00, 0.00, 0.00]";
		System.out.println(actStockReservationPopGridBody);
		System.out.println(expStockReservationPopGridBody);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockRes_QtyToReserveRow1));
		stockRes_QtyToReserveRow1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_PickBtn));
		res_PickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_OkBtn));
		res_OkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 68, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));

		String actReserveColumn	=select1stRow_8thColumn.getText();
		String expReserveColumn	="Reserved";

		System.out.println("Reservation Column  : "+actReserveColumn+"  Value Expected  "+expReserveColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRsaveBtn));
		MRsaveBtn.click();

		
		String expMessage1="Voucher saved successfully";
		String actMessage = checkValidationMessage(expMessage1);
		
		String expMessage2 =": 2";
		
		

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage2);

		if (actQtyInVoucher.equalsIgnoreCase(expQtyInVoucher) && actStockReservationPopGridBody.equalsIgnoreCase(expStockReservationPopGridBody) &&
				actReserveColumn.equalsIgnoreCase(expReserveColumn) && actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2)) 
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









	public static boolean checkEditingSalesOrderVoucher2AndAdding2ndRow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocumentNo = documentNumberTxt.getAttribute("value");
		String expDocumentNo = excelReader.getCellData(xlSheetName, 69, 6);
		excelReader.setCellData(xlfile, xlSheetName, 69, 7, actDocumentNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomerAccount = newCashBankAccountTxt.getAttribute("value");
		String expCustomerAccount = excelReader.getCellData(xlSheetName, 70, 6);
		excelReader.setCellData(xlfile, xlSheetName, 70, 7, actCustomerAccount);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actWarehouse = select1stRow_1stColumn.getText();
		String expWarehouse = excelReader.getCellData(xlSheetName, 71, 6);
		excelReader.setCellData(xlfile, xlSheetName, 71, 7, actWarehouse);

		String actItemName = select1stRow_2ndColumn.getText();
		String expItemName = excelReader.getCellData(xlSheetName, 72, 6);
		excelReader.setCellData(xlfile, xlSheetName, 72, 7, actItemName);

		String actUnits = select1stRow_3rdColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 73, 6);
		excelReader.setCellData(xlfile, xlSheetName, 73, 7, actUnits);

		String actAQ = select1stRow_5thColumn.getText();
		String expAQ = excelReader.getCellData(xlSheetName, 74, 6);
		excelReader.setCellData(xlfile, xlSheetName, 74, 7, actAQ);

		String actFQ = select1stRow_6thColumn.getText();
		String expFQ = excelReader.getCellData(xlSheetName, 75, 6);
		excelReader.setCellData(xlfile, xlSheetName, 75, 7, actFQ);

		String actQuantity = select1stRow_7thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 76, 6);
		excelReader.setCellData(xlfile, xlSheetName, 76, 7, actQuantity);

		String actReserve = select1stRow_8thColumn.getText();
		String expReserve = excelReader.getCellData(xlSheetName, 77, 6);
		excelReader.setCellData(xlfile, xlSheetName, 77, 7, actReserve);

		String actRate = select1stRow_10thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 78, 6);
		excelReader.setCellData(xlfile, xlSheetName, 78, 7, actRate);

		String actGross = select1stRow_11thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 79, 6);
		excelReader.setCellData(xlfile, xlSheetName, 79, 7, actGross);

		System.out.println("Document Number   : "+actDocumentNo       +" Value Expected "+expDocumentNo);
		System.out.println("Customer Account  : "+actCustomerAccount  +" Value Expected "+expCustomerAccount);
		System.out.println("Warehouse         : "+actWarehouse        +" Value Expected "+expWarehouse);
		System.out.println("Item Name         : "+actItemName         +" Value Expected "+expItemName);
		System.out.println("Units             : "+actUnits            +" Value Expected "+expUnits);
		System.out.println("AQ                : "+actAQ               +" Value Expected "+expAQ);
		System.out.println("FQ                : "+actFQ               +" Value Expected "+expFQ);
		System.out.println("Quantity          : "+actQuantity         +" Value Expected "+expQuantity);
		System.out.println("Reserve           : "+actReserve          +" Value Expected "+expReserve);
		System.out.println("Rate              : "+actRate             +" Value Expected "+expRate);
		System.out.println("Gross             : "+actGross            +" Value Expected "+expGross);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();

		
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 70, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
		select2ndRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
		so_enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 71, 5));
		so_enter_AQTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_FQTxt));
		so_enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 72, 5));
		so_enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		String actQtyInVoucher	=enter_Quantity.getAttribute("value");
		String expQtyInVoucher	="10.00";
		enter_Quantity.sendKeys(Keys.TAB);
		

		System.out.println("Qty Preload is displaying as Value In Sales Order Vouche  : "+actQtyInVoucher  +" Value Expected : "+expQtyInVoucher);


		int stockReservationPopGridBodyCount=stockReservationPopGridBody.size();

		ArrayList<String> stockReservationPopGridBodyList = new ArrayList<String>();

		for (int i = 0; i < stockReservationPopGridBodyCount; i++) 
		{
			String data=stockReservationPopGridBody.get(i).getText();
			stockReservationPopGridBodyList.add(data);
		}

		String actStockReservationPopGridBody = stockReservationPopGridBodyList.toString();

		String expStockReservationPopGridBody = "[,  , HYDERABAD, Stock, 0.00, 84.00, 12.00, 96.00, 0.00, 0.00]";
		
		System.out.println(actStockReservationPopGridBody);
		System.out.println(expStockReservationPopGridBody);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockRes_QtyToReserveRow1));
		stockRes_QtyToReserveRow1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_PickBtn));
		res_PickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_OkBtn));
		res_OkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 73, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));

		String actReserveColumn	=select1stRow_8thColumn.getText();
		String expReserveColumn	="Reserved";
		System.out.println("Reservation Column  : "+actReserveColumn+"  Value Expected  "+expReserveColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRsaveBtn));
		MRsaveBtn.click();
		
		Thread.sleep(4000);

	
		
		String expMessage1="Voucher saved successfully";
		String actMessage = checkValidationMessage(expMessage1);
		
		
		String expMessage2 =": 2";
		

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage2);

		if (actDocumentNo.equalsIgnoreCase(expDocumentNo) && actCustomerAccount.equalsIgnoreCase(expCustomerAccount) &&
				actWarehouse.equalsIgnoreCase(expWarehouse) && actItemName.equalsIgnoreCase(expItemName) && actUnits.equalsIgnoreCase(expUnits) &&
				actAQ.equalsIgnoreCase(expAQ) && actFQ.equalsIgnoreCase(expFQ) && actQuantity.equalsIgnoreCase(expQuantity) && 
				actReserve.equalsIgnoreCase(expReserve) && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross) &&

				actQtyInVoucher.equalsIgnoreCase(expQtyInVoucher) && actStockReservationPopGridBody.equalsIgnoreCase(expStockReservationPopGridBody) &&
				actReserveColumn.equalsIgnoreCase(expReserveColumn) && actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2))
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







	public boolean checkSavingSalesOrderVoucher3WithStockItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 85, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 86, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 87, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
		so_enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 88, 5));
		so_enter_AQTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_FQTxt));
		so_enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 89, 5));
		so_enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		String actQtyInVoucher	=enter_Quantity.getAttribute("value");
		String expQtyInVoucher	="12.00";
		enter_Quantity.sendKeys(Keys.TAB);
	
		
		System.out.println("Qty Preload is displaying as Value In Sales Order Vouche  : "+actQtyInVoucher  +" Value Expected : "+expQtyInVoucher);


		int stockReservationPopGridBodyCount=stockReservationPopGridBody.size();

		ArrayList<String> stockReservationPopGridBodyList = new ArrayList<String>();

		for (int i = 0; i < stockReservationPopGridBodyCount; i++) 
		{
			String data=stockReservationPopGridBody.get(i).getText();
			stockReservationPopGridBodyList.add(data);
		}

		String actStockReservationPopGridBody = stockReservationPopGridBodyList.toString();

		String expStockReservationPopGridBody = "[, , HYDERABAD, Stock, 0.00, 74.00, 0.00, 96.00, 0.00, 0.00]";
		
		System.out.println(actStockReservationPopGridBody);
		System.out.println(expStockReservationPopGridBody);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockRes_QtyToReserveRow1));
		stockRes_QtyToReserveRow1.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_PickBtn));
		res_PickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_OkBtn));
		res_OkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 90, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));

		String actReserveColumn	=select1stRow_8thColumn.getText();
		String expReserveColumn	="Reserved";
		
		System.out.println("Reservation Column  : "+actReserveColumn+"  Value Expected  "+expReserveColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRsaveBtn));
		MRsaveBtn.click();

		
		String expMessage1="Voucher saved successfully";
		String actMessage = checkValidationMessage(expMessage1);
		
	
		String expMessage2 =": 3";
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
		
		
		if (actQtyInVoucher.equalsIgnoreCase(expQtyInVoucher) && actStockReservationPopGridBody.equalsIgnoreCase(expStockReservationPopGridBody) &&
				actReserveColumn.equalsIgnoreCase(expReserveColumn) && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
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









	public static boolean checkVerifyingSalesOrderVoucher3AndDeleting() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocumentNo = documentNumberTxt.getAttribute("value");
		String expDocumentNo = excelReader.getCellData(xlSheetName, 91, 6);
		excelReader.setCellData(xlfile, xlSheetName, 91, 7, actDocumentNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomerAccount = newCashBankAccountTxt.getAttribute("value");
		String expCustomerAccount = excelReader.getCellData(xlSheetName, 92, 6);
		excelReader.setCellData(xlfile, xlSheetName, 92, 7, actCustomerAccount);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actWarehouse = select1stRow_1stColumn.getText();
		String expWarehouse = excelReader.getCellData(xlSheetName, 93, 6);
		excelReader.setCellData(xlfile, xlSheetName, 93, 7, actWarehouse);
		
		String actItemName = select1stRow_2ndColumn.getText();
		String expItemName = excelReader.getCellData(xlSheetName, 94, 6);
		excelReader.setCellData(xlfile, xlSheetName, 94, 7, actItemName);

		String actUnits = select1stRow_3rdColumn.getText();
		String expUnits = excelReader.getCellData(xlSheetName, 95, 6);
		excelReader.setCellData(xlfile, xlSheetName, 95, 7, actUnits);

		String actAQ = select1stRow_5thColumn.getText();
		String expAQ = excelReader.getCellData(xlSheetName, 96, 6);
		excelReader.setCellData(xlfile, xlSheetName, 96, 7, actAQ);
		
		String actFQ = select1stRow_6thColumn.getText();
		String expFQ = excelReader.getCellData(xlSheetName, 97, 6);
		excelReader.setCellData(xlfile, xlSheetName, 97, 7, actFQ);

		String actQuantity = select1stRow_7thColumn.getText();
		String expQuantity = excelReader.getCellData(xlSheetName, 98, 6);
		excelReader.setCellData(xlfile, xlSheetName, 98, 7, actQuantity);

		String actReserve = select1stRow_8thColumn.getText();
		String expReserve = excelReader.getCellData(xlSheetName, 99, 6);
		excelReader.setCellData(xlfile, xlSheetName, 99, 7, actReserve);

		String actRate = select1stRow_10thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 100, 6);
		excelReader.setCellData(xlfile, xlSheetName, 100, 7, actRate);

		String actGross = select1stRow_11thColumn.getText();
		String expGross = excelReader.getCellData(xlSheetName, 101, 6);
		excelReader.setCellData(xlfile, xlSheetName, 101, 7, actGross);

		System.out.println("Document Number   : "+actDocumentNo       +" Value Expected "+expDocumentNo);
		System.out.println("Customer Account  : "+actCustomerAccount  +" Value Expected "+expCustomerAccount);
		System.out.println("Warehouse         : "+actWarehouse        +" Value Expected "+expWarehouse);
		System.out.println("Item Name         : "+actItemName         +" Value Expected "+expItemName);
		System.out.println("Units             : "+actUnits            +" Value Expected "+expUnits);
		System.out.println("AQ                : "+actAQ               +" Value Expected "+expAQ);
		System.out.println("FQ                : "+actFQ               +" Value Expected "+expFQ);
		System.out.println("Quantity          : "+actQuantity         +" Value Expected "+expQuantity);
		System.out.println("Reserve           : "+actReserve          +" Value Expected "+expReserve);
		System.out.println("Rate              : "+actRate             +" Value Expected "+expRate);
		System.out.println("Gross             : "+actGross            +" Value Expected "+expGross);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

	/*	getWaitForAlert();
		getAlert().accept();*/
		
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucher_ConfirmYesBtn)); 
		  voucher_ConfirmYesBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 102, 6);
		String actMessage  = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 102, 7, actMessage);
		
		System.out.println("Deleting Message  : "+actMessage+"  Value Expected  "+expMessage);

		if (actDocumentNo.equalsIgnoreCase(expDocumentNo) && actCustomerAccount.equalsIgnoreCase(expCustomerAccount) &&
				actWarehouse.equalsIgnoreCase(expWarehouse) && actItemName.equalsIgnoreCase(expItemName) && actUnits.equalsIgnoreCase(expUnits) &&
				actAQ.equalsIgnoreCase(expAQ) && actFQ.equalsIgnoreCase(expFQ) && actQuantity.equalsIgnoreCase(expQuantity) && 
				actReserve.equalsIgnoreCase(expReserve) && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross) &&

				actMessage.equalsIgnoreCase(expMessage))
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







	public boolean checkSavingSalesOrderVoucherWithBinItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 103, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 104, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 105, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
		so_enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 106, 5));
		so_enter_AQTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_FQTxt));
		so_enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		String actQtyInVoucher	=enter_Quantity.getAttribute("value");
		String expQtyInVoucher	="6.00";
		enter_Quantity.sendKeys(Keys.TAB);
		

		System.out.println("Qty Preload is displaying as Value In Sales Order Vouche  : "+actQtyInVoucher  +" Value Expected : "+expQtyInVoucher);


		int binReservationPopGridBodyCount=binReservationPopGridBody.size();

		Set<String> binReservationPopGridBodyList=new HashSet<String>();

		for (int i = 0; i < binReservationPopGridBodyCount; i++) 
		{
			String data=binReservationPopGridBody.get(i).getText();
			binReservationPopGridBodyList.add(data);
		}

		String actbinReservationPopGridBody = binReservationPopGridBodyList.toString();

		String expbinReservationPopGridBody = "[, 0.00, 0, 1, Bin5, 2, Bin4, 51.00, 3, HYDERABAD, 24.00, 4, 5, 6, 7, 8, Available]";
		System.out.println("binReservationPopGridBody Actual   : " +actbinReservationPopGridBody);
		System.out.println("binReservationPopGridBody Expected : " +expbinReservationPopGridBody);

		int binReservationPopGridBinCount = binReservationPopGridBin.size();

		ArrayList<String> binReservationPopGridBinArray = new ArrayList<String> ();

		for(int i=0;i<binReservationPopGridBinCount;i++)
		{
			String data = binReservationPopGridBin.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 107, 5)))
			{
				binReservationPopGridQtyToRelease.get(i).click();

				break;
			}

		}
		
		bin1strow.click();//Bin4


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_PickBtn));
		res_PickBtn.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_OkBtn));
		res_OkBtn.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 108, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));

		String actReserveColumn	=select1stRow_8thColumn.getText();
		String expReserveColumn	="Reserved";

		System.out.println("Reservation Column  : "+actReserveColumn+"  Value Expected  "+expReserveColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		String expMessage1="Voucher saved successfully";
		String actMessage = checkValidationMessage(expMessage1);
		
		
		String expMessage2 =": 3";
		
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
		


		if (actQtyInVoucher.equalsIgnoreCase(expQtyInVoucher) && actbinReservationPopGridBody.equalsIgnoreCase(expbinReservationPopGridBody)
				&& actReserveColumn.equalsIgnoreCase(expReserveColumn) && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 103, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 103, 8, resFail);
			return false;
		}
	}


	
	

	public static boolean checkEditingSalesOrderVoucher3AndAdding2ndRow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocumentNo = documentNumberTxt.getAttribute("value");
		String expDocumentNo = excelReader.getCellData(xlSheetName, 109, 6);
		excelReader.setCellData(xlfile, xlSheetName, 109, 7, actDocumentNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomerAccount = newCashBankAccountTxt.getAttribute("value");
		String expCustomerAccount = excelReader.getCellData(xlSheetName, 110, 6);
		excelReader.setCellData(xlfile, xlSheetName, 110, 7, actCustomerAccount);
		
		int voucherGridBodyListCount = voucherGridBodyList.size();

		ArrayList<String> voucherGridBodyListArray = new ArrayList<String>();

		for(int i=0;i<voucherGridBodyListCount;i++)
		{

			String data= voucherGridBodyList.get(i).getText();

			voucherGridBodyListArray.add(data);
		}

		String actvoucherGridBodyList=voucherGridBodyListArray.toString();

		String expvoucherGridBodyList="[1, HYDERABAD, FIFO COGS ITEM, Pcs, 0.00, 6.00, 0.00, 6.00, Reserved, , 10.00, 60.00, 0.00, 2, , , , , , , , , , , , , 3, , , , , , , , , , , , , 4, , , , , , , , , , , , , 5, , , , , , , , , , , , , 6, , , , , , , , , , , , , 7, , , , , , , , , , , , , 8, , , , , , , , , , , , ]";

		System.out.println("voucherGridBodyList Actual   : " + actvoucherGridBodyList);
		System.out.println("voucherGridBodyList Expected : " + expvoucherGridBodyList);



		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 109, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 110, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
		select2ndRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
		so_enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 111, 5));
		so_enter_AQTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_FQTxt));
		so_enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 112, 5));
		so_enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		String actQtyInVoucher	=enter_Quantity.getAttribute("value");
		String expQtyInVoucher	="6.00";
		enter_Quantity.sendKeys(Keys.TAB);

		System.out.println("Qty Preload is displaying as Value In Sales Order Vouche  : "+actQtyInVoucher  +" Value Expected : "+expQtyInVoucher);


		int binReservationPopGridBodyCount=binReservationPopGridBody.size();

		Set<String> binReservationPopGridBodyList=new HashSet<String>();

		for (int i = 0; i < binReservationPopGridBodyCount; i++) 
		{
			String data=binReservationPopGridBody.get(i).getText();
			binReservationPopGridBodyList.add(data);
		}

		String actbinReservationPopGridBody = binReservationPopGridBodyList.toString();

		String expbinReservationPopGridBody = "[, 6.00, 0.00, 0, 1, Bin5, 2, Bin4, 51.00, 3, HYDERABAD, 24.00, 4, 5, 6, 7, 45.00, 8, Available]";
		
		System.out.println("binReservationPopGridBody Actual   : " +actbinReservationPopGridBody);
		System.out.println("binReservationPopGridBody Expected : " +expbinReservationPopGridBody);

		int binReservationPopGridBinCount = binReservationPopGridBin.size();

		ArrayList<String> binReservationPopGridBinArray = new ArrayList<String> ();

		for(int i=0;i<binReservationPopGridBinCount;i++)
		{
			String data = binReservationPopGridBin.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 113, 5)))
			{
				binReservationPopGridQtyToRelease.get(i).click();

				break;
			}

		}

		bin2ndrow.click();//bin5

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_PickBtn));
		res_PickBtn.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_OkBtn));
		res_OkBtn.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 114, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));

		String actReserveColumn	=select1stRow_8thColumn.getText();
		String expReserveColumn	="Reserved";

		System.out.println("Reservation Column  : "+actReserveColumn+"  Value Expected  "+expReserveColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();


		
		
		String expMessage1="Voucher saved successfully";
		String actMessage = checkValidationMessage(expMessage1);
		
	
		String expMessage2 =": 3";
		
		
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
	


		if (actDocumentNo.equalsIgnoreCase(expDocumentNo) && actCustomerAccount.equalsIgnoreCase(expCustomerAccount) 
				&& actQtyInVoucher.equalsIgnoreCase(expQtyInVoucher)  && actReserveColumn.equalsIgnoreCase(expReserveColumn) 
				&& actvoucherGridBodyList.equalsIgnoreCase(expvoucherGridBodyList) && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 109, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 109, 8, resFail);
			return false;
		}
	}





	public boolean checkSavingSalesOrderVoucher4WithBinItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.click();
		customerAccountTxt.clear();
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 117, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 118, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 119, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
		so_enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 120, 5));
		so_enter_AQTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_FQTxt));
		so_enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		String actQtyInVoucher	=enter_Quantity.getAttribute("value");
		String expQtyInVoucher	="3.00";
		enter_Quantity.sendKeys(Keys.TAB);

		System.out.println("Qty Preload is displaying as Value In Sales Order Vouche  : "+actQtyInVoucher  +" Value Expected : "+expQtyInVoucher);

		int binReservationPopGridBodyCount=binReservationPopGridBody.size();

		Set<String> binReservationPopGridBodyList=new HashSet<String>();

		for (int i = 0; i < binReservationPopGridBodyCount; i++) 
		{
			String data=binReservationPopGridBody.get(i).getText();
			binReservationPopGridBodyList.add(data);
		}

		String actbinReservationPopGridBody = binReservationPopGridBodyList.toString();

		String expbinReservationPopGridBody = "[, 0.00, 18.00, 0, 1, Bin5, 2, Bin4, 51.00, 3, HYDERABAD, 24.00, 4, 5, 6, 7, 45.00, 8, Available]";

		System.out.println("binReservationPopGridBody Actual   : " +actbinReservationPopGridBody);
		System.out.println("binReservationPopGridBody Expected : " +expbinReservationPopGridBody);

		int binReservationPopGridBinCount = binReservationPopGridBin.size();

		ArrayList<String> binReservationPopGridBinArray = new ArrayList<String> ();

		for(int i=0;i<binReservationPopGridBinCount;i++)
		{
			String data = binReservationPopGridBin.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 121, 5)))
			{
				binReservationPopGridQtyToRelease.get(i).click();

				break;
			}

		}
		
		bin1strow.click();//Bin4

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_PickBtn));
		res_PickBtn.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_OkBtn));
		res_OkBtn.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 122, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));

		String actReserveColumn	=select1stRow_8thColumn.getText();
		String expReserveColumn	="Reserved";

		System.out.println("Reservation Column  : "+actReserveColumn+"  Value Expected  "+expReserveColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		Thread.sleep(2000);

		
		String expMessage1="Voucher saved successfully";
		String actMessage = checkValidationMessage(expMessage1);
		
	
		String expMessage2 =": 4";
		
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
	


		if (actQtyInVoucher.equalsIgnoreCase(expQtyInVoucher) && actbinReservationPopGridBody.equalsIgnoreCase(expbinReservationPopGridBody)
				&& actReserveColumn.equalsIgnoreCase(expReserveColumn) && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 117, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 117, 8, resFail);
			return false;
		}
	}




	public static boolean checkVerifyingSalesOrderVoucher4AndDeleting() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocumentNo = documentNumberTxt.getAttribute("value");
		String expDocumentNo = excelReader.getCellData(xlSheetName, 123, 6);
		excelReader.setCellData(xlfile, xlSheetName, 123, 7, actDocumentNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomerAccount = newCashBankAccountTxt.getAttribute("value");
		String expCustomerAccount = excelReader.getCellData(xlSheetName, 124, 6);
		excelReader.setCellData(xlfile, xlSheetName, 124, 7, actCustomerAccount);

		int voucherGridBodyListCount = voucherGridBodyList.size();

		ArrayList<String> voucherGridBodyListArray = new ArrayList<String>();

		for(int i=0;i<voucherGridBodyListCount;i++)
		{

			String data= voucherGridBodyList.get(i).getText();

			voucherGridBodyListArray.add(data);
		}

		String actvoucherGridBodyList=voucherGridBodyListArray.toString();

		String expvoucherGridBodyList="[1, HYDERABAD, FIFO COGS ITEM, Pcs, 0.00, 3.00, 0.00, 3.00, Reserved, , 10.00, 30.00, 0.00, 2, , , , , , , , , , , , , 3, , , , , , , , , , , , , 4, , , , , , , , , , , , , 5, , , , , , , , , , , , , 6, , , , , , , , , , , , , 7, , , , , , , , , , , , , 8, , , , , , , , , , , , ]";
		
		System.out.println("voucherGridBodyList Actual   : " + actvoucherGridBodyList);
		System.out.println("voucherGridBodyList Expected : " + expvoucherGridBodyList);


		System.out.println("Document Number   : "+actDocumentNo       +" Value Expected "+expDocumentNo);
		System.out.println("Customer Account  : "+actCustomerAccount  +" Value Expected "+expCustomerAccount);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		/*getWaitForAlert();
		getAlert().accept();*/
		
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucher_ConfirmYesBtn)); 
		  voucher_ConfirmYesBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 126, 6);
		String actMessage  = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 126, 7, actMessage);
		
		System.out.println("Deleting Message  : "+actMessage+"  Value Expected  "+expMessage);

		if (actDocumentNo.equalsIgnoreCase(expDocumentNo) && actCustomerAccount.equalsIgnoreCase(expCustomerAccount) 
				&& actvoucherGridBodyList.equalsIgnoreCase(expvoucherGridBodyList) && actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 123, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 123, 8, resFail);
			return false;
		}
	}









	public boolean checkSavingSalesOrderVoucherWithRmaItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 127, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 128, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 129, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
		so_enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 130, 5));
		so_enter_AQTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_FQTxt));
		so_enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		String actQtyInVoucher	=enter_Quantity.getAttribute("value");
		String expQtyInVoucher	=excelReader.getCellData(xlSheetName, 127, 6);
		enter_Quantity.sendKeys(Keys.TAB);
		excelReader.setCellData(xlfile, xlSheetName, 127, 7, actQtyInVoucher);

		System.out.println("Qty Preload is displaying as Value In Sales Order Vouche  : "+actQtyInVoucher  +" Value Expected : "+expQtyInVoucher);
		
		Thread.sleep(1000);

		int RmaWarehouseListCount = RmaWarehouseList.size();

		ArrayList<String> RmaWarehouseListArray = new ArrayList<String>();

		for (int i=0;i<RmaWarehouseListCount;i++)
		{
			String data = RmaWarehouseList.get(i).getText();

			RmaWarehouseListArray.add(data);
		}

		String actRmaWarehouseList = RmaWarehouseListArray.toString();
		String expRmaWarehouseList = "[HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD]";

		System.out.println("RmaWarehouseList Actual   : "+actRmaWarehouseList);
		System.out.println("RmaWarehouseList Expected : "+expRmaWarehouseList);

		int RmaRmaNoListCount = RmaRmaNoList.size();

		ArrayList<String> RmaRmaNoListArray = new ArrayList<String>();

		for (int i=0;i<RmaRmaNoListCount;i++)
		{
			String data = RmaRmaNoList.get(i).getText();

			RmaRmaNoListArray.add(data);
		}

		String actRmaRmaNoList = RmaRmaNoListArray.toString();
		String expRmaRmaNoList = "[RMA#R2OSV2, RMA#R2OSV3, RMA#R2OSV4, RMA#R2OSV5, RMA#R2OSV6, RMA#R2OSV7, RMA#R2OSV8, RMA#R2OSV9, RMA#R1OSV10, RMA#R1OSV11, RMA#R1OSV12, RMA#R2OSV10, RMA#R2OSV11, RMA#R2OSV12, RMA#R1PVVAT1, RMA#R1PVVAT2, RMA#R1PVVAT3, RMA#R1PVVAT4, RMA#R1PVVAT5, RMA#R1PVVAT6, RMA#R1PVVAT7, RMA#R1PVVAT8, RMA#R1PVVAT9, RMA#R2PVVAT1, RMA#R2PVVAT2, RMA#R2PVVAT3, RMA#R2PVVAT4, RMA#R2PVVAT5, RMA#R2PVVAT6, RMA#R2PVVAT7, RMA#R2PVVAT8, RMA#R2PVVAT9, Ex Rma Update, RMA#R1PVVAT10, RMA#R1PVVAT11, RMA#R1PVVAT12, RMA#R2PVVAT10, Ex Rma Update1, Ex Rma Update2, Ex Rma Update3, Ex Rma Update4, Ex Rma Update5, Ex Rma Update6, Ex Rma Update7, Ex Rma Update8, Ex Rma Update9]";

		System.out.println("RmaRmaNoList Actual   : "+actRmaRmaNoList);
		System.out.println("RmaRmaNoList Expected : "+expRmaRmaNoList);

		boolean actSelected = false;

		int RmaSelectedListCount = RmaSelectedList.size();

		for(int i=0;i<RmaSelectedListCount;i++)
		{
			actSelected = RmaSelectedList.get(i).isSelected();
		}

		boolean expSelected = false;

		System.out.println("actSelected : "+actSelected);
		System.out.println("expSelected : "+expSelected);

		int ReservePopUpRmaNoListCount = ReservePopUpRmaNoList.size();

		for(int i=0;i<ReservePopUpRmaNoListCount;i++)
		{
			String data = ReservePopUpRmaNoList.get(i).getText();
			System.err.println("DATA :"+data);

			if(data.equalsIgnoreCase("RMA#R2OSV2") || data.equalsIgnoreCase("RMA#R2OSV3") 
					|| data.equalsIgnoreCase("RMA#R2OSV4") || data.equalsIgnoreCase("RMA#R2OSV5") 
					|| data.equalsIgnoreCase("RMA#R2OSV6") || data.equalsIgnoreCase("RMA#R2OSV7") )
			{
				ReservePopUpRmaNoChkBox.get(i).click();
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_OkBtn));
		res_OkBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 137, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));

		String actReserveColumn	=select1stRow_8thColumn.getText();
		String expReserveColumn	="Reserved";

		System.out.println("Reservation Column  : "+actReserveColumn+"  Value Expected  "+expReserveColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		
	
		
		String expMessage1="Voucher saved successfully";
		String actMessage = checkValidationMessage(expMessage1);
		
	
		String expMessage2 =": 4";
		
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
	

		if (actQtyInVoucher.equalsIgnoreCase(expQtyInVoucher) && actRmaWarehouseList.equalsIgnoreCase(expRmaWarehouseList) && actRmaRmaNoList.equalsIgnoreCase(expRmaRmaNoList)
				&& actSelected==expSelected && actReserveColumn.equalsIgnoreCase(expReserveColumn) && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 127, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 127, 8, resFail);
			return false;
		}
	}

	
	


	public static boolean checkEditingSalesOrderVoucher4AndAdding2ndRow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocumentNo = documentNumberTxt.getAttribute("value");
		String expDocumentNo = excelReader.getCellData(xlSheetName, 138, 6);
		excelReader.setCellData(xlfile, xlSheetName, 138, 7, actDocumentNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomerAccount = newCashBankAccountTxt.getAttribute("value");
		String expCustomerAccount = excelReader.getCellData(xlSheetName, 139, 6);
		excelReader.setCellData(xlfile, xlSheetName, 139, 7, actCustomerAccount);

		int voucherGridBodyListCount = voucherGridBodyList.size();

		ArrayList<String> voucherGridBodyListArray = new ArrayList<String>();

		for(int i=0;i<voucherGridBodyListCount;i++)
		{
			String data= voucherGridBodyList.get(i).getText();

			voucherGridBodyListArray.add(data);
		}

		String actvoucherGridBodyList=voucherGridBodyListArray.toString();

		String expvoucherGridBodyList="[1, HYDERABAD, WA COGS ITEM, Pcs, 0.00, 6.00, 0.00, 6.00, Reserved, , 10.00, 60.00, 0.00, 2, , , , , , , , , , , , , 3, , , , , , , , , , , , , 4, , , , , , , , , , , , , 5, , , , , , , , , , , , ]";

		System.out.println("voucherGridBodyList Actual   : " + actvoucherGridBodyList);
		System.out.println("voucherGridBodyList Expected : " + expvoucherGridBodyList);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 138, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 139, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
		select2ndRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
		so_enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 140, 5));
		so_enter_AQTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_FQTxt));
		so_enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		String actQtyInVoucher	=enter_Quantity.getAttribute("value");
		String expQtyInVoucher	=excelReader.getCellData(xlSheetName, 141, 6);
		enter_Quantity.sendKeys(Keys.TAB);
		excelReader.setCellData(xlfile, xlSheetName, 141, 7, actQtyInVoucher);

		if(validationConfirmationMessage.getText().isEmpty()==false)
		{
			errorMessageCloseBtn.click();
		}

		System.out.println("Qty Preload is displaying as Value In Sales Order Vouche  : "+actQtyInVoucher  +" Value Expected : "+expQtyInVoucher);
		Thread.sleep(1000);

		int RmaWarehouseListCount = RmaWarehouseList.size();

		ArrayList<String> RmaWarehouseListArray = new ArrayList<String>();

		for (int i=0;i<RmaWarehouseListCount;i++)
		{
			String data = RmaWarehouseList.get(i).getText();

			RmaWarehouseListArray.add(data);
		}

		String actRmaWarehouseList = RmaWarehouseListArray.toString();
		String expRmaWarehouseList = "[HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD]";

		System.out.println("RmaWarehouseList Actual   : "+actRmaWarehouseList);
		System.out.println("RmaWarehouseList Expected : "+expRmaWarehouseList);

		int RmaRmaNoListCount = RmaRmaNoList.size();

		ArrayList<String> RmaRmaNoListArray = new ArrayList<String>();

		for (int i=0;i<RmaRmaNoListCount;i++)
		{
			String data = RmaRmaNoList.get(i).getText();

			RmaRmaNoListArray.add(data);
		}

		String actRmaRmaNoList = RmaRmaNoListArray.toString();
		String expRmaRmaNoList = "[RMA#R2OSV8, RMA#R2OSV9, RMA#R1OSV10, RMA#R1OSV11, RMA#R1OSV12, RMA#R2OSV10, RMA#R2OSV11, RMA#R2OSV12, RMA#R1PVVAT1, RMA#R1PVVAT2, RMA#R1PVVAT3, RMA#R1PVVAT4, RMA#R1PVVAT5, RMA#R1PVVAT6, RMA#R1PVVAT7, RMA#R1PVVAT8, RMA#R1PVVAT9, RMA#R2PVVAT1, RMA#R2PVVAT2, RMA#R2PVVAT3, RMA#R2PVVAT4, RMA#R2PVVAT5, RMA#R2PVVAT6, RMA#R2PVVAT7, RMA#R2PVVAT8, RMA#R2PVVAT9, Ex Rma Update, RMA#R1PVVAT10, RMA#R1PVVAT11, RMA#R1PVVAT12, RMA#R2PVVAT10, Ex Rma Update1, Ex Rma Update2, Ex Rma Update3, Ex Rma Update4, Ex Rma Update5, Ex Rma Update6, Ex Rma Update7, Ex Rma Update8, Ex Rma Update9]";

		System.out.println("RmaRmaNoList Actual   : "+actRmaRmaNoList);
		System.out.println("RmaRmaNoList Expected : "+expRmaRmaNoList);

		boolean actSelected = false;

		int RmaSelectedListCount = RmaSelectedList.size();

		for(int i=0;i<RmaSelectedListCount;i++)
		{
			actSelected = RmaSelectedList.get(i).isSelected();
		}

		boolean expSelected = false;

		System.out.println("actSelected : "+actSelected);
		System.out.println("expSelected : "+expSelected);

		int ReservePopUpRmaNoListCount = ReservePopUpRmaNoList.size();

		for(int i=0;i<ReservePopUpRmaNoListCount;i++)
		{
			String data = ReservePopUpRmaNoList.get(i).getText();
			System.err.println("DATA :"+data);

			if(data.equalsIgnoreCase("RMA#R2OSV8") || data.equalsIgnoreCase("RMA#R2OSV9") 
					|| data.equalsIgnoreCase("RMA#R2OSV10"))
			{
				ReservePopUpRmaNoChkBox.get(i).click();
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_OkBtn));
		res_OkBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 144, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_8thColumn));

		String actReserveColumn	=select2ndRow_8thColumn.getText();
		String expReserveColumn	="Reserved";

		System.out.println("Reservation Column  : "+actReserveColumn+"  Value Expected  "+expReserveColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		if(validationConfirmationMessage.getText().isEmpty()==false)
		{
			errorMessageCloseBtn.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		Thread.sleep(2000);

		String expMessage1="Voucher saved successfully";
		String actMessage = checkValidationMessage(expMessage1);
		
		
		String expMessage2 =": 4";
		
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
	


		if (actQtyInVoucher.equalsIgnoreCase(expQtyInVoucher) && actRmaWarehouseList.equalsIgnoreCase(expRmaWarehouseList) && actRmaRmaNoList.equalsIgnoreCase(expRmaRmaNoList)
				&& actSelected==expSelected && actReserveColumn.equalsIgnoreCase(expReserveColumn) && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 138, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 138, 8, resFail);
			return false;
		}
	}


	
	
	
	


	public boolean checkSavingSalesOrderVoucher5WithRmaItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 148, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 149, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 150, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
		so_enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 151, 5));
		so_enter_AQTxt.sendKeys(Keys.TAB);

		checkValidationMessage("");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_FQTxt));
		so_enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		String actQtyInVoucher	=enter_Quantity.getAttribute("value");
		String expQtyInVoucher	="3.00";
		enter_Quantity.sendKeys(Keys.TAB);
		

		System.out.println("Qty Preload is displaying as Value In Sales Order Vouche  : "+actQtyInVoucher  +" Value Expected : "+expQtyInVoucher);
		Thread.sleep(1000);

		int RmaWarehouseListCount = RmaWarehouseList.size();

		ArrayList<String> RmaWarehouseListArray = new ArrayList<String>();

		for (int i=0;i<RmaWarehouseListCount;i++)
		{
			String data = RmaWarehouseList.get(i).getText();

			RmaWarehouseListArray.add(data);
		}

		String actRmaWarehouseList = RmaWarehouseListArray.toString();
		String expRmaWarehouseList = "[HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD, HYDERABAD]";

		System.out.println("RmaWarehouseList Actual   : "+actRmaWarehouseList);
		System.out.println("RmaWarehouseList Expected : "+expRmaWarehouseList);

		int RmaRmaNoListCount = RmaRmaNoList.size();

		ArrayList<String> RmaRmaNoListArray = new ArrayList<String>();

		for (int i=0;i<RmaRmaNoListCount;i++)
		{
			String data = RmaRmaNoList.get(i).getText();

			RmaRmaNoListArray.add(data);
		}

		String actRmaRmaNoList = RmaRmaNoListArray.toString();
		String expRmaRmaNoList = "[RMA#R1OSV10, RMA#R1OSV11, RMA#R1OSV12, RMA#R2OSV11, RMA#R2OSV12, RMA#R1PVVAT1, RMA#R1PVVAT2, RMA#R1PVVAT3, RMA#R1PVVAT4, RMA#R1PVVAT5, RMA#R1PVVAT6, RMA#R1PVVAT7, RMA#R1PVVAT8, RMA#R1PVVAT9, RMA#R2PVVAT1, RMA#R2PVVAT2, RMA#R2PVVAT3, RMA#R2PVVAT4, RMA#R2PVVAT5, RMA#R2PVVAT6, RMA#R2PVVAT7, RMA#R2PVVAT8, RMA#R2PVVAT9, Ex Rma Update, RMA#R1PVVAT10, RMA#R1PVVAT11, RMA#R1PVVAT12, RMA#R2PVVAT10, Ex Rma Update1, Ex Rma Update2, Ex Rma Update3, Ex Rma Update4, Ex Rma Update5, Ex Rma Update6, Ex Rma Update7, Ex Rma Update8, Ex Rma Update9]";

		System.out.println("RmaRmaNoList Actual   : "+actRmaRmaNoList);
		System.out.println("RmaRmaNoList Expected : "+expRmaRmaNoList);

		boolean actSelected = false;

		int RmaSelectedListCount = RmaSelectedList.size();

		for(int i=0;i<RmaSelectedListCount;i++)
		{
			actSelected = RmaSelectedList.get(i).isSelected();
		}

		boolean expSelected = false;

		System.out.println("actSelected : "+actSelected);
		System.out.println("expSelected : "+expSelected);

		int ReservePopUpRmaNoListCount = ReservePopUpRmaNoList.size();

		for(int i=0;i<ReservePopUpRmaNoListCount;i++)
		{
			String data = ReservePopUpRmaNoList.get(i).getText();
			System.err.println("DATA :"+data);

			if(data.equalsIgnoreCase("RMA#R1OSV10") || data.equalsIgnoreCase("RMA#R1OSV11") 
					|| data.equalsIgnoreCase("RMA#R1OSV12"))
			{
				ReservePopUpRmaNoChkBox.get(i).click();
			}
		}


		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_OkBtn));
		res_OkBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 155, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));

		String actReserveColumn	=select1stRow_8thColumn.getText();
		String expReserveColumn	="Reserved";

		System.out.println("Reservation Column  : "+actReserveColumn+"  Value Expected  "+expReserveColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

	
		String expMessage1="Voucher saved successfully";
		String actMessage = checkValidationMessage(expMessage1);
		
	
		String expMessage2 =": 5";
		
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
	

		if (actQtyInVoucher.equalsIgnoreCase(expQtyInVoucher) && actRmaWarehouseList.equalsIgnoreCase(expRmaWarehouseList) 
				&& actRmaRmaNoList.equalsIgnoreCase(expRmaRmaNoList) && actSelected==expSelected
				&& actReserveColumn.equalsIgnoreCase(expReserveColumn) && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 148, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 148, 8, resFail);
			return false;
		}
	}

	
	
	
	

	public static boolean checkVerifyingSalesOrderVoucher5AndDeleting() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("");

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocumentNo = documentNumberTxt.getAttribute("value");
		String expDocumentNo = excelReader.getCellData(xlSheetName, 156, 6);
		excelReader.setCellData(xlfile, xlSheetName, 156, 7, actDocumentNo);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomerAccount = newCashBankAccountTxt.getAttribute("value");
		String expCustomerAccount = excelReader.getCellData(xlSheetName, 157, 6);
		excelReader.setCellData(xlfile, xlSheetName, 157, 7, actCustomerAccount);

		int voucherGridBodyListCount = voucherGridBodyList.size();

		ArrayList<String> voucherGridBodyListArray = new ArrayList<String>();

		for(int i=0;i<voucherGridBodyListCount;i++)
		{
			String data= voucherGridBodyList.get(i).getText();

			voucherGridBodyListArray.add(data);
		}

		String actvoucherGridBodyList=voucherGridBodyListArray.toString();

		String expvoucherGridBodyList="[1, HYDERABAD, WA COGS ITEM, Pcs, 0.00, 3.00, 0.00, 3.00, Reserved, , 10.00, 30.00, 0.00, 2, , , , , , , , , , , , , 3, , , , , , , , , , , , , 4, , , , , , , , , , , , , 5, , , , , , , , , , , , , 6, , , , , , , , , , , , , 7, , , , , , , , , , , , , 8, , , , , , , , , , , , ]";		

		System.out.println("voucherGridBodyList Actual   : " + actvoucherGridBodyList);
		System.out.println("voucherGridBodyList Expected : " + expvoucherGridBodyList);


		System.out.println("Document Number   : "+actDocumentNo       +" Value Expected "+expDocumentNo);
		System.out.println("Customer Account  : "+actCustomerAccount  +" Value Expected "+expCustomerAccount);

	/*	for (int i = 0; i < 2; i++)
		{
			if (validationConfirmationMessage.getText().isEmpty()==false) 
			{
				errorMessageCloseBtn.click();
			}	
		}
*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();

		/*getWaitForAlert();
		getAlert().accept();*/
		
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucher_ConfirmYesBtn)); 
		  voucher_ConfirmYesBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 159, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 159, 7, actMessage);

		System.out.println("Deleting Message  : "+actMessage+"  Value Expected  "+expMessage);

		Thread.sleep(2000);

		if (actDocumentNo.equalsIgnoreCase(expDocumentNo) && actCustomerAccount.equalsIgnoreCase(expCustomerAccount) 
				&& actvoucherGridBodyList.equalsIgnoreCase(expvoucherGridBodyList) && actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 156, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 156, 8, resFail);
			return false;
		}
	}



	
	
	


	public boolean checkSavingAndSuspendingSalesOrderVoucher5FromEntryPage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 160, 5));
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 161, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 162, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
		so_enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 163, 5));
		so_enter_AQTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_FQTxt));
		so_enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 164, 5));
		so_enter_FQTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockRes_QtyToReserveRow1));
		stockRes_QtyToReserveRow1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_PickBtn));
		res_PickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_OkBtn));
		res_OkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 165, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));

		String actReserveColumn	=select1stRow_8thColumn.getText();
		String expReserveColumn	=excelReader.getCellData(xlSheetName, 160, 6);
		excelReader.setCellData(xlfile, xlSheetName, 160, 7, actReserveColumn);

		System.out.println("Reservation Column  : "+actReserveColumn+"  Value Expected  "+expReserveColumn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRsaveBtn));
		MRsaveBtn.click();

		
		
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": 5";
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		checkValidationMessage("Voucher loaded successfully");
		
		Thread.sleep(6000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
		new_SuspendBtn.click();

		String expSuspendMessage1="Voucher saved successfully";
		String expSuspendMessage2=": 5";

		String actSuspendMessage=checkValidationMessage(expSuspendMessage1);
		
		excelReader.setCellData(xlfile, xlSheetName, 163, 7, actSuspendMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
		Thread.sleep(6000);
		
		String actSuspendStatus = voucherHomeRow1SuspendedStatus.getText();
		String expSuspendStatus = "True";
		
		System.out.println("Suspended Status : "+actSuspendStatus+"  Value Expected  "+expSuspendStatus);

		if (actReserveColumn.equalsIgnoreCase(expReserveColumn) && actSuspendMessage.startsWith(expSuspendMessage1) && actSuspendMessage.endsWith(expSuspendMessage2)
				&& actSuspendStatus.equalsIgnoreCase(expSuspendStatus) && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 160, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 160, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	public boolean checkResavingSalesOrderVoucher5AndSuspendingFromHomepage() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 166, 5);

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
		expMessage.add("Voucher saved successfully : 5");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 166, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 167, 6));
		excelReader.setCellData(xlfile, xlSheetName, 166, 7, actMessage);
		
		String expMessage = expMsg.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
	/*	String expMessage1 = excelReader.getCellData(xlSheetName, 166, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 167, 6);
		
		
		excelReader.setCellData(xlfile, xlSheetName, 166, 7, actMessage);*/
		
		String expMessage1="Voucher saved successfully";
		String actMessage = checkValidationMessage(expMessage1);
		
		//String expMessage2 = excelReader.getCellData(xlSheetName, 26, 6);
		String expMessage2 =": 5";
		
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
		//actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		
		Thread.sleep(2000);

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 167, 5);

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
		
		String expSuspendMessage = excelReader.getCellData(xlSheetName, 168, 6);
		String actSuspendMessage = checkValidationMessage(expSuspendMessage);
		excelReader.setCellData(xlfile, xlSheetName, 168, 7, actSuspendMessage);
		Thread.sleep(2000);
		
		String actSuspendStatus = voucherHomeRow1SuspendedStatus.getText();
		String expSuspendStatus = excelReader.getCellData(xlSheetName, 169, 6);
		excelReader.setCellData(xlfile, xlSheetName, 169, 7, actSuspendStatus);
		
		System.out.println("Suspended Status : "+actSuspendStatus+"  Value Expected  "+expSuspendStatus);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2) && actSuspendMessage.equalsIgnoreCase(expSuspendMessage)
				&& actSuspendStatus.equalsIgnoreCase(expSuspendStatus))
		{
			excelReader.setCellData(xlfile, xlSheetName, 166, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 166, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	public boolean checkDeleteOptionInSalesOrderHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 170, 5);

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

		String expAlertMessage=excelReader.getCellData(xlSheetName, 170, 6);

		excelReader.setCellData(xlfile, xlSheetName, 170, 7, actAlertMessage);

		System.out.println("Alert Message on Deleteing the Voucher :  "+actAlertMessage +"  "+expAlertMessage);

		Thread.sleep(3000);

		getAlert().accept();

		String expMessage = excelReader.getCellData(xlSheetName, 171, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 171, 7, actMessage);

		System.out.println("Message : "+actMessage+" Value Expected : "+expMessage);

		if(actAlertMessage.contains(expAlertMessage) && actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 170, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 170, 8, resFail);
			return false;
		}
	}
	
	
	
	



	public boolean checkVerifingDetailsOfSavedSalesOrderVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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

			System.out.println(i+" "+data);

			if (data.isEmpty()==false && i!=9 && i!=10 && i!=23 && i!=24 && i!=37 && i!=38 && i!=51 && i!=52) 
			{
				actVouchersList.add(data);
			}
		}


		ArrayList<String>  expVouchersList= new ArrayList<String>(); 

		expVouchersList.add(excelReader.getCellData(xlSheetName, 172, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 173, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 174, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 175, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 176, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 177, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 178, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 179, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 180, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 181, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 182, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 183, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 184, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 185, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 186, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 187, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 188, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 189, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 190, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 191, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 192, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 193, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 194, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 195, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 196, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 197, 6));

		excelReader.setCellData(xlfile, xlSheetName, 172, 7, actVouchersList.toString());
		
		System.out.println("**********************checkVerifingDetailsOfSavedSalesOrderVoucherInHomepage*********************");

		System.out.println(actVouchersList);
		System.out.println(expVouchersList);

		if (actVouchersList.equals(expVouchersList)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 172, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 172, 8, resFail);
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

	        //getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	        //companyLogo.click();

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


	public boolean checkLogoutSalesOrdervoucherPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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






	public SalesOrderVoucherPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
