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
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashSet;
import java.util.List;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
//import org.bson.assertions.Assertions;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;

import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

import com.testautomationguru.utility.PDFUtil;

public class SellerPriceBookValidationPage extends BaseEngine{
	
	
	@FindBy(xpath="//span[@class='icon-financial icon-font1']")
	private static WebElement  finacinalsMenu;

	@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[12]/a")
	private static WebElement  settingsBtnSalesInvoice;

	
	@FindBy(xpath="//*[@id='id_transactionentry_more']")
	private static WebElement  salesInvoice_MoreBtn;
	
	
	@FindBy(xpath="//*[@id='id_transactionentry_settings']")
	private static WebElement  salesInvoice_SettingsBtn;
	
	
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
	
	
	public static void checkOpenSalesInvoiceVoucher() throws InterruptedException
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
	}
	
	
	public static boolean checkSellerPriceBookwithNoTagRateValueatSalesInvoiceVoucher() throws InterruptedException
	{
		

		checkOpenSalesInvoiceVoucher();
		Thread.sleep(2000);
	
	if(checkAddingFormulaforAccount())
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	}

	
	Thread.sleep(2000);
	
	
	
	
	//BR COGS ITEM
	String actRate=checkRateValue("BR COGS ITEM");
	String expRate="10.00";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt=enter_ItemTxt.getAttribute("value");
	

	//WA COGS ITEM
	
	String actRate1=checkRateValue("WA COGS ITEM");
	String expRate1="20.00";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt1=enter_ItemTxt.getAttribute("value");
	


	//FIFO COGS ITEM
	
	String actRate2=checkRateValue("FIFO COGS ITEM");
	String expRate2="30.00";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt2=enter_ItemTxt.getAttribute("value");
	
	
	
	//STD RATE COGS ITEM
	
		String actRate3=checkRateValue("BIN RAW MATERIAL ITEM");
		String expRate3="0.00";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		String item_Txt3=enter_ItemTxt.getAttribute("value");
		
	
	
	
	System.out.println("Item Name: 	"+ item_Txt				+	"Actual Rate :		" +	actRate 		+ "Value Expected:		"	+expRate);
	System.out.println("Item Name: 	"+ item_Txt1			+	"Actual Rate :		" +	actRate1 		+ "Value Expected:		"	+expRate1);
	System.out.println("Item Name: 	"+ item_Txt2			+	"Actual Rate :		" +	actRate2 		+ "Value Expected:		"	+expRate2);
	System.out.println("Item Name: 	"+ item_Txt3			+	"Actual Rate :		" +	actRate3 		+ "Value Expected:		"	+expRate3);
	
	if(actRate.equalsIgnoreCase(expRate) && actRate1.equalsIgnoreCase(expRate1) && actRate2.equalsIgnoreCase(expRate2) && actRate3.equalsIgnoreCase(expRate3))

	{

	return true;
	}
	else
	{
	return false;
	}
		
	}
	
	
	public static String checkRateValue(String item) throws InterruptedException
	{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	
		enter_ItemTxt.sendKeys(item);
		Thread.sleep(2000);

		/*int itemcount1=itemListCount.size();

		
		for(int i=0 ; i < itemcount1 ;i++)
		{
		String data=itemListCount.get(i).getText();

		if(data.equalsIgnoreCase(item))
		{
		itemListCount.get(i).click();

		break;
		}
		}*/

		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		Thread.sleep(1000);

		String actRate=enter_Rate.getAttribute("value");
		return actRate;
		
		
	}
	
	
	
	public static String checkRateValue(String item,int qty) throws InterruptedException
	{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	
		enter_ItemTxt.sendKeys(item);
		Thread.sleep(2000);
	
		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
		select1stRow_2ndColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		//enter_UnitTxt.click();
		enter_UnitTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		//Thread.sleep(1500);
		enter_UnitTxt.sendKeys(Keys.DELETE);
		Thread.sleep(1500);
		enter_UnitTxt.sendKeys(Keys.ESCAPE);
		Thread.sleep(2000);
		enter_UnitTxt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.sendKeys(Keys.END);
		enter_Quantity.sendKeys(Keys.SHIFT,Keys.HOME);
		
		enter_Quantity.sendKeys(Integer.toString(qty));
		Thread.sleep(2000);
		enter_Quantity.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		Thread.sleep(1000);

		String actRate=enter_Rate.getAttribute("value");
		return actRate;
		
		
	}
	
	
	
	
	public static String checkRateValue(String Customer,String item) throws InterruptedException
	{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.click();
	
		customerAccountTxt.sendKeys(Keys.END);
		customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		customerAccountTxt.sendKeys(Customer);
		/*//customerAccountTxt.sendKeys(Keys.SPACE);

		int customercount=customerAccountListCount.size();

		System.err.println(customercount);

		for(int i=0 ; i < customercount ;i++)
		{
		String data=customerAccountListCount.get(i).getText();

		if(data.equalsIgnoreCase(Customer))
		{
		customerAccountListCount.get(i).click();
		Thread.sleep(2000);

		break;
		}
		}
*/
		customerAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		/*enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(Keys.SPACE);*/
		
		enter_ItemTxt.sendKeys(item);
		Thread.sleep(2000);

	/*	int itemcount1=itemListCount.size();

		
		for(int i=0 ; i < itemcount1 ;i++)
		{
		String data=itemListCount.get(i).getText();

		if(data.equalsIgnoreCase(item))
		{
		itemListCount.get(i).click();

		break;
		}
		}*/

		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_UnitTxt.sendKeys(Keys.DELETE);
		enter_UnitTxt.sendKeys(Keys.ESCAPE);
		Thread.sleep(2000);
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		Thread.sleep(1000);

		String actRate=enter_Rate.getAttribute("value");
	

		return actRate;
		
		
	}
	
	public static String checkRateValueforDepartment(String Dept,String item) throws InterruptedException
	{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.END);
		departmentTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		departmentTxt.sendKeys(Keys.SPACE);

		departmentTxt.sendKeys(Dept);
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);	
		Thread.sleep(2000);
		
		
		if(getIsAlertPresent())
		{
			getWaitForAlert();
			Thread.sleep(2000);
			getAlert().accept();
			Thread.sleep(2000);
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		/*enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(Keys.SPACE);*/
		Thread.sleep(2000);
		
		enter_ItemTxt.sendKeys(item);
		
/*
		int itemcount1=itemListCount.size();

		
		for(int i=0 ; i < itemcount1 ;i++)
		{
		String data=itemListCount.get(i).getText();

		if(data.equalsIgnoreCase(item))
		{
		itemListCount.get(i).click();

		break;
		}
		}*/

		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		Thread.sleep(1000);

		String actRate=enter_Rate.getAttribute("value");
	

		return actRate;
		
		
	}
	
	
	public static boolean checkAddingFormulaforAccount() throws InterruptedException
	{
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
		editScreenFormulaTxtBox.sendKeys("&salrt(0)");
		Thread.sleep(1000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFormulaOkBtn));
		editScreenFormulaOkBtn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsUpdateBtn));
		settingsUpdateBtn.click();
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsCloseBtn));
		settingsCloseBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		
		if(newBtn.isDisplayed()==true)
		{
			return true;
		}
		else
		{
			return false;
		}
		
		
	}
	
	
	
	public static boolean checkAddingFormulaforGroupAccount() throws InterruptedException
	{
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
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsUpdateBtn));
		settingsUpdateBtn.click();
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsCloseBtn));
		settingsCloseBtn.click();

		Thread.sleep(2000);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		
		if(newBtn.isDisplayed()==true)
		{
			return true;
		}
		else
		{
			return false;
		}
		
		
	}
	
	
	public static boolean checkAddingFormulaforUnits() throws InterruptedException
	{
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
		editScreenFormulaTxtBox.sendKeys("&((qtyb>0)&&(qtyb<=12)?salrt(0):((qtyb>12)&&(qtyb<=24)?salrt(1):((qtyb>24)&&(qtyb<=48)?salrt(2):salrt(3))))");
		Thread.sleep(1000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFormulaOkBtn));
		editScreenFormulaOkBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsUpdateBtn));
		settingsUpdateBtn.click();
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsCloseBtn));
		settingsCloseBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		
		if(newBtn.isDisplayed()==true)
		{
			return true;
		}
		else
		{
			return false;
		}
		
		
	}
	
	
	
	public static boolean checkAddingFormulaforAbbrv() throws InterruptedException
	{
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
		editScreenFormulaTxtBox.sendKeys("&((qty>0)&&(qty<=10)?srtABVR1(0):((qty>10)&&(qty<=25)?srtABVR2(1):((qty>25)&&(qty<=500)?srtABVR3(2):srtABVR3(3))))");
		Thread.sleep(1000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFormulaOkBtn));
		editScreenFormulaOkBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsUpdateBtn));
		settingsUpdateBtn.click();
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsCloseBtn));
		settingsCloseBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		
		if(newBtn.isDisplayed()==true)
		{
			return true;
		}
		else
		{
			return false;
		}
		
		
	}
	
	
	
	
	
	
	@FindBy(xpath="//*[@id='id_transactionentry_new']/a/span")
	public static WebElement salesInvoice_newBtn;
	
	public static boolean checkRatewithPricebookAccountOneatSalesInvoiceVoucher() throws InterruptedException
	{
		
		checkOpenSalesInvoiceVoucher();
		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_newBtn));
		salesInvoice_newBtn.click();
		Thread.sleep(2000);
		
		
		getWaitForAlert();
		Thread.sleep(1500);
		getAlert().accept();
		Thread.sleep(2000);*/
		
		
		if(checkAddingFormulaforAccount())
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		}

		
		Thread.sleep(2000);
		
		//BR COGS ITEM Pricebook Account One
		String actRateCust1=checkRateValue("Pricebook Account One","BR COGS ITEM");
		String expRateCust1="100.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		String item_Txt1=enter_ItemTxt.getAttribute("value");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String custName1=customerAccountTxt.getAttribute("value");
		
			//WA COGS ITEM Pricebook Account One
				String actRateCust2=checkRateValue("Pricebook Account One","WA COGS ITEM");
				String expRateCust2="200.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt2=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName2=customerAccountTxt.getAttribute("value");
		
				//FIFO COGS ITEM Pricebook Account One
				String actRateCust3=checkRateValue("Pricebook Account One","FIFO COGS ITEM");
				String expRateCust3="300.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt3=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName3=customerAccountTxt.getAttribute("value");
		
				//STD RATE COGS ITEM Pricebook Account One
				String actRateCust4=checkRateValue("Pricebook Account One","BIN RAW MATERIAL ITEM");
				String expRateCust4="0.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt4=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName4=customerAccountTxt.getAttribute("value");
				
				
		
		
		System.out.println("Account Name:	"+ custName1	+	"Item Name: 	"+ item_Txt1			+	"Actual Rate :		" +	actRateCust1 		+ "Value Expected:		"	+expRateCust1);
		System.out.println("Account Name:	"+ custName2	+	"Item Name: 	"+ item_Txt2			+	"Actual Rate :		" +	actRateCust2 		+ "Value Expected:		"	+expRateCust2);
		System.out.println("Account Name:	"+ custName3	+	"Item Name: 	"+ item_Txt3			+	"Actual Rate :		" +	actRateCust3 		+ "Value Expected:		"	+expRateCust3);
		System.out.println("Account Name:	"+ custName4	+	"Item Name: 	"+ item_Txt4			+	"Actual Rate :		" +	actRateCust4 		+ "Value Expected:		"	+expRateCust4);
		
		
		if(actRateCust1.equalsIgnoreCase(expRateCust1) && actRateCust2.equalsIgnoreCase(expRateCust2) 
				&& actRateCust3.equalsIgnoreCase(expRateCust3) && actRateCust4.equalsIgnoreCase(expRateCust4))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
	
	
	
	public static boolean checkRatewithPricebookAccountTwoatSalesInvoiceVoucher() throws InterruptedException
	{
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_newBtn));
		salesInvoice_newBtn.click();
		Thread.sleep(2000);
		
		
		getWaitForAlert();
		Thread.sleep(1500);
		getAlert().accept();
		Thread.sleep(2000);
		
		//BR COGS ITEM Pricebook Account One
		String actRateCust1=checkRateValue("Pricebook Account Two","BR COGS ITEM");
		String expRateCust1="500.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		String item_Txt1=enter_ItemTxt.getAttribute("value");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String custName1=customerAccountTxt.getAttribute("value");
		
			//WA COGS ITEM Pricebook Account One
				String actRateCust2=checkRateValue("Pricebook Account Two","WA COGS ITEM");
				String expRateCust2="700.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt2=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName2=customerAccountTxt.getAttribute("value");
		
				//FIFO COGS ITEM Pricebook Account One
				String actRateCust3=checkRateValue("Pricebook Account Two","FIFO COGS ITEM");
				String expRateCust3="600.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt3=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName3=customerAccountTxt.getAttribute("value");
		
				//STD RATE COGS ITEM Pricebook Account One
				String actRateCust4=checkRateValue("Pricebook Account Two","BIN RAW MATERIAL ITEM");
				String expRateCust4="0.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt4=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName4=customerAccountTxt.getAttribute("value");
				
				
		
		
		System.out.println("Account Name:	"+ custName1	+	"Item Name: 	"+ item_Txt1			+	"Actual Rate :		" +	actRateCust1 		+ "Value Expected:		"	+expRateCust1);
		System.out.println("Account Name:	"+ custName2	+	"Item Name: 	"+ item_Txt2			+	"Actual Rate :		" +	actRateCust2 		+ "Value Expected:		"	+expRateCust2);
		System.out.println("Account Name:	"+ custName3	+	"Item Name: 	"+ item_Txt3			+	"Actual Rate :		" +	actRateCust3 		+ "Value Expected:		"	+expRateCust3);
		System.out.println("Account Name:	"+ custName4	+	"Item Name: 	"+ item_Txt4			+	"Actual Rate :		" +	actRateCust4 		+ "Value Expected:		"	+expRateCust4);
		
		
		if(actRateCust1.equalsIgnoreCase(expRateCust1) && actRateCust2.equalsIgnoreCase(expRateCust2) 
				&& actRateCust3.equalsIgnoreCase(expRateCust3) && actRateCust4.equalsIgnoreCase(expRateCust4))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
	
	
	
	public static boolean checkRatewithPricebookGroupAccountOneatSalesInvoiceVoucher() throws InterruptedException
	{
		Thread.sleep(2000);
		
		
		checkOpenSalesInvoiceVoucher();
		
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_newBtn));
		salesInvoice_newBtn.click();
		Thread.sleep(2000);
		
		
		getWaitForAlert();
		Thread.sleep(1500);
		getAlert().accept();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_MoreBtn));
		salesInvoice_MoreBtn.click();
		Thread.sleep(1500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_SettingsBtn));
		salesInvoice_SettingsBtn.click();*/
		Thread.sleep(2000);
		
		
		if(checkAddingFormulaforGroupAccount())
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		}

		
		Thread.sleep(2000);
		
		
		
		//BR COGS ITEM Pricebook Account One
		String actRateCust1=checkRateValue("Group One Account One","BR COGS ITEM");
		String expRateCust1="2000.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		String item_Txt1=enter_ItemTxt.getAttribute("value");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String custName1=customerAccountTxt.getAttribute("value");
		
		
		
		
		
		
		
			//WA COGS ITEM Pricebook Account One
				String actRateCust2=checkRateValue("Group One Account One","WA COGS ITEM");
				String expRateCust2="1000.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt2=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName2=customerAccountTxt.getAttribute("value");
				
				
				//FIFO COGS ITEM Pricebook Account One
				String actRateCust3=checkRateValue("Group One Account One","FIFO COGS ITEM");
				String expRateCust3="3000.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt3=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName3=customerAccountTxt.getAttribute("value");
				
				
				
				//STD RATE COGS ITEM Pricebook Account One
				String actRateCust4=checkRateValue("Group One Account One","STD RATE COGS ITEM");
				String expRateCust4="0.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt4=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName4=customerAccountTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_newBtn));
				salesInvoice_newBtn.click();
				Thread.sleep(2000);
				
				
				getWaitForAlert();
				Thread.sleep(1500);
				getAlert().accept();
				Thread.sleep(2000);
				
				
				//BR COGS iTEM
				
				String actRateCust11=checkRateValue("Group One Account Two","BR COGS ITEM");
				String expRateCust11="2000.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt11=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName11=customerAccountTxt.getAttribute("value");
				
				//WA COGS ITEM
				String actRateCust21=checkRateValue("Group One Account Two","WA COGS ITEM");
				String expRateCust21="1000.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt21=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName21=customerAccountTxt.getAttribute("value");
				
				//FIFO COGS ITEM
				String actRateCust31=checkRateValue("Group One Account Two","FIFO COGS ITEM");
				String expRateCust31="3000.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt31=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName31=customerAccountTxt.getAttribute("value");
				
				//STD RATE COGS ITEM		
				String actRateCust41=checkRateValue("Group One Account Two","STD RATE COGS ITEM");
				String expRateCust41="0.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt41=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName41=customerAccountTxt.getAttribute("value");
				
				
				
				
		
				
		
		
		System.out.println("Account Name:	"+ custName1	+	"Item Name: 	"+ item_Txt1			+	"Actual Rate :		" +	actRateCust1 		+ "Value Expected:		"	+expRateCust1);
		System.out.println("Account Name:	"+ custName1	+	"Item Name: 	"+ item_Txt11			+	"Actual Rate :		" +	actRateCust11		+ "Value Expected:		"	+expRateCust11);
		
		System.out.println("Account Name:	"+ custName2	+	"Item Name: 	"+ item_Txt2			+	"Actual Rate :		" +	actRateCust2 		+ "Value Expected:		"	+expRateCust2);
		System.out.println("Account Name:	"+ custName2	+	"Item Name: 	"+ item_Txt21			+	"Actual Rate :		" +	actRateCust21 		+ "Value Expected:		"	+expRateCust21);
		
		System.out.println("Account Name:	"+ custName3	+	"Item Name: 	"+ item_Txt3			+	"Actual Rate :		" +	actRateCust3 		+ "Value Expected:		"	+expRateCust3);
		System.out.println("Account Name:	"+ custName3	+	"Item Name: 	"+ item_Txt31			+	"Actual Rate :		" +	actRateCust31 		+ "Value Expected:		"	+expRateCust31);
		
		System.out.println("Account Name:	"+ custName4	+	"Item Name: 	"+ item_Txt4			+	"Actual Rate :		" +	actRateCust4 		+ "Value Expected:		"	+expRateCust4);
		System.out.println("Account Name:	"+ custName4	+	"Item Name: 	"+ item_Txt41			+	"Actual Rate :		" +	actRateCust41 		+ "Value Expected:		"	+expRateCust41);
		
		
		
		
		
		
		if(actRateCust1.equalsIgnoreCase(expRateCust1) && actRateCust2.equalsIgnoreCase(expRateCust2) 
				&& actRateCust3.equalsIgnoreCase(expRateCust3) && actRateCust4.equalsIgnoreCase(expRateCust4))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
	
	
	public static boolean checkRatewithPricebookGroupAccountTwoatSalesInvoiceVoucher() throws InterruptedException
	{
		
		checkOpenSalesInvoiceVoucher();
		
		Thread.sleep(2000);
		
		
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_newBtn));
		salesInvoice_newBtn.click();
		Thread.sleep(2000);
		
		
		getWaitForAlert();
		Thread.sleep(1500);
		getAlert().accept();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_MoreBtn));
		salesInvoice_MoreBtn.click();
		Thread.sleep(1500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_SettingsBtn));
		salesInvoice_SettingsBtn.click();
		Thread.sleep(2000);
		*/
		
		if(checkAddingFormulaforGroupAccount())
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		}

		
		Thread.sleep(2000);
		
		
		
		//BR COGS ITEM Pricebook Account One
		String actRateCust1=checkRateValue("Group Two Account One","BR COGS ITEM");
		String expRateCust1="7000.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		String item_Txt1=enter_ItemTxt.getAttribute("value");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String custName1=customerAccountTxt.getAttribute("value");
		
		
		
		
		
			//WA COGS ITEM Pricebook Account One
				String actRateCust2=checkRateValue("Group Two Account One","WA COGS ITEM");
				String expRateCust2="5000.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt2=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName2=customerAccountTxt.getAttribute("value");
				
				
				
				
		
				//FIFO COGS ITEM Pricebook Account One
				String actRateCust3=checkRateValue("Group Two Account One","FIFO COGS ITEM");
				String expRateCust3="6000.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt3=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName3=customerAccountTxt.getAttribute("value");
				
				
				
		
				//STD RATE COGS ITEM Pricebook Account One
				String actRateCust4=checkRateValue("Group Two Account One","STD RATE COGS ITEM");
				String expRateCust4="0.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt4=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName4=customerAccountTxt.getAttribute("value");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_newBtn));
				salesInvoice_newBtn.click();
				Thread.sleep(2000);
				
				
				getWaitForAlert();
				Thread.sleep(1500);
				getAlert().accept();
				Thread.sleep(2000);
				
				
				
				//BR COGS ITEM
				
				String actRateCust11=checkRateValue("Group Two Account Two","BR COGS ITEM");
				String expRateCust11="7000.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt11=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName11=customerAccountTxt.getAttribute("value");
				
				
				//WA COGS ITEM
				String actRateCust21=checkRateValue("Group Two Account Two","WA COGS ITEM");
				String expRateCust21="5000.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt21=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName21=customerAccountTxt.getAttribute("value");
				
				//FIFO COGS ITEM
				String actRateCust31=checkRateValue("Group Two Account Two","FIFO COGS ITEM");
				String expRateCust31="6000.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt31=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName31=customerAccountTxt.getAttribute("value");
				
				
				//STD RATE ITEM COGS
				
				String actRateCust41=checkRateValue("Group Two Account Two","STD RATE COGS ITEM");
				String expRateCust41="0.00";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				String item_Txt41=enter_ItemTxt.getAttribute("value");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String custName41=customerAccountTxt.getAttribute("value");
				
				
				
		
		
				System.out.println("Account Name:	"+ custName1	+	"Item Name: 	"+ item_Txt1			+	"Actual Rate :		" +	actRateCust1 		+ "Value Expected:		"	+expRateCust1);
				System.out.println("Account Name:	"+ custName1	+	"Item Name: 	"+ item_Txt11			+	"Actual Rate :		" +	actRateCust11		+ "Value Expected:		"	+expRateCust11);
				
				System.out.println("Account Name:	"+ custName2	+	"Item Name: 	"+ item_Txt2			+	"Actual Rate :		" +	actRateCust2 		+ "Value Expected:		"	+expRateCust2);
				System.out.println("Account Name:	"+ custName2	+	"Item Name: 	"+ item_Txt21			+	"Actual Rate :		" +	actRateCust21 		+ "Value Expected:		"	+expRateCust21);
				
				System.out.println("Account Name:	"+ custName3	+	"Item Name: 	"+ item_Txt3			+	"Actual Rate :		" +	actRateCust3 		+ "Value Expected:		"	+expRateCust3);
				System.out.println("Account Name:	"+ custName3	+	"Item Name: 	"+ item_Txt31			+	"Actual Rate :		" +	actRateCust31 		+ "Value Expected:		"	+expRateCust31);
				
				System.out.println("Account Name:	"+ custName4	+	"Item Name: 	"+ item_Txt4			+	"Actual Rate :		" +	actRateCust4 		+ "Value Expected:		"	+expRateCust4);
				System.out.println("Account Name:	"+ custName4	+	"Item Name: 	"+ item_Txt41			+	"Actual Rate :		" +	actRateCust41 		+ "Value Expected:		"	+expRateCust41);
				
		
		if(actRateCust1.equalsIgnoreCase(expRateCust1) && actRateCust2.equalsIgnoreCase(expRateCust2) 
				&& actRateCust3.equalsIgnoreCase(expRateCust3) && actRateCust4.equalsIgnoreCase(expRateCust4))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
	
	
public static boolean checkRatewithPricebookGroupAccountOneandGroupItem() throws InterruptedException
{
	
	checkOpenSalesInvoiceVoucher();
	Thread.sleep(3000);
	
	
/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_newBtn));
	salesInvoice_newBtn.click();
	Thread.sleep(2000);
	
	if(getIsAlertPresent())
	{
	getWaitForAlert();
	Thread.sleep(1500);
	getAlert().accept();
	Thread.sleep(2000);
	
	}
	*/
	if(checkAddingFormulaforGroupAccount())
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	}

	
	Thread.sleep(2000);
	
	
	
	
	//Group One Item One
	
	String actRateCust1=checkRateValue("Group One Account One","Group One Item One");
	String expRateCust1="15.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt1=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName1=customerAccountTxt.getAttribute("value");
	
	
	//WA COGS ITEM
	String actRateCust2=checkRateValue("Group One Account One","Group One Item Two");
	String expRateCust2="15.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt2=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName2=customerAccountTxt.getAttribute("value");
	
	//FIFO COGS ITEM
	String actRateCust3=checkRateValue("Group One Account One","Group Two Item One");
	String expRateCust3="60.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt3=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName3=customerAccountTxt.getAttribute("value");
	
	
	//STD RATE ITEM COGS
	
	String actRateCust4=checkRateValue("Group One Account One","Group Two Item Two");
	String expRateCust4="60.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt4=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName4=customerAccountTxt.getAttribute("value");
	
	
	//Group One Account Two
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_newBtn));
	salesInvoice_newBtn.click();
	Thread.sleep(2000);
	
	if(getIsAlertPresent())
	{
	getWaitForAlert();
	Thread.sleep(1500);
	getAlert().accept();
	Thread.sleep(2000);
	}
	
	

	
	String actRateCust11=checkRateValue("Group One Account Two","Group One Item One");
	String expRateCust11="15.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt11=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName11=customerAccountTxt.getAttribute("value");
	
	
	
	String actRateCust21=checkRateValue("Group One Account Two","Group One Item Two");
	String expRateCust21="15.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt21=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName21=customerAccountTxt.getAttribute("value");
	
	
	String actRateCust31=checkRateValue("Group One Account Two","Group Two Item One");
	String expRateCust31="60.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt31=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName31=customerAccountTxt.getAttribute("value");
	
	
	
	
	String actRateCust41=checkRateValue("Group One Account Two","Group Two Item Two");
	String expRateCust41="60.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt41=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName41=customerAccountTxt.getAttribute("value");
	
	
	
	
	
	System.out.println("Account Name:	"+ custName1	+	"Item Name: 	"+ item_Txt1			+	"Actual Rate :		" +	actRateCust1 		+ "Value Expected:		"	+expRateCust1);
	System.out.println("Account Name:	"+ custName11	+	"Item Name: 	"+ item_Txt11			+	"Actual Rate :		" +	actRateCust11 		+ "Value Expected:		"	+expRateCust11);
	
	System.out.println("Account Name:	"+ custName2	+	"Item Name: 	"+ item_Txt2			+	"Actual Rate :		" +	actRateCust2 		+ "Value Expected:		"	+expRateCust2);
	System.out.println("Account Name:	"+ custName21	+	"Item Name: 	"+ item_Txt21			+	"Actual Rate :		" +	actRateCust21 		+ "Value Expected:		"	+expRateCust21);
	
	System.out.println("Account Name:	"+ custName3	+	"Item Name: 	"+ item_Txt3			+	"Actual Rate :		" +	actRateCust3 		+ "Value Expected:		"	+expRateCust3);
	System.out.println("Account Name:	"+ custName31	+	"Item Name: 	"+ item_Txt31			+	"Actual Rate :		" +	actRateCust31 		+ "Value Expected:		"	+expRateCust31);
	
	System.out.println("Account Name:	"+ custName4	+	"Item Name: 	"+ item_Txt4			+	"Actual Rate :		" +	actRateCust4 		+ "Value Expected:		"	+expRateCust4);
	System.out.println("Account Name:	"+ custName41	+	"Item Name: 	"+ item_Txt41			+	"Actual Rate :		" +	actRateCust41 		+ "Value Expected:		"	+expRateCust41);
	
	if(actRateCust1.equalsIgnoreCase(expRateCust1) && actRateCust2.equalsIgnoreCase(expRateCust2) 
			&& actRateCust3.equalsIgnoreCase(expRateCust3) && actRateCust4.equalsIgnoreCase(expRateCust4)
			&& actRateCust11.equalsIgnoreCase(expRateCust11) && actRateCust21.equalsIgnoreCase(expRateCust21) 
			&& actRateCust31.equalsIgnoreCase(expRateCust31) && actRateCust41.equalsIgnoreCase(expRateCust41))
	{
		return true;
	}
	else
	{
		return false;
	}
	
	
	
	
	
}
	
	
public static boolean checkRatewithPricebookGroupAccountTwoandGroupItem() throws InterruptedException
{
/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_newBtn));
	salesInvoice_newBtn.click();
	Thread.sleep(2000);
	
	if(getIsAlertPresent())
	{
	getWaitForAlert();
	Thread.sleep(1500);
	getAlert().accept();
	Thread.sleep(2000);
	}*/
	
	checkOpenSalesInvoiceVoucher();
	Thread.sleep(2000);
	
	
	if(checkAddingFormulaforGroupAccount())
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	}

	
	Thread.sleep(2000);
	
	String actRateCust1=checkRateValue("Group Two Account One","Group One Item One");
	String expRateCust1="10.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt1=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName1=customerAccountTxt.getAttribute("value");
	
	
	
	String actRateCust2=checkRateValue("Group Two Account One","Group One Item Two");
	String expRateCust2="10.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt2=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName2=customerAccountTxt.getAttribute("value");
	
	
	String actRateCust3=checkRateValue("Group Two Account One","Group Two Item One");
	String expRateCust3="25.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt3=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName3=customerAccountTxt.getAttribute("value");
	
	
	
	
	String actRateCust4=checkRateValue("Group Two Account One","Group Two Item Two");
	String expRateCust4="25.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt4=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName4=customerAccountTxt.getAttribute("value");
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_newBtn));
	salesInvoice_newBtn.click();
	Thread.sleep(2000);
	
	if(getIsAlertPresent())
	{
	getWaitForAlert();
	Thread.sleep(1500);
	getAlert().accept();
	Thread.sleep(2000);
	}
	
	
	
	
	String actRateCust11=checkRateValue("Group Two Account Two","Group One Item One");
	String expRateCust11="10.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt11=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName11=customerAccountTxt.getAttribute("value");
	
	
	
	String actRateCust21=checkRateValue("Group Two Account Two","Group One Item Two");
	String expRateCust21="10.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt21=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName21=customerAccountTxt.getAttribute("value");
	
	
	String actRateCust31=checkRateValue("Group Two Account Two","Group Two Item One");
	String expRateCust31="25.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt31=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName31=customerAccountTxt.getAttribute("value");
	
	
	
	
	String actRateCust41=checkRateValue("Group Two Account Two","Group Two Item Two");
	String expRateCust41="25.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt41=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String custName41=customerAccountTxt.getAttribute("value");
	
	
	
	
	
	System.out.println("Account Name:	"+ custName1	+	"Item Name: 	"+ item_Txt1			+	"Actual Rate :		" +	actRateCust1 		+ "Value Expected:		"	+expRateCust1);
	System.out.println("Account Name:	"+ custName11	+	"Item Name: 	"+ item_Txt11			+	"Actual Rate :		" +	actRateCust11 		+ "Value Expected:		"	+expRateCust11);
	
	System.out.println("Account Name:	"+ custName2	+	"Item Name: 	"+ item_Txt2			+	"Actual Rate :		" +	actRateCust2 		+ "Value Expected:		"	+expRateCust2);
	System.out.println("Account Name:	"+ custName21	+	"Item Name: 	"+ item_Txt21			+	"Actual Rate :		" +	actRateCust21 		+ "Value Expected:		"	+expRateCust21);
	
	
	System.out.println("Account Name:	"+ custName3	+	"Item Name: 	"+ item_Txt3			+	"Actual Rate :		" +	actRateCust3 		+ "Value Expected:		"	+expRateCust3);
	System.out.println("Account Name:	"+ custName31	+	"Item Name: 	"+ item_Txt31			+	"Actual Rate :		" +	actRateCust31 		+ "Value Expected:		"	+expRateCust31);
	
	
	System.out.println("Account Name:	"+ custName4	+	"Item Name: 	"+ item_Txt4			+	"Actual Rate :		" +	actRateCust4 		+ "Value Expected:		"	+expRateCust4);
	System.out.println("Account Name:	"+ custName41	+	"Item Name: 	"+ item_Txt41			+	"Actual Rate :		" +	actRateCust41		+ "Value Expected:		"	+expRateCust41);
	
	if(actRateCust1.equalsIgnoreCase(expRateCust1) && actRateCust2.equalsIgnoreCase(expRateCust2) 
			&& actRateCust3.equalsIgnoreCase(expRateCust3) && actRateCust4.equalsIgnoreCase(expRateCust4)
			&& actRateCust11.equalsIgnoreCase(expRateCust11) && actRateCust21.equalsIgnoreCase(expRateCust21) 
			&& actRateCust31.equalsIgnoreCase(expRateCust31) && actRateCust41.equalsIgnoreCase(expRateCust41))
	{
		return true;
	}
	else
	{
		return false;
	}

	
}
		
	
public static boolean checkRateLoadingforPricebookDepartmentGroupinSalesInvoice() throws InterruptedException
{

	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_newBtn));
	salesInvoice_newBtn.click();
	Thread.sleep(2000);
	
	if(getIsAlertPresent())
	{
	getWaitForAlert();
	Thread.sleep(1500);
	getAlert().accept();
	Thread.sleep(2000);
	}*/
	
	checkOpenSalesInvoiceVoucher();
	Thread.sleep(2000);
	
	if(checkAddingFormulaforGroupAccount())
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	}

	
	Thread.sleep(2000);
	
	
	
	
	String actRateCust1=checkRateValueforDepartment("Group One Dept One","Pricebook Item1");
	String expRateCust1="1000.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt1=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String custName1=departmentTxt.getAttribute("value");
	
	
	
	String actRateCust2=checkRateValueforDepartment("Group One Dept One","Pricebook Item2");
	String expRateCust2="0.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt2=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String custName2=departmentTxt.getAttribute("value");
	
	
	String actRateCust3=checkRateValueforDepartment("Group One Dept Two","Pricebook Item1");
	String expRateCust3="1000.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt3=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String custName3=departmentTxt.getAttribute("value");
	
	
	
	
	String actRateCust4=checkRateValueforDepartment("Group One Dept Two","Pricebook Item2");
	String expRateCust4="0.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt4=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String custName4=departmentTxt.getAttribute("value");
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_newBtn));
	salesInvoice_newBtn.click();
	Thread.sleep(2000);
	
	if(getIsAlertPresent())
	{
	getWaitForAlert();
	Thread.sleep(1500);
	getAlert().accept();
	Thread.sleep(2000);
	}*/
	
	
	
	
	String actRateCust11=checkRateValueforDepartment("Group Two Dept One","Pricebook Item1");
	String expRateCust11="0.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt11=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String custName11=departmentTxt.getAttribute("value");
	
	
	
	String actRateCust21=checkRateValueforDepartment("Group Two Dept One","Pricebook Item2");
	String expRateCust21="5000.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt21=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String custName21=departmentTxt.getAttribute("value");
	
	
	String actRateCust31=checkRateValueforDepartment("Group Two Dept Two","Pricebook Item1");
	String expRateCust31="0.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt31=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String custName31=departmentTxt.getAttribute("value");
	
	
	
	
	String actRateCust41=checkRateValueforDepartment("Group Two Dept Two","Pricebook Item2");
	String expRateCust41="5000.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt41=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String custName41=departmentTxt.getAttribute("value");
	
	
	
	
	
	
	
	
	System.out.println("Account Name:	"+ custName1	+	"Item Name: 	"+ item_Txt1			+	"Actual Rate :		" +	actRateCust1 		+ "Value Expected:		"	+expRateCust1);
	System.out.println("Account Name:	"+ custName11	+	"Item Name: 	"+ item_Txt11			+	"Actual Rate :		" +	actRateCust11 		+ "Value Expected:		"	+expRateCust11);
	
	System.out.println("Account Name:	"+ custName2	+	"Item Name: 	"+ item_Txt2			+	"Actual Rate :		" +	actRateCust2 		+ "Value Expected:		"	+expRateCust2);
	System.out.println("Account Name:	"+ custName21	+	"Item Name: 	"+ item_Txt21			+	"Actual Rate :		" +	actRateCust21 		+ "Value Expected:		"	+expRateCust21);
	
	
	System.out.println("Account Name:	"+ custName3	+	"Item Name: 	"+ item_Txt3			+	"Actual Rate :		" +	actRateCust3 		+ "Value Expected:		"	+expRateCust3);
	System.out.println("Account Name:	"+ custName31	+	"Item Name: 	"+ item_Txt31			+	"Actual Rate :		" +	actRateCust31 		+ "Value Expected:		"	+expRateCust31);
	
	
	System.out.println("Account Name:	"+ custName4	+	"Item Name: 	"+ item_Txt4			+	"Actual Rate :		" +	actRateCust4 		+ "Value Expected:		"	+expRateCust4);
	System.out.println("Account Name:	"+ custName41	+	"Item Name: 	"+ item_Txt41			+	"Actual Rate :		" +	actRateCust41		+ "Value Expected:		"	+expRateCust41);
	
	if(actRateCust1.equalsIgnoreCase(expRateCust1) && actRateCust2.equalsIgnoreCase(expRateCust2) 
			&& actRateCust3.equalsIgnoreCase(expRateCust3) && actRateCust4.equalsIgnoreCase(expRateCust4)
			&& actRateCust11.equalsIgnoreCase(expRateCust11) && actRateCust21.equalsIgnoreCase(expRateCust21) 
			&& actRateCust31.equalsIgnoreCase(expRateCust31) && actRateCust41.equalsIgnoreCase(expRateCust41))
	{
		return true;
	}
	else
	{
		return false;
	}

	

}


public static boolean checkRateLoadingforItemhavingMaxandMinQuantityinSalesInvoiceVoucher() throws InterruptedException
{
	
	checkOpenSalesInvoiceVoucher();
	Thread.sleep(2000);
	
	if(checkAddingFormulaforGroupAccount())
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	}
	
	
	
	//Lass than Min Quantity
	
	String actRateCust1=checkRateValue("RMA FITEM",0);
	String expRateCust1="0.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt1=enter_ItemTxt.getAttribute("value");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt1=enter_Quantity.getAttribute("value");
	int q=new Double(qty_Txt1).intValue();
	
	
	//Min Quantity
	
	
		
		String actRateCust2=checkRateValue("RMA FITEM",(q+1));
		String expRateCust2="100.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		String item_Txt2=enter_ItemTxt.getAttribute("value");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		String qty_Txt2=enter_Quantity.getAttribute("value");
		
		int q1=new Double(qty_Txt2).intValue();
		System.err.println(q);
		
	//Max Quantity
		
		String actRateCust3=checkRateValue("RMA FITEM",200);
		String expRateCust3="100.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		String item_Txt3=enter_ItemTxt.getAttribute("value");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		String qty_Txt3=enter_Quantity.getAttribute("value");
		
	
		//greater than  Max Quantity
		
		String actRateCust4=checkRateValue("RMA FITEM",202);
		String expRateCust4="0.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		String item_Txt4=enter_ItemTxt.getAttribute("value");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		String qty_Txt4=enter_Quantity.getAttribute("value");
		
	
		System.out.println("Item Name:	"	+ item_Txt1	+	"Quantity :	"	+ 	qty_Txt1			+	"Actual Rate :		" +	actRateCust1 		+ "Value Expected:		"	+expRateCust1);
		System.out.println("Item Name:	"	+ item_Txt2	+	"Quantity :	"	+ 	qty_Txt2			+	"Actual Rate :		" +	actRateCust2 		+ "Value Expected:		"	+expRateCust2);
		System.out.println("Item Name:	"	+ item_Txt3	+	"Quantity :	"	+ 	qty_Txt3			+	"Actual Rate :		" +	actRateCust3 		+ "Value Expected:		"	+expRateCust3);
		System.out.println("Item Name:	"	+ item_Txt4	+	"Quantity :	"	+ 	qty_Txt4			+	"Actual Rate :		" +	actRateCust4 		+ "Value Expected:		"	+expRateCust4);
	
	
		
		if(actRateCust1.equalsIgnoreCase(expRateCust1) && actRateCust2.equalsIgnoreCase(expRateCust2) 
				&& actRateCust3.equalsIgnoreCase(expRateCust3) && actRateCust4.equalsIgnoreCase(expRateCust4))
		{
			return true;
		}
		else
		{
			return false;
		}
		
}


/*
	public static boolean checkquantiy() throws InterruptedException
	{
		
		boolean result=checkRateLoadingforItemhavingMaxandMinQuantityinSalesInvoiceVoucher("Pricebook Item1");
		return result;
	}
*/

		@FindBy(xpath="//*[@id='StartingDateH']")
		private static WebElement buyerPriceBook_StartDate;
		
		
		@FindBy(xpath="//*[@id='EndingDateH']")
		private static WebElement buyerPriceBook_EndDate;
		
		@FindBy(xpath="//*[@id='EndingDateH_checkbox']")
		private static WebElement buyerPriceBook_EndDateChkBox;
		
		@FindBy(xpath="//*[@id='PriceBookWebGrid_val1']")
		private static WebElement val1Value;
		

			public boolean checkCreatingSellerPriceBookforDateRange() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
			
				getDriver().navigate().refresh();
				Thread.sleep(2000);
			
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
			buyerPriceBook_PriceBook.sendKeys("Seller for DateRange");
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
			itemValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			itemValue.sendKeys("Pricebook Item1");
			Thread.sleep(2000);
			
			/*int count = itemComboBoxList.size();
			
			for(int i=0;i<count;i++)
			{
			String data = itemComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Pricebook Item1"))
			{
			itemComboBoxList.get(i).click();
			break;
			}
			}*/
			
			itemValue.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
			rateValue.sendKeys("10.10");
			rateValue.sendKeys(Keys.TAB);
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
			//val1Value.click();
			val1Value.sendKeys("100.25");
			val1Value.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
			itemSecondCellSelect.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
			itemValue.click();
			itemValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			itemValue.sendKeys("Pricebook Item2");
			Thread.sleep(2000);
			
			/*int count1 = itemComboBoxList.size();
			
			for(int i=0;i<count1;i++)
			{
			String data = itemComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Pricebook Item2"))
			{
			itemComboBoxList.get(i).click();
			break;
			}
			}
			*/
			itemValue.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
			rateValue.sendKeys("15.10");
			rateValue.sendKeys(Keys.TAB);
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
			//val1Value.click();
			val1Value.sendKeys("120.25");
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
			buyerPriceBook_PriceBook.sendKeys("Seller2 for DateRange");
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
			itemValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			itemValue.sendKeys("Pricebook Item2");
			Thread.sleep(2000);
			/*int count11 = itemComboBoxList.size();
			
			for(int i=0;i<count11;i++)
			{
			String data = itemComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Pricebook Item2"))
			{
			itemComboBoxList.get(i).click();
			break;
			}
			}*/
			
			itemValue.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
			rateValue.sendKeys("15.25");
			rateValue.sendKeys(Keys.TAB);
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
			//val1Value.click();
			val1Value.sendKeys("150.75");
			val1Value.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
			buyerPriceBookSaveBtn.click();
			
			System.out.println("*********************************checkSavingBuyerPriceBook********************************");
			
			String expMessage1="Price Book created successfully";
			
			String actMessage1=checkValidationMessage(expMessage1);
			
			
			
			
			if (actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1) ) 
			{
			System.out.println("Pass: Saved Successfully");
			
			return true;
			} 
			else 
			{
			System.out.println("Fail:  Saved Successfully");
			
			return false;
			}
			
			
			
			}


			
	public static boolean checkRateLoadingforItemhavingDateRangeinSalesInvoiceVoucher() throws InterruptedException
	{
		checkOpenSalesInvoiceVoucher();
		Thread.sleep(2000);
		
		if(checkAddingFormulaforAccount())
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		}

		
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
		Thread.sleep(2000);
		
		String actRate=checkRateValue("Pricebook Item1");
		String expRate="10.10";
		
	//checking future date	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dateTxt));
		dateTxt.click();
		dateTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
		
		cal11.add(Calendar.DATE, 10);
		d1=cal11.getTime();
		dateTxt.sendKeys(format.format(d1));
		Thread.sleep(1500);
		dateTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		if(getIsAlertPresent())
		{
			getWaitForAlert();
			Thread.sleep(1500);
			getAlert().accept();
			Thread.sleep(2000);
		}
		
		String actRate1=checkRateValue("Pricebook Item1");
		String expRate1="0.00";
		
		
		
		
		Calendar cal1=Calendar.getInstance();
		cal1.add(Calendar.YEAR,0);
		cal1.set(Calendar.DAY_OF_YEAR,-1);
		Date start=cal1.getTime();
		
		System.err.println(format.format(start));
		
	//checking past date	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dateTxt));
		dateTxt.click();
		dateTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
		dateTxt.sendKeys(format.format(start));
		Thread.sleep(1500);
		dateTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		if(getIsAlertPresent())
		{
			getWaitForAlert();
			Thread.sleep(1500);
			getAlert().accept();
			Thread.sleep(2000);
		}
		
		String actRate2=checkRateValue("Pricebook Item1");
		String expRate2="0.00";
		
		System.out.println("***********Checking Date Range for Priceboo Item1******************");
		System.out.println("Actual Rate	"	+	actRate		+	"Vale Expected	"	+	expRate);
		System.out.println("Actual Rate	"	+	actRate1	+	"Vale Expected	"	+	expRate1);
		System.out.println("Actual Rate	"	+	actRate2	+	"Vale Expected	"	+	expRate2);	
		
		
		//checking future date	
		
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dateTxt));
				dateTxt.click();
				dateTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
				
				
				cal11.add(Calendar.DATE, 10);
				d1=cal11.getTime();
				dateTxt.sendKeys(format.format(d1));
				Thread.sleep(1500);
				dateTxt.sendKeys(Keys.TAB);
				Thread.sleep(2000);
				
				if(getIsAlertPresent())
				{
					getWaitForAlert();
					Thread.sleep(1500);
					getAlert().accept();
					Thread.sleep(2000);
				}
				
				String actRate3=checkRateValue("Pricebook Item2");
				String expRate3="15.25";
				
				
				
				
				Calendar cal=Calendar.getInstance();
				cal1.add(Calendar.YEAR,0);
				cal1.set(Calendar.DAY_OF_YEAR,-1);
				Date start1=cal1.getTime();
				
				System.err.println(format.format(start));
				
			//checking past date	
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dateTxt));
				dateTxt.click();
				dateTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
				
				dateTxt.sendKeys(format.format(start1));
				Thread.sleep(1500);
				dateTxt.sendKeys(Keys.TAB);
				Thread.sleep(2000);
				
				if(getIsAlertPresent())
				{
					getWaitForAlert();
					Thread.sleep(1500);
					getAlert().accept();
					Thread.sleep(2000);
				}
				
				String actRate4=checkRateValue("Pricebook Item2");
				String expRate4="0.00";
				
				
				System.out.println("***********Checking Date Range for Priceboo Item2******************");
				System.out.println("Actual Rate	"	+	actRate3		+	"Vale Expected	"	+	expRate3);
				System.out.println("Actual Rate	"	+	actRate4		+	"Vale Expected	"	+	expRate4);
					
		
		
		if(actRate.equalsIgnoreCase(expRate) && actRate1.equalsIgnoreCase(expRate1) && actRate2.equalsIgnoreCase(expRate2)&&
				actRate3.equalsIgnoreCase(expRate3) && actRate4.equalsIgnoreCase(expRate4) )
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
			


///////Filter///////////////////
	
	@FindBy(xpath="//*[@id='FilterProduct']")
	public static WebElement filter_Item;
	
	@FindBy(xpath="//*[@id='FilterCurrency']")
	public static WebElement filter_Currency;
	
	
	@FindBy(xpath="//*[@id='FilterCustomer']")
	public static WebElement filter_Customer;
	
	@FindBy(xpath="//*[@id='FilterDepartment']")
	public static WebElement filter_Department;
	
	
	
	public static boolean checkItemFilterinSellerPricebook() throws InterruptedException 
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
	
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
	buyerPriceBook_PriceBook.sendKeys("focus");
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);*/
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();

	Thread.sleep(3000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Item));
	filter_Item.click();
	filter_Item.sendKeys("ND 150");
	Thread.sleep(1500);
	filter_Item.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterBtn));
	priceBookFilterBtn.click();

	Thread.sleep(3000);
	
	Calendar cal=Calendar.getInstance();
	SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
	String presentDate = currentDate.format(cal.getTime());
	

	int count = buyerPriceBook_TableRow1List.size();

	ArrayList<String> pricebookRow1List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow1List.get(i).getText();

		pricebookRow1List.add(data);
	}

	String actPriceBookRow1 = pricebookRow1List.toString();
	String expPriceBookRow1 ="[1, ND 150, , , 01/04/2022, 31/03/2023, AIA ENGINEERING LTD.(U12), Indian Rupees, 0, 0, , 157.700000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow1);
	System.out.println(expPriceBookRow1);


	ArrayList<String> pricebookRow2List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow2List.get(i).getText();

		pricebookRow2List.add(data);
	}

	String actPriceBookRow2 = pricebookRow2List.toString();
	String expPriceBookRow2 ="[2, ND 150, , , 01/04/2022, 31/03/2023, AIA ENGINEERING LTD.(U12), , 0, 0, , 157.700000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow2);
	System.out.println(expPriceBookRow2);
	
	ArrayList<String> pricebookRow3List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow3List.get(i).getText();

		pricebookRow3List.add(data);
	}

	String actPriceBookRow3 = pricebookRow3List.toString();
	String expPriceBookRow3 ="[3, ND 150, , , 01/04/2022, 31/03/2023, SUNRISE INDUSTRIES, Indian Rupees, 0, 0, , 173.050000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow3);
	System.out.println(expPriceBookRow3);
	
	///Validation
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Item));
	filter_Item.click();
	boolean result=false;
	
	for(WebElement p:filter_ItemList)
	{
		
		if(!(p.getText().isEmpty()))
		if((filter_Item.getAttribute("value")).equalsIgnoreCase(p.getText()))
		{
			result=true;
		}
		
		
	}
	
	
	if(actPriceBookRow1.equalsIgnoreCase(expPriceBookRow1) && actPriceBookRow2.equalsIgnoreCase(expPriceBookRow2)
			&& actPriceBookRow3.equalsIgnoreCase(expPriceBookRow3) && result)
	{
		return true;
	}
	else
	{
		return false;
	}
	
	}
	
	
	@FindBy(xpath="//tbody[@id='PriceBookWebGrid_body']/tr/td[2]")
	public static List<WebElement> filter_ItemList;
	
	
	@FindBy(xpath="//tbody[@id='PriceBookWebGrid_body']/tr/td[7]")
	public static List<WebElement> filter_CustomerList;
	
	
	@FindBy(xpath="//tbody[@id='PriceBookWebGrid_body']/tr/td[8]")
	public static List<WebElement> filter_CurrencyList;
	
	
	
public static boolean checkCustomerFilterinSellerPricebook() throws InterruptedException
{
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookClearBtn));
	buyerPriceBookClearBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys("focus");
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();

	Thread.sleep(3000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Customer));
	filter_Customer.click();
	filter_Customer.sendKeys("SURYA ALLOYS");
	Thread.sleep(1500);
	filter_Customer.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterBtn));
	priceBookFilterBtn.click();

	Thread.sleep(3000);
	
	Calendar cal=Calendar.getInstance();
	SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
	String presentDate = currentDate.format(cal.getTime());
	

	int count = buyerPriceBook_TableRow1List.size();

	ArrayList<String> pricebookRow1List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow1List.get(i).getText();

		pricebookRow1List.add(data);
	}

	String actPriceBookRow1 = pricebookRow1List.toString();
	String expPriceBookRow1 ="[1, V134, , , 01/04/2022, 31/03/2023, SURYA ALLOYS, , 0, 0, , 106.450000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow1);
	System.out.println(expPriceBookRow1);


	ArrayList<String> pricebookRow2List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow2List.get(i).getText();

		pricebookRow2List.add(data);
	}

	String actPriceBookRow2 = pricebookRow2List.toString();
	String expPriceBookRow2 ="[2, V680, , , 01/04/2022, 31/03/2023, SURYA ALLOYS, , 0, 0, , 421.600000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow2);
	System.out.println(expPriceBookRow2);
	
	ArrayList<String> pricebookRow3List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow3List.get(i).getText();

		pricebookRow3List.add(data);
	}

	String actPriceBookRow3 = pricebookRow3List.toString();
	String expPriceBookRow3 ="[3, ND 180, , , 01/02/2023, 31/03/2023, SURYA ALLOYS, Indian Rupees, 0, 0, , 123.650000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow3);
	System.out.println(expPriceBookRow3);
	
	
	ArrayList<String> pricebookRow4List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow4List.get(i).getText();

		pricebookRow4List.add(data);
	}

	String actPriceBookRow4 = pricebookRow4List.toString();
	String expPriceBookRow4 ="[4, V97, , , 01/02/2023, 31/03/2023, SURYA ALLOYS, Indian Rupees, 0, 0, , 200.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow4);
	System.out.println(expPriceBookRow4);
	
	///Validation
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Customer));
	filter_Customer.click();
	boolean result=false;
	
	for(WebElement p:filter_CustomerList)
	{
		
		if(!(p.getText().isEmpty()))
		if((filter_Customer.getAttribute("value")).equalsIgnoreCase(p.getText()))
		{
			result=true;
		}
		
		
	}

	if(actPriceBookRow1.equalsIgnoreCase(expPriceBookRow1) && actPriceBookRow2.equalsIgnoreCase(expPriceBookRow2)
			&& actPriceBookRow3.equalsIgnoreCase(expPriceBookRow3) && result)
	{
		return true;
	}
	else
	{
		return false;
	}
	
}
	
	

public static boolean checkCurrencyFilterinSellerPricebook() throws InterruptedException
{

	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookClearBtn));
	buyerPriceBookClearBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys("focus");
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();

	Thread.sleep(3000);
	
	

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Currency));
	filter_Currency.click();
	filter_Currency.sendKeys("Indian Rupees");
	Thread.sleep(2000);
	filter_Currency.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterBtn));
	priceBookFilterBtn.click();

	Thread.sleep(3000);
	
	Calendar cal=Calendar.getInstance();
	SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
	String presentDate = currentDate.format(cal.getTime());
	

	int count = buyerPriceBook_TableRow1List.size();

	ArrayList<String> pricebookRow1List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow1List.get(i).getText();

		pricebookRow1List.add(data);
	}

	String actPriceBookRow1 = pricebookRow1List.toString();
	String expPriceBookRow1 ="[1, ND 150, , , 01/04/2022, 31/03/2023, AIA ENGINEERING LTD.(U12), Indian Rupees, 0, 0, , 157.700000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow1);
	System.out.println(expPriceBookRow1);


	ArrayList<String> pricebookRow2List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow2List.get(i).getText();

		pricebookRow2List.add(data);
	}

	String actPriceBookRow2 = pricebookRow2List.toString();
	String expPriceBookRow2 ="[2, ND 180, , , 01/02/2023, 31/03/2023, SURYA ALLOYS, Indian Rupees, 0, 0, , 123.650000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow2);
	System.out.println(expPriceBookRow2);
	
	ArrayList<String> pricebookRow3List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow3List.get(i).getText();

		pricebookRow3List.add(data);
	}

	String actPriceBookRow3 = pricebookRow3List.toString();
	String expPriceBookRow3 ="[3, V28, , , 01/02/2023, 31/03/2023, AIA ENGINEERING LTD.(U1), Indian Rupees, 0, 0, , 109.230000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow3);
	System.out.println(expPriceBookRow3);
	
	ArrayList<String> pricebookRow4List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow4List.get(i).getText();

		pricebookRow4List.add(data);
	}

	String actPriceBookRow4 = pricebookRow4List.toString();
	String expPriceBookRow4 ="[4, V97, , , 01/02/2023, 31/03/2023, SURYA ALLOYS, Indian Rupees, 0, 0, , 200.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow4);
	System.out.println(expPriceBookRow4);


	ArrayList<String> pricebookRow5List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow5List.get(i).getText();

		pricebookRow5List.add(data);
	}

	String actPriceBookRow5 = pricebookRow5List.toString();
	String expPriceBookRow5 ="[5, ND 150, , , 01/04/2022, 31/03/2023, SUNRISE INDUSTRIES, Indian Rupees, 0, 0, , 173.050000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow5);
	System.out.println(expPriceBookRow5);
	
	ArrayList<String> pricebookRow6List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow6List.get(i).getText();

		pricebookRow6List.add(data);
	}

	String actPriceBookRow6 = pricebookRow6List.toString();
	String expPriceBookRow6 ="[6, V28, , , 01/02/2023, 31/03/2023, ARIHANT TECHNOCAST INDIA PRIVATE LTD, Indian Rupees, 0, 0, , 109.230000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow6);
	System.out.println(expPriceBookRow6);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Currency));
	filter_Currency.click();
	
	boolean result=false;
	
	for(WebElement p:filter_CurrencyList)
	{
		
		if(!(p.getText().isEmpty()))
		if((filter_Currency.getAttribute("value")).equalsIgnoreCase(p.getText()))
		{
			result=true;
		}
		
		
	}
	
	
	
	
	if(actPriceBookRow1.equalsIgnoreCase(expPriceBookRow1) && actPriceBookRow2.equalsIgnoreCase(expPriceBookRow2)
			&& actPriceBookRow3.equalsIgnoreCase(expPriceBookRow3) && result
			&& actPriceBookRow4.equalsIgnoreCase(expPriceBookRow4) && actPriceBookRow5.equalsIgnoreCase(expPriceBookRow5)
			&& actPriceBookRow6.equalsIgnoreCase(expPriceBookRow6) )
	{
		return true;
	}
	else
	{
		return false;
	}
	

}


public static boolean checkItemandCurrencyFilterinSellerPricebook() throws InterruptedException
{


	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookClearBtn));
	buyerPriceBookClearBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys("focus");
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();

	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Item));
	filter_Item.click();
	filter_Item.sendKeys("V28");
	Thread.sleep(2000);
	filter_Item.sendKeys(Keys.TAB);


	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Currency));
	filter_Currency.click();
	filter_Currency.sendKeys("Indian Rupees");
	Thread.sleep(2000);
	filter_Currency.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterBtn));
	priceBookFilterBtn.click();

	Thread.sleep(3000);
	
	Calendar cal=Calendar.getInstance();
	SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
	String presentDate = currentDate.format(cal.getTime());
	

	int count = buyerPriceBook_TableRow1List.size();

	ArrayList<String> pricebookRow1List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow1List.get(i).getText();

		pricebookRow1List.add(data);
	}

	String actPriceBookRow1 = pricebookRow1List.toString();
	String expPriceBookRow1 ="[1, V28, , , 01/02/2023, 31/03/2023, AIA ENGINEERING LTD.(U1), Indian Rupees, 0, 0, , 109.230000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow1);
	System.out.println(expPriceBookRow1);


	ArrayList<String> pricebookRow2List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow2List.get(i).getText();

		pricebookRow2List.add(data);
	}

	String actPriceBookRow2 = pricebookRow2List.toString();
	String expPriceBookRow2 ="[2, V28, , , 01/02/2023, 31/03/2023, ARIHANT TECHNOCAST INDIA PRIVATE LTD, Indian Rupees, 0, 0, , 109.230000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow2);
	System.out.println(expPriceBookRow2);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Item));
	filter_Item.click();
	
	boolean result=false;
	
	for(WebElement p:filter_ItemList)
	{
		
		if(!(p.getText().isEmpty()))
		if((filter_Item.getAttribute("value")).equalsIgnoreCase(p.getText()))
		{
			result=true;
		}
		
		
	}
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Currency));
	filter_Currency.click();
	
	boolean result1=false;
	
	for(WebElement p:filter_CurrencyList)
	{
		
		if(!(p.getText().isEmpty()))
		if((filter_Currency.getAttribute("value")).equalsIgnoreCase(p.getText()))
		{
			result1=true;
		}
		
		
	}
	
	
	
	
	if(actPriceBookRow1.equalsIgnoreCase(expPriceBookRow1) && actPriceBookRow2.equalsIgnoreCase(expPriceBookRow2)
			 && result && result1)
	{
		return true;
	}
	else
	{
		return false;
	}
	


}




public static boolean checkItemandCustomerFilterinSellerPricebook() throws InterruptedException
{



	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookClearBtn));
	buyerPriceBookClearBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys("focus");
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();

	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Item));
	filter_Item.click();
	filter_Item.sendKeys("ND 150");
	Thread.sleep(2000);
	filter_Item.sendKeys(Keys.TAB);


	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Customer));
	filter_Customer.click();
	filter_Customer.sendKeys("SUNRISE INDUSTRIES");
	Thread.sleep(2000);
	filter_Customer.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterBtn));
	priceBookFilterBtn.click();

	Thread.sleep(3000);
	
	Calendar cal=Calendar.getInstance();
	SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
	String presentDate = currentDate.format(cal.getTime());
	

	int count = buyerPriceBook_TableRow1List.size();

	ArrayList<String> pricebookRow1List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow1List.get(i).getText();

		pricebookRow1List.add(data);
	}

	String actPriceBookRow1 = pricebookRow1List.toString();
	String expPriceBookRow1 ="[1, ND 150, , , 01/04/2022, 31/03/2023, SUNRISE INDUSTRIES, Indian Rupees, 0, 0, , 173.050000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow1);
	System.out.println(expPriceBookRow1);


	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Item));
	filter_Item.click();
	
	boolean result=false;
	
	for(WebElement p:filter_ItemList)
	{
		
		if(!(p.getText().isEmpty()))
		if((filter_Item.getAttribute("value")).equalsIgnoreCase(p.getText()))
		{
			result=true;
		}
		
		
	}
	
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Customer));
	filter_Customer.click();
	
	boolean result1=false;
	
	for(WebElement p:filter_CustomerList)
	{
		
		if(!(p.getText().isEmpty()))
		if((filter_Customer.getAttribute("value")).equalsIgnoreCase(p.getText()))
		{
			result1=true;
		}
		
		
	}
	
	
	
	
	if(actPriceBookRow1.equalsIgnoreCase(expPriceBookRow1) && result && result1)
	{
		return true;
	}
	else
	{
		return false;
	}
	



}


public static boolean checkItemCurrencyandCustomerFilterinSellerPricebook() throws InterruptedException
{




	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookClearBtn));
	buyerPriceBookClearBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys("focus");
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();

	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Item));
	filter_Item.click();
	filter_Item.sendKeys("V28");
	Thread.sleep(2000);
	filter_Item.sendKeys(Keys.TAB);

	Thread.sleep(1500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Currency));
	filter_Currency.click();
	filter_Currency.sendKeys("Indian Rupees");
	Thread.sleep(2000);
	filter_Currency.sendKeys(Keys.TAB);
	
	Thread.sleep(1500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Customer));
	filter_Customer.click();
	filter_Customer.sendKeys("AIA ENGINEERING LTD.(U1)");
	Thread.sleep(2000);
	filter_Customer.sendKeys(Keys.TAB);
	Thread.sleep(1500);
	
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterBtn));
	priceBookFilterBtn.click();

	Thread.sleep(3000);
	
	Calendar cal=Calendar.getInstance();
	SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
	String presentDate = currentDate.format(cal.getTime());
	

	int count = buyerPriceBook_TableRow1List.size();

	ArrayList<String> pricebookRow1List = new ArrayList<String>();

	for (int i = 0; i < count; i++) 
	{
		String data = buyerPriceBook_TableRow1List.get(i).getText();

		pricebookRow1List.add(data);
	}

	String actPriceBookRow1 = pricebookRow1List.toString();
	String expPriceBookRow1 ="[1, V28, , , 01/02/2023, 31/03/2023, AIA ENGINEERING LTD.(U1), Indian Rupees, 0, 0, , 109.230000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";

	System.out.println(actPriceBookRow1);
	System.out.println(expPriceBookRow1);


	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Item));
	filter_Item.click();
	
	boolean result=false;
	
	for(WebElement p:filter_ItemList)
	{
		
		if(!(p.getText().isEmpty()))
		if((filter_Item.getAttribute("value")).equalsIgnoreCase(p.getText()))
		{
			result=true;
		}
		
		
	}
	
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Currency));
	filter_Currency.click();
	
	boolean result2=false;
	
	for(WebElement p:filter_CurrencyList)
	{
		
		if(!(p.getText().isEmpty()))
		if((filter_Currency.getAttribute("value")).equalsIgnoreCase(p.getText()))
		{
			result2=true;
		}
		
		
	}
	
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Customer));
	filter_Customer.click();
	
	boolean result1=false;
	
	for(WebElement p:filter_CustomerList)
	{
		
		if(!(p.getText().isEmpty()))
		if((filter_Customer.getAttribute("value")).equalsIgnoreCase(p.getText()))
		{
			result1=true;
		}
		
		
	}
	
	
	
	
	if(actPriceBookRow1.equalsIgnoreCase(expPriceBookRow1) && result && result1 && result2)
	{
		return true;
	}
	else
	{
		return false;
	}
	




}


public boolean checkItemCurrencyCustomerDepartmentFilterinSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookClearBtn));
	buyerPriceBookClearBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys("focus");
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();

	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Item));
	filter_Item.click();
	filter_Item.sendKeys("V28");
	Thread.sleep(2000);
	filter_Item.sendKeys(Keys.TAB);

	Thread.sleep(1500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Currency));
	filter_Currency.click();
	filter_Currency.sendKeys("Indian Rupees");
	Thread.sleep(2000);
	filter_Currency.sendKeys(Keys.TAB);
	
	Thread.sleep(1500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Customer));
	filter_Customer.click();
	filter_Customer.sendKeys("AIA ENGINEERING LTD.(U1)");
	Thread.sleep(2000);
	filter_Customer.sendKeys(Keys.TAB);
	Thread.sleep(1500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Department));
	filter_Department.click();
	filter_Department.sendKeys("AMERICA");
	Thread.sleep(2000);
	filter_Department.sendKeys(Keys.TAB);
	Thread.sleep(1500);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();

	Thread.sleep(3000);
	
	String expMsg="Price Book is empty";
	String actMsg=checkValidationMessage(expMsg);
	if(actMsg.equalsIgnoreCase(expMsg))
	{
		return true;
	}
	else
	{
		return false;
	}
	
	
	
	
}



public static boolean checkExportFileafterApplyingFilter() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
{
	
				getDriver().navigate().refresh();
				Thread.sleep(2000);
			
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
				buyerPriceBook_PriceBook.click();
				buyerPriceBook_PriceBook.sendKeys("focus");
				buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
				
				
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
				priceBookFilterAndLoadBtn.click();
			
				Thread.sleep(3000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Item));
				filter_Item.click();
				filter_Item.sendKeys("V28");
				Thread.sleep(2000);
				filter_Item.sendKeys(Keys.TAB);
			
			
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filter_Currency));
				filter_Currency.click();
				filter_Currency.sendKeys("Indian Rupees");
				Thread.sleep(2000);
				filter_Currency.sendKeys(Keys.TAB);
			
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterBtn));
				priceBookFilterBtn.click();
			
				Thread.sleep(3000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookExportToExcel));
				priceBookExportToExcel.click();
				Thread.sleep(3000);
			
				String expMessage="";
			
				String actMessage=checkValidationMessage(expMessage);
				
			
				
				File Efile=new File(getBaseDir()+"\\autoIt\\ExportFiles\\focus.xlsx");
				
				if(Efile.exists())
				{
					Efile.delete();
				}
				
				
			
				Thread.sleep(5000);
			
				Robot robot = new Robot();
				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_J);
				robot.keyRelease(KeyEvent.VK_J);
				robot.keyRelease(KeyEvent.VK_CONTROL);
			
				Thread.sleep(3000);
			
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
				Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\focus.exe");
				
				Thread.sleep(30000);
				
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
				Thread.sleep(4000);
				
				
				robot.keyPress(KeyEvent.VK_ALT);
				robot.keyPress(KeyEvent.VK_F4);
				
				robot.keyRelease(KeyEvent.VK_ALT);
				robot.keyRelease(KeyEvent.VK_F4);
				Thread.sleep(3000);
				
				
				ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
					
				int actOpenWindowsCount = getDriver().getWindowHandles().size();
				int expOpenWindowsCount = 2;
				
				System.out.println("Number of Windows  : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
				
				Thread.sleep(1000);
			
			 	
			 	getDriver().switchTo().window(openTabs.get(1)).close();
			 	Thread.sleep(1000);
			 	getDriver().switchTo().window(openTabs.get(0));
			 	
			 	
			 	 ExcelReader excelReader = new ExcelReader(POJOUtility.getExcelPath());
			 	Calendar cal=Calendar.getInstance();
				SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
				String presentDate = currentDate.format(cal.getTime());
				
			
			  	String actExcelfile = getBaseDir()+"\\autoIt\\ExportFiles\\focus.xlsx";
			  
			  	String beforePath="//tbody[@id='PriceBookWebGrid_body']/tr[";
				String afterPath="]//td";
				
				
			//	int count = buyerPriceBook_TableRow1List.size();
			  	
			  	boolean result=false;
			  	//Reading Exported Excel File Data
			  	ArrayList<String[]>rowsFromExcel=ExcelReader.readContentFromExcel(actExcelfile);
			  	
			  	
				ArrayList<String> pricebookRow1List = new ArrayList<String>();
				
				System.out.println("Size:"+rowsFromExcel.size());
				
		    	for(int i=1;i<rowsFromExcel.size();i++)
		    	{
		    		ArrayList<String> exportedExcelRow1List = new ArrayList<String>();
		    		for(int j=0;j<rowsFromExcel.get(i).length;j++)
		    		{
		    			
		    			//System.out.print((rowsFromExcel.get(i)[j])+"\t");
		    			
		    			exportedExcelRow1List.add(rowsFromExcel.get(i)[j]);
		    		}
		    		
		    		String act=exportedExcelRow1List.toString();
		    		// System.out.println("Excel" +act);
		    		 
		    		 
		    		 int reportsRow1ListCount = (getDriver().findElements(By.xpath(beforePath + i +"]//td"))).size();
		    			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		    			for(int k=1;k<reportsRow1ListCount;k++)
		    			{
		    				String data = (getDriver().findElements(By.xpath(beforePath + i +"]//td"))).get(k).getText();
		    				reportsRow1ListArray.add(data);
		    			}
		    			String actRow1List = reportsRow1ListArray.toString();
		    		 
		    			// System.out.println(actRow1List);
		    			 
		    			/*// System.out.println(CollectionUtils.isEqualCollection(exportedExcelRow1List, reportsRow1ListArray)); 
		    			 
		    			 Collections.sort(exportedExcelRow1List);
		    			 Collections.sort(reportsRow1ListArray);
		    			 Assert.assertEquals((exportedExcelRow1List.equals(reportsRow1ListArray)),true);
		    			 
		    			 if( Assert.assertTrue(exportedExcelRow1List.equals(reportsRow1ListArray))
		    			 {
		    				 result= true;
		    			 }*/
		    			 
		    			 List<String> list1 = new ArrayList<String>(exportedExcelRow1List);
		    			 List<String> list2 = new ArrayList<String>(reportsRow1ListArray);
		    			 boolean r= list1.retainAll(list2);
			    			
			    			System.err.println("Common in both List		"	+	list1);
			    			
			    			 if(r==true)
			    			 {
			    				 result =true;
			    			 }
			    		 
		    		 
		    		
		    	}
		    	
		
		    	
		    	if(result==true)
		    	{
			  		return true;
		    	}
		    	else
		    	{
		    		return false;
		    	}
}



///Creating Extra Master

		@FindBy(xpath="//*[@id='1112']")
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
		
		
		@FindBy(xpath = "//*[@id='16']")
		public static WebElement settingsMenu;
		
		@FindBy(xpath="//select[@id='SellingRate']")
		public static WebElement sellingRatesSelect;

		public void checkSettingsofBuyerPriceBook(String buyerName) throws InterruptedException
		{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsMenu));
		settingsMenu.click();

		Thread.sleep(1500);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(preferencesBtn));
		preferencesBtn.click();
		Thread.sleep(2000);

	


	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sellingRatesSelect));
		sellingRatesSelect.click();
		Thread.sleep(2000);

		Select s = new Select(sellingRatesSelect);
		s.selectByVisibleText(buyerName);
		
		Thread.sleep(2000);

		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		Thread.sleep(1000);

		getWaitForAlert();

		getAlert().accept();
		Thread.sleep(2000);

		

		}


@FindBy(xpath="//tbody[@id='PriceBookWebGrid_body']")
public static WebElement tableBodyGrid;



@FindBy(xpath="//*[@id='PriceBookWebGrid_option_9_table_body']/tr/td[2]")
private static List<WebElement> statesComboBoxList;

@FindBy(xpath="//input[@id='PriceBookWebGrid_option_9']")
private static WebElement stateValue; 



public boolean checkCreatingSellerPriceBookforState() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		checkSettingsofBuyerPriceBook("State");
		Thread.sleep(2000);
		
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
		buyerPriceBook_PriceBook.sendKeys("SellerState1");
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
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.SPACE);
		Thread.sleep(2000);
		itemValue.sendKeys("V28");
		Thread.sleep(1500);
		
	/*	int count = itemComboBoxList.size();
		
		for(int i=0;i<count;i++)
		{
		String data = itemComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("V28"))
		{
		itemComboBoxList.get(i).click();
		break;
		}
		}*/
		
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
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		itemValue.sendKeys(Keys.SPACE);
		Thread.sleep(2000);
		itemValue.sendKeys("RAW MATERIAL");
		Thread.sleep(2000);
		
		/*int count2 = itemComboBoxList.size();
		
		for(int i=0;i<count2;i++)
		{
		String data = itemComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("RAW MATERIAL"))
		{
		itemComboBoxList.get(i).click();
		break;
		}
		}
		*/
		itemValue.sendKeys(Keys.TAB);
		
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stateValue));
		stateValue.click();
		stateValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		stateValue.sendKeys(Keys.SPACE);
		
		int count11 = statesComboBoxList.size();
		
		for(int i=0;i<count11;i++)
		{
		String data = statesComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("Telangana"))
		{
		statesComboBoxList.get(i).click();
		Thread.sleep(1000);
		break;
		}
		}
		
		stateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("525.23");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("625.89");
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



			public static boolean checkCreatedMasterRateValueinSalesInvoiceVoucher() throws InterruptedException
			{
				checkOpenSalesInvoiceVoucher();
				Thread.sleep(2000);
				
				checkAddingExtraFieldinSalesInvoiceVoucher();
				Thread.sleep(2000);
			
			if(checkAddingFormulaforAccount())
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
				newBtn.click();
			}
			
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StateTxt));
			StateTxt.click();
			StateTxt.sendKeys("AP");
			Thread.sleep(1000);
			StateTxt.sendKeys(Keys.TAB);

			Thread.sleep(1500);
			
			
			
			
			//V28 ITEM
			String actRate=checkRateValue("V28");
			String expRate="1000.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			String item_Txt=enter_ItemTxt.getAttribute("value");
			
			System.out.println("Actual :" + actRate	+"Expected :"	+expRate);
			
		
			
			//Group Item
			
			
			checkOpenSalesInvoiceVoucher();
			Thread.sleep(2000);
			
			
		
		if(checkAddingFormulaforGroupAccount())
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		}
		
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StateTxt));
		StateTxt.click();
		StateTxt.sendKeys("Telangana");
		Thread.sleep(1000);
		StateTxt.sendKeys(Keys.TAB);

		Thread.sleep(1500);
		
		
		
		
		//Dummy ITEM
		String actRate1=checkRateValue("Dummy");
		String expRate1="525.23";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		String item_Txt1=enter_ItemTxt.getAttribute("value");
		
		System.out.println("Actual :" + actRate		+"Expected :"	+expRate);
		System.out.println("Actual :" + actRate1	+"Expected :"	+expRate1);
			
			
			
			if(actRate.equalsIgnoreCase(expRate))
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

			
			
			
			public static void checkAddingExtraFieldinSalesInvoiceVoucher() throws InterruptedException
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

			
			
//////////////////////Footer values/////////////////////////////////////////////
			
			
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
			
			
			
			
			
			public boolean checkAdvExcelImportingSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
			{
				ExcelReader excelReader = new ExcelReader(POJOUtility.getExcelPath());
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			  	homeMenu.click();
			  	
			  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
			  	mastersMenu.click();
			  
			  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		      	homeMasterItemMenu.click();
		      	
		      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
		      	homeMasterItem_sellerPriceBookMenu.click();
			  	
			  	Thread.sleep(2000);
				
				
				String excelfile = getBaseDir()+"\\autoIt\\ImportFiles\\TestSmokeSellingRatesImport.xlsx";
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
				
				Thread.sleep(3000);
			  	
			  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
				buyerPriceBook_PriceBook.click();
				buyerPriceBook_PriceBook.sendKeys(Keys.END);
				Thread.sleep(2000);
				buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
				Thread.sleep(2000);
				buyerPriceBook_PriceBook.sendKeys("TestSmokeSellingRatesAdvExcelImport");
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
				
				Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\FocusSellerImportPrice.exe");
				
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
				pricebook.selectByVisibleText("TestSmokeSellingRatesAdvExcelImport");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAdvAccountHeading));
				sheetAdvAccountHeading.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvbuyerPriceBook_MappingDrpdwn));
				pricebook.selectByVisibleText("Customer");
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAdvPOHeading));
				sheetAdvPOHeading.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvbuyerPriceBook_MappingDrpdwn));
				pricebook.selectByVisibleText("So");
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAdvPTHeading));
				sheetAdvPTHeading.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvbuyerPriceBook_MappingDrpdwn));
				pricebook.selectByVisibleText("St");
				
				Thread.sleep(3000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvImportDataBtn));
				AdvImportDataBtn.click();
			  	
			  	String expMessage="Price book imported successfully";
				
				String actMessage=checkValidationMessage(expMessage);
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
				buyerPriceBook_PriceBook.click();
				buyerPriceBook_PriceBook.sendKeys(Keys.END);
				buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
				buyerPriceBook_PriceBook.sendKeys("TestSmokeSellingRatesAdvExcelImport");
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
				String expPriceBookRow1 ="[1, ITEMS GROUP, , , "+presentDate+", "+endDate+", , Indian Rupees, 0, 0, , 300.000000, 200.000000, 100.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";
				
				System.out.println(actPriceBookRow1);
				System.out.println(expPriceBookRow1);
				
				
				ArrayList<String> pricebookRow2List = new ArrayList<String>();
				
				for (int i = 0; i < count; i++) 
				{
					String data = buyerPriceBook_TableRow2List.get(i).getText();
					
					pricebookRow2List.add(data);
				}
				
				String actPriceBookRow2 = pricebookRow2List.toString();
				String expPriceBookRow2 ="[2, BR COGS ITEM, , , "+presentDate+", "+endDate+", , Indian Rupees, 0, 0, Pcs, 500.000000, 400.000000, 300.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";
				
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
			buyerPriceBook_PriceBook.sendKeys("TestSmokeSellingRatesAdvExcelImport");
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

			
			
			public boolean checkExportingXMLBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
			{
				File Efile=new File(getBaseDir()+"\\autoIt\\ExportFiles\\TestSmokeSellerRates.xml");
				
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
		      	
		      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
		      	homeMasterItem_sellerPriceBookMenu.click();
			  	
			  	Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
				buyerPriceBook_PriceBook.sendKeys("TestSmokeSellingRates");
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
				
				getDriver().navigate().refresh();
				Thread.sleep(2000);
				ExcelReader excelReader = new ExcelReader(POJOUtility.getExcelPath());
				
			/*	File Efile=new File(getBaseDir()+"\\autoIt\\ExportFiles\\TestSmokeSellingRatesImport.xml");
				
				if(Efile.exists())
				{
					Efile.delete();
				}
				*/
				Thread.sleep(2000);
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			  	homeMenu.click();
			  	
			  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
			  	mastersMenu.click();
			  
			  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		      	homeMasterItemMenu.click();
		      	
		      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
			  	homeMasterItem_sellerPriceBookMenu.click();
			  	
			  	Thread.sleep(2000);
				
				
				
			  	
			  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
				buyerPriceBook_PriceBook.click();
				buyerPriceBook_PriceBook.sendKeys(Keys.END);
				Thread.sleep(2000);
				buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
				Thread.sleep(2000);
				buyerPriceBook_PriceBook.sendKeys("TestSmokeSellingRatesXMLImport");
				Thread.sleep(2000);
				buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
				//buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
				
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
				itemFirstCellSelect.click();
				Thread.sleep(2000);
				
			  	
			  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PricebookImportFromXML));
			  	PricebookImportFromXML.click();
			  	
			  /*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvbrowseButton));
				AdvbrowseButton.click();
				
				*/Thread.sleep(2000);
				
				Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\FocusSellerImportPriceXML.exe");
				
				Thread.sleep(3000);
				
				
			  	
			  	String expMessage="Price book imported successfully";
				
				String actMessage=checkValidationMessage(expMessage);
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
				buyerPriceBook_PriceBook.click();
				buyerPriceBook_PriceBook.sendKeys(Keys.END);
				buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
				buyerPriceBook_PriceBook.sendKeys("TestSmokeSellingRatesXMLImport");
				buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
				
				Thread.sleep(3000);
				
				getAction().moveToElement(priceBookFilterAndLoadBtn).build().perform();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
				priceBookFilterAndLoadBtn.click();
				
				Thread.sleep(3000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
				buyerPriceBook_CustamizeSettingBtn.click();
				
			 	Thread.sleep(2000);
			 	
			 	
			 	
			 	int actCount = buyerPriceBook_CustomizeCheckBoxList.size();
				int expCount = 6;
				
				for (int i = 0; i < actCount; i++) 
				{
					if (buyerPriceBook_CustomizeCheckBoxList.get(i).isSelected()==false) 
					{
						buyerPriceBook_CustomizeCheckBoxList.get(i).click();
					}
				}
				
				Thread.sleep(4000);
			 	
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
				buyerPriceBook_CustamizeSettingBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
				itemFirstCellSelect.click();
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
				itemValue.click();
			 	Thread.sleep(2000);
				
				
				int count = buyerPriceBook_TableRow1List.size();
				
				ArrayList<String> pricebookRow1List = new ArrayList<String>();
				
				for (int i = 0; i < count; i++) 
				{
					String data = buyerPriceBook_TableRow1List.get(i).getText();
					
					pricebookRow1List.add(data);
				}
				
				String actPriceBookRow1 = pricebookRow1List.toString();
				String expPriceBookRow1 ="[1, , , AMERICA, 10/09/2020, , , Indian Rupees, 0, 0, , 200.000000, 100.000000, 50.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";
				
				System.out.println(actPriceBookRow1);
				System.out.println(expPriceBookRow1);
				
				
				ArrayList<String> pricebookRow2List = new ArrayList<String>();
				
				for (int i = 0; i < count; i++) 
				{
					String data = buyerPriceBook_TableRow2List.get(i).getText();
					
					pricebookRow2List.add(data);
				}
				
				String actPriceBookRow2 = pricebookRow2List.toString();
				String expPriceBookRow2 ="[2, BR COGS ITEM, , AMERICA, , , Indian Rupees, 0, 0, Dozs, 300.000000, 200.000000, 100.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, 0.000000, , , , Authorized, ]";
				
				System.out.println(actPriceBookRow2);
				System.out.println(expPriceBookRow2);
				
				
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
			
			
			
	
			
			public boolean checkPrintRateHistorySellerPriceBook() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException, AWTException
			{
			
				
				getDriver().navigate().refresh();
				Thread.sleep(2000);

				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		  	homeMenu.click();
		  	
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		  	mastersMenu.click();
		  
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	      	homeMasterItemMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
		  	homeMasterItem_sellerPriceBookMenu.click();
		  	
		  	Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			buyerPriceBook_PriceBook.sendKeys("Pricebook With Account Group");
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
			PRHItemTxt.sendKeys("WA COGS ITEM" );
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
			PRHSaveTxt.sendKeys("SellerLayout");
			
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
			PRHNewLayoutTxt.sendKeys("SellerLayout");
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
			
			
			File Efile=new File(getBaseDir()+"\\autoIt\\ExportFiles\\SellerPriceBookToPDF.pdf");
			
			if(Efile.exists())
			{
				Efile.delete();
			}
			
			Thread.sleep(2000);
			
			Robot robot = new Robot();
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_S);
			robot.keyRelease(KeyEvent.VK_S);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			
			Thread.sleep(5000);
				
			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\SavingSellerPriceBookPdf.exe");
			
			Thread.sleep(8000);
			
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			
			Thread.sleep(3000);
			
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
			
			
			
		 	String actPDF = getBaseDir()+"\\autoIt\\ExportFiles\\SellingPriceBookToPDF.pdf";
			String expPDF = getBaseDir()+"\\autoIt\\ImportFiles\\SellingPriceBookToPDF.pdf";
			
			PDFUtil pdfutil = new PDFUtil();
			
			boolean result = pdfutil.compare(actPDF, expPDF);
			
			String actData = pdfutil.getText(actPDF);
			String expData = pdfutil.getText(expPDF).replaceAll("02/02/2023", currentDate);;
			
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
		
		
		public boolean checkViewPriceCustomizationBtnSellerPriceBook() throws InterruptedException
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
			
			
			

			
	///Delivery Notes
			
			
			
			public boolean checkSavingDeliveryNoteVoucher1WithNoTagsAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
			{
				getDriver().navigate().refresh();
							
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
				inventoryTransactionsMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsSalesMenu));
				inventoryTransactionsSalesMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deliveryNotesBtn));
				deliveryNotesBtn.click();

				Thread.sleep(3000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
				newBtn.click();

				checkValidationMessage("Screen opened");

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				customerAccountTxt.sendKeys("Bank");
				Thread.sleep(2000);
				customerAccountTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
				warehouseTxt.sendKeys("HYDERABAD");
				Thread.sleep(2000);
				warehouseTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
				departmentTxt.sendKeys("DUBAI");
				Thread.sleep(2000);
				departmentTxt.sendKeys(Keys.TAB);
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();
				
				
				String actRow1List = checkRowsValueforItem("BR COGS ITEM");
				String expRow1List = "[BR COGS ITEM, , 1.00, , 5.00, 5.00, 10.00, 10.10, 0.00, 0.00, 10.20, 10.30, 10.40, 10.50, 10.60, 10.70, 10.80, 10.90, 11.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, , , ]";

				Thread.sleep(1000);
				
				
	//RMA
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();
				Thread.sleep(2000);
				
				String actRow2List = checkRowsValueforItem("WA COGS ITEM");
				String expRow2List = "[WA COGS ITEM, , 1.00, , 5.00, 5.00, 20.00, 20.10, 0.00, 0.00, 20.20, 20.30, 20.40, 20.50, 20.60, 20.70, 20.80, 20.90, 21.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, , , ]";

				Thread.sleep(1000);

				

				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
	

				if(actRow1List.equalsIgnoreCase(expRow1List) )
				{
					
					return true;
				}
				else
				{
					
					return false;
				}
			}


@FindBy(xpath="//*[@id='id_transactionentry_new']/a/span")
public static WebElement deliveryNotes_NewBtn;




			public boolean checkSavingDeliveryNoteVoucher2WithPricebookAccountOneAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
			{
				 Thread.sleep(2000);
				
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deliveryNotes_NewBtn));
				 deliveryNotes_NewBtn.click();
				 Thread.sleep(2000);
				 
				 getWaitForAlert();
				 Thread.sleep(2000);
				 getAlert().accept();
				 Thread.sleep(2000);
				 
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				customerAccountTxt.click();
				customerAccountTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
				customerAccountTxt.sendKeys("Pricebook Account One");
				Thread.sleep(2000);
				customerAccountTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
				warehouseTxt.sendKeys("HYDERABAD");
				Thread.sleep(2000);
				warehouseTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
				departmentTxt.sendKeys("DUBAI");
				Thread.sleep(2000);
				departmentTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();
				
				
				
				
				String actRow1List = checkRowsValueforItem("BR COGS ITEM");
				String expRow1List ="[BR COGS ITEM, , 1.00, , 5.00, 5.00, 100.00, 100.10, 100.00, 100.10, 100.20, 100.30, 100.40, 100.50, 100.60, 100.70, 100.80, 100.90, 102.00, 100.20, 100.30, 100.40, 100.50, 100.60, 100.70, 100.80, 100.90, 102.00, , , ]";
				
				
				System.out.println(actRow1List);
				System.out.println(expRow1List);

				
				if(actRow1List.equalsIgnoreCase(expRow1List) )
				{
					
					return true;
				}
				else
				{
					
					return false;
				}
			}



			


			public boolean checkSavingDeliveryNoteVoucher3WithPricebookAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
			{

				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deliveryNotes_NewBtn));
				 deliveryNotes_NewBtn.click();
				 Thread.sleep(2000);
				 
				 getWaitForAlert();
				 Thread.sleep(2000);
				 getAlert().accept();
				 Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				customerAccountTxt.sendKeys("Pricebook Account Two");
				Thread.sleep(2000);
				customerAccountTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
				warehouseTxt.sendKeys("HYDERABAD");
				Thread.sleep(2000);
				warehouseTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
				departmentTxt.sendKeys("DUBAI");
				Thread.sleep(2000);
				departmentTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();
				String actRow1List = checkRowsValueforItem("BR COGS ITEM");
				String expRow1List ="[BR COGS ITEM, , 1.00, , 5.00, 5.00, 500.00, 500.10, 500.00, 500.10, 500.20, 500.30, 500.40, 500.50, 500.60, 500.70, 500.80, 500.90, 501.00, 500.20, 500.30, 500.40, 500.50, 500.60, 500.70, 500.80, 500.90, 501.00, , , ]";
				System.out.println(actRow1List);
				System.out.println(expRow1List);

				
				if(actRow1List.equalsIgnoreCase(expRow1List) )
				{
					
					return true;
				}
				else
				{
					
					return false;
				}
			}







			public boolean checkSavingDeliveryNoteVoucher4WithGroupOneAccountOneAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
			{
				
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deliveryNotes_NewBtn));
				 deliveryNotes_NewBtn.click();
				 Thread.sleep(2000);
				 
				 getWaitForAlert();
				 Thread.sleep(2000);
				 getAlert().accept();
				 Thread.sleep(2000);
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				customerAccountTxt.sendKeys("Group One Account One");
				Thread.sleep(2000);
				customerAccountTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
				warehouseTxt.sendKeys("HYDERABAD");
				Thread.sleep(2000);
				warehouseTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
				departmentTxt.sendKeys("DUBAI");
				Thread.sleep(2000);
				departmentTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();
				String actRow1List = checkRowsValueforItem("BR COGS ITEM");
				String expRow1List ="[BR COGS ITEM, , 1.00, , 5.00, 5.00, 10.00, 10.10, 2,000.00, 2,000.10, 10.20, 10.30, 10.40, 10.50, 10.60, 10.70, 10.80, 10.90, 11.00, 2,000.20, 2,000.30, 2,000.40, 2,000.50, 2,000.60, 2,000.70, 2,000.80, 2,000.90, 2,001.00, , , ]";

				if(actRow1List.equalsIgnoreCase(expRow1List) )
				{
					
					return true;
				}
				else
				{
					
					return false;
				}
			}








			public boolean checkSavingDeliveryNoteVoucher5WithGroupOneAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
			{
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deliveryNotes_NewBtn));
				 deliveryNotes_NewBtn.click();
				 Thread.sleep(2000);
				 
				 getWaitForAlert();
				 Thread.sleep(2000);
				 getAlert().accept();
				 Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				customerAccountTxt.sendKeys("Group One Account Two");
				Thread.sleep(2000);
				customerAccountTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
				warehouseTxt.sendKeys("HYDERABAD");
				Thread.sleep(2000);
				warehouseTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
				departmentTxt.sendKeys("DUBAI");
				Thread.sleep(2000);
				departmentTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();
				String actRow1List = checkRowsValueforItem("BR COGS ITEM");
				String expRow1List ="[BR COGS ITEM, , 1.00, , 5.00, 5.00, 10.00, 10.10, 2,000.00, 2,000.10, 10.20, 10.30, 10.40, 10.50, 10.60, 10.70, 10.80, 10.90, 11.00, 2,000.20, 2,000.30, 2,000.40, 2,000.50, 2,000.60, 2,000.70, 2,000.80, 2,000.90, 2,001.00, , , ]";
				Thread.sleep(1000);

				System.out.println(actRow1List);
				System.out.println(expRow1List);

				
				if(actRow1List.equalsIgnoreCase(expRow1List))
				{
					
					return true;
				}
				else
				{
					
					return false;
				}
			}






			public boolean checkSavingDeliveryNoteVoucher6WithGroupTwoAccountOneAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
			{
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deliveryNotes_NewBtn));
				 deliveryNotes_NewBtn.click();
				 Thread.sleep(2000);
				 
				 getWaitForAlert();
				 Thread.sleep(2000);
				 getAlert().accept();
				 Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				customerAccountTxt.sendKeys("Group Two Account One");
				Thread.sleep(2000);
				customerAccountTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
				warehouseTxt.sendKeys("HYDERABAD");
				Thread.sleep(2000);
				warehouseTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
				departmentTxt.sendKeys("DUBAI");
				Thread.sleep(2000);
				departmentTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();
				String actRow1List = checkRowsValueforItem("BR COGS ITEM");
				String expRow1List ="[BR COGS ITEM, , 1.00, , 5.00, 5.00, 10.00, 10.10, 7,000.00, 7,000.10, 10.20, 10.30, 10.40, 10.50, 10.60, 10.70, 10.80, 10.90, 11.00, 7,000.20, 7,000.30, 7,000.40, 7,000.50, 7,000.60, 7,000.70, 7,000.80, 7,000.90, 7,001.00, , , ]";

				Thread.sleep(1000);

				System.out.println(actRow1List);
				System.out.println(expRow1List);

				
				if(actRow1List.equalsIgnoreCase(expRow1List))
				{
					
					return true;
				}
				else
				{
					
					return false;
				}
			}








			public boolean checkSavingDeliveryNoteVoucher7WithGroupTwoAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
			{
				Thread.sleep(2000); 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deliveryNotes_NewBtn));
				 deliveryNotes_NewBtn.click();
				 Thread.sleep(2000);
				 
				 getWaitForAlert();
				 Thread.sleep(2000);
				 getAlert().accept();
				 Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				customerAccountTxt.sendKeys("Group Two Account Two");
				Thread.sleep(2000);
				customerAccountTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
				warehouseTxt.sendKeys("HYDERABAD");
				Thread.sleep(2000);
				warehouseTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
				departmentTxt.sendKeys("DUBAI");
				Thread.sleep(2000);
				departmentTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();
				String actRow1List = checkRowsValueforItem("BR COGS ITEM");
				String expRow1List ="[BR COGS ITEM, , 1.00, , 5.00, 5.00, 10.00, 10.10, 7,000.00, 7,000.10, 10.20, 10.30, 10.40, 10.50, 10.60, 10.70, 10.80, 10.90, 11.00, 7,000.20, 7,000.30, 7,000.40, 7,000.50, 7,000.60, 7,000.70, 7,000.80, 7,000.90, 7,001.00, , , ]";

				Thread.sleep(1000);

				System.out.println(actRow1List);
				System.out.println(expRow1List);

			
				if(actRow1List.equalsIgnoreCase(expRow1List))
				{
					
					return true;
				}
				else
				{
					
					return false;
				}
			}


	public String checkRowsValueforItem(String itemName) throws InterruptedException
	{
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(itemName);
		Thread.sleep(2000);
		
		enter_ItemTxt.sendKeys(Keys.TAB);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_UnitTxt.sendKeys(Keys.DELETE);
		Thread.sleep(1500);
		enter_UnitTxt.sendKeys(Keys.ESCAPE);
		Thread.sleep(2000);
		enter_UnitTxt.sendKeys(Keys.TAB);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys("1");
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys("5");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt0Txt));
		enter_salrt0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrt1Txt));
		enter_salrt1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp0Txt));
		enter_salrtGrp0Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp1Txt));
		enter_salrtGrp1Txt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_val2Txt));
		enter_val2Txt.sendKeys(Keys.TAB);
		
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_val3Txt));
		enter_val3Txt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_val4Txt));
		enter_val4Txt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_val5Txt));
		enter_val5Txt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_val6Txt));
		enter_val6Txt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_val7Txt));
		enter_val7Txt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_val8Txt));
		enter_val8Txt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_val9Txt));
		enter_val9Txt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_val10Txt));
		enter_val10Txt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp2Txt));
		enter_salrtGrp2Txt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp3Txt));
		enter_salrtGrp3Txt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp4Txt));
		enter_salrtGrp4Txt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp5Txt));
		enter_salrtGrp5Txt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp6Txt));
		enter_salrtGrp6Txt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp7Txt));
		enter_salrtGrp7Txt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp8Txt));
		enter_salrtGrp8Txt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp9Txt));
		enter_salrtGrp9Txt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_salrtGrp10Txt));
		enter_salrtGrp10Txt.sendKeys(Keys.TAB);
		
		
	
		System.out.println(enter_ItemTxt.getAttribute("value"));
		
		
		switch(itemName)
		{
			case "BR COGS ITEM" :
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchCloseBtn));
				batchCloseBtn.click();

				Thread.sleep(2000);
				break;
				
			case "WA COGS ITEM" :
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMACancelBtn));
				RMACancelBtn.click();

				Thread.sleep(2000);
				break;
							
		}
		
		ArrayList<String> row1 = new ArrayList<String>();

		int count = voucherBodyRow1List.size();

		for (int j = 1; j < count; j++) 
		{
			String data = voucherBodyRow1List.get(j).getText();
			row1.add(data);
		}

		String actRow1List = row1.toString();
		
		return actRow1List;
		
	}

	@FindBy(xpath="//*[@id='div_Clear']")		
				
	public static WebElement RMACancelBtn;		
			
	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_table_body']/tr/td[2]")	
	public static List<WebElement> itemList;
			
	
public boolean checkRateValueofanItemBasedonUnitsinSalesInvoiceVoucher() throws InterruptedException
{
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	
	checkOpenSalesInvoiceVoucher();
	Thread.sleep(2000);
	
	
	if(checkAddingFormulaforUnits())
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	}
	
	Thread.sleep(2000);
	
	
	//BR COGS ITEM with quantiy as 12 and Units as Pcs
			String actRateUnit1=checkRateValueforUnits("STD RATE COGS ITEM","Pcs",12);
			String expRateUnit1="111.00";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			String item_Txt1=enter_ItemTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
			select1stRow_2ndColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			String unit_Txt1=enter_UnitTxt.getAttribute("value");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			String qty_Txt1=enter_Quantity.getAttribute("value");
			
			System.out.println("************ BR COGS ITEM Pcs***********************");
			System.out.println("Item 		"		+ item_Txt1	);
			System.out.println("Units 		"		+ unit_Txt1);
			System.out.println("Quantity 	"		+ qty_Txt1	);
			System.out.println("Rate 		"		+ actRateUnit1	+ "Expected 	:"	+	expRateUnit1);
			
			
			
			//BR COGS ITEM with quantiy as 2 and Units as Dozs
			String actRateUnit2=checkRateValueforUnits("STD RATE COGS ITEM","Dozs",2);
			String expRateUnit2="1333.20";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			String item_Txt2=enter_ItemTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
			select1stRow_2ndColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			String unit_Txt2=enter_UnitTxt.getAttribute("value");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			String qty_Txt2=enter_Quantity.getAttribute("value");
			
			
			System.out.println("************ BR COGS ITEM Dozs***********************");
			System.out.println("Item 		"		+ item_Txt2	);
			System.out.println("Units 		"		+ unit_Txt2);
			System.out.println("Quantity 	"		+ qty_Txt2	);
			System.out.println("Rate 		"		+ actRateUnit2	+ "Expected 	:"	+	expRateUnit2);
			
			
			
			
	if(actRateUnit1.equalsIgnoreCase(expRateUnit1)&&actRateUnit2.equalsIgnoreCase(expRateUnit2))
	{
		return true;
	}
	else
	{
		return false;
	}
	
}


	public boolean checkRatevalueforGroupItemsbasedonUnitsinSalesInvoiceVoucher() throws InterruptedException
	{
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_newBtn));
		salesInvoice_newBtn.click();
		Thread.sleep(2000);
		
		if(getIsAlertPresent())
		{
			getWaitForAlert();
			Thread.sleep(2000);
			getAlert().accept();
			Thread.sleep(2000);
		}
	
	//RMA FITEM with quantiy as 2 and Units as Dozs
			String actRateUnit1=checkRateValueforUnits("RMA FITEM","Dozs",2);
			String expRateUnit1="222.10";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			String item_Txt1=enter_ItemTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
			select1stRow_2ndColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			String unit_Txt1=enter_UnitTxt.getAttribute("value");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			String qty_Txt1=enter_Quantity.getAttribute("value");
			
			System.out.println("************ RMA FITEM Dozs***********************");
			System.out.println("Item 		"		+ item_Txt1	);
			System.out.println("Units 		"		+ unit_Txt1);
			System.out.println("Quantity 	"		+ qty_Txt1	);
			System.out.println("Rate 		"		+ actRateUnit1	+ "Expected 	:"	+	expRateUnit1);
			
			
			
			//Group ITEM with quantiy as 156 and Units as Pcs
			String actRateUnit2=checkRateValueforUnits("Group One Item One","Pcs",156);
			String expRateUnit2="335.30";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			String item_Txt2=enter_ItemTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
			select1stRow_2ndColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			String unit_Txt2=enter_UnitTxt.getAttribute("value");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			String qty_Txt2=enter_Quantity.getAttribute("value");
			
			
			System.out.println("************ RMA FITEM Dozs***********************");
			System.out.println("Item 		"		+ item_Txt2	);
			System.out.println("Units 		"		+ unit_Txt2);
			System.out.println("Quantity 	"		+ qty_Txt2	);
			System.out.println("Rate 		"		+ actRateUnit2	+ "Expected 	:"	+	expRateUnit2);
			
			
			
			//Group ITEM with quantiy as 10 and Units as Pcs
			String actRateUnit3=checkRateValueforUnits("Group One Item Two","Pcs",10);
			String expRateUnit3="335.00";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			String item_Txt3=enter_ItemTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
			select1stRow_2ndColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			String unit_Txt3=enter_UnitTxt.getAttribute("value");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			String qty_Txt3=enter_Quantity.getAttribute("value");
			
			
			System.out.println("************ Group Item***********************");
			System.out.println("Item 		"		+ item_Txt3	);
			System.out.println("Units 		"		+ unit_Txt3);
			System.out.println("Quantity 	"		+ qty_Txt3	);
			System.out.println("Rate 		"		+ actRateUnit3	+ "Expected 	:"	+	expRateUnit3);
			
			
			
			//Group ITEM with quantiy as 15 and Units as Pcs
			String actRateUnit4=checkRateValueforUnits("Group Two Item One","Pcs",15);
			String expRateUnit4="0.00";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			String item_Txt4=enter_ItemTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
			select1stRow_2ndColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			String unit_Txt4=enter_UnitTxt.getAttribute("value");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			String qty_Txt4=enter_Quantity.getAttribute("value");
			
			
			System.out.println("************ Group Item***********************");
			System.out.println("Item 		"		+ item_Txt4	);
			System.out.println("Units 		"		+ unit_Txt4);
			System.out.println("Quantity 	"		+ qty_Txt4	);
			System.out.println("Rate 		"		+ actRateUnit4	+ "Expected 	:"	+	expRateUnit4);
			
			
			
			
			
	if(actRateUnit1.equalsIgnoreCase(expRateUnit1)&&actRateUnit2.equalsIgnoreCase(expRateUnit2)
			&&actRateUnit3.equalsIgnoreCase(expRateUnit3)&&actRateUnit4.equalsIgnoreCase(expRateUnit4))
	{
		return true;
	}
	else
	{
		return false;
	}
	
}


	
	
	public boolean checkRateValuewhenquantiyisZerobasedonUnitsinSalesInvoiceVoucher() throws InterruptedException
	{
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoice_newBtn));
		salesInvoice_newBtn.click();
		Thread.sleep(2000);
		
		if(getIsAlertPresent())
		{
			getWaitForAlert();
			Thread.sleep(2000);
			getAlert().accept();
			Thread.sleep(2000);
		}
	
	//RMA FITEM with quantiy as 0 and Units as Dozs
			String actRateUnit1=checkRateValueforUnits("RMA FITEM","Dozs",0);
			String expRateUnit1="222.30";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			String item_Txt1=enter_ItemTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
			select1stRow_2ndColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			String unit_Txt1=enter_UnitTxt.getAttribute("value");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			String qty_Txt1=enter_Quantity.getAttribute("value");
			
			System.out.println("************ RMA FITEM Dozs***********************");
			System.out.println("Item 		"		+ item_Txt1	);
			System.out.println("Units 		"		+ unit_Txt1);
			System.out.println("Quantity 	"		+ qty_Txt1	);
			System.out.println("Rate 		"		+ actRateUnit1	+ "Expected 	:"	+	expRateUnit1);
			
			
			if(actRateUnit1.equalsIgnoreCase(expRateUnit1))
			{
				return true;
			}
			else
			{
				return false;
			}
			
	}
	

public static String checkRateValueforUnits(String item,String unit,int qty) throws InterruptedException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.click();
	enter_ItemTxt.sendKeys(Keys.END);
	enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	enter_ItemTxt.sendKeys(item);
	Thread.sleep(2000);
	
	enter_ItemTxt.sendKeys(Keys.TAB);

	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
	select1stRow_2ndColumn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Units));
	enter_Units.sendKeys(Keys.END);
	enter_Units.sendKeys(Keys.SHIFT,Keys.HOME);

	enter_Units.sendKeys(unit);
	Thread.sleep(2000);
	enter_Units.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	enter_Quantity.sendKeys(Keys.END);
	enter_Quantity.sendKeys(Keys.SHIFT,Keys.HOME);
	enter_Quantity.sendKeys(Integer.toString(qty));
	enter_Quantity.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	select1stRow_11thColumn.click();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	enter_Rate.click();
	Thread.sleep(1000);

	String actRate=enter_Rate.getAttribute("value");
	return actRate;
	
	
}


public static String checkRateValueforQuantity(String item,int qty) throws InterruptedException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.click();
	enter_ItemTxt.sendKeys(Keys.END);
	enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	enter_ItemTxt.sendKeys(item);
	Thread.sleep(2000);
	
	enter_ItemTxt.sendKeys(Keys.TAB);

	Thread.sleep(3000);
	
	

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	enter_Quantity.sendKeys(Keys.END);
	enter_Quantity.sendKeys(Keys.SHIFT,Keys.HOME);
	enter_Quantity.sendKeys(Integer.toString(qty));
	enter_Quantity.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	select1stRow_11thColumn.click();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	enter_Rate.click();
	Thread.sleep(1000);

	String actRate=enter_Rate.getAttribute("value");
	return actRate;
	
	
}


		public boolean checkSavingaPricebookforCurrency() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			
		
			
			getDriver().navigate().refresh();
			Thread.sleep(2000);
		
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
		buyerPriceBook_PriceBook.sendKeys("Seller for Currency");
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
		buyerPriceBook_CustamizeSettingBtn.click();
		if(buyerPriceBook_CurrencyCheckBox.isSelected()==false)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CurrencyCheckBoxSelected));
			buyerPriceBook_CurrencyCheckBoxSelected.click();
			
		}
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
		buyerPriceBook_CustamizeSettingBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		itemValue.sendKeys("RMA BATCH ITEM");
		Thread.sleep(2000);
		
		
		
		itemValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyOptionValue));
		currencyOptionValue.sendKeys("Indian Rupees");
		Thread.sleep(2000);
		currencyOptionValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		//val1Value.click();
		rateValue.sendKeys("18.25");
		rateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("10.25");
		val1Value.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		itemValue.sendKeys("Batch WA ITEM");
		Thread.sleep(2000);
		
		
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyOptionValue));
		currencyOptionValue.sendKeys("Australia Dollars");
		Thread.sleep(2000);
		currencyOptionValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("2.56");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("20.25");
		val1Value.sendKeys(Keys.TAB);
		
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		System.out.println("*********************************checkSavingSellerPriceBook********************************");
		
		String expMessage="Price Book created successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		
		
		
		if (actMessage.equalsIgnoreCase(expMessage)  ) 
		{
		System.out.println("Pass: Saved Successfully");
		
		return true;
		} 
		else 
		{
		System.out.println("Fail:  Saved Successfully");
		
		return false;
		}
		
		
		
		}

		public boolean checkRateValueisLoadingforCurrencyandIteminSalesInvoiceVoucher() throws InterruptedException
		{
			
			checkOpenSalesInvoiceVoucher();
			Thread.sleep(3000);
			
			if(checkAddingFormulaforAccount())
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
				newBtn.click();
			}

			
			
		
		
			Thread.sleep(2000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
			voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			voucherHeaderCurrency.sendKeys("INR");
			voucherHeaderCurrency.sendKeys(Keys.TAB);
			Thread.sleep(2000);
			
			
			//Batch WA Item
			String actRate=checkRateValue("RMA BATCH ITEM");
			String expRate="18.25";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
			voucherHeaderCurrency.click();
			String actCurrency=voucherHeaderCurrency.getAttribute("value");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			String item_Txt=enter_ItemTxt.getAttribute("value");
			
			

			Thread.sleep(2000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
			voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			voucherHeaderCurrency.sendKeys("USD");
			Thread.sleep(2000);
			voucherHeaderCurrency.sendKeys(Keys.TAB);
			Thread.sleep(2000);
			
			
				getWaitForAlert();
				Thread.sleep(2000);
				getAlert().accept();
			
			
			//BATCH WA ITEM
			String actRate1=checkRateValue("BATCH WA ITEM");
			String expRate1="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
			voucherHeaderCurrency.click();
			String actCurrency1=voucherHeaderCurrency.getAttribute("value");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			String item_Txt1=enter_ItemTxt.getAttribute("value");
			
			
			
			System.out.println("Currency	:	"	+	actCurrency);
			
			System.out.println("Item		:	"	+	item_Txt);
			
			System.out.println("Rate		:	"	+	actRate		+	"Expected	:"		+	expRate);
			
			System.out.println("Currency	:	"	+	actCurrency1);
			
			System.out.println("Item		:	"	+	item_Txt1);
			
			System.out.println("Rate		:	"	+	actRate1		+	"Expected	:"		+	expRate1);
			
			if(actRate.equalsIgnoreCase(expRate) && actRate1.equalsIgnoreCase(expRate1) )
			{
				return true;
			}
			else
			{
				return false;
			}
			
			
}
		
		
public boolean checkRateValueinSalesInvoiceVoucherBasedonAbbrevationofPricebookforLessQuantity() throws InterruptedException
{
	
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	
	checkOpenSalesInvoiceVoucher();
	Thread.sleep(2000);
	
	
	if(checkAddingFormulaforAbbrv())
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	}
	
	Thread.sleep(2000);
	
	String actRate=checkRateValueforQuantity("BATCH BIN FINISHED GOODS ITEM",6);
	String expRate="40.00";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt1=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt1=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt1	);
	
	System.out.println("Quantity 	"		+ qty_Txt1	);
	System.out.println("Rate 		"		+ actRate	+ "Expected 	:"	+	expRate);
	
	
	
	String actRate1=checkRateValueforQuantity("BATCH BIN WITH NO RESERVATION ITEM",4);
	String expRate1="60.00";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt2=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt2=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt2	);
	
	System.out.println("Quantity 	"		+ qty_Txt2	);
	System.out.println("Rate 		"		+ actRate1	+ "Expected 	:"	+	expRate1);
	
	
	String actRate2=checkRateValueforQuantity("BATCH BIN WITH NO STOCK UPDATE ITEM",3);
	String expRate2="70.00";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt3=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt3=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt3	);
	
	System.out.println("Quantity 	"		+ qty_Txt3	);
	System.out.println("Rate 		"		+ actRate2	+ "Expected 	:"	+	expRate2);
	
	String actRate3=checkRateValueforQuantity("BATCH IGNORE EXP LIFO ITEM",8);
	String expRate3="90.00";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt4=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt4=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt4	);
	
	System.out.println("Quantity 	"		+ qty_Txt4	);
	System.out.println("Rate 		"		+ actRate3	+ "Expected 	:"	+	expRate3);
	
	
	String actRate4=checkRateValueforQuantity("BATCH FIFO ITEM",1);
	String expRate4="50.00";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt5=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt5=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt5	);
	
	System.out.println("Quantity 	"		+ qty_Txt5	);
	System.out.println("Rate 		"		+ actRate4	+ "Expected 	:"	+	expRate4);
	
	String actRate5=checkRateValueforQuantity("BIN FINISHED GOODS ITEM",6);
	String expRate5="0.00";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt6=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt6=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt6	);
	
	System.out.println("Quantity 	"		+ qty_Txt6	);
	System.out.println("Rate 		"		+ actRate5	+ "Expected 	:"	+	expRate5);
	
	
	
	
	if(actRate.equalsIgnoreCase(expRate) && actRate1.equalsIgnoreCase(expRate1) && actRate2.equalsIgnoreCase(expRate2) 
			&&actRate3.equalsIgnoreCase(expRate3)&& actRate4.equalsIgnoreCase(expRate4) && actRate5.equalsIgnoreCase(expRate5))
	{
		return true;
	}
	else
	{
		return false;
	}
		
	
}
		


public boolean checkRateValueinSalesInvoiceVoucherBasedonAbbrevationofPricebookforMediumQuantity() throws InterruptedException
{
	
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	
	checkOpenSalesInvoiceVoucher();
	Thread.sleep(2000);
	
	
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	
	
	Thread.sleep(2000);
	
	String actRate=checkRateValueforQuantity("BATCH BIN FINISHED GOODS ITEM",15);
	String expRate="120.10";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt1=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt1=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt1	);
	
	System.out.println("Quantity 	"		+ qty_Txt1	);
	System.out.println("Rate 		"		+ actRate	+ "Expected 	:"	+	expRate);
	
	
	
	String actRate1=checkRateValueforQuantity("BATCH BIN WITH NO RESERVATION ITEM",19);
	String expRate1="130.10";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt2=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt2=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt2	);
	
	System.out.println("Quantity 	"		+ qty_Txt2	);
	System.out.println("Rate 		"		+ actRate1	+ "Expected 	:"	+	expRate1);
	
	
	String actRate2=checkRateValueforQuantity("BATCH BIN WITH NO STOCK UPDATE ITEM",20);
	String expRate2="140.10";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt3=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt3=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt3	);
	
	System.out.println("Quantity 	"		+ qty_Txt3	);
	System.out.println("Rate 		"		+ actRate2	+ "Expected 	:"	+	expRate2);
	
	String actRate3=checkRateValueforQuantity("BATCH IGNORE EXP LIFO ITEM",24);
	String expRate3="150.10";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt4=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt4=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt4	);
	
	System.out.println("Quantity 	"		+ qty_Txt4	);
	System.out.println("Rate 		"		+ actRate3	+ "Expected 	:"	+	expRate3);
	
	
	String actRate4=checkRateValueforQuantity("BATCH FIFO ITEM",11);
	String expRate4="170.10";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt5=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt5=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt5	);
	
	System.out.println("Quantity 	"		+ qty_Txt5	);
	System.out.println("Rate 		"		+ actRate4	+ "Expected 	:"	+	expRate4);
	
	String actRate5=checkRateValueforQuantity("BIN FINISHED GOODS ITEM",25);
	String expRate5="0.00";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt6=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt6=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt6	);
	
	System.out.println("Quantity 	"		+ qty_Txt6	);
	System.out.println("Rate 		"		+ actRate5	+ "Expected 	:"	+	expRate5);
	
	
	
	
	if(actRate.equalsIgnoreCase(expRate) && actRate1.equalsIgnoreCase(expRate1) && actRate2.equalsIgnoreCase(expRate2) 
			&&actRate3.equalsIgnoreCase(expRate3)&& actRate4.equalsIgnoreCase(expRate4) && actRate5.equalsIgnoreCase(expRate5))
	{
		return true;
	}
	else
	{
		return false;
	}
		
	
}
		

public boolean checkRateValueinSalesInvoiceVoucherBasedonAbbrevationofPricebookforHighQuantity() throws InterruptedException
{
	
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	
	checkOpenSalesInvoiceVoucher();
	Thread.sleep(2000);
	
	
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	
	
	Thread.sleep(2000);
	
	String actRate=checkRateValueforQuantity("BATCH BIN FINISHED GOODS ITEM",556);
	String expRate="1111.30";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt1=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt1=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt1	);
	
	System.out.println("Quantity 	"		+ qty_Txt1	);
	System.out.println("Rate 		"		+ actRate	+ "Expected 	:"	+	expRate);
	
	
	
	String actRate1=checkRateValueforQuantity("BATCH BIN WITH NO RESERVATION ITEM",878);
	String expRate1="2222.30";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt2=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt2=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt2	);
	
	System.out.println("Quantity 	"		+ qty_Txt2	);
	System.out.println("Rate 		"		+ actRate1	+ "Expected 	:"	+	expRate1);
	
	
	String actRate2=checkRateValueforQuantity("BATCH BIN WITH NO STOCK UPDATE ITEM",613);
	String expRate2="3333.30";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt3=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt3=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt3	);
	
	System.out.println("Quantity 	"		+ qty_Txt3	);
	System.out.println("Rate 		"		+ actRate2	+ "Expected 	:"	+	expRate2);
	
	String actRate3=checkRateValueforQuantity("BATCH IGNORE EXP LIFO ITEM",893);
	String expRate3="4444.30";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt4=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt4=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt4	);
	
	System.out.println("Quantity 	"		+ qty_Txt4	);
	System.out.println("Rate 		"		+ actRate3	+ "Expected 	:"	+	expRate3);
	
	
	String actRate4=checkRateValueforQuantity("BATCH FIFO ITEM",999);
	String expRate4="5555.30";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt5=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt5=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt5	);
	
	System.out.println("Quantity 	"		+ qty_Txt5	);
	System.out.println("Rate 		"		+ actRate4	+ "Expected 	:"	+	expRate4);
	
	String actRate5=checkRateValueforQuantity("BIN FINISHED GOODS ITEM",1000);
	String expRate5="0.00";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt6=enter_ItemTxt.getAttribute("value");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	String qty_Txt6=enter_Quantity.getAttribute("value");
	
	
	
	System.out.println("Item 		"		+ item_Txt6	);
	
	System.out.println("Quantity 	"		+ qty_Txt6	);
	System.out.println("Rate 		"		+ actRate5	+ "Expected 	:"	+	expRate5);
	
	
	
	
	if(actRate.equalsIgnoreCase(expRate) && actRate1.equalsIgnoreCase(expRate1) && actRate2.equalsIgnoreCase(expRate2) 
			&&actRate3.equalsIgnoreCase(expRate3)&& actRate4.equalsIgnoreCase(expRate4) && actRate5.equalsIgnoreCase(expRate5))
	{
		return true;
	}
	else
	{
		return false;
	}
		
	
}
		

@FindBy(xpath="//input[@id='id_body_23_search_popup_input']")
private static WebElement searchPopUpSearchTXT;

@FindBy(xpath="//input[@id='id_body_23_search_popup_category']")
private static WebElement searchPopUpCategoryTxt;

@FindBy(xpath="//button[contains(text(),'Refresh')]")
private static WebElement searchPopUpRefreshbtn;

@FindBy(xpath="//input[@id='id_body_23_search_popup_grid_control_heading_ctrl_1']")
private static WebElement searchPopUpSelectAllChkBox;

@FindBy(xpath="//*[@id='id_body_23_search_popup']/div/div/div[3]/div/div[2]/input[1]")
private static WebElement searchPopUpOkBtn;


/*public boolean checkRateValueFromAbbrevationofPricebook() throws InterruptedException
{
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	
	checkOpenSalesInvoiceVoucher();
	Thread.sleep(2000);
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	
	
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.sendKeys(Keys.F5);
	
	Thread.sleep(2000);
	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchPopUpSearchTXT));
	searchPopUpSearchTXT.click();
	searchPopUpSearchTXT.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	searchPopUpSearchTXT.sendKeys("batch");

	Thread.sleep(2000);
	searchPopUpSearchTXT.sendKeys(Keys.TAB);
	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchPopUpRefreshbtn));
	searchPopUpRefreshbtn.click();
	
	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchPopUpSelectAllChkBox));
	searchPopUpSelectAllChkBox.click();
	
	Thread.sleep(1999);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchPopUpOkBtn));
	searchPopUpOkBtn.click();
	
	Thread.sleep(2000);
	
	
	
}*/
		
		

	public SellerPriceBookValidationPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}

	
	
	

}
