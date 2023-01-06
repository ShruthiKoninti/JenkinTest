
package com.focus.Pages;

import java.awt.AWTException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import javax.security.auth.login.CredentialExpiredException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.omg.Messaging.SyncScopeHelper;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;

import com.focus.Pages.LoginPage;
import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

public class MasterPreloadPage extends BaseEngine  
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "MasterPreloadPage";
	
	public static boolean checkCreatingCostCenterWithStockCostCenter() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	    homeMenu.click();
		
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	    mastersMenu.click();
	     
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(costcenter));
		costcenter.click();
	    	  
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
	    masterNewBtn.click();
		 
		System.out.println("********************************checkCreatingCostCenterWithStockCostCenter*****************************");
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.clear();
		nameTxt.sendKeys(excelReader.getCellData(xlSheetName, 7, 5));
		nameTxt.sendKeys(Keys.TAB);
		
		//Thread.sleep(1000);
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.clear();
		codeTxt.sendKeys(excelReader.getCellData(xlSheetName, 8, 5));
		codeTxt.sendKeys(Keys.TAB);
		
		//Thread.sleep(1000);
		 
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		 Thread.sleep(2000);
		String expMessage=excelReader.getCellData(xlSheetName, 7, 6);
			
		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 7, 7, actMessage);
		
		System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
		
		//Thread.sleep(2000);
			
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 7, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 7, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	public static boolean checkCreatingCostCenterWithSTDRateCostCenter() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		System.out.println("********************************checkCreatingCostCenterWithSTDRateCostCenter*****************************");
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.clear();
		nameTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		nameTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(1000);
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.clear();
		codeTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		codeTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(1000);
		 
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		Thread.sleep(2000);
		 
		String expMessage=excelReader.getCellData(xlSheetName, 9, 6);
			
		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, actMessage);
		
		System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
		
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
	
	
	
	
	
	
	
	
	
	
	static String ItemCostCenterVariableName;
	
	public  boolean checkAddingExtraFieldInGeneralTabAsCostCenter() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
    {
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
	   
		 getDriver().navigate().refresh();
		 Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
      	 homeMenu.click();
      	 Thread.sleep(2000);
      	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
      	 mastersMenu.click();
      	 Thread.sleep(2000);
      	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
      	 homeMasterItemMenu.click();
      	 Thread.sleep(2000);
    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
    	 homeMasterItem_ItemMenu.click();
    	
    	 Thread.sleep(2000);
    	//getAction().moveToElement(mastersSlider).build().perform();
    	
    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersSlider));
    	 mastersSlider.click();
    	 
    	 
    	  	 Thread.sleep(2000);
    	 
    	 getAction().moveToElement(masterCustamizemasterBtn).build().perform();
    	 Thread.sleep(2000);
    	 
    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCustamizemasterBtn));
    	 masterCustamizemasterBtn.click();
    	 
    	 Thread.sleep(2000);
    	 
    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(generalHeaderDetailsTab));
    	 generalHeaderDetailsTab.click();
    	 
    	 Thread.sleep(1000);
    	 
    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountAddBtn));
    	 accountAddBtn.click();
    	 
    	 Thread.sleep(2000);
    	 
    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_FieldDetailsCaption));
    	 extraFields_FieldDetailsCaption.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
    	 extraFields_FieldDetailsCaption.sendKeys(Keys.TAB);
    	 
    	 Select master = new Select(extraFields_FieldDetailsDatTypeDropdown);
    	 master.selectByVisibleText(excelReader.getCellData(xlSheetName, 12, 5));
    	 
    	 Thread.sleep(2000);
    	 
    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterToLinkCombo));
    	 masterToLinkCombo.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
    	 Thread.sleep(2000);
    	 masterToLinkCombo.sendKeys(Keys.TAB);
    	 
    	 Thread.sleep(2000);
    	 
    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFieldOkBtn));
    	 extraFieldOkBtn.click();
    	 
    	 String expMessage = excelReader.getCellData(xlSheetName, 11, 6);
    	 String actMessage = checkValidationMessage(expMessage);
    	 excelReader.setCellData(xlfile, xlSheetName, 11, 7, actMessage);
    	 
    	 Thread.sleep(1000);
    	 
    	 int count = customizeMasterFieldCaptionList.size();
    	 
    	 for (int i = 0; i < count; i++) 
    	 {
			String data = customizeMasterFieldCaptionList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 14, 5))) 
			{
				ItemCostCenterVariableName = customizeMasterVariableNameList.get(i).getText();
				break;
			}
		 }
    	 
    	 System.err.println(ItemCostCenterVariableName);
    	// ((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", accountSaveBtn);
    	 
    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountSaveBtn));
    	 accountSaveBtn.click();
    	 Thread.sleep(2000);
    	 
    	 String expMessage1 = excelReader.getCellData(xlSheetName, 12, 6);
    	 String actMessage1 = checkValidationMessage(expMessage1);
    	 excelReader.setCellData(xlfile, xlSheetName, 12, 7, actMessage1);
    	 
    	// Thread.sleep(1000);
    	 
      	 if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1))
		 {	
      		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeMasterCloseBtn));
      		 customizeMasterCloseBtn.click();
      		excelReader.setCellData(xlfile, xlSheetName, 11, 8, resPass);
      		 return true;
		 }	 
		 else
		 {
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeMasterCloseBtn));
      		 customizeMasterCloseBtn.click();
      		 excelReader.setCellData(xlfile, xlSheetName, 11, 8, resFail);
      		 return false;
		 }
     }
	
	
	
	
	
	
	
	
	// Linking Created Cost Centers to Items
	
	public  boolean checkAssigningCostCenterToStockItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
    {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
      	homeMenu.click();
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
      	mastersMenu.click();
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
      	homeMasterItemMenu.click();
      	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
    	homeMasterItem_ItemMenu.click();
    	
    	Thread.sleep(2000);
    	
    	int count = masterItemGridBodyName.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = masterItemGridBodyName.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 15, 5))) 
			{
				getAction().doubleClick(masterGridBodyChkbox.get(i)).build().perform();
				System.err.println("**Stock Item is Opened**");
				break;
			}
		}
		
		Thread.sleep(4000);
		
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", Item_CostCenterCombo);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Item_CostCenterCombo));
		Item_CostCenterCombo.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));
		Thread.sleep(2000);
		Item_CostCenterCombo.sendKeys(Keys.TAB);
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
		Thread.sleep(1000);
		
		System.out.println("*************************checkAssigningCostCenterToStockItem***************************");
		 
		String expMessage=excelReader.getCellData(xlSheetName, 15, 6);
			
		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 15, 7, actMessage);
		 
		System.out.println("Message  : "+actMessage+" Value Expected : "+expMessage);
		
		Thread.sleep(2000);
			
		if(actMessage.equalsIgnoreCase(expMessage))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 15, 8, resPass);
      		return true;
		}	 
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 15, 8, resFail);
			return false;
		}
     }
	
	
	
	
	
	
	public  boolean checkAssigningCostCenterToSTDRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
    {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
    	int count = masterItemGridBodyName.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = masterItemGridBodyName.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 17, 5))) 
			{
				getAction().doubleClick(masterGridBodyChkbox.get(i)).build().perform();
				System.err.println("**STD RATE COGS ITEM is Opened**");
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Item_CostCenterCombo));
		Item_CostCenterCombo.sendKeys(excelReader.getCellData(xlSheetName, 18, 5));
		Thread.sleep(2000);
		Item_CostCenterCombo.sendKeys(Keys.TAB);
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
		Thread.sleep(1000);
		
		System.out.println("*************************checkAssigningCostCenterToSTDRateItem***************************");
		
		String expMessage=excelReader.getCellData(xlSheetName, 17, 6);
			
		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 17, 7, actMessage);
		 
		System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{	
			excelReader.setCellData(xlfile, xlSheetName, 17, 8, resPass);
      		return true;
		}	 
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 17, 8, resFail);
			return false;
		}
     }
	
	
	
	
	@FindBy(xpath="//*[@id='768']")
	public static WebElement documentCustomizationPurchaseVoucher;
	
	
	// Creating ExtraField in Edit Layout in Body of PVVAT and Assigning Formula.
	
	public static boolean checkAddingFieldsInEditLayoutBodyTabAsCostCenterBodyInPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionBtn));
		configureTransactionBtn.click();
		
		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionDocumentCustomiztionMinus));
		configureTransactionDocumentCustomiztionMinus.click();
		
		Thread.sleep(2000);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomizationPurchaseVoucherVAT));
		documentCustomizationPurchaseVoucher.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTab));
		editLayoutTab.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBodyTab));
		editLayoutBodyTab.click();
		
		Thread.sleep(1000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutAddFieldsBtn));
		editLayoutAddFieldsBtn.click();
		
		getWaitForAlert();
		
		String actAlertTxt	= getAlert().getText();
		String expAlertTxt	= excelReader.getCellData(xlSheetName, 19, 6);
		excelReader.setCellData(xlfile, xlSheetName, 19, 7, actAlertTxt);
		
		getAlert().accept();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCaptionTxt));
		editLayoutCaptionTxt.sendKeys(excelReader.getCellData(xlSheetName, 19, 5));
		Thread.sleep(1000);
		editLayoutCaptionTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		Select dataType = new Select(editLayoutDataTypeDropdown);
		dataType.selectByVisibleText(excelReader.getCellData(xlSheetName, 20, 5));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutDefaultValueTxt));
		editLayoutDefaultValueTxt.sendKeys(/*"&#MPF2053"+*/excelReader.getCellData(xlSheetName, 21, 5));
		Thread.sleep(1000);
		editLayoutDefaultValueTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutMasterToLinkTxt));
		editLayoutMasterToLinkTxt.sendKeys(excelReader.getCellData(xlSheetName, 22, 5));
		Thread.sleep(2000);
		editLayoutMasterToLinkTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutApplyBtn));
		editLayoutApplyBtn.click();		
		
		String expMessage = excelReader.getCellData(xlSheetName, 21, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actMessage);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtnInPvVAT));
		updateBtnInPvVAT.click();	
		
		String actMessage1 = checkValidationMessage(expMessage);
		
		Thread.sleep(1000);
		
		if (actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 19, 8, resPass);
			return  true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 19, 8, resFail);
			return false;
		}
	}
	
	
	
	@FindBy(xpath="//*[@id='id_header_1']")
	public static WebElement documentTxt;
	
	
	// Validating Master Preload in PVVAT
	
	public static boolean CheckSavingPurchaseVoucherWithStockItemAndSTDRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionsExpandBtn));
		transactionsExpandBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesExpandBtn));
		purchasesExpandBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesVoucherBtn));
		purchasesVoucherBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		    	
		Thread.sleep(2000);
		checkValidationMessage("Screen opened");
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentTxt));
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
		purchaseAccountTxt.click();
		purchaseAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 23, 5));
		Thread.sleep(2000);
		purchaseAccountTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 24, 5));
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.click();
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 25, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 26, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
    	select1stRow_1stColumn.click();
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
    	enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 27, 5));
    	Thread.sleep(2000);
    	enter_ItemTxt.sendKeys(Keys.TAB);
    	
    	Thread.sleep(2000);
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
    	enter_UnitTxt.sendKeys(Keys.TAB);
       	
     	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
     	select1stRow_4thColumn.click();
       	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
       	enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 28, 5));
    	enter_Quantity.sendKeys(Keys.TAB);
       	
    	Thread.sleep(3000);
       	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
    	enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 29, 5));
    	enter_Rate.sendKeys(Keys.TAB);
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
    	enter_Gross.sendKeys(Keys.TAB);
    	
    	Thread.sleep(1000);
    	
       	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
    	String actR1CostCenterBody = select1stRow_9thColumn.getText();
    	String expR1CostCenterBody = excelReader.getCellData(xlSheetName, 23, 6);
    	excelReader.setCellData(xlfile, xlSheetName, 23, 7, actR1CostCenterBody);
    	
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
    	select2ndRow_1stColumn.click();
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
    	enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 30, 5));
    	Thread.sleep(2000);
    	enter_ItemTxt.sendKeys(Keys.TAB);
    	
    	Thread.sleep(2000);
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
    	enter_UnitTxt.sendKeys(Keys.TAB);
       	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
    	select2ndRow_4thColumn.click();
       	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
       	enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 31, 5));
    	enter_Quantity.sendKeys(Keys.TAB);
       	
    	Thread.sleep(3000);
       	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
    	enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 32, 5));
    	enter_Rate.sendKeys(Keys.TAB);
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
    	enter_Gross.sendKeys(Keys.TAB);
    	
    	Thread.sleep(1000);
    	
       	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
    	String actR2CostCenterBody = select2ndRow_9thColumn.getText();
    	String expR2CostCenterBody = excelReader.getCellData(xlSheetName, 24, 6);
    	excelReader.setCellData(xlfile, xlSheetName, 24, 7, actR2CostCenterBody);

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();
		
			 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();
		
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		/*HashSet<String> actMsg = new HashSet();
		
		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMsg.add(data);
		}*/
		
		/*HashSet<String> expMessage = new HashSet();
		
		expMessage.add("Voucher saved successfully : 1");
		expMessage.add("Saving in background.");*/
		
		
		/*String expMessage = excelReader.getCellData(xlSheetName, 25, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 25, 7, actMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 25, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 26, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 25, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		
		
		System.out.println("*********************************CheckSavingPurchaseVoucherVATWithStockItemAndSTDRateItem**************************************");
		
		System.out.println("R1CostCenterBody  : "+actR1CostCenterBody+"  Value Expected  "+expR1CostCenterBody);
		System.out.println("R2CostCenterBody  : "+actR2CostCenterBody+"  Value Expected  "+expR2CostCenterBody);
    	
		if (actR1CostCenterBody.equalsIgnoreCase(expR1CostCenterBody) && actR2CostCenterBody.equalsIgnoreCase(expR2CostCenterBody)  
				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 23, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 23, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	public static boolean CheckCostCenterInSavedVoucherWithStockItemAndSTDRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();
		
		checkValidationMessage("Voucher Loaded Successfully");
		
		Thread.sleep(3000);
		
       	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
    	String actR1CostCenterBody = select1stRow_9thColumn.getText();
    	String expR1CostCenterBody = excelReader.getCellData(xlSheetName, 33, 6);
    	excelReader.setCellData(xlfile, xlSheetName, 33, 7, actR1CostCenterBody);
    	
       	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
    	String actR2CostCenterBody = select2ndRow_9thColumn.getText();
    	String expR2CostCenterBody = excelReader.getCellData(xlSheetName, 34, 6);
    	excelReader.setCellData(xlfile, xlSheetName, 34, 7, actR1CostCenterBody);
    	
		System.out.println("*********************************CheckCostCenterInSavedVoucherWithStockItemAndSTDRateItem**************************************");
		
		System.out.println("R1CostCenterBody  : "+actR1CostCenterBody+"  Value Expected  "+expR1CostCenterBody);
		System.out.println("R2CostCenterBody  : "+actR2CostCenterBody+"  Value Expected  "+expR2CostCenterBody);
    	
		if (actR1CostCenterBody.equalsIgnoreCase(expR1CostCenterBody) && actR2CostCenterBody.equalsIgnoreCase(expR2CostCenterBody))
		{
			excelReader.setCellData(xlfile, xlSheetName, 33, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 33, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	public static boolean CheckEditingPurchaseVoucher1WithStockItemAndSTDRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
    	enter_ItemTxt.click();
    	enter_ItemTxt.sendKeys(Keys.END);
    	enter_ItemTxt.sendKeys(Keys.SHIFT, Keys.HOME);
    	enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 35, 5));
    	Thread.sleep(2000);
    	enter_ItemTxt.sendKeys(Keys.TAB);
    	
    	Thread.sleep(2000);
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
    	enter_UnitTxt.sendKeys(Keys.TAB);
       	
     	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
     	select1stRow_6thColumn.click();
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
    	enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 36, 5));
    	enter_Rate.sendKeys(Keys.TAB);
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
    	enter_Gross.sendKeys(Keys.TAB);
    	
    	Thread.sleep(1000);
    	
       	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
    	String actR1CostCenterBody = select1stRow_9thColumn.getText();
    	String expR1CostCenterBody = excelReader.getCellData(xlSheetName, 35, 6);
       	excelReader.setCellData(xlfile, xlSheetName, 35, 7, actR1CostCenterBody);
    	
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
    	select2ndRow_1stColumn.click();
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
    	enter_ItemTxt.click();
    	enter_ItemTxt.sendKeys(Keys.END);
    	enter_ItemTxt.sendKeys(Keys.SHIFT, Keys.HOME);
    	enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 37, 5));
    	Thread.sleep(2000);
    	enter_ItemTxt.sendKeys(Keys.TAB);
    	
    	Thread.sleep(2000);
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
    	enter_UnitTxt.sendKeys(Keys.TAB);
       	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_6thColumn));
    	select2ndRow_6thColumn.click();
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
    	enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 38, 5));
    	enter_Rate.sendKeys(Keys.TAB);
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
    	enter_Gross.sendKeys(Keys.TAB);
    	
    	Thread.sleep(1000);
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
       	getAction().moveToElement(select2ndRow_9thColumn).build().perform();
    	String actR2CostCenterBody = select2ndRow_9thColumn.getText();
    	String expR2CostCenterBody = excelReader.getCellData(xlSheetName, 36, 6);
    	excelReader.setCellData(xlfile, xlSheetName, 36, 7, actR2CostCenterBody);
    	
    	Thread.sleep(2000);
    	
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
    	select3rdRow_1stColumn.click();
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
    	enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 39, 5));
    	Thread.sleep(2000);
    	enter_ItemTxt.sendKeys(Keys.TAB);
    	
    	Thread.sleep(2000);

    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
    	enter_UnitTxt.sendKeys(Keys.TAB);

    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_4thColumn));
    	select3rdRow_4thColumn.click();
       	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
       	enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 40, 5));
    	enter_Quantity.sendKeys(Keys.TAB);
       	
    	Thread.sleep(3000);
       	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
    	enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 41, 5));
    	enter_Rate.sendKeys(Keys.TAB);
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
    	enter_Gross.sendKeys(Keys.TAB);
    	
    	Thread.sleep(1000);
       	
       	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_9thColumn));
    	String actR3CostCenterBody = Boolean.toString(select3rdRow_9thColumn.getText().isEmpty());
    	String expR3CostCenterBody = excelReader.getCellData(xlSheetName, 37, 6);
    	excelReader.setCellData(xlfile, xlSheetName, 37, 7, actR3CostCenterBody);
    	
    	String thirdRowCostCenterBody = select3rdRow_9thColumn.getText();
    	
       	
    	System.out.println("*********************************CheckEditingPurchaseVoucherVAT3WithStockItemAndSTDRateItem**************************************");
		
		System.out.println("R1CostCenterBody  : "+actR1CostCenterBody+"  Value Expected  "+expR1CostCenterBody);
		System.out.println("R2CostCenterBody  : "+actR2CostCenterBody+"  Value Expected  "+expR2CostCenterBody);
		System.out.println("R3CostCenterBody  : "+actR3CostCenterBody+"  Value Expected  "+expR3CostCenterBody);
    	System.err.println("*"+thirdRowCostCenterBody+"*");
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		Thread.sleep(2000);
			 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		/*HashSet<String> actMsg = new HashSet();
		
		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMsg.add(data);
		}*/
		
		/*HashSet<String> expMessage = new HashSet();
		
		expMessage.add("Voucher saved successfully : 1");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = actMsg.toString();
		String expMessage = excelReader.getCellData(xlSheetName, 38, 6);
		excelReader.setCellData(xlfile, xlSheetName, 38, 7, actMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
		 
*/    
String expMessage1 = " Voucher saved successfully";/*excelReader.getCellData(xlSheetName, 38, 6)*/;
		
		String actMessage = checkValidationMessage(expMessage1);
		/*String expMessage2 = excelReader.getCellData(xlSheetName, 39, 6);
		*/
		String expMessage2=": 1";
		//excelReader.setCellData(xlfile, xlSheetName, 38, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		
		if (actR1CostCenterBody.equalsIgnoreCase(expR1CostCenterBody) && actR2CostCenterBody.equalsIgnoreCase(expR2CostCenterBody) 
    			&& actR3CostCenterBody.equalsIgnoreCase(expR3CostCenterBody) && actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2)) 
    	{
    		excelReader.setCellData(xlfile, xlSheetName, 35, 8, resPass);
			return true;
		}
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 35, 8, resFail);
			return false;
		}
	}
	
	
	
	
	public boolean checkSavedPurchaseVoucherAfterEditingVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();
		
		checkValidationMessage("Voucher Loaded Successfully");
		
       	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
    	String actR1CostCenterBody = select1stRow_9thColumn.getText();
    	String expR1CostCenterBody = excelReader.getCellData(xlSheetName, 42, 6);
    	excelReader.setCellData(xlfile, xlSheetName, 42, 7, actR1CostCenterBody);
    	
       	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
    	String actR2CostCenterBody = select2ndRow_9thColumn.getText();
    	String expR2CostCenterBody = excelReader.getCellData(xlSheetName, 43, 6);
    	excelReader.setCellData(xlfile, xlSheetName, 43, 7, actR2CostCenterBody);
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_9thColumn));
    	String  actR3CostCenterBody = Boolean.toString(select3rdRow_9thColumn.getText().isEmpty());
    	String  expR3CostCenterBody = excelReader.getCellData(xlSheetName, 44, 6);
    	excelReader.setCellData(xlfile, xlSheetName, 44, 7, actR3CostCenterBody);
    	
		System.out.println("*********************************checkSavedPVVATAfterEditingVoucher**************************************");
		
		System.out.println("R1CostCenterBody  : "+actR1CostCenterBody+"  Value Expected  "+expR1CostCenterBody);
		System.out.println("R2CostCenterBody  : "+actR2CostCenterBody+"  Value Expected  "+expR2CostCenterBody);
		System.out.println("R3CostCenterBody  : "+actR3CostCenterBody+"  Value Expected  "+expR3CostCenterBody);
    	
		if (actR1CostCenterBody.equalsIgnoreCase(expR1CostCenterBody) && actR2CostCenterBody.equalsIgnoreCase(expR2CostCenterBody) && 
				actR3CostCenterBody.equalsIgnoreCase(expR3CostCenterBody))
		{
			excelReader.setCellData(xlfile, xlSheetName, 42, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 42, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	@FindBy(xpath="(//*[@id='panelsStayOpen-headingOne']/button)[2]")
	public static WebElement inventorySettingsBtn;
	
	
	public static boolean checkUpdatingItemRateCannotBeBelowCostForSalesInvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		//getDriver().navigate().refresh();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();
					
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
					
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
		financialTransactionSalesMenu.click();
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
		salesInvoiceVATVoucher.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsBtn));
		settingsBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryoptionsTab));
		inventoryoptionsTab.click();
		
		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventorySettingsBtn));
		inventorySettingsBtn.click();
		Thread.sleep(2000);*/
		
		Select ItemRate = new Select(inventoryOptionsTabItemRateDrpdwn);
		ItemRate.selectByVisibleText(excelReader.getCellData(xlSheetName, 45, 5));
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		
		System.out.println("******************checkUpdatingItemRateCannotBeBelowCostForSalesInvoiceVAT*****************");
		
		String expMessage = excelReader.getCellData(xlSheetName, 45, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 45, 7, actMessage);
		
		Thread.sleep(2000);
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 45, 8, resPass);
			return true;
		}
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 45, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	public static boolean checkValidatingCannotBeBelowCostDropdownAndSavingVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		    	
		checkValidationMessage("Screen opened");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 46, 5));
		Thread.sleep(2000);	
	  	customerAccountTxt.sendKeys(Keys.TAB);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 47, 5));
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
		salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 48, 5));
		Thread.sleep(2000);
		salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 49, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 50, 5));
		Thread.sleep(2000);		
	  	pvWareHouseTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	  	enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 51, 5));
	  	Thread.sleep(2000);	
	  	enter_ItemTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 52, 5));
		enter_AQTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
		enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 53, 5));
		enter_FQTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(Keys.TAB);
		
		Thread.sleep(1000);

		if (validationConfirmationMessage.getText().isEmpty()==false) 
		{
			System.err.println(errorMessage.getText());
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
			errorMessageCloseBtn.click();
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 54, 5));
		enter_Rate.sendKeys(Keys.TAB);
		
		String expMessage = excelReader.getCellData(xlSheetName, 46, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 46, 7, actMessage);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		String actRate = select1stRow_14thColumn.getText();
		String expRate = excelReader.getCellData(xlSheetName, 48, 6);
		excelReader.setCellData(xlfile, xlSheetName, 48, 7, actRate);
		
		System.out.println("Rate Cannot Below   : "+actMessage +"  Value Expected  "+expMessage);
		System.out.println("Rate After Changing : "+actRate    +"  Value Expected  "+expRate);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		if (errorMessage.getText().equalsIgnoreCase(excelReader.getCellData(xlSheetName, 49, 6))) 
		{
			System.err.println(errorMessage.getText());
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
			errorMessageCloseBtn.click();
		}
		
		/*HashSet<String> actMsg = new HashSet();
		
		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMsg.add(data);
		}
		
		String actSavingMessage = actMsg.toString();
		
		HashSet<String> expSavingMessage = new HashSet();
		
		expSavingMessage.add("Voucher saved successfully : 1");
		expSavingMessage.add("Saving in background.");
		
		String expSavingMessage = excelReader.getCellData(xlSheetName, 51, 6);
		excelReader.setCellData(xlfile, xlSheetName, 51, 7, actSavingMessage);
		
		System.out.println("Actual Message    : "+actSavingMessage);
		System.out.println("Expected Message  : "+expSavingMessage);
		*/
		String expMessage5 = excelReader.getCellData(xlSheetName, 49, 6);
		
		String actMessage5 = checkValidationMessage(expMessage5);
		
		//String expMessage1 = excelReader.getCellData(xlSheetName, 51, 6);
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage1 = checkValidationMessage(expMessage1);
		//String expMessage2 = excelReader.getCellData(xlSheetName, 52, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 51, 7, actMessage);
		excelReader.setCellData(xlfile, xlSheetName, 49, 7, actMessage5);
		String expMessage2 =": 1";
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		
		
		
		if(actMessage.equalsIgnoreCase(expMessage) && actRate.equalsIgnoreCase(expRate) 
				&& actMessage1.startsWith(expMessage1)&& actMessage1.endsWith(expMessage2))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(1000);
			excelReader.setCellData(xlfile, xlSheetName, 46, 8, resPass);			
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(1000);
			excelReader.setCellData(xlfile, xlSheetName, 46, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	public static boolean checkUpdatingItemRateCannotBeChangedForSalesInvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsBtn));
		settingsBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryoptionsTab));
		inventoryoptionsTab.click();
		
		Thread.sleep(2000);
		
		Select ItemRate = new Select(inventoryOptionsTabItemRateDrpdwn);
		ItemRate.selectByVisibleText(excelReader.getCellData(xlSheetName, 55, 5));
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		
		System.out.println("******************checkUpdatingItemRateCannotBeChangedForSalesInvoiceVAT*****************");
		
		String expMessage = excelReader.getCellData(xlSheetName, 55, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 55, 7, actMessage);
		
		Thread.sleep(2000);
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 55, 8, resPass);
			return true;
		}
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 55, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	public static boolean checkValidatingCannotBeChangedInSavedVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		int count = grid_VoucherNoList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();
			
			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 56, 5))) 
			{
				getAction().doubleClick(grid_CheckBoxList.get(i)).build().perform();
			}
		}
		
		checkValidationMessage("Screen opened");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		String actRateColumn = select1stRow_14thColumn.getAttribute("style");
		String expRateColumn = excelReader.getCellData(xlSheetName, 56, 6);
		excelReader.setCellData(xlfile, xlSheetName, 56, 7, actRateColumn);
		
		Thread.sleep(1000);
		
		System.out.println(actRateColumn);
		System.out.println(expRateColumn);
		
		if(actRateColumn.equalsIgnoreCase(expRateColumn))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(1000);
			excelReader.setCellData(xlfile, xlSheetName, 56, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(1000);
			excelReader.setCellData(xlfile, xlSheetName, 56, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	public static boolean checkUpdatingItemRateEditableForSalesInvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsBtn));
		settingsBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryoptionsTab));
		inventoryoptionsTab.click();
		
		Thread.sleep(2000);
		
		Select ItemRate = new Select(inventoryOptionsTabItemRateDrpdwn);
		ItemRate.selectByVisibleText(excelReader.getCellData(xlSheetName, 60, 5));
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		
		System.out.println("******************checkUpdatingItemRateCannotBeChangedForSalesInvoiceVAT*****************");
		
		String expMessage = excelReader.getCellData(xlSheetName, 60, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 60, 7, actMessage);
		
		Thread.sleep(2000);
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 60, 8, resPass);
			return true;
		}
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationCloseBtn));
			customizationCloseBtn.click();
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 60, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	
	public boolean checkEraseAllTransactionsAfterCompletingAllValidations() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataMangementMenu));
		dataMangementMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseAll));
		eraseAll.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseTranscationsRadio));
		eraseTranscationsRadio.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseAllOkBtn));
		eraseAllOkBtn.click();
		
		if(getIsAlertPresent())
		{
			getWaitForAlert();
			
			getAlert().accept();
		}
		
		String expMessage = excelReader.getCellData(xlSheetName, 61, 6);
		
		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 61, 7, actMessage);
		
		System.out.println("************************************* checkEraseAllTransactionsAfterCompletingAllValidations  *********************************");
		System.out.println("Message  :  "+actMessage +" Value Expected : "+expMessage);
		
		if (actMessage.equalsIgnoreCase(expMessage))
		{
			System.out.println("Voucher Saved");
			excelReader.setCellData(xlfile, xlSheetName, 61, 8, resPass);
			return true;
		}
		else
		{
			System.out.println("Voucher Not Saved");
			excelReader.setCellData(xlfile, xlSheetName, 61, 8, resFail);
			return false;
		}
	}
	
	
	
	
	public static boolean checkRestoreAutomationCompany() throws InterruptedException, IOException, AWTException
	{
		String actMessage=BaseEngine.restoreCompany("SanityUptoMastersMounted","Automation Company");
		String expMessage="Restore company code : 010";
		
		System.err.println("Actual Text :"	+actMessage);
		System.err.println("Expected Text :"	+expMessage);
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	 public MasterPreloadPage(WebDriver driver)
	 {
	    PageFactory.initElements(driver, this);
	 }
	
		
}
