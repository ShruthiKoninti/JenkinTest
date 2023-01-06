

package com.focus.Pages;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.lang.reflect.Method;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import com.focus.base.BaseEngine;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.collections4.bag.SynchronizedSortedBag;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ddf.EscherColorRef.SysIndexProcedure;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.Keys;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import com.focus.base.BaseEngine;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.remote.server.handler.GetAllWindowHandles;
import org.openqa.selenium.remote.server.handler.GetWindowSize;


public class MasterBuyingAndSellingPage extends BaseEngine 
{	
	@FindBy(xpath="//*[@id='1']/div/div")
	private static WebElement  homeMenu;   

	@FindBy(xpath="//*[@id='1000']/span")
	private static WebElement  homeMasterMenu;

	@FindBy(xpath="//*[@id='221']/span")
	private static WebElement  homeMasterItemMenu;
	
	@FindBy(xpath="//*[@id='80']/span")
	private static WebElement  homeMasterItem_buyerPriceBookMenu;
	
	@FindBy(xpath="//*[@id='ctrlOptionProPriceBookH']")
	private static WebElement buyerPriceBook_PriceBook;
	
	@FindBy(xpath="//*[@id='ctrlOptionProPriceBookH_input_settings']/span")
	private static WebElement buyerPriceBook_PriceBookSettingBtn;

	@FindBy(xpath="//*[@id='txtAbbreviation']")
	private static WebElement buyerPriceBook_AbbreviationBtn;
	
	@FindBy(xpath="//*[@id='PriceBookContainer']/div[2]/nav/div/div[1]/a/div/span[2]")
	private static WebElement buyerPriceBookLabel;
	
	@FindBy(xpath="//*[@id='btnClear']")
	private static WebElement buyerPriceBookClearBtn;
	
	@FindBy(xpath="//*[@id='btnPaste']")
	private static WebElement buyerPriceBookPasteBtn;
	
	@FindBy(xpath="//*[@id='btnCopy']")
	private static WebElement buyerPriceBookCopyBtn;
	
	@FindBy(xpath="//*[@id='btnSelectAll']")
	private static WebElement buyerPriceBookSelectAllBtn;
	
	@FindBy(xpath="//*[@id='btnClose']")
	private static WebElement buyerPriceBookCloseBtn;
	
	@FindBy(xpath="//*[@id='btnSave']")
	private static WebElement buyerPriceBookSaveBtn;
	
	@FindBy(xpath="//span[@id='btnClose']")
	private static WebElement buyerPriceBookCloseAfterImport;
	
	
	@FindBy(xpath="//*[@id='btnDelete']")
	private static WebElement buyerPriceBookDeleteBtn;
	
	@FindBy(xpath="//*[@id='StartingDateH']")
	private static WebElement buyerPriceBookStaringDate;
	
	@FindBy(xpath="//*[@id='EndingDateH']")
	private static WebElement buyerPriceBookENdingDate;

	@FindBy(xpath="//*[@id='InActive']")
	private static WebElement buyerPriceBookInActiveCheckBox;
	
	@FindBy(xpath="//*[@id='FilterProduct']")
	private static WebElement buyerPriceBookFilter_Item;
	
	@FindBy(xpath="//*[@id='FilterCurrency']")
	private static WebElement buyerPriceBookFilter_Currency;
	
	@FindBy(xpath="//*[@id='FilterCustomer']")
	private static WebElement buyerPriceBookFilter_Customer;
	
	@FindBy(xpath="//*[@id='FilterDepartment']")
	private static WebElement buyerPriceBookFilter_Department;
	
	@FindBy(xpath="//*[@id='btnFilterClear']")
	private static WebElement buyerPriceBookFilter_ClearBtn;
	
	@FindBy(xpath="//*[@id='btnFilter']")
	private static WebElement buyerPriceBookFilter_FilterBtn;
	
	@FindBy(xpath="//*[@id='btnLoad']")
	private static WebElement buyerPriceBookFilter_FilterAndLoadBtn;
	
	@FindBy(xpath="//*[@id='btnAdvExcelImport']")
	private static WebElement buyerPriceBook_AdvanceExcelImportBtn;
	
	@FindBy(xpath="//*[@id='btnPBPrintRateHistory']")
	private static WebElement buyerPriceBook_PrintRateHistoryBtn;
	
	@FindBy(xpath="//*[@id='btnPBExportToXML']")
	private static WebElement buyerPriceBook_ExportToXmlBtn;
	
	@FindBy(xpath="//span[contains(text(),'Import From XML')]")
	private static WebElement buyerPriceBook_ImportFromXmlBtn;
	
	@FindBy(xpath="//form[@id='formLoadDataXML']//button[@id='btnExporttoExcel']")
	private static WebElement buyerPriceBook_ExportToExcel;
	
	@FindBy(xpath="//button[@id='btnPBImportFromExcel']")
	private static WebElement buyerPriceBook_ImportFromExcel;
	
	@FindBy(xpath="//button[@id='btnViewPrice']")
	private static WebElement buyerPriceBook_ViewPrice;
	
	@FindBy(xpath="//button[@id='btnUpdatePriceList']")
	private static WebElement buyerPriceBook_UpdatePriceList;
	
	@FindBy(xpath="//*[@id='toggleDropDown']/i")
	private static WebElement buyerPriceBook_CustamizeSettingBtn;
	
	@FindBy(xpath="//*[@id='chkDateRange']")
	private static WebElement buyerPriceBook_CustamizeDateRangeCheckBox;
	
	@FindBy(xpath="//*[@id='chkCustomer']")
	private static WebElement buyerPriceBook_CustomizeVendorCheckBox;
	
	@FindBy(xpath="//*[@id='chkDepartment']")
	private static WebElement buyerPriceBook_WarehouseCheckBox;

	@FindBy(xpath="//*[@id='chkQtyRange']")
	private static WebElement buyerPriceBook_QualityRangeCheckBox;
	
	@FindBy(xpath="//*[@id='chkCurrency']")
	private static WebElement buyerPriceBook_CurrencyCheckBox;
	
	@FindBy(xpath="//*[@id='chkUnit']")
	private static WebElement buyerPriceBook_UnitCheckBox;
	
	//Dash board Page
	@FindBy(xpath="//*[@id='dashName']")
    private static WebElement labelDashboard ;
	
	@FindBy(xpath="//*[@id='Select_dash']")
	private static WebElement selectDashboard ;
		
	@FindBy(xpath="//*[@id='Dashboard_AddDash']")
    private static WebElement newAddDashBoard;
	 
    @FindBy(xpath="//*[@id='Dashboard_Dash_Config']")
    private static WebElement dashboardCustomizationSettings;
		
    //Logout and Login Screen
	
	@FindBy(xpath="//*[@id='txtUsername']")
	private static WebElement username;

	@FindBy(id="txtPassword")
	private static WebElement password;

	@FindBy(id="btnSignin")
	private static WebElement signIn;

    @FindBy(id="ddlCompany")
    private static WebElement companyDropDownList;

	@FindBy(xpath="//span[@class='hidden-xs']")
	private static WebElement userNameDisplay;

	@FindBy(xpath="//*[@id='companyLogo']")
	private static WebElement companyLogo;

	@FindBy(xpath="//*[@id='ulCompanyDetails_HomePage']/li[1]")
	private static WebElement companyName;

    @FindBy(xpath="//*[@id='userprofile']/li/span[2]")
	private static WebElement logoutOption;	 
	 
    @FindBy(xpath="//th[@id='PriceBookWebGrid_control_heading_1']//div[contains(text(),'Item')]")
    private static WebElement  gridItemColoumn;
    
    @FindBy(xpath="//th[@id='PriceBookWebGrid_control_heading_3']//div[contains(text(),'Warehouse')]")
    private static WebElement  gridWareHouseColoumn;
    
    @FindBy(xpath="//th[@id='PriceBookWebGrid_control_heading_11']//div[contains(text(),'Rate')]")
    private static WebElement  gridRateColoumn;
    
    @FindBy(xpath="//th[@id='PriceBookWebGrid_control_heading_12']//div[contains(text(),'Po')]")
    private static WebElement  gridPoColoumn;
    
    @FindBy(xpath="//div[contains(text(),'Pt')]")
    private static WebElement  gridPtColoumn;
    
    @FindBy(xpath="//input[@id='donotshow']")
    private static WebElement doNotShowCheckbox;
    
    @FindBy(xpath="//span[@class='pull-right']")
    private static WebElement closeBtnInDemoPopupScreen;
    
    
    
    public static void checkPopUpWindow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
    {
    	//Thread.sleep(5000);
	        
    	try 
    	{
    		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(doNotShowCheckbox));
    		doNotShowCheckbox.click();
    		
    		//Thread.sleep(2000);
    		
    		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtnInDemoPopupScreen));
    		closeBtnInDemoPopupScreen.click(); 
    		
    		System.err.println("POP UP DISPLAYED AND CLOSED SUCCESSFULLY");
    	} 
    	catch (Exception e)
    	{
    		System.err.println("NO POP UP DISPLAYED");
    	}
    	//Thread.sleep(4000);
	}
			 
			
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

	
	@FindBy(xpath="//button[contains(text(),'Ok')]")
	private static WebElement loginRefreshOkBtn;
	
	
	public static void checkRefershPopOnlogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		try 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginRefreshOkBtn));
			loginRefreshOkBtn.click();
		} 
		catch (Exception e)
		{
			System.err.println("NO ALERT POP UP DISPLAYED");	
		}	
		//Thread.sleep(4000);
	} 
	
		
	private static String xlfile;
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	
	private static int cSize;
	
	public boolean checkBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		LoginPage lp=new LoginPage(getDriver()); 
			
		String unamelt="su";
		
		String pawslt="su";
		
		lp.enterUserName(unamelt);
		
		lp.enterPassword(pawslt);
		
		String compname="Automation Company";
	 	
		 Select oSelect = new Select(companyDropDownList);
		 
		 List <WebElement> elementCount = oSelect.getOptions();
		
		 int cqSize = elementCount.size();
		 
		 int zqSize=cSize+1;
		 
		 System.out.println("CompanyDropdownList Count :"+cqSize);
		 
		 System.out.println("Company dropdown is :"+ zqSize);
	 
	 
	 //Select dropdown= new Select(lp.companyDropDownList);
	  int i;
	  
	  //List<WebElement> list = dropdown.getOptions();

		//List<String> text = new ArrayList<>();
		for(i=0; i<elementCount.size(); i++) 
		{
		
		  elementCount.get(i).getText();
  	  String optionName = elementCount.get(i).getText();
  	  if(optionName.toUpperCase().startsWith(compname.toUpperCase()))
  	  {
  		  System.out.println("q"+elementCount.get(i).getText());
  		  elementCount.get(i).click();
  		  
  	  }
    
    }
		
		lp.clickOnSignInBtn();
		
		//checkRefershPopOnlogin();
		
		//checkPopUpWindow();
		
		Thread.sleep(5000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	  	homeMenu.click();
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterMenu));
	  	homeMasterMenu.click();
	  
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	  	homeMasterItemMenu.click();
	  	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
	  	homeMasterItem_buyerPriceBookMenu.click();
	  	  	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		
		boolean actbuyerPriceBook_PriceBook 				=buyerPriceBook_PriceBook.isDisplayed();
		boolean actbuyerPriceBook_AbbreviationBtn 			=buyerPriceBook_AbbreviationBtn.isDisplayed();
		boolean actbuyerPriceBookLabel 						=buyerPriceBookLabel.isDisplayed();
		boolean actbuyerPriceBookClearBtn 					=buyerPriceBookClearBtn.isDisplayed();
		boolean actbuyerPriceBookPasteBtn 					=buyerPriceBookPasteBtn.isDisplayed();
		boolean actbuyerPriceBookCopyBtn 					=buyerPriceBookCopyBtn.isDisplayed();
		boolean actbuyerPriceBookSelectAllBtn 				=buyerPriceBookSelectAllBtn.isDisplayed();
		boolean actbuyerPriceBookCloseBtn 					=buyerPriceBookCloseBtn.isDisplayed();
		boolean actbuyerPriceBookSaveBtn 					=buyerPriceBookSaveBtn.isDisplayed();
		boolean actbuyerPriceBookDeleteBtn 					=buyerPriceBookDeleteBtn.isDisplayed();
		boolean actbuyerPriceBookStaringDate 				=buyerPriceBookStaringDate.isDisplayed();
		boolean actbuyerPriceBookENdingDate 				=buyerPriceBookENdingDate.isDisplayed();
		boolean actbuyerPriceBookInActiveCheckBox 			=buyerPriceBookInActiveCheckBox.isDisplayed();
		boolean actbuyerPriceBookFilter_Item 				=buyerPriceBookFilter_Item.isDisplayed();
		boolean actbuyerPriceBookFilter_Currency 			=buyerPriceBookFilter_Currency.isDisplayed();
		boolean actbuyerPriceBookFilter_Customer 			=buyerPriceBookFilter_Customer.isDisplayed();
		boolean actbuyerPriceBookFilter_Department 			=buyerPriceBookFilter_Department.isDisplayed();
		boolean actbuyerPriceBookFilter_ClearBtn 			=buyerPriceBookFilter_ClearBtn.isDisplayed();
		boolean actbuyerPriceBookFilter_FilterBtn 			=buyerPriceBookFilter_FilterBtn.isDisplayed();
		boolean actbuyerPriceBookFilter_FilterAndLoadBtn 	=buyerPriceBookFilter_FilterAndLoadBtn.isDisplayed();
		boolean actbuyerPriceBook_AdvanceExcelImportBtn 	=buyerPriceBook_AdvanceExcelImportBtn.isDisplayed();
		boolean actbuyerPriceBook_PrintRateHistoryBtn 		=buyerPriceBook_PrintRateHistoryBtn.isDisplayed();
		boolean actbuyerPriceBook_ExportToXmlBtn 			=buyerPriceBook_ExportToXmlBtn.isDisplayed();
		boolean actbuyerPriceBook_ImportFromXmlBtn 			=buyerPriceBook_ImportFromXmlBtn.isDisplayed();
		boolean actbuyerPriceBook_ExportToExcel 			=buyerPriceBook_ExportToExcel.isDisplayed();
		boolean actbuyerPriceBook_ImportFromExcel 			=buyerPriceBook_ImportFromExcel.isDisplayed();
		boolean actbuyerPriceBook_ViewPrice 				=buyerPriceBook_ViewPrice.isDisplayed();
		boolean actbuyerPriceBook_UpdatePriceList 			=buyerPriceBook_UpdatePriceList.isDisplayed();
		boolean actgridItemColoumn 							=gridItemColoumn.isDisplayed();
		boolean actgridRateColoumn 							=gridRateColoumn.isDisplayed();
		boolean actgridPoColoumn 							=gridPoColoumn.isDisplayed();
		boolean actgridPtColoumn 							=gridPtColoumn.isDisplayed();
		
		boolean expbuyerPriceBook_PriceBook 				=true;
		boolean expbuyerPriceBook_AbbreviationBtn 			=true;
		boolean expbuyerPriceBookLabel 						=true;
		boolean expbuyerPriceBookClearBtn 					=true;
		boolean expbuyerPriceBookPasteBtn 					=true;
		boolean expbuyerPriceBookCopyBtn 					=true;
		boolean expbuyerPriceBookSelectAllBtn 				=true;
		boolean expbuyerPriceBookCloseBtn 					=true;
		boolean expbuyerPriceBookSaveBtn 					=true;
		boolean expbuyerPriceBookDeleteBtn 					=true;
		boolean expbuyerPriceBookStaringDate 				=true;
		boolean expbuyerPriceBookENdingDate 				=true;
		boolean expbuyerPriceBookInActiveCheckBox 			=true;
		boolean expbuyerPriceBookFilter_Item 				=true;
		boolean expbuyerPriceBookFilter_Currency 			=true;
		boolean expbuyerPriceBookFilter_Customer 			=true;
		boolean expbuyerPriceBookFilter_Department 			=true;
		boolean expbuyerPriceBookFilter_ClearBtn 			=true;
		boolean expbuyerPriceBookFilter_FilterBtn 			=true;
		boolean expbuyerPriceBookFilter_FilterAndLoadBtn 	=true;
		boolean expbuyerPriceBook_AdvanceExcelImportBtn 	=true;
		boolean expbuyerPriceBook_PrintRateHistoryBtn 		=true;
		boolean expbuyerPriceBook_ExportToXmlBtn 			=true;
		boolean expbuyerPriceBook_ImportFromXmlBtn 			=true;
		boolean expbuyerPriceBook_ExportToExcel 			=true;
		boolean expbuyerPriceBook_ImportFromExcel 			=true;
		boolean expbuyerPriceBook_ViewPrice 				=true;
		boolean expbuyerPriceBook_UpdatePriceList 			=true;
		boolean expgridItemColoumn 							=true;
		boolean expgridRateColoumn 							=true;
		boolean expgridPoColoumn 							=true;
		boolean expgridPtColoumn 							=true;
		
		System.out.println("********************************checkBuyerPriceBook*************************************");
		
		System.out.println("buyerPriceBook_PriceBook              : "+actbuyerPriceBook_PriceBook              +" Value Expected : "+expbuyerPriceBook_PriceBook);
		System.out.println("buyerPriceBook_AbbreviationBtn        : "+actbuyerPriceBook_AbbreviationBtn        +" Value Expected : "+expbuyerPriceBook_AbbreviationBtn);
		System.out.println("buyerPriceBookLabel                   : "+actbuyerPriceBookLabel                   +" Value Expected : "+expbuyerPriceBookLabel);
		System.out.println("buyerPriceBookClearBtn                : "+actbuyerPriceBookClearBtn                +" Value Expected : "+expbuyerPriceBookClearBtn);
		System.out.println("buyerPriceBookPasteBtn                : "+actbuyerPriceBookPasteBtn                +" Value Expected : "+expbuyerPriceBookPasteBtn);
		System.out.println("buyerPriceBookCopyBtn                 : "+actbuyerPriceBookCopyBtn                 +" Value Expected : "+expbuyerPriceBookCopyBtn);
		System.out.println("buyerPriceBookSelectAllBtn            : "+actbuyerPriceBookSelectAllBtn            +" Value Expected : "+expbuyerPriceBookSelectAllBtn);
		System.out.println("buyerPriceBookCloseBtn                : "+actbuyerPriceBookCloseBtn                +" Value Expected : "+expbuyerPriceBookCloseBtn);
		System.out.println("buyerPriceBookSaveBtn                 : "+actbuyerPriceBookSaveBtn                 +" Value Expected : "+expbuyerPriceBookSaveBtn);
		System.out.println("buyerPriceBookDeleteBtn               : "+actbuyerPriceBookDeleteBtn               +" Value Expected : "+expbuyerPriceBookDeleteBtn);
		System.out.println("buyerPriceBookStaringDate             : "+actbuyerPriceBookStaringDate             +" Value Expected : "+expbuyerPriceBookStaringDate);
		System.out.println("buyerPriceBookENdingDate              : "+actbuyerPriceBookENdingDate              +" Value Expected : "+expbuyerPriceBookENdingDate);
		System.out.println("buyerPriceBookInActiveCheckBox        : "+actbuyerPriceBookInActiveCheckBox        +" Value Expected : "+expbuyerPriceBookInActiveCheckBox);
		System.out.println("buyerPriceBookFilter_Item             : "+actbuyerPriceBookFilter_Item             +" Value Expected : "+expbuyerPriceBookFilter_Item);
		System.out.println("buyerPriceBookFilter_Currency         : "+actbuyerPriceBookFilter_Currency         +" Value Expected : "+expbuyerPriceBookFilter_Currency);
		System.out.println("buyerPriceBookFilter_Customer         : "+actbuyerPriceBookFilter_Customer         +" Value Expected : "+expbuyerPriceBookFilter_Customer);
		System.out.println("buyerPriceBookFilter_Department       : "+actbuyerPriceBookFilter_Department       +" Value Expected : "+expbuyerPriceBookFilter_Department);
		System.out.println("buyerPriceBookFilter_ClearBtn         : "+actbuyerPriceBookFilter_ClearBtn         +" Value Expected : "+expbuyerPriceBookFilter_ClearBtn);
		System.out.println("buyerPriceBookFilter_FilterBtn        : "+actbuyerPriceBookFilter_FilterBtn        +" Value Expected : "+expbuyerPriceBookFilter_FilterBtn);
		System.out.println("buyerPriceBookFilter_FilterAndLoadBtn : "+actbuyerPriceBookFilter_FilterAndLoadBtn +" Value Expected : "+expbuyerPriceBookFilter_FilterAndLoadBtn);
		System.out.println("buyerPriceBook_AdvanceExcelImportBtn  : "+actbuyerPriceBook_AdvanceExcelImportBtn  +" Value Expected : "+expbuyerPriceBook_AdvanceExcelImportBtn);
		System.out.println("buyerPriceBook_PrintRateHistoryBtn    : "+actbuyerPriceBook_PrintRateHistoryBtn    +" Value Expected : "+expbuyerPriceBook_PrintRateHistoryBtn);
		System.out.println("buyerPriceBook_ExportToXmlBtn         : "+actbuyerPriceBook_ExportToXmlBtn         +" Value Expected : "+expbuyerPriceBook_ExportToXmlBtn);
		System.out.println("buyerPriceBook_ImportFromXmlBtn       : "+actbuyerPriceBook_ImportFromXmlBtn       +" Value Expected : "+expbuyerPriceBook_ImportFromXmlBtn);
		System.out.println("buyerPriceBook_ExportToExcel          : "+actbuyerPriceBook_ExportToExcel          +" Value Expected : "+expbuyerPriceBook_ExportToExcel);
		System.out.println("buyerPriceBook_ImportFromExcel        : "+actbuyerPriceBook_ImportFromExcel        +" Value Expected : "+expbuyerPriceBook_ImportFromExcel);
		System.out.println("buyerPriceBook_ViewPrice              : "+actbuyerPriceBook_ViewPrice              +" Value Expected : "+expbuyerPriceBook_ViewPrice);
		System.out.println("buyerPriceBook_UpdatePriceList        : "+actbuyerPriceBook_UpdatePriceList        +" Value Expected : "+expbuyerPriceBook_UpdatePriceList);
		System.out.println("gridItemColoumn                       : "+actgridItemColoumn                       +" Value Expected : "+expgridItemColoumn);
		System.out.println("gridRateColoumn                       : "+actgridRateColoumn                       +" Value Expected : "+expgridRateColoumn);
		System.out.println("gridPoColoumn                         : "+actgridPoColoumn                         +" Value Expected : "+expgridPoColoumn);
		System.out.println("gridPtColoumn                         : "+actgridPtColoumn                         +" Value Expected : "+expgridPtColoumn);
		
	  	if (actbuyerPriceBook_PriceBook==expbuyerPriceBook_PriceBook && actbuyerPriceBook_AbbreviationBtn==expbuyerPriceBook_AbbreviationBtn
	  			&& actbuyerPriceBookLabel==expbuyerPriceBookLabel && actbuyerPriceBookClearBtn==expbuyerPriceBookClearBtn
	  			&& actbuyerPriceBookPasteBtn==expbuyerPriceBookPasteBtn && actbuyerPriceBookCopyBtn==expbuyerPriceBookCopyBtn
	  			&& actbuyerPriceBookSelectAllBtn==expbuyerPriceBookSelectAllBtn && actbuyerPriceBookCloseBtn==expbuyerPriceBookCloseBtn
	  			&& actbuyerPriceBookSaveBtn==expbuyerPriceBookSaveBtn && actbuyerPriceBookDeleteBtn==expbuyerPriceBookDeleteBtn
	  			&& actbuyerPriceBookStaringDate==expbuyerPriceBookStaringDate && actbuyerPriceBookENdingDate==expbuyerPriceBookENdingDate
	  			&& actbuyerPriceBookInActiveCheckBox==expbuyerPriceBookInActiveCheckBox && actbuyerPriceBookFilter_Item==expbuyerPriceBookFilter_Item
	  			&& actbuyerPriceBookFilter_Currency==expbuyerPriceBookFilter_Currency && actbuyerPriceBookFilter_Customer==expbuyerPriceBookFilter_Customer
	  			&& actbuyerPriceBookFilter_Department==expbuyerPriceBookFilter_Department && actbuyerPriceBookFilter_ClearBtn==expbuyerPriceBookFilter_ClearBtn
	  			&& actbuyerPriceBookFilter_FilterBtn==expbuyerPriceBookFilter_FilterBtn && actbuyerPriceBookFilter_FilterAndLoadBtn==expbuyerPriceBookFilter_FilterAndLoadBtn
	  			&& actbuyerPriceBook_AdvanceExcelImportBtn==expbuyerPriceBook_AdvanceExcelImportBtn && actbuyerPriceBook_PrintRateHistoryBtn==expbuyerPriceBook_PrintRateHistoryBtn
	  			&& actbuyerPriceBook_ExportToXmlBtn==expbuyerPriceBook_ExportToXmlBtn && actbuyerPriceBook_ImportFromXmlBtn==expbuyerPriceBook_ImportFromXmlBtn
	  			&& actbuyerPriceBook_ExportToExcel==expbuyerPriceBook_ExportToExcel && actbuyerPriceBook_ImportFromExcel==expbuyerPriceBook_ImportFromExcel
	  			&& actbuyerPriceBook_ViewPrice==expbuyerPriceBook_ViewPrice && actbuyerPriceBook_UpdatePriceList==expbuyerPriceBook_UpdatePriceList
	  			&& actgridItemColoumn==expgridItemColoumn && actgridRateColoumn==expgridRateColoumn && actgridPoColoumn==expgridPoColoumn
	  			&& actgridPtColoumn==expgridPtColoumn)
	  	{	
	  		System.out.println("Pass: Dispalyed Buyer Price Book Screen ");
            excelReader.setCellData(xlfile, "Sheet1", 797, 9, resPass);
	  		return true;
		} 
	  	else 
	  	{
	  		System.out.println("Fail: Dispalyed Buyer Price Book Screen ");
	  		excelReader.setCellData(xlfile, "Sheet1", 797, 9, resFail);
	  		return false;
		}
	}
	

	@FindBy(xpath="//a[@id='16']//div[@class='form-group theme_button_color']")
	private static WebElement  settingsmenuBtn;
	
	@FindBy(xpath="//span[contains(text(),'Configure Transactions')]")
	private static WebElement  configureTransactionBtn;
	
	@FindBy(xpath="//select[@id='SellingRate']")
	private static WebElement  sellingsRateDropdown;
	
	@FindBy(xpath="//select[@id='BuyingRates']")
	private static WebElement  buyingRatesdropdown;
	
	@FindBy(xpath="//span[@id='updateButton']")
	private static WebElement  updateBtn;
	
	@FindBy(xpath="//i[@class='icon-close icon-font6']")
	private static WebElement  closeBtn;

	
	public boolean checkCustomizationSettingsOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	  	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
		buyerPriceBook_CustamizeSettingBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeDateRangeCheckBox));
		
		boolean actbuyerPriceBook_CustamizeDateRangeCheckBox   	=buyerPriceBook_CustamizeDateRangeCheckBox.isDisplayed();
		boolean actbuyerPriceBook_CustomizeVendorCheckBox   	=buyerPriceBook_CustomizeVendorCheckBox.isDisplayed();
		boolean actbuyerPriceBook_WarehouseCheckBox   			=buyerPriceBook_WarehouseCheckBox.isDisplayed();
		boolean actbuyerPriceBook_QualityRangeCheckBox   		=buyerPriceBook_QualityRangeCheckBox.isDisplayed();
		boolean actbuyerPriceBook_CurrencyCheckBox   			=buyerPriceBook_CurrencyCheckBox.isDisplayed();
		boolean actbuyerPriceBook_UnitCheckBox   				=buyerPriceBook_UnitCheckBox.isDisplayed();
		
		boolean expbuyerPriceBook_CustamizeDateRangeCheckBox   	=true;
		boolean expbuyerPriceBook_CustomizeVendorCheckBox   	=true;
		boolean expbuyerPriceBook_WarehouseCheckBox   			=true;
		boolean expbuyerPriceBook_QualityRangeCheckBox   		=true;
		boolean expbuyerPriceBook_CurrencyCheckBox   			=true;
		boolean expbuyerPriceBook_UnitCheckBox   				=true;
		
		System.out.println("**************************checkCustomizationSettingsOption**************************");
		
		System.out.println("buyerPriceBook_CustamizeDateRangeCheckBox  : "+actbuyerPriceBook_CustamizeDateRangeCheckBox +" Value Expected : "+expbuyerPriceBook_CustamizeDateRangeCheckBox);
		System.out.println("buyerPriceBook_CustomizeVendorCheckBox     : "+actbuyerPriceBook_CustomizeVendorCheckBox    +" Value Expected : "+expbuyerPriceBook_CustomizeVendorCheckBox);
		System.out.println("buyerPriceBook_WarehouseCheckBox           : "+actbuyerPriceBook_WarehouseCheckBox          +" Value Expected : "+expbuyerPriceBook_WarehouseCheckBox);
		System.out.println("buyerPriceBook_QualityRangeCheckBox        : "+actbuyerPriceBook_QualityRangeCheckBox       +" Value Expected : "+expbuyerPriceBook_QualityRangeCheckBox);
		System.out.println("buyerPriceBook_CurrencyCheckBox            : "+actbuyerPriceBook_CurrencyCheckBox           +" Value Expected : "+expbuyerPriceBook_CurrencyCheckBox);
		System.out.println("buyerPriceBook_UnitCheckBox                : "+actbuyerPriceBook_UnitCheckBox               +" Value Expected : "+expbuyerPriceBook_UnitCheckBox);
		
		if (actbuyerPriceBook_CustamizeDateRangeCheckBox==expbuyerPriceBook_CustamizeDateRangeCheckBox
				&& actbuyerPriceBook_CustomizeVendorCheckBox==expbuyerPriceBook_CustomizeVendorCheckBox
				&& actbuyerPriceBook_WarehouseCheckBox==expbuyerPriceBook_WarehouseCheckBox 
				&& actbuyerPriceBook_QualityRangeCheckBox==expbuyerPriceBook_QualityRangeCheckBox
				&& actbuyerPriceBook_CurrencyCheckBox==expbuyerPriceBook_CurrencyCheckBox
				&& actbuyerPriceBook_UnitCheckBox==expbuyerPriceBook_UnitCheckBox) 
		{
			System.out.println("Pass: Customize Screen Displayed");
			excelReader.setCellData(xlfile, "Sheet1", 798, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: Customize Screen Displayed");
			excelReader.setCellData(xlfile, "Sheet1", 798, 9, resFail);
			return false;
		}
	}
	
	
	@FindBy(xpath="//div[@class='theme_color font-6']")
	public static WebElement errorMessage;
	 
	@FindBy(xpath="//span[@class='icon-reject2 theme_color']")
	public static WebElement errorMessageCloseBtn;
	
	@FindBy(xpath="//i[@id='ToggleFilter']")
	private static WebElement beforeClickOnToggle; 	
			
	@FindBy(xpath="//i[@id='TogglePBTable']")
	private static WebElement afterClickOnToggle; 
	
	
/*	@FindBy(xpath="//th[@id='PriceBookWebGrid_control_heading_1']//div[contains(text(),'Item')]")
	private static WebElement item; 	
			
	@FindBy(xpath="//th[@id='PriceBookWebGrid_control_heading_3']//div[contains(text(),'Warehouse')]")
	private static WebElement warehouse; 	*/
			
	// Check Customize Settings Options
	
	@FindBy(xpath="//div[contains(text(),'Starting date')]")
	private static WebElement startingdate; 	
			
	@FindBy(xpath="//div[contains(text(),'Ending date')]")
	private static WebElement endingdate; 	
			
	@FindBy(xpath="//th[@id='PriceBookWebGrid_control_heading_6']//div[contains(text(),'Vendor')]")
	private static WebElement vendor; 	
	
	@FindBy(xpath="//th[@id='PriceBookWebGrid_control_heading_6']//div[contains(text(),'Customer')]")
	private static WebElement customer; 
			
	@FindBy(xpath="//th[@id='PriceBookWebGrid_control_heading_7']//div[contains(text(),'Currency')]")
	private static WebElement currency; 	
			
	@FindBy(xpath="//div[contains(text(),'Min Qty')]")
	private static WebElement minQty; 	
			
	@FindBy(xpath="//div[contains(text(),'Max Qty')]")
	private static WebElement maxQty; 	
			
	@FindBy(xpath="//div[contains(text(),'Unit')]")
	private static WebElement unit; 	
			
	@FindBy(xpath="//th[@id='PriceBookWebGrid_control_heading_11']//div[contains(text(),'Rate')]")
	private static WebElement rate; 	
			
	@FindBy(xpath="//th[@id='PriceBookWebGrid_control_heading_12']//div[contains(text(),'Po')]")
	private static WebElement po; 	
			
	@FindBy(xpath="//div[contains(text(),'So')]")
	private static WebElement so; 	


	@FindBy(xpath="//th[@id='PriceBookWebGrid_control_heading_13']//div[contains(text(),'St')]")
	private static WebElement st; 	
	
	@FindBy(xpath="//div[contains(text(),'Pt')]")
	private static WebElement pt; 	
			
	@FindBy(xpath="//div[contains(text(),'Val 3')]")
	private static WebElement Val3; 	
			
	@FindBy(xpath="//div[contains(text(),'Val 4')]")
	private static WebElement Val4; 	
			
	@FindBy(xpath="//div[contains(text(),'Val 5')]")
	private static WebElement Val5; 	
			
	@FindBy(xpath="//div[contains(text(),'Val 6')]")
	private static WebElement Val6; 	
			
	@FindBy(xpath="//div[contains(text(),'Val 7')]")
	private static WebElement Val7; 	
			
	@FindBy(xpath="//div[contains(text(),'Val 8')]")
	private static WebElement Val8; 	
			
	@FindBy(xpath="//div[contains(text(),'Val 9')]")
	private static WebElement Val9; 	
			
	@FindBy(xpath="//div[contains(text(),'Val 10')]")
	private static WebElement Val10; 	
			
	@FindBy(xpath="//div[contains(text(),'Val 11')]")
	private static WebElement Val1; 	
			
	@FindBy(xpath="//div[contains(text(),'Val 12')]")
	private static WebElement Val12; 	
			
	@FindBy(xpath="//div[contains(text(),'Val 13')]")
	private static WebElement Val13; 	
			
	// Grid Options display

	//First Row

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-1']")
	private static WebElement itemFirstCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-3']")
	private static WebElement wareHouseFirstCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-4']")
	private static WebElement satrtingDateFirstCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-5']")
	private static WebElement endingDateFirstCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-6']")
	private static WebElement vendorFirstCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-7']")
	private static WebElement currencyFirstCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-8']")
	private static WebElement minQtyFirstCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-9']")
	private static WebElement maxQtyFirstCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-10']")
	private static WebElement unitFirstCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-11']")
	private static WebElement rateFirstCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-12']")
	private static WebElement PoFirstCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-13']")
	private static WebElement PtFirstCellSelect; 	
			
	
	//Second Row
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_2-1']")
	private static WebElement itemSecondCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_2-3']")
	private static WebElement wareHouseSecondCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_2-4']")
	private static WebElement satrtingDateSecondCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_2-5']")
	private static WebElement endingDateSecondCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_2-6']")
	private static WebElement vendorSecondCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_2-7']")
	private static WebElement currencySecondCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_2-8']")
	private static WebElement minQtySecondCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_2-9']")
	private static WebElement maxQtySecondCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_2-10']")
	private static WebElement unitSecondCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_2-11']")
	private static WebElement rateSecondCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_2-12']")
	private static WebElement PoSecondCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_2-13']")
	private static WebElement PtSecondCellSelect; 	
			

	//Third Row
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-1']")
	private static WebElement itemThirdCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-3']")
	private static WebElement wareHouseThirdCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-4']")
	private static WebElement satrtingDateThirdCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-5']")
	private static WebElement endingDateThirdCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-6']")
	private static WebElement vendorThirdCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-7']")
	private static WebElement currencyThirdCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-8']")
	private static WebElement minQtyThirdCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-9']")
	private static WebElement maxQtyThirdCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-10']")
	private static WebElement unitThirdCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-11']")
	private static WebElement rateThirdCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-12']")
	private static WebElement PoThirdCellSelect; 	
			
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-13']")
	private static WebElement PtThirdCellSelect; 	
			
	
	//To get Value From Cell
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_2']")
	private static WebElement itemValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_4']")
	private static WebElement wareHouseValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_StartingDate']")
	private static WebElement startingDateValue; 	
	
	@FindBy(xpath="//*[@id='PriceBookWebGrid_option_EndingDate_checkbox']")
	private static WebElement endingDateChkBox; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_EndingDate']")
	private static WebElement endingDateValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_1']")
	private static WebElement vendorValue; 	
			
	@FindBy(xpath="//input[@id='Currency_option']")
	private static WebElement currencyOptionValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_minqty']")
	private static WebElement minQtyValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_maxqty']")
	private static WebElement maxQtyValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_11']")
	private static WebElement unitValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_val0']")
	private static WebElement rateValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_val1']")
	private static WebElement PoValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_val2']")
	private static WebElement PtValue; 	
	
	

	
	public boolean checkCustomizeColumnsInGrid() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	  	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeDateRangeCheckBox));
		buyerPriceBook_CustamizeDateRangeCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustomizeVendorCheckBox));
		buyerPriceBook_CustomizeVendorCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_WarehouseCheckBox));
		buyerPriceBook_WarehouseCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_QualityRangeCheckBox));
		buyerPriceBook_QualityRangeCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CurrencyCheckBox));
		buyerPriceBook_CurrencyCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_UnitCheckBox));
		buyerPriceBook_UnitCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
		buyerPriceBook_CustamizeSettingBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(beforeClickOnToggle));
		beforeClickOnToggle.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemColoumn));
		
		boolean actgridItemColoumn        =gridItemColoumn.isDisplayed();
		boolean actgridWareHouseColoumn   =gridWareHouseColoumn.isDisplayed();
		boolean actstartingdate   		  =startingdate.isDisplayed();
		boolean actendingdate             =endingdate.isDisplayed();
		boolean actvendor                 =vendor.isDisplayed();
		boolean actcurrency               =currency.isDisplayed();
		
		boolean expgridItemColoumn        =true;
		boolean expgridWareHouseColoumn   =true;
		boolean expstartingdate   		  =true;
		boolean expendingdate             =true;
		boolean expvendor                 =true;
		boolean expcurrency               =true;
		
		System.out.println("*************************checkCustomizeColumnsInGrid*****************************");
		
		System.out.println("gridItemColoumn      : "+actgridItemColoumn      +" Value Expected : "+expgridItemColoumn);
		System.out.println("gridWareHouseColoumn : "+actgridWareHouseColoumn +" Value Expected : "+expgridWareHouseColoumn);
		System.out.println("startingdate         : "+actstartingdate         +" Value Expected : "+expstartingdate);
		System.out.println("endingdate           : "+actendingdate           +" Value Expected : "+expendingdate);
		System.out.println("vendor               : "+actvendor               +" Value Expected : "+expvendor);
		System.out.println("currency             : "+actcurrency             +" Value Expected : "+expcurrency);
		
		
		if (actgridItemColoumn==expgridItemColoumn && actgridWareHouseColoumn==expgridWareHouseColoumn && actstartingdate==expstartingdate 
				&& actendingdate==expendingdate && actvendor==expvendor && actcurrency==expcurrency) 
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyFirstCellSelect));
			maxQtyFirstCellSelect.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
			maxQtyValue.click();
			maxQtyValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQty));
			
			boolean actminQty =minQty.isDisplayed();
			boolean actmaxQty =maxQty.isDisplayed();
			boolean actunit   =unit.isDisplayed();
			boolean actrate   =rate.isDisplayed();
			boolean actpo     =po.isDisplayed();
			boolean actpt     =pt.isDisplayed();
			
			boolean expminQty =true;
			boolean expmaxQty =true;
			boolean expunit   =true;
			boolean exprate   =true;
			boolean exppo     =true;
			boolean exppt     =true;
			
			System.out.println("minQty  : "+actminQty+" Value Expected : "+expminQty);
			System.out.println("maxQty  : "+actmaxQty+" Value Expected : "+expmaxQty);
			System.out.println("unit    : "+actunit  +" Value Expected : "+expunit);
			System.out.println("rate    : "+actrate  +" Value Expected : "+exprate);
			System.out.println("po      : "+actpo    +" Value Expected : "+exppo);
			System.out.println("pt      : "+actpt    +" Value Expected : "+exppt);
			
			if (actminQty==expminQty && actmaxQty==expmaxQty && actunit==expunit && actrate==exprate && actpo==exppo && actpt==exppt) 
			{
				System.out.println("Pass: Customize Screen Displayed");
				excelReader.setCellData(xlfile, "Sheet1", 799, 9, resPass);
				return true;
			} 
			else 
			{
				System.out.println("Fail: Customize Screen Displayed");
				excelReader.setCellData(xlfile, "Sheet1", 799, 9, resFail);
				return false;
			}
		} 
		else 
		{
			System.out.println("Fail: Customize Screen Displayed");
			excelReader.setCellData(xlfile, "Sheet1", 799, 9, resFail);
			return false;
		}
	}
	
	
	
	public boolean checkCustomizeColumnsAfterInputofBuyerPriceName() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	  	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitFirstCellSelect));
		unitFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.click();
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemColoumn));
		
		boolean actgridItemColoumn        =gridItemColoumn.isDisplayed();
		boolean actgridWareHouseColoumn   =gridWareHouseColoumn.isDisplayed();
		boolean actstartingdate   		  =startingdate.isDisplayed();
		boolean actendingdate             =endingdate.isDisplayed();
		boolean actvendor                 =vendor.isDisplayed();
		boolean actcurrency               =currency.isDisplayed();
		
		boolean expgridItemColoumn        =true;
		boolean expgridWareHouseColoumn   =true;
		boolean expstartingdate   		  =true;
		boolean expendingdate             =true;
		boolean expvendor                 =true;
		boolean expcurrency               =true;
		
		System.out.println("*************************checkCustomizeColumnsAfterInputofBuyerPriceName*****************************");
		
		System.out.println("gridItemColoumn      : "+actgridItemColoumn      +" Value Expected : "+expgridItemColoumn);
		System.out.println("gridWareHouseColoumn : "+actgridWareHouseColoumn +" Value Expected : "+expgridWareHouseColoumn);
		System.out.println("startingdate         : "+actstartingdate         +" Value Expected : "+expstartingdate);
		System.out.println("endingdate           : "+actendingdate           +" Value Expected : "+expendingdate);
		System.out.println("vendor               : "+actvendor               +" Value Expected : "+expvendor);
		System.out.println("currency             : "+actcurrency             +" Value Expected : "+expcurrency);
		
		
		if (actgridItemColoumn==expgridItemColoumn && actgridWareHouseColoumn==expgridWareHouseColoumn && actstartingdate==expstartingdate 
				&& actendingdate==expendingdate && actvendor==expvendor && actcurrency==expcurrency) 
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyFirstCellSelect));
			maxQtyFirstCellSelect.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
			maxQtyValue.click();
			maxQtyValue.sendKeys(Keys.TAB);
			
			boolean actminQty =minQty.isDisplayed();
			boolean actmaxQty =maxQty.isDisplayed();
			boolean actunit   =unit.isDisplayed();
			boolean actrate   =rate.isDisplayed();
			boolean actpo     =po.isDisplayed();
			boolean actpt     =pt.isDisplayed();
			
			boolean expminQty =true;
			boolean expmaxQty =true;
			boolean expunit   =true;
			boolean exprate   =true;
			boolean exppo     =true;
			boolean exppt     =true;
			
			System.out.println("minQty  : "+actminQty+" Value Expected : "+expminQty);
			System.out.println("maxQty  : "+actmaxQty+" Value Expected : "+expmaxQty);
			System.out.println("unit    : "+actunit  +" Value Expected : "+expunit);
			System.out.println("rate    : "+actrate  +" Value Expected : "+exprate);
			System.out.println("po      : "+actpo    +" Value Expected : "+exppo);
			System.out.println("pt      : "+actpt    +" Value Expected : "+exppt);
			
			if (actminQty==expminQty && actmaxQty==expmaxQty && actunit==expunit && actrate==exprate && actpo==exppo && actpt==exppt)
			{
				System.out.println("Pass: Customize Screen Displayed");
				excelReader.setCellData(xlfile, "Sheet1", 799, 9, resPass);
				return true;
			} 
			else 
			{
				System.out.println("Fail: Customize Screen Displayed");
				excelReader.setCellData(xlfile, "Sheet1", 799, 9, resFail);
				return false;
			}
		} 
		else 
		{
			System.out.println("Fail: Customize Screen Displayed");
			excelReader.setCellData(xlfile, "Sheet1", 799, 9, resFail);
			return false;
		}
	}
	
	
	public boolean checkMandatoryFields() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	  			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.clear();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		String expMessage="Enter price book name";
		
		System.out.println("*****************************checkMandatoryFields***********************");
		
		String actMessage=checkValidationMessage(expMessage);
		
	   if (actMessage.equalsIgnoreCase(expMessage)) 
	   {
			System.out.println("Pass: displayed Mandatory Message as Enter Name of the Price Book");
			excelReader.setCellData(xlfile, "Sheet1", 800, 9, resPass);
			return true;
	   } 
	   else 
	   {	
		   System.out.println("Fail: displayed Mandatory Message as Enter Name of the Price Book");
		   excelReader.setCellData(xlfile, "Sheet1", 800, 9, resFail);
		   return false;
	   }
	}
	
	
	public boolean checkProductNameAsMandatory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys("Test Buying Price Hyd Wh");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		String expMessage="Select Item";
		
		System.out.println("***************************checkProductNameAsMandatory*************************");
		
		String actMessage=checkValidationMessage(expMessage);
	    
	   if (actMessage.equalsIgnoreCase(expMessage)) 
	   {
		   System.out.println("Pass: displayed message as Select Product");
		   excelReader.setCellData(xlfile, "Sheet1", 801, 9, resPass);
		   return true;
	   } 
	   else 
	   {
		   System.out.println("Fail:displayed message as Select Product");
		   excelReader.setCellData(xlfile, "Sheet1", 801, 9, resFail);
		   return false;
	   }
	}
	
	
	/*@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-1']")
	private static WebElement gridItemFirstRow;
	
	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_2']")
	private static WebElement gridItemFirstRowafterSelecting;

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-3']")
	private static WebElement gridWarehouseFirstRow;
	
	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_4']")
	private static WebElement gridWarehouseFirstRowAfterSelecting;
	
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-11']")
	private static WebElement gridRateFirstRow;
	
	@FindBy(xpath="//input[@id='PriceBookWebGrid_val0']")
	private static WebElement gridRateFirstRowAfterSelecting;
	
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_2-1']")
	private static WebElement gridItemSecondRow;
	
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_2-3']")
	private static WebElement gridWarehouseSecondRow;
		
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_2-11']")
	private static WebElement gridRateSecondRow;

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-1']")
	private static WebElement gridItemThirdRow;
	
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-3']")
	private static WebElement gridWarehouseThirdRow;
	
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-11']")
	private static WebElement gridRateThirdRow;

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_4-1']")
	private static WebElement gridItemFourthRow;
	
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_4-3']")
	private static WebElement gridWarehouseFourthRow;
	
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_4-11']")
	private static WebElement gridRateFourthRow;

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_5-1']")
	private static WebElement gridItemFifthRow;
	
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_5-3']")
	private static WebElement gridWarehouseFifthRow;
	
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_5-11']")
	private static WebElement gridRateFifthRow;*/
	
	@FindBy(xpath="//*[@id='ctrlOptionProPriceBookH_table_body']/tr/td[2]")
	private static List<WebElement> pricebookComboBoxList;
	
	@FindBy(xpath="//*[@id='PriceBookWebGrid_option_2_table_body']/tr/td[2]")
	private static List<WebElement> itemComboBoxList;
	
	@FindBy(xpath="//*[@id='PriceBookWebGrid_option_4_table_body']/tr/td[2]")
	private static List<WebElement> warehouseComboBoxList;
	
	@FindBy(xpath="//*[@id='PriceBookWebGrid_option_1_table_body']/tr/td[2]")
	private static List<WebElement> vendorComboBoxList;
	
	@FindBy(xpath="//*[@id='Currency_option_table_body']/tr/td[2]")
	private static List<WebElement> currencyComboBoxList;
	
	@FindBy(xpath="//*[@id='PriceBookWebGrid_option_11_table_body']/tr/td[2]")
	private static List<WebElement> unitComboBoxList;
	
	public boolean checkRateAsMandatory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.SPACE);
		
		int count = itemComboBoxList.size();
		   
		for(int i=0;i<count;i++)
		{
			String data = itemComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("ITEMS GROUP"))
			{
				itemComboBoxList.get(i).click();
				break;
			}
		}
		   
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		String expMessage="Enter Rate In Row Number 1";
		
		System.out.println("*******************************************checkRateAsMandatory******************************************");
		
		String actMessage=checkValidationMessage(expMessage);
	    
	    if (actMessage.equalsIgnoreCase(expMessage)) 
	    {
			System.out.println("Pass: displayed message as Select Product");
            excelReader.setCellData(xlfile, "Sheet1", 802, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail:displayed message as Select Product");
             excelReader.setCellData(xlfile, "Sheet1", 802, 9, resFail);
			return false;
		}
	}
	
	
	
	
	public boolean checkSaveBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseFirstCellSelect));
		wareHouseFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
		wareHouseValue.click();
		wareHouseValue.sendKeys(Keys.SPACE);
		
		int warehousecount = warehouseComboBoxList.size();
		   
		for(int i=0;i<warehousecount;i++)
		{
			String data = warehouseComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("HYDERABAD"))
			{
				warehouseComboBoxList.get(i).click();
				break;
			}
		}
		//*[@id='PriceBookWebGrid_option_EndingDate_checkbox']
		wareHouseValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingDateValue));
		startingDateValue.click();
		startingDateValue.sendKeys(Keys.TAB);
		
		if(endingDateChkBox.isSelected()==true)
		{
			endingDateChkBox.click();
		}
		
		//afterClickOnToggle.click();
		
		if(endingDateChkBox.isSelected()==false)
		{
			endingDateChkBox.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(endingDateValue));
		endingDateValue.click();
		endingDateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
		vendorValue.click();
		vendorValue.sendKeys(Keys.SPACE);
		
		int vendorcount = vendorComboBoxList.size();
		   
		for(int i=0;i<vendorcount;i++)
		{
			String data = vendorComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Vendor B"))
			{
				vendorComboBoxList.get(i).click();
				break;
			}
		}
		   
		vendorValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyOptionValue));
		currencyOptionValue.click();
		currencyOptionValue.sendKeys("I");
		
		int currencycount = currencyComboBoxList.size();
		   
		for(int i=0;i<currencycount;i++)
		{
			String data = currencyComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Indian Rupees"))
			{
				currencyComboBoxList.get(i).click();
				break;
			}
		}
		   
		currencyOptionValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateFirstCellSelect));
		rateFirstCellSelect.click();
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("100");
		rateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PoValue));
		PoValue.sendKeys("101");
		PoValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		String expMessage="Price Book created successfully";
		
		System.out.println("*********************************checkSaveBuyerPriceBook********************************");
		
		String actMessage=checkValidationMessage(expMessage);

	   if (actMessage.equalsIgnoreCase(expMessage)) 
	   {
		   System.out.println("Pass: Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");
		   excelReader.setCellData(xlfile, "Sheet1", 803, 9, resPass);
		   return true;
	   } 
	   else 
	   {
		   System.out.println("Fail:  Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");
		   excelReader.setCellData(xlfile, "Sheet1", 803, 9, resFail);
		   return false;
	   }	
	}
	

	public boolean checkEditBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(afterClickOnToggle));
		afterClickOnToggle.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.END);
		buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		
		int pricebookcount = pricebookComboBoxList.size();
		   
		for(int i=0;i<pricebookcount;i++)
		{
			String data = pricebookComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Test Buying Price Hyd Wh"))
			{
				pricebookComboBoxList.get(i).click();
				break;
			}
		}
		   
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.click();
		 
		String actbuyerPriceBook_PriceBook =buyerPriceBook_PriceBook.getAttribute("value");
		String expbuyerPriceBook_PriceBook ="Test Buying Price Hyd Wh";
	     
		System.out.println("************************************checkEditBuyerPriceBook**********************");
		
		System.out.println("buyerPriceBook_PriceBook  : "+actbuyerPriceBook_PriceBook+" Value Expeted : "+expbuyerPriceBook_PriceBook);
	    
		if(actbuyerPriceBook_PriceBook.equalsIgnoreCase(expbuyerPriceBook_PriceBook)) 
		{
			System.out.println("Pass: Edit Buyer PriceBook");
			excelReader.setCellData(xlfile, "Sheet1", 804, 9, resPass);
	        return true;
		} 
		else 
		{
			System.out.println("Fail: Edit Buyer PriceBook");
			excelReader.setCellData(xlfile, "Sheet1", 804, 9, resFail);
	        return false;
		}
	}
	
	
	public boolean checkSavedBuyerRatesInformation() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
	    buyerPriceBookFilter_FilterAndLoadBtn.click();
	 
	    //Thread.sleep(2000);
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(beforeClickOnToggle));
	    beforeClickOnToggle.click();
	    
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitFirstCellSelect));
	    unitFirstCellSelect.click();

	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		
		String actitemValue =itemValue.getAttribute("value");
	    String expitemValue ="ITEMS GROUP";    
	    
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseFirstCellSelect));
	    wareHouseFirstCellSelect.click();
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
		wareHouseValue.click();
	
		String actwareHouseValue =wareHouseValue.getAttribute("value");
	    String expwareHouseValue ="HYDERABAD";
		
	    System.out.println("******************************checkSavedBuyerRatesInformation****************************");
	    
	    System.out.println("itemValue        : "+actitemValue      +" Value Expected : "+expitemValue);
	    System.out.println("wareHouseValue   : "+actwareHouseValue +" Value Expected : "+expwareHouseValue);
	    
	    if(actitemValue.equalsIgnoreCase(expitemValue) && actwareHouseValue.equalsIgnoreCase(expwareHouseValue)) 
	    {
			System.out.println("Pass : Allow to select and HYD WH should be display in the Warehouse Combo box");
			excelReader.setCellData(xlfile, "Sheet1", 805, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: Allow to select and HYD WH should be display in the Warehouse Combo box");
			excelReader.setCellData(xlfile, "Sheet1", 805, 9, resFail);
			return false;
		}
	}
	
	
	public boolean checkUpdateBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.SPACE);
		
		int itemcount = itemComboBoxList.size();
		   
		for(int i=0;i<itemcount;i++)
		{
			String data = itemComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("BATCH BR ITEM"))
			{
				itemComboBoxList.get(i).click();
				break;
			}
		}
		   
		itemValue.sendKeys(Keys.TAB);
		
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
		wareHouseValue.click();
		wareHouseValue.sendKeys(Keys.END);
		wareHouseValue.sendKeys(Keys.SHIFT,Keys.HOME);
		wareHouseValue.sendKeys(Keys.SPACE);
		
		int warehousecount = warehouseComboBoxList.size();
		   
		for(int i=0;i<warehousecount;i++)
		{
			String data = warehouseComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("HYDERABAD"))
			{
				warehouseComboBoxList.get(i).click();
				break;
			}
		}
		   
		wareHouseValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingDateValue));
		startingDateValue.click();
		startingDateValue.sendKeys(Keys.TAB);
		

		if(endingDateChkBox.isSelected()==true)
		{
			endingDateChkBox.click();
		}
		
		if(endingDateChkBox.isSelected()==false)
		{
			endingDateChkBox.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(endingDateValue));
		endingDateValue.click();
		endingDateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
		vendorValue.click();
		vendorValue.sendKeys(Keys.END);
		vendorValue.sendKeys(Keys.SHIFT,Keys.HOME);
		vendorValue.sendKeys(Keys.SPACE);
		
		int vendorcount = vendorComboBoxList.size();
		   
		for(int i=0;i<vendorcount;i++)
		{
			String data = vendorComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Vendor B"))
			{
				vendorComboBoxList.get(i).click();
				break;
			}
		}
		   
		vendorValue.sendKeys(Keys.TAB);
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyOptionValue));
		currencyOptionValue.click();
		currencyOptionValue.sendKeys(Keys.END);
		currencyOptionValue.sendKeys(Keys.SHIFT,Keys.HOME);
		currencyOptionValue.sendKeys("I");
		
		int currencycount = currencyComboBoxList.size();
		   
		for(int i=0;i<currencycount;i++)
		{
			String data = currencyComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Indian Rupees"))
			{
				currencyComboBoxList.get(i).click();
				break;
			}
		}
		   
		currencyOptionValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
		minQtyValue.click();
		minQtyValue.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtySecondCellSelect));
		maxQtySecondCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
		maxQtyValue.click();
		maxQtyValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitValue));
		unitValue.click();
		unitValue.sendKeys(Keys.END);
		unitValue.sendKeys(Keys.SHIFT,Keys.HOME);
		unitValue.sendKeys(Keys.SPACE);
		
		int unitcount = unitComboBoxList.size();
		   
		for(int i=0;i<unitcount;i++)
		{
			String data = unitComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Dozs"))
			{
				unitComboBoxList.get(i).click();
				break;
			}
		}
		   
		unitValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("200");
		rateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PoValue));
		PoValue.sendKeys("202");
		PoValue.sendKeys(Keys.TAB);
		
		
		//Input in Third
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitSecondCellSelect));
		unitSecondCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
		itemThirdCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.END);
		itemValue.sendKeys(Keys.SHIFT,Keys.HOME);
		itemValue.sendKeys(Keys.SPACE);
		
		int itemcount1 = itemComboBoxList.size();
		   
		for(int i=0;i<itemcount1;i++)
		{
			String data = itemComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("BIN RAW MATERIAL ITEM"))
			{
				itemComboBoxList.get(i).click();
				break;
			}
		}
		   
		itemValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
		wareHouseValue.click();
		wareHouseValue.sendKeys(Keys.END);
		wareHouseValue.sendKeys(Keys.SHIFT,Keys.HOME);
		wareHouseValue.sendKeys(Keys.SPACE);
		
		int warehousecount1 = warehouseComboBoxList.size();
		   
		for(int i=0;i<warehousecount1;i++)
		{
			String data = warehouseComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("SECUNDERABAD"))
			{
				warehouseComboBoxList.get(i).click();
				break;
			}
		}
		   
		wareHouseValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingDateValue));
		startingDateValue.click();
		startingDateValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(endingDateValue));
		endingDateValue.click();
		endingDateValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
		vendorValue.click();
		vendorValue.sendKeys(Keys.END);
		vendorValue.sendKeys(Keys.SHIFT,Keys.HOME);
		vendorValue.sendKeys(Keys.SPACE);
		
		int vendorcount1 = vendorComboBoxList.size();
		   
		for(int i=0;i<vendorcount1;i++)
		{
			String data = vendorComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Vendor B"))
			{
				vendorComboBoxList.get(i).click();
				break;
			}
		}
		   
		vendorValue.sendKeys(Keys.TAB);
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyOptionValue));
		currencyOptionValue.click();
		currencyOptionValue.sendKeys(Keys.END);
		currencyOptionValue.sendKeys(Keys.SHIFT,Keys.HOME);
		currencyOptionValue.sendKeys("I");
		
		int currencycount1 = currencyComboBoxList.size();
		   
		for(int i=0;i<currencycount1;i++)
		{
			String data = currencyComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Indian Rupees"))
			{
				currencyComboBoxList.get(i).click();
				break;
			}
		}
		   
		currencyOptionValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
		minQtyValue.click();
		minQtyValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
		maxQtyValue.click();
		maxQtyValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitValue));
		unitValue.click();
		unitValue.sendKeys(Keys.END);
		unitValue.sendKeys(Keys.SHIFT,Keys.HOME);
		unitValue.sendKeys(Keys.SPACE);
		
		int unitcount1 = unitComboBoxList.size();
		   
		for(int i=0;i<unitcount1;i++)
		{
			String data = unitComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Dozs"))
			{
				unitComboBoxList.get(i).click();
				break;
			}
		}
		   
		unitValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("300");
		rateValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PoValue));
		PoValue.sendKeys("303");
		PoValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		String expMessage="Price Book updated successfully";
		
		System.out.println("****************************checkUpdateBuyerPriceBook********************************");
		
		String actMessage=checkValidationMessage(expMessage);
	    
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			System.out.println("Pass: Update Buyer PriceBook and Saved Successfully");
			excelReader.setCellData(xlfile, "Sheet1", 806, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail:   Update Buyer PriceBook and Saved Successfully");
			excelReader.setCellData(xlfile, "Sheet1", 806, 9, resFail);
			return false;
		}
	}
	
	
	@FindBy(xpath="//h4[@id='dashName']")
	private static WebElement dashBoardLabel;
	
	public boolean checkCloseInBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCloseBtn));
		buyerPriceBookCloseBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
		
		boolean actlabelDashboard                  =labelDashboard.isDisplayed();
		boolean actselectDashboard                 =selectDashboard.isDisplayed();
		boolean actnewAddDashBoard                 =newAddDashBoard.isDisplayed();
		boolean actdashboardCustomizationSettings  =dashboardCustomizationSettings.isDisplayed();
		
		boolean explabelDashboard                  =true;
		boolean expselectDashboard                 =true;
		boolean expnewAddDashBoard                 =true;
		boolean expdashboardCustomizationSettings  =true;
		
		System.out.println("******************************checkCloseInBuyerPriceBook************************");
		
		System.out.println("labelDashboard                  : "+actlabelDashboard                 +" Value Expected : "+explabelDashboard);
		System.out.println("selectDashboard                 : "+actselectDashboard                +" Value Expected : "+expselectDashboard);
		System.out.println("newAddDashBoard                 : "+actnewAddDashBoard                +" Value Expected : "+expnewAddDashBoard);
		System.out.println("dashboardCustomizationSettings  : "+actdashboardCustomizationSettings +" Value Expected : "+expdashboardCustomizationSettings);
		
		if(labelDashboard.isDisplayed() && selectDashboard.isDisplayed() && newAddDashBoard.isDisplayed() && dashboardCustomizationSettings.isDisplayed())
		{
			System.out.println("Pass: Close Button in Buyer Price Book Screen");
			excelReader.setCellData(xlfile, "Sheet1", 807, 9, resPass);
			return true;
		}
		else 
		{
			System.out.println("Fail: Close Button in Buyer Price Book Screen");
			excelReader.setCellData(xlfile, "Sheet1", 807, 9, resFail);
			return false;
		}
	}
	
	
	public boolean checkClearOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	  	homeMenu.click();
	  	 	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterMenu));
	  	homeMasterMenu.click();
	  
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	  	homeMasterItemMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
	  	homeMasterItem_buyerPriceBookMenu.click();

	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	  	buyerPriceBook_PriceBook.click();
	  	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	  	buyerPriceBook_AbbreviationBtn.click();
		 
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
	  	buyerPriceBookFilter_FilterAndLoadBtn.click();
		
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookClearBtn));
	  	buyerPriceBookClearBtn.click();
		 
	  	getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(itemThirdCellSelect));
		
	  	boolean actitemThirdCellSelect=itemThirdCellSelect.getText().isEmpty();
	  	boolean expitemThirdCellSelect=true;
		 
	  	System.out.println("**************************checkClearOption********************");
		 
	  	System.out.println("itemThirdCellSelect : "+actitemThirdCellSelect+" Value Expected : "+expitemThirdCellSelect);
	  	
	  	if (actitemThirdCellSelect==expitemThirdCellSelect) 
  		{
	  		System.out.println("Fail: Clear Option Working Successfully");
  			excelReader.setCellData(xlfile, "Sheet1", 808, 9, resPass);
  			return true;
  		} 
  		else 
  		{
  			System.out.println("Pass:  Clear Option Working Successfully");
  			excelReader.setCellData(xlfile, "Sheet1", 808, 9, resFail);
  			return false;
  		}
	  	
	}	
	
	   
	public boolean checkCopyandPaste() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.END);
		buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		
		int pricebookcount = pricebookComboBoxList.size();
		   
		for(int i=0;i<pricebookcount;i++)
		{
			String data = pricebookComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Test Buying Price Hyd Wh"))
			{
				pricebookComboBoxList.get(i).click();
				break;
			}
		}
		   
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
		buyerPriceBookFilter_FilterAndLoadBtn.click();
		
		//Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCopyBtn));
		buyerPriceBookCopyBtn.click();			 
			 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookClearBtn));
		buyerPriceBookClearBtn.click();
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookPasteBtn));
		buyerPriceBookPasteBtn.click();
			 		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
			    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
						
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(itemValue));
		
		String actitemValue =itemValue.getAttribute("value");
		String expitemValue ="ITEMS GROUP";
		
		System.out.println("********************************checkCopyandPaste******************************");
		
		System.out.println("itemValue  : "+actitemValue+" Value Expected : "+expitemValue);
		
		if (actitemValue.equalsIgnoreCase(expitemValue)) 
		{
			System.out.println("Pass: Copy And Paste Option ");
			excelReader.setCellData(xlfile, "Sheet1", 809, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: Copy And Paste Option");
			excelReader.setCellData(xlfile, "Sheet1", 809, 9, resFail);
			return false;
		}
	}
	   
	
	public boolean checkSaveOptionAfterPerformingPasteOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		String expMessage="Enter price book name";
		
		System.out.println("****************************checkSaveOptionAfterPerformingPasteOption**************************");
		
		String actMessage=checkValidationMessage(expMessage);
	    
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			System.out.println("Pass: Update Buyer PriceBook and Saved Successfully");
			excelReader.setCellData(xlfile, "Sheet1", 810, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail:   Update Buyer PriceBook and Saved Successfully");
			excelReader.setCellData(xlfile, "Sheet1", 810, 9, resFail);
			return false;
		}
	}
	
	
	
	public boolean checkSaveOptionAfterPerformingPasteOptionWithPricebookName() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.END);
		buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		
		int pricebookcount = pricebookComboBoxList.size();
		   
		for(int i=0;i<pricebookcount;i++)
		{
			String data = pricebookComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Test Buying Price Hyd Wh"))
			{
				pricebookComboBoxList.get(i).click();
				break;
			}
		}
		   
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.click();
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
		buyerPriceBookFilter_FilterAndLoadBtn.click();
	   
		//Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCopyBtn));
		buyerPriceBookCopyBtn.click();			 
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookClearBtn));
		buyerPriceBookClearBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys("BuyingCopy");
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookPasteBtn));
		buyerPriceBookPasteBtn.click();
		 
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		String expMessage="Price Book created successfully";
		
		System.out.println("**********************checkSaveOptionAfterPerformingPasteOptionWithPricebookName**********************");
		
		String actMessage=checkValidationMessage(expMessage);
	    
	   if (actMessage.equalsIgnoreCase(expMessage)) 
	   {
			System.out.println("Pass: Update Buyer PriceBook and Saved Successfully");
			excelReader.setCellData(xlfile, "Sheet1", 811, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail:   Update Buyer PriceBook and Saved Successfully");
			excelReader.setCellData(xlfile, "Sheet1", 811, 9, resFail);
			return false;
		}
	}
	
	
	
	public boolean checkPriceBookThroughCopyandPaste() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.END);
		buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		
		int pricebookcount = pricebookComboBoxList.size();
		   
		for(int i=0;i<pricebookcount;i++)
		{
			String data = pricebookComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("BuyingCopy"))
			{
				pricebookComboBoxList.get(i).click();
				break;
			}
		}
		   
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(buyerPriceBook_PriceBook));
		String actbuyerPriceBook_PriceBook =buyerPriceBook_PriceBook.getAttribute("value");
	    String expbuyerPriceBook_PriceBook ="BuyingCopy";
	    
	    System.out.println("*********************************checkPriceBookThroughCopyandPaste******************************");
	    
	    System.out.println("buyerPriceBook_PriceBook  : "+actbuyerPriceBook_PriceBook+" Value Expected : "+expbuyerPriceBook_PriceBook);
	    
	    if (actbuyerPriceBook_PriceBook.equalsIgnoreCase(expbuyerPriceBook_PriceBook)) 
	    {
	    	System.out.println("Pass: Price Book Through Copy and Paste");
			excelReader.setCellData(xlfile, "Sheet1", 812, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail:Price Book Through Copy and Paste");
			excelReader.setCellData(xlfile, "Sheet1", 812, 9, resFail);
			return false;
		}
	}
	
	
	 
	public boolean clickCopyAndPasteInformatioinOnLoadButton() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
		buyerPriceBookFilter_FilterAndLoadBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
				
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(itemValue));
		String actitemValue  =itemValue.getAttribute("value");
		String expitemValue  ="ITEMS GROUP";
		
		System.out.println("**************************clickCopyAndPasteInformatioinOnLoadButton********************************");
		    
		System.out.println("itemValue  : "+actitemValue+" Value Expected : "+expitemValue);
		    
		if (actitemValue.equalsIgnoreCase(expitemValue)) 
		{
			System.out.println("Pass: Copy And Paste Option ");
			excelReader.setCellData(xlfile, "Sheet1", 813, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: Copy And Paste Option");
			excelReader.setCellData(xlfile, "Sheet1", 813, 9, resFail);
			return false;
		}
	}
	
	
	public boolean checkDeleteOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		//Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCopyBtn));
		buyerPriceBookCopyBtn.click();			 
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookClearBtn));
		buyerPriceBookClearBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys("Test Hyd Wh");
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookPasteBtn));
		buyerPriceBookPasteBtn.click();
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.END);
		itemValue.sendKeys(Keys.SHIFT,Keys.HOME);
		itemValue.sendKeys(Keys.SPACE);
		
		int count = itemComboBoxList.size();
		   
		for(int i=0;i<count;i++)
		{
			String data = itemComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("BATCH FIFO ITEM"))
			{
				itemComboBoxList.get(i).click();
				break;
			}
		}
		   
		itemValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
		wareHouseValue.click();
		wareHouseValue.sendKeys(Keys.END);
		wareHouseValue.sendKeys(Keys.SHIFT,Keys.HOME);
		wareHouseValue.sendKeys(Keys.SPACE);
		
		int warehousecount = warehouseComboBoxList.size();
		   
		for(int i=0;i<warehousecount;i++)
		{
			String data = warehouseComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("HYDERABAD"))
			{
				warehouseComboBoxList.get(i).click();
				break;
			}
		}
		   
		wareHouseValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingDateValue));
		startingDateValue.click();
		startingDateValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(endingDateValue));
		endingDateValue.click();
		endingDateValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
		vendorValue.click();
		vendorValue.sendKeys(Keys.END);
		vendorValue.sendKeys(Keys.SHIFT,Keys.HOME);
		vendorValue.sendKeys(Keys.SPACE);
		
		int vendorcount1 = vendorComboBoxList.size();
		   
		for(int i=0;i<vendorcount1;i++)
		{
			String data = vendorComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Vendor B"))
			{
				vendorComboBoxList.get(i).click();
				break;
			}
		}
		   
		vendorValue.sendKeys(Keys.TAB);
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyOptionValue));
		currencyOptionValue.click();
		currencyOptionValue.sendKeys(Keys.END);
		currencyOptionValue.sendKeys(Keys.SHIFT,Keys.HOME);
		currencyOptionValue.sendKeys("I");
		
		int currencycount1 = currencyComboBoxList.size();
		   
		for(int i=0;i<currencycount1;i++)
		{
			String data = currencyComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Indian Rupees"))
			{
				currencyComboBoxList.get(i).click();
				break;
			}
		}
		   
		Thread.sleep(2000);
		currencyOptionValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
		minQtyValue.sendKeys("9");
		minQtyValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
		maxQtyValue.sendKeys("39");
		maxQtyValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitValue));
		unitValue.click();
		unitValue.sendKeys(Keys.END);
		unitValue.sendKeys(Keys.SHIFT,Keys.HOME);
		unitValue.sendKeys(Keys.SPACE);
		
		int unitcount1 = unitComboBoxList.size();
		   
		for(int i=0;i<unitcount1;i++)
		{
			String data = unitComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Dozs"))
			{
				unitComboBoxList.get(i).click();
				break;
			}
		}
		   
		unitValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("20");
		rateValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PoValue));
		PoValue.sendKeys("29");
		PoValue.sendKeys(Keys.TAB);
		 
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		String expMessage="Price Book created successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.END);
		buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		
		int pricebookcount = pricebookComboBoxList.size();
		   
		for(int i=0;i<pricebookcount;i++)
		{
			String data = pricebookComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("BuyingCopy"))
			{
				pricebookComboBoxList.get(i).click();
				break;
			}
		}
		   
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.click();
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
		buyerPriceBookFilter_FilterAndLoadBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookDeleteBtn));
		buyerPriceBookDeleteBtn.click();
		
		getWaitForAlert();
		 
		String actAlertMgs =getAlert().getText();
		String expAlertMgs ="Do you want to delete Total Price book?";
		
		System.out.println("***************************checkDeleteOption************************");
		 
		System.out.println("AlertMgs  : "+actAlertMgs+" Value Expected : "+expAlertMgs);
		
		if(actAlertMgs.equalsIgnoreCase(expAlertMgs)) 
		{
			System.out.println("Pass: Created price Book from Dropdown and Click on Delete button");
			excelReader.setCellData(xlfile, "Sheet1", 814, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: Created price Book from Dropdown and Click on Delete button");
            excelReader.setCellData(xlfile, "Sheet1", 814, 9, resFail);
			return false;
		}
	}
	
	
	public boolean checkNoOptionOnClickOnDelete() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getAlert().dismiss();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(buyerPriceBook_PriceBook));
		String actbuyerPriceBook_PriceBook =buyerPriceBook_PriceBook.getAttribute("value");
	    String expbuyerPriceBook_PriceBook ="BuyingCopy";
	    
	    System.out.println("**************************checkNoOptionOnClickOnDelete*****************************");
	    
	    System.out.println("buyerPriceBook_PriceBook  : "+actbuyerPriceBook_PriceBook+" Value Expected : "+expbuyerPriceBook_PriceBook);
	    
	   if(actbuyerPriceBook_PriceBook.equalsIgnoreCase(expbuyerPriceBook_PriceBook)) 
	   {
			System.out.println("Pass:No Button in Delete option");
			excelReader.setCellData(xlfile, "Sheet1", 815, 9, resPass);
			return true;
	   } 
	   else 
	   {
		   System.out.println("Fail: No Button in Delete option");
		   excelReader.setCellData(xlfile, "Sheet1", 815, 9, resFail);
		   return false;
	   }
	}
	
	
	public boolean checkYesOptionOnClickDelete() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookDeleteBtn));
		buyerPriceBookDeleteBtn.click();
	
		getWaitForAlert();
		
		getAlert().accept();

		String expMessage="Pricebook details deleted successfully.";
		
		System.out.println("******************************checkYesOptionOnClickDelete*****************************");
		
		String actMessage=checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			System.out.println("Pass: Yes Button in Delete option ");
			excelReader.setCellData(xlfile, "Sheet1", 816, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: Yes Button in Delete option ");
		    excelReader.setCellData(xlfile, "Sheet1", 816, 9, resFail);
			return false;
		}
	}
	
	public boolean checkExportExcelWithoutSelectingPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_ExportToExcel));
		buyerPriceBook_ExportToExcel.click();
		 
		String expMessage="Select Price book to export";
		
		System.out.println("*****************************checkExportExcelWithoutSelectingPriceBook****************************");
			
		String actMessage=checkValidationMessage(expMessage);
		    
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			System.out.println("Pass: Yes Button in Delete option ");
			excelReader.setCellData(xlfile, "Sheet1", 817, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: Yes Button in Delete option ");
			excelReader.setCellData(xlfile, "Sheet1", 817, 9, resFail);
			return false;
		}
	}
	
	
	
	
	public boolean checkExportExcel() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		File buyingRatesFile=new File("C:\\Users\\Administrator\\Downloads\\Test Hyd Wh.xlsx");
		
		try 
		{
			if (buyingRatesFile.exists())
			{
				System.out.println("File Present");
				buyingRatesFile.delete();
				System.out.println("File Deleted");
			} 
			else
			{
				System.out.println(" File Not Present");
			}
		} 
		catch (Exception e)
		{
			e.printStackTrace();
		}	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.END);
		buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		
		int pricebookcount = pricebookComboBoxList.size();
		   
		for(int i=0;i<pricebookcount;i++)
		{
			String data = pricebookComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Test Hyd Wh"))
			{
				pricebookComboBoxList.get(i).click();
				break;
			}
		}
		   
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_ExportToExcel));
		buyerPriceBook_ExportToExcel.click();
		
		String expMessage="Pricebook exported successfully";
		
		System.out.println("*****************************checkExportExcel*****************************");
		
		String actMessage=checkValidationMessage(expMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCloseBtn));
		buyerPriceBookCloseBtn.click();
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			System.out.println("Pass: Yes Button in Delete option ");
			excelReader.setCellData(xlfile, "Sheet1", 818, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: Yes Button in Delete option ");
			excelReader.setCellData(xlfile, "Sheet1", 818, 9, resFail);
			return false;
		}
	}
	
	
	@FindBy(xpath="//span[@class='btn btn-sm theme_background-color fileUpload theme_color']")
	private static WebElement browseButton;
	
	@FindBy(xpath="//input[@id='rdoFilterName']")
	private static WebElement nameRadioButton;
	
	@FindBy(xpath="//input[@id='rdoFilterCode']")
	private static WebElement codeRadioButton;
	
	@FindBy(xpath="//input[@id='rdoFilterBoth']")
	private static WebElement bothRadioButton;
	
	
	@FindBy(xpath="//input[@id='chkOverride']")
	private static WebElement overrideCheckbox;

	@FindBy(xpath="//button[@id='btnImportData']")
	private static WebElement importButton;

	@FindBy(xpath="//button[@id='btnCloseImport']")
	private static WebElement closeButtonInImport;

	
	public boolean checkFileExportedIntoFile() throws EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		File buyingRatesFile=new File("C:\\Users\\Administrator\\Downloads\\Test Hyd Wh.xlsx");

		boolean actFile =buyingRatesFile.exists();
		boolean expFile =true;
		
		System.out.println("*********************checkFileExportedIntoFile*******************");
		
		System.out.println("File  : "+actFile+" Value Expected : "+expFile);
		
		if(actFile==expFile)
		{
			excelReader.setCellData(xlfile, "Sheet1", 819, 9, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, "Sheet1", 819, 9, resFail);
			return false;
		}
	}
	

	
	public boolean checkImportFromExcelOptionWithoutPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		Thread.sleep(2000);
		
	  	Robot robot = new Robot();                          
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_J);
		robot.keyRelease(KeyEvent.VK_CONTROL); 
		robot.keyRelease(KeyEvent.VK_J);
		
		Thread.sleep(2000);
		
		robot.keyPress(KeyEvent.VK_CONTROL); 
		robot.keyPress(KeyEvent.VK_W); 
		robot.keyRelease(KeyEvent.VK_CONTROL); 
		robot.keyRelease(KeyEvent.VK_W);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	  	homeMenu.click();
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterMenu));
	  	homeMasterMenu.click();
	  
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	  	homeMasterItemMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
	  	homeMasterItem_buyerPriceBookMenu.click();
	  	 	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_ImportFromExcel));
	  	buyerPriceBook_ImportFromExcel.click();
		
	  	String expMessage="Enter price book name";
	  	
	  	System.out.println("****************************checkImportFromExcelOptionWithoutPriceBook***************************");
		
	  	String actMessage=checkValidationMessage(expMessage);
		    
	  	if (actMessage.equalsIgnoreCase(expMessage)) 
	  	{
	  		System.out.println("Pass: Yes Button in Delete option ");
	  		excelReader.setCellData(xlfile, "Sheet1", 820, 9, resPass);
	  		return true;
	  	} 
	  	else 
	  	{
	  		System.out.println("Fail: Yes Button in Delete option ");
	  		excelReader.setCellData(xlfile, "Sheet1", 820, 9, resFail);
	  		return false;
	  	}
	}
	
	public boolean checkImportFromExcelOptionWithPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	  
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys("Test Buying Price Import");
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.click();
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_ImportFromExcel));
	  	buyerPriceBook_ImportFromExcel.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(browseButton));
		
		boolean actbrowseButton 		=browseButton.isDisplayed();
		boolean actnameRadioButton 		=nameRadioButton.isDisplayed();
		boolean actcodeRadioButton 		=codeRadioButton.isDisplayed();
		boolean actbothRadioButton 		=bothRadioButton.isDisplayed();
		boolean actoverrideCheckbox 	=overrideCheckbox.isDisplayed();
		boolean actimportButton 		=importButton.isDisplayed();
		boolean actcloseButtonInImport 	=closeButtonInImport.isDisplayed();
		
		boolean expbrowseButton 		=true;
		boolean expnameRadioButton 		=true;
		boolean expcodeRadioButton 		=true;
		boolean expbothRadioButton 		=true;
		boolean expoverrideCheckbox 	=true;
		boolean expimportButton 		=true;
		boolean expcloseButtonInImport 	=true;
		
		System.out.println("*****************************checkImportFromExcelOptionWithPriceBook*********************************");
		
		System.out.println("browseButton         : "+actbrowseButton        +" Value Expected : "+expbrowseButton);
		System.out.println("nameRadioButton      : "+actnameRadioButton     +" Value Expected : "+expnameRadioButton);
		System.out.println("codeRadioButton      : "+actcodeRadioButton     +" Value Expected : "+expcodeRadioButton);
		System.out.println("bothRadioButton      : "+actbothRadioButton     +" Value Expected : "+expbothRadioButton);
		System.out.println("overrideCheckbox     : "+actoverrideCheckbox    +" Value Expected : "+expoverrideCheckbox);
		System.out.println("importButton         : "+actimportButton        +" Value Expected : "+expimportButton);
		System.out.println("closeButtonInImport  : "+actcloseButtonInImport +" Value Expected : "+expcloseButtonInImport);
		
		if (actbrowseButton==expbrowseButton && actnameRadioButton==expnameRadioButton && actcodeRadioButton==expcodeRadioButton
				&& actbothRadioButton==expbothRadioButton && actoverrideCheckbox==expoverrideCheckbox && actimportButton==expimportButton
				&& actcloseButtonInImport==expcloseButtonInImport)
        {
			System.out.println("Pass: Advance Import Option in Buyer Price Book Screen");
			excelReader.setCellData(xlfile, "Sheet1", 821, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: Advance Import Option in Buyer Price Book Screen");
			excelReader.setCellData(xlfile, "Sheet1", 821, 9, resFail);
			return false;
		}
	}
	
	
	
	public boolean checkImportButtonWithoutPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importButton));
	  	importButton.click();
		
	  	String expMessage="Please select the file to import.";
	  	
	  	System.out.println("***************************checkImportButtonWithoutPriceBook*************************");
		
	  	String actMessage=checkValidationMessage(expMessage);
	  	
	  	if (actMessage.equalsIgnoreCase(expMessage)) 
	  	{
	  		System.out.println("Pass: Yes Button in Delete option ");
	  		excelReader.setCellData(xlfile, "Sheet1", 822, 9, resPass);
	  		return true;
	  	} 
	  	else 
	  	{
	  		System.out.println("Fail: Yes Button in Delete option ");
	  		excelReader.setCellData(xlfile, "Sheet1", 822, 9, resFail);
	  		return false;
	  	}
	}
	
	
	
	public boolean checkSelectTheBrowseFile() throws IOException, InterruptedException, EncryptedDocumentException, InvalidFormatException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(browseButton));
		browseButton.click();
		
		Thread.sleep(2000);
		Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\FocusBuyerImportPrice.exe");
		Thread.sleep(3000);
		
		getWaitForAlert();
		 
		String expAlertMessage ="Do you want to load fields mapping?";
		
		String actAlertMessage =getAlert().getText();
	     
		System.out.println("************************checkSelectTheBrowseFile***********************");
		
		System.out.println("AlertMessage  : "+actAlertMessage+" Value Expected : "+expAlertMessage);
	     
		if(actAlertMessage.equalsIgnoreCase(expAlertMessage))
		{ 
			excelReader.setCellData(xlfile, "Sheet1", 823, 9, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, "Sheet1", 823, 9, resFail);
			return false;
		}
	}
		
	
		
	@FindBy(xpath="//span[contains(text(),'Select Sheet')]")
	private static WebElement sheetTitle;
	
	@FindBy(xpath="//li[@class='clsSheetName']")
	private static WebElement sheetName;
		
	@FindBy(xpath="//button[@id='btnSheetOk']")
	private static WebElement sheetOkButton;
		
	@FindBy(xpath="//button[@id='btnSheetClose']")
	private static WebElement sheetCloseButton;
			
		
	public boolean checkSheetPopUpToSelect() throws IOException, InterruptedException, EncryptedDocumentException, InvalidFormatException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getAlert().accept();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetTitle));
		
		boolean actsheetTitle 		=sheetTitle.isDisplayed();
		boolean actsheetName 		=sheetName.isDisplayed();
		boolean actsheetOkButton 	=sheetOkButton.isDisplayed();
		boolean actsheetCloseButton =sheetCloseButton.isDisplayed();
		
		boolean expsheetTitle 		=true;
		boolean expsheetName 		=true;
		boolean expsheetOkButton 	=true;
		boolean expsheetCloseButton =true;
		
		System.out.println("****************************checkSheetPopUpToSelect******************************");
		
		System.out.println("sheetTitle        : "+actsheetTitle       +" Value Expected : "+expsheetTitle);
		System.out.println("sheetName         : "+actsheetName        +" Value Expected : "+expsheetName);
		System.out.println("sheetOkButton     : "+actsheetOkButton    +" Value Expected : "+expsheetOkButton);
		System.out.println("sheetCloseButton  : "+actsheetCloseButton +" Value Expected : "+expsheetCloseButton);
	  	
		if(actsheetTitle==expsheetTitle && actsheetName==expsheetName && actsheetOkButton==expsheetOkButton && actsheetCloseButton==expsheetCloseButton)
		{ 
			excelReader.setCellData(xlfile, "Sheet1", 824, 9, resPass);
			return true;
		}	
		else
		{
			excelReader.setCellData(xlfile, "Sheet1", 824, 9, resFail);
			return false;
		}
	}
	
	public boolean checkOkButtonWithoutSelectingSheetInSheetPopUpToSelect() throws IOException, InterruptedException, EncryptedDocumentException, InvalidFormatException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetOkButton));
		sheetOkButton.click();
		
		getWaitForAlert();
		 
		String expAlertMessage ="Select Sheet";
		 
		String actAlertMessage =getAlert().getText();
	     
		System.out.println("************************checkOkButtonWithoutSelectingSheetInSheetPopUpToSelect**************************");
		
		System.out.println("AlertMessage  : "+actAlertMessage+" Value Expected : "+expAlertMessage);
		
		if(actAlertMessage.equalsIgnoreCase(expAlertMessage))
		{ 
			getAlert().accept();
			excelReader.setCellData(xlfile, "Sheet1", 825, 9, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, "Sheet1", 825, 9, resFail);
			return false;
		}
	}
		
	
	public boolean checkOkButtonWithSelectingSheetInSheetPopUpToSelect() throws IOException, InterruptedException, EncryptedDocumentException, InvalidFormatException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetName));
		sheetName.click();
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetOkButton));
		sheetOkButton.click();
	
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(browseButton));
		
		boolean actbrowseButton 		=browseButton.isDisplayed();
		boolean actnameRadioButton 		=nameRadioButton.isDisplayed();
		boolean actcodeRadioButton 		=codeRadioButton.isDisplayed();
		boolean actbothRadioButton 		=bothRadioButton.isDisplayed();
		boolean actoverrideCheckbox 	=overrideCheckbox.isDisplayed();
		boolean actimportButton 		=importButton.isDisplayed();
		boolean actcloseButtonInImport 	=closeButtonInImport.isDisplayed();
		
		boolean expbrowseButton 		=true;
		boolean expnameRadioButton 		=true;
		boolean expcodeRadioButton 		=true;
		boolean expbothRadioButton 		=true;
		boolean expoverrideCheckbox 	=true;
		boolean expimportButton 		=true;
		boolean expcloseButtonInImport 	=true;
		
		System.out.println("*****************************checkOkButtonWithSelectingSheetInSheetPopUpToSelect*********************************");
		
		System.out.println("browseButton         : "+actbrowseButton        +" Value Expected : "+expbrowseButton);
		System.out.println("nameRadioButton      : "+actnameRadioButton     +" Value Expected : "+expnameRadioButton);
		System.out.println("codeRadioButton      : "+actcodeRadioButton     +" Value Expected : "+expcodeRadioButton);
		System.out.println("bothRadioButton      : "+actbothRadioButton     +" Value Expected : "+expbothRadioButton);
		System.out.println("overrideCheckbox     : "+actoverrideCheckbox    +" Value Expected : "+expoverrideCheckbox);
		System.out.println("importButton         : "+actimportButton        +" Value Expected : "+expimportButton);
		System.out.println("closeButtonInImport  : "+actcloseButtonInImport +" Value Expected : "+expcloseButtonInImport);
		
		if (actbrowseButton==expbrowseButton && actnameRadioButton==expnameRadioButton && actcodeRadioButton==expcodeRadioButton
				&& actbothRadioButton==expbothRadioButton && actoverrideCheckbox==expoverrideCheckbox && actimportButton==expimportButton
				&& actcloseButtonInImport==expcloseButtonInImport)
		{
			System.out.println("Pass: Advance Import Option in Buyer Price Book Screen");
			excelReader.setCellData(xlfile, "Sheet1", 826, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: Advance Import Option in Buyer Price Book Screen");
			excelReader.setCellData(xlfile, "Sheet1", 826, 9, resFail);
			return false;
		}
	}
		
		
	public boolean checkMandatoryFieldsWithoutSelectExternalFields() throws IOException, InterruptedException, EncryptedDocumentException, InvalidFormatException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importButton));
	  	importButton.click();
		
	  	String expMessage="Item must be selected";
	  	
	  	System.out.println("***********************checkMandatoryFieldsWithoutSelectExternalFields**********************");
		
	  	String actMessage=checkValidationMessage(expMessage);
	
	  	if (actMessage.equalsIgnoreCase(expMessage)) 
	  	{
	  		System.out.println("Pass: Yes Button in Delete option ");
	  		excelReader.setCellData(xlfile, "Sheet1", 827, 9, resPass);
	  		return true;
	  	} 
	  	else 
	  	{
	  		System.out.println("Fail: Yes Button in Delete option ");
	  		excelReader.setCellData(xlfile, "Sheet1", 827, 9, resFail);
	  		return false;
	  	}
	}
		
	
	
	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[1]//td[3]")
	private static WebElement sheetItemHeading;
	
	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList0']")
	private static WebElement sheetItem;
			
	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[2]//td[3]")
	private static WebElement sheetAccountHeading;
	
	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList1']")
	private static WebElement sheetAccount;
	
	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[3]//td[3]")
	private static WebElement sheetTagHeading;
	
	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList2']")
	private static WebElement sheetTag;
	
	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[4]//td[3]")
	private static WebElement sheetCurrencyHeading;
	
	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList3']")
	private static WebElement sheetCurrency;
	
	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[7]//td[3]")
	private static WebElement sheetMinQtyHeading;
	
	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList6']")
	private static WebElement sheetMinQty;
	
	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[8]//td[3]")
	private static WebElement sheetMaxQtyHeading;
	
	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList7']")
	private static WebElement sheetMaxQty;
	
	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[9]//td[3]")
	private static WebElement sheetUnitHeading;
	
	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList8']")
	private static WebElement sheetUnitQty;
	
	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[10]//td[3]")
	private static WebElement sheetRateHeading;
	
	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList9']")
	private static WebElement sheetRate;
	
	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[11]//td[3]")
	private static WebElement sheetPOHeading;
	
	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList10']")
	private static WebElement sheetPO;
	
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_1-10']")
	private static WebElement unitValueImport;
	
	
	public boolean checkCountOfExtraFields() throws IOException, InterruptedException, EncryptedDocumentException, InvalidFormatException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetItemHeading));
		sheetItemHeading.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetItem));
		sheetItem.click();
	
		Select oSelect = new Select(sheetItem);
		List <WebElement> elementCount = oSelect.getOptions();
		
		int actsheetItem = elementCount.size();
		int expsheetItem =26;
		
		System.out.println("************************checkCountOfExtraFields*******************************");
		
		System.out.println("sheetItem  : "+actsheetItem+" Value Expected : "+expsheetItem);
		
		if(actsheetItem==expsheetItem) 
		{
			System.out.println("Pass: Yes Button in Delete option ");
			excelReader.setCellData(xlfile, "Sheet1", 828, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: Yes Button in Delete option ");
			excelReader.setCellData(xlfile, "Sheet1", 828, 9, resFail);
			return false;
		}
	}
		
		
	
	public boolean checkImportDataOption() throws IOException, InterruptedException, EncryptedDocumentException, InvalidFormatException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetItem));
		sheetItem.click();
		sheetItem.sendKeys("ProductName");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAccountHeading));
		sheetAccountHeading.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetAccount));
		sheetAccount.click();
		sheetAccount.sendKeys("AccountName");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetTagHeading));
		sheetTagHeading.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetTag));
		sheetTag.click();
		sheetTag.sendKeys("TagName");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetCurrencyHeading));
		sheetCurrencyHeading.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetCurrency));
		sheetCurrency.click();
		sheetCurrency.sendKeys("CurrencyName");
		
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
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetRateHeading));
		sheetRateHeading.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetRate));
		sheetRate.click();
		sheetRate.sendKeys("Rate");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetPOHeading));
		sheetPOHeading.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sheetPO));
		sheetPO.click();
		sheetPO.sendKeys("PO");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importButton));
	  	importButton.click();
	  	
	  	String expMessage="Price book imported successfully";
	  	
	  	System.out.println("*************************checkImportDataOption****************************");
		
	  	String actMessage=checkValidationMessage(expMessage);
	  	
	  	//Thread.sleep(2000);
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCloseAfterImport));
	  	buyerPriceBookCloseAfterImport.click();
		    
	  	if (actMessage.equalsIgnoreCase(expMessage)) 
	  	{
	  		System.out.println("Pass: Yes Button in Delete option ");
	  		excelReader.setCellData(xlfile, "Sheet1", 829, 9, resPass);
	  		return true;
	  	} 
	  	else 
	  	{
	  		System.out.println("Fail: Yes Button in Delete option ");
	  		excelReader.setCellData(xlfile, "Sheet1", 829, 9, resFail);
	  		return false;
	  	}
	}
		
	
		
	public boolean checkImportPriceBook() throws IOException, InterruptedException, EncryptedDocumentException, InvalidFormatException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	  	homeMenu.click();
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterMenu));
	  	homeMasterMenu.click();
	  
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	  	homeMasterItemMenu.click();
	  	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
	  	homeMasterItem_buyerPriceBookMenu.click();
	  	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.END);
		buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		
		int pricebookcount = pricebookComboBoxList.size();
		   
		for(int i=0;i<pricebookcount;i++)
		{
			String data = pricebookComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Test Buying Price Import"))
			{
				pricebookComboBoxList.get(i).click();
				break;
			}
		}
		   
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.click();
		 
		String actbuyerPriceBook_PriceBook =buyerPriceBook_PriceBook.getAttribute("value");
	    String expbuyerPriceBook_PriceBook ="Test Buying Price Import";
	    
	    System.out.println("*******************************checkImportPriceBook********************************");
	   
	    System.out.println("buyerPriceBook_PriceBook : "+actbuyerPriceBook_PriceBook+" Value Expected : "+expbuyerPriceBook_PriceBook);
	    
	    if(actbuyerPriceBook_PriceBook.equalsIgnoreCase(expbuyerPriceBook_PriceBook)) 
	    {
			System.out.println("Pass: Edit Buyer PriceBook");
			excelReader.setCellData(xlfile, "Sheet1", 830, 9, resPass);	
	        return true;
	    } 
	    else 
	    {
	    	System.out.println("Fail: Edit Buyer PriceBook");
	    	excelReader.setCellData(xlfile, "Sheet1", 830, 9, resFail);
			return false;
	    }
	}
	
	
	
	public boolean checkImportPriceBookInformation() throws IOException, InterruptedException, EncryptedDocumentException, InvalidFormatException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
		buyerPriceBook_CustamizeSettingBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeDateRangeCheckBox));
		buyerPriceBook_CustamizeDateRangeCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustomizeVendorCheckBox));
		buyerPriceBook_CustomizeVendorCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_WarehouseCheckBox));
		buyerPriceBook_WarehouseCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_QualityRangeCheckBox));
		buyerPriceBook_QualityRangeCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CurrencyCheckBox));
		buyerPriceBook_CurrencyCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_UnitCheckBox));
		buyerPriceBook_UnitCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
		buyerPriceBook_CustamizeSettingBtn.click();
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
		buyerPriceBookFilter_FilterAndLoadBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(beforeClickOnToggle));
		beforeClickOnToggle.click();
		    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
			
		String actitemValue  =itemValue.getAttribute("value");
		String expitemValue  ="BATCH WA ITEM";
		    
		itemValue.sendKeys(Keys.TAB);
		    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
		wareHouseValue.click();
		
		String actwareHouseValue =wareHouseValue.getAttribute("value");
		String expwareHouseValue ="SECUNDERABAD";
		    
		wareHouseValue.sendKeys(Keys.TAB);
			
		startingDateValue.sendKeys(Keys.TAB);
		
		if(endingDateChkBox.isSelected()==true)
		{
			endingDateChkBox.click();
		}
		
		if(endingDateChkBox.isSelected()==false)
		{
			endingDateChkBox.click();
		}
		
		endingDateValue.sendKeys(Keys.TAB);
		    	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
		vendorValue.click();
			
		String actvendorValue =vendorValue.getAttribute("value");
		String expvendorValue ="Vendor B";
		    
		vendorValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyOptionValue));
		currencyOptionValue.click();
			
		String actcurrencyOptionValue =currencyOptionValue.getAttribute("value");
		String expcurrencyOptionValue ="Indian Rupees";
		    
		currencyOptionValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
		minQtyValue.click();
		
		String actminQtyValue =minQtyValue.getAttribute("value");
		String expminQtyValue ="9";
		    
		minQtyValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
		maxQtyValue.click();
			
		String actmaxQtyValue =maxQtyValue.getAttribute("value");
		String expmaxQtyValue ="39";
			    
		maxQtyValue.sendKeys(Keys.TAB);
			
		unitValue.sendKeys(Keys.TAB);
		
		rateValue.sendKeys(Keys.TAB);
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PoValue));
		PoValue.click();
			
		String actPoValue =PoValue.getAttribute("value");
		String expPoValue ="789.000000";   //798.00
		PoValue.sendKeys(Keys.SHIFT,Keys.TAB);
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitFirstCellSelect));
		unitFirstCellSelect.click();

		String actunitValue =unitValue.getAttribute("value");
		String expunitValue ="Pcs";
		
		unitValue.sendKeys(Keys.TAB);
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.click();
			
		String actrateValue =rateValue.getAttribute("value");
		String exprateValue ="123.000000";

		rateValue.sendKeys(Keys.TAB);
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateSecondCellSelect));
		rateSecondCellSelect.click();
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.click();
			
		String actrateValueTwo =rateValue.getAttribute("value");
		String exprateValueTwo ="1000.000000";

		rateValue.sendKeys(Keys.TAB);
		
		System.out.println("****************************checkImportPriceBookInformation**************************");
	
		System.out.println("itemValue            : "+actitemValue           +" Value Expected : "+expitemValue);
		System.out.println("wareHouseValue       : "+actwareHouseValue      +" Value Expected : "+expwareHouseValue);
		System.out.println("vendorValue          : "+actvendorValue         +" Value Expected : "+expvendorValue);
		System.out.println("currencyOptionValue  : "+actcurrencyOptionValue +" Value Expected : "+expcurrencyOptionValue);
		System.out.println("minQtyValue          : "+actminQtyValue         +" Value Expected : "+expminQtyValue);
		System.out.println("maxQtyValue          : "+actmaxQtyValue         +" Value Expected : "+expmaxQtyValue);
		System.out.println("PoValue              : "+actPoValue             +" Value Expected : "+expPoValue);
		System.out.println("unitValue            : "+actunitValue           +" Value Expected : "+expunitValue);
		System.out.println("rateValue            : "+actrateValue           +" Value Expected : "+exprateValue);
		System.out.println("rateValueTwo         : "+actrateValueTwo        +" Value Expected : "+exprateValueTwo);
				
		if (actitemValue.equalsIgnoreCase(expitemValue) && actwareHouseValue.equalsIgnoreCase(expwareHouseValue) && actvendorValue.equalsIgnoreCase(expvendorValue)
				&& actcurrencyOptionValue.equalsIgnoreCase(expcurrencyOptionValue) && actminQtyValue.equalsIgnoreCase(expminQtyValue)
				&& actmaxQtyValue.equalsIgnoreCase(expmaxQtyValue) && actPoValue.equalsIgnoreCase(expPoValue) && actunitValue.equalsIgnoreCase(expunitValue)
				&& actrateValue.equalsIgnoreCase(exprateValue) && actrateValueTwo.equalsIgnoreCase(exprateValueTwo)) 
		{
			System.out.println("Pass : Allow to select and HYD WH should be display in the Warehouse Combo box");
			excelReader.setCellData(xlfile, "Sheet1", 831, 9, resPass);	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCloseAfterImport));
			buyerPriceBookCloseAfterImport.click();
			return true;
		} 
		else 
		{
			System.out.println("Fail: Allow to select and HYD WH should be display in the Warehouse Combo box");
			excelReader.setCellData(xlfile, "Sheet1", 831, 9, resFail);	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCloseAfterImport));
			buyerPriceBookCloseAfterImport.click();
			return false;
		}
	}
	
		
	
	
	

	
	
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	///////////////////End of Buyer Screen////////////////////////////////////////////////////////////////////////
	
	////////////////////////////////////////////////////////////////////////////////////////////////////////
	
	
	
///////////////////////////////////**************************************************///////////////////////////////	
	
	
	
	
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////	
	
	
	

	@FindBy(xpath="//*[@id='1']/div/div")
	private static WebElement  HomeMenu;   


	@FindBy(xpath="//*[@id='1000']/span")
	private static WebElement  HomeMasterMenu;


	@FindBy(xpath="//*[@id='221']/span")
	private static WebElement  HomeMasterItemMenu;
	
	@FindBy(xpath="//span[contains(text(),'Seller Price Book')]")
	private static WebElement  homeMasterItemSellerPriceBookMenu;
	
	
	
	@FindBy(xpath="//*[@id='72']/span")
	private static WebElement  HomeMasterItemUnit_ConversionMenu;
	
	
	@FindBy(xpath="//*[@id='ctrlOptionProPriceBookH']")
	private static WebElement  HomeMasterItemSellerPriceBook_PriceBookBtn;
	
	@FindBy(xpath="//*[@id='ctrlOptionProPriceBookH_input_settings']/span")
	private static WebElement  HomeMasterItemSellerPriceBook_PriceBookSettingBtn;
	
	
	@FindBy(xpath="//*[@id='txtAbbreviation']")
	private static WebElement  HomeMasterItemSellerPriceBook_AbbreviationBtn;
	
	@FindBy(xpath="//*[@id='PriceBookContainer']/div[2]/nav/div/div[1]/a/div/span[2]")
	private static WebElement  HomeMasterItemSellerPriceBookLabel;
	
	@FindBy(xpath="//*[@id='btnClear']")
	private static WebElement  HomeMasterItemSellerPriceBookClearBtn;
	
	@FindBy(xpath="//*[@id='btnPaste']")
	private static WebElement  HomeMasterItemSellerPriceBookPasteBtn;
	
	@FindBy(xpath="//*[@id='btnCopy']")
	private static WebElement  HomeMasterItemSellerPriceBookCopyBtn;
	
	@FindBy(xpath="//*[@id='btnSelectAll']")
	private static WebElement  HomeMasterItemSellerPriceBookSelectAllBtn;
	
	@FindBy(xpath="//*[@id='btnClose']")
	private static WebElement  HomeMasterItemSellerPriceBookCloseBtn;
	
	@FindBy(xpath="//*[@id='btnSave']")
	private static WebElement  HomeMasterItemSellerPriceBookSaveBtn;
	
	@FindBy(xpath="//*[@id='btnDelete']")
	private static WebElement  HomeMasterItemSellerPriceBookDeleteBtn;
	
	
	@FindBy(xpath="//*[@id='StartingDateH']")
	private static WebElement  HomeMasterItemSellerPriceBookStaringDateBtn;
	
	@FindBy(xpath="//*[@id='EndingDateH']")
	private static WebElement  HomeMasterItemSellerPriceBookENdingDateBtn;
	
	
	@FindBy(xpath="//*[@id='InActive']")
	private static WebElement  HomeMasterItemSellerPriceBookInActiveCheckBox;
	
	@FindBy(xpath="//*[@id='FilterProduct']")
	private static WebElement  HomeMasterItemSellerPriceBookFilter_ItemBtn;
	
	@FindBy(xpath="//*[@id='FilterCurrency']")
	private static WebElement  HomeMasterItemSellerPriceBookFilter_CurrencyBtn;
	
	
	@FindBy(xpath="//*[@id='FilterCustomer']")
	private static WebElement  HomeMasterItemSellerPriceBookFilter_CustomerBtn;
	
	
	@FindBy(xpath="//*[@id='FilterDepartment']")
	private static WebElement  HomeMasterItemSellerPriceBookFilter_DepartmentBtn;
	
	
	@FindBy(xpath="//*[@id='btnFilterClear']")
	private static WebElement  HomeMasterItemSellerPriceBookFilter_ClearBtn;
	
	
	@FindBy(xpath="//*[@id='btnFilter']")
	private static WebElement  HomeMasterItemSellerPriceBookFilter_FilterBtn;
	
	
	@FindBy(xpath="//*[@id='btnLoad']")
	private static WebElement  HomeMasterItemSellerPriceBookFilter_FilterAndLoadBtn;
	
	
	@FindBy(xpath="//*[@id='btnAdvExcelImport']")
	private static WebElement  HomeMasterItemSellerPriceBook_AdvanceExcelImportBtn;
	
	
	@FindBy(xpath="//*[@id='btnPBPrintRateHistory']")
	private static WebElement  HomeMasterItemSellerPriceBook_PrintRateHistoryBtn;
	
	
	@FindBy(xpath="//*[@id='btnPBExportToXML']")
	private static WebElement  HomeMasterItemSellerPriceBook_ExportToXmlBtn;
	
	
	@FindBy(xpath="//*[@id='uploadBtn']")
	private static WebElement  HomeMasterItemSellerPriceBook_ImportFromXmlBtn;
	
	@FindBy(xpath="//*[@id='btnExporttoExcel']")
	private static WebElement  HomeMasterItemSellerPriceBook_ExportToExcel;
	
	@FindBy(xpath="//*[@id='btnPBImportFromExcel']")
	private static WebElement  HomeMasterItemSellerPriceBook_ImportFromExcel;
	
	@FindBy(xpath="//*[@id='btnViewPrice']")
	private static WebElement  HomeMasterItemSellerPriceBook_ViewPrice;
	
	@FindBy(xpath="//*[@id='btnUpdatePriceList']")
	private static WebElement  HomeMasterItemSellerPriceBook_UpdatePriceList;
	
	
	@FindBy(xpath="//*[@id='toggleDropDown']/i")
	private static WebElement  HomeMasterItemSellerPriceBook_CustamizeSettingBtn;
	
	@FindBy(xpath="//*[@id='chkDateRange']")
	private static WebElement  HomeMasterItemSellerPriceBook_CustamizeDateRangeCheckBoxBtn;
	
	@FindBy(xpath="//*[@id='chkCustomer']")
	private static WebElement  HomeMasterItemSellerPriceBook_CustamizeCustomerCheckBoxBtn;
	
	@FindBy(xpath="//*[@id='chkDepartment']")
	private static WebElement  HomeMasterItemSellerPriceBook_DepartmentCheckBoxBtn;
	
	
	@FindBy(xpath="//*[@id='chkQtyRange']")
	private static WebElement  HomeMasterItemSellerPriceBook_QualityRangeCheckBoxBtn;
	
	@FindBy(xpath="//*[@id='chkCurrency']")
	private static WebElement  HomeMasterItemSellerPriceBook_CurrencyCheckBoxBtn;
	
	
	@FindBy(xpath="//*[@id='chkUnit']")
	private static WebElement  HomeMasterItemSellerPriceBook_UnitCheckBoxBtn;
	
	
	public boolean checkSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		//Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	  	homeMenu.click();
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterMenu));
	  	homeMasterMenu.click();
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	  	homeMasterItemMenu.click();
		
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemSellerPriceBookMenu));
	  	homeMasterItemSellerPriceBookMenu.click();
	  	Thread.sleep(2000);
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		
		boolean actbuyerPriceBook_PriceBook 				=buyerPriceBook_PriceBook.isDisplayed();
		boolean actbuyerPriceBook_AbbreviationBtn 			=buyerPriceBook_AbbreviationBtn.isDisplayed();
		boolean actbuyerPriceBookLabel 						=buyerPriceBookLabel.isDisplayed();
		boolean actbuyerPriceBookClearBtn 					=buyerPriceBookClearBtn.isDisplayed();
		boolean actbuyerPriceBookPasteBtn 					=buyerPriceBookPasteBtn.isDisplayed();
		boolean actbuyerPriceBookCopyBtn 					=buyerPriceBookCopyBtn.isDisplayed();
		boolean actbuyerPriceBookSelectAllBtn 				=buyerPriceBookSelectAllBtn.isDisplayed();
		boolean actbuyerPriceBookCloseBtn 					=buyerPriceBookCloseBtn.isDisplayed();
		boolean actbuyerPriceBookSaveBtn 					=buyerPriceBookSaveBtn.isDisplayed();
		boolean actbuyerPriceBookDeleteBtn 					=buyerPriceBookDeleteBtn.isDisplayed();
		boolean actbuyerPriceBookStaringDate 				=buyerPriceBookStaringDate.isDisplayed();
		boolean actbuyerPriceBookENdingDate 				=buyerPriceBookENdingDate.isDisplayed();
		boolean actbuyerPriceBookInActiveCheckBox 			=buyerPriceBookInActiveCheckBox.isDisplayed();
		boolean actbuyerPriceBookFilter_Item 				=buyerPriceBookFilter_Item.isDisplayed();
		boolean actbuyerPriceBookFilter_Currency 			=buyerPriceBookFilter_Currency.isDisplayed();
		boolean actbuyerPriceBookFilter_Customer 			=buyerPriceBookFilter_Customer.isDisplayed();
		boolean actbuyerPriceBookFilter_Department 			=buyerPriceBookFilter_Department.isDisplayed();
		boolean actbuyerPriceBookFilter_ClearBtn 			=buyerPriceBookFilter_ClearBtn.isDisplayed();
		boolean actbuyerPriceBookFilter_FilterBtn 			=buyerPriceBookFilter_FilterBtn.isDisplayed();
		boolean actbuyerPriceBookFilter_FilterAndLoadBtn 	=buyerPriceBookFilter_FilterAndLoadBtn.isDisplayed();
		boolean actbuyerPriceBook_AdvanceExcelImportBtn 	=buyerPriceBook_AdvanceExcelImportBtn.isDisplayed();
		boolean actbuyerPriceBook_PrintRateHistoryBtn 		=buyerPriceBook_PrintRateHistoryBtn.isDisplayed();
		boolean actbuyerPriceBook_ExportToXmlBtn 			=buyerPriceBook_ExportToXmlBtn.isDisplayed();
		boolean actbuyerPriceBook_ImportFromXmlBtn 			=buyerPriceBook_ImportFromXmlBtn.isDisplayed();
		boolean actbuyerPriceBook_ExportToExcel 			=buyerPriceBook_ExportToExcel.isDisplayed();
		boolean actbuyerPriceBook_ImportFromExcel 			=buyerPriceBook_ImportFromExcel.isDisplayed();
		boolean actbuyerPriceBook_ViewPrice 				=buyerPriceBook_ViewPrice.isDisplayed();
		boolean actbuyerPriceBook_UpdatePriceList 			=buyerPriceBook_UpdatePriceList.isDisplayed();
		
		boolean expbuyerPriceBook_PriceBook 				=true;
		boolean expbuyerPriceBook_AbbreviationBtn 			=true;
		boolean expbuyerPriceBookLabel 						=true;
		boolean expbuyerPriceBookClearBtn 					=true;
		boolean expbuyerPriceBookPasteBtn 					=true;
		boolean expbuyerPriceBookCopyBtn 					=true;
		boolean expbuyerPriceBookSelectAllBtn 				=true;
		boolean expbuyerPriceBookCloseBtn 					=true;
		boolean expbuyerPriceBookSaveBtn 					=true;
		boolean expbuyerPriceBookDeleteBtn 					=true;
		boolean expbuyerPriceBookStaringDate 				=true;
		boolean expbuyerPriceBookENdingDate 				=true;
		boolean expbuyerPriceBookInActiveCheckBox 			=true;
		boolean expbuyerPriceBookFilter_Item 				=true;
		boolean expbuyerPriceBookFilter_Currency 			=true;
		boolean expbuyerPriceBookFilter_Customer 			=true;
		boolean expbuyerPriceBookFilter_Department 			=true;
		boolean expbuyerPriceBookFilter_ClearBtn 			=true;
		boolean expbuyerPriceBookFilter_FilterBtn 			=true;
		boolean expbuyerPriceBookFilter_FilterAndLoadBtn 	=true;
		boolean expbuyerPriceBook_AdvanceExcelImportBtn 	=true;
		boolean expbuyerPriceBook_PrintRateHistoryBtn 		=true;
		boolean expbuyerPriceBook_ExportToXmlBtn 			=true;
		boolean expbuyerPriceBook_ImportFromXmlBtn 			=true;
		boolean expbuyerPriceBook_ExportToExcel 			=true;
		boolean expbuyerPriceBook_ImportFromExcel 			=true;
		boolean expbuyerPriceBook_ViewPrice 				=true;
		boolean expbuyerPriceBook_UpdatePriceList 			=true;
		
		System.out.println("********************************checkSellerPriceBook*************************************");
		
		System.out.println("buyerPriceBook_PriceBook              : "+actbuyerPriceBook_PriceBook              +" Value Expected : "+expbuyerPriceBook_PriceBook);
		System.out.println("buyerPriceBook_AbbreviationBtn        : "+actbuyerPriceBook_AbbreviationBtn        +" Value Expected : "+expbuyerPriceBook_AbbreviationBtn);
		System.out.println("buyerPriceBookLabel                   : "+actbuyerPriceBookLabel                   +" Value Expected : "+expbuyerPriceBookLabel);
		System.out.println("buyerPriceBookClearBtn                : "+actbuyerPriceBookClearBtn                +" Value Expected : "+expbuyerPriceBookClearBtn);
		System.out.println("buyerPriceBookPasteBtn                : "+actbuyerPriceBookPasteBtn                +" Value Expected : "+expbuyerPriceBookPasteBtn);
		System.out.println("buyerPriceBookCopyBtn                 : "+actbuyerPriceBookCopyBtn                 +" Value Expected : "+expbuyerPriceBookCopyBtn);
		System.out.println("buyerPriceBookSelectAllBtn            : "+actbuyerPriceBookSelectAllBtn            +" Value Expected : "+expbuyerPriceBookSelectAllBtn);
		System.out.println("buyerPriceBookCloseBtn                : "+actbuyerPriceBookCloseBtn                +" Value Expected : "+expbuyerPriceBookCloseBtn);
		System.out.println("buyerPriceBookSaveBtn                 : "+actbuyerPriceBookSaveBtn                 +" Value Expected : "+expbuyerPriceBookSaveBtn);
		System.out.println("buyerPriceBookDeleteBtn               : "+actbuyerPriceBookDeleteBtn               +" Value Expected : "+expbuyerPriceBookDeleteBtn);
		System.out.println("buyerPriceBookStaringDate             : "+actbuyerPriceBookStaringDate             +" Value Expected : "+expbuyerPriceBookStaringDate);
		System.out.println("buyerPriceBookENdingDate              : "+actbuyerPriceBookENdingDate              +" Value Expected : "+expbuyerPriceBookENdingDate);
		System.out.println("buyerPriceBookInActiveCheckBox        : "+actbuyerPriceBookInActiveCheckBox        +" Value Expected : "+expbuyerPriceBookInActiveCheckBox);
		System.out.println("buyerPriceBookFilter_Item             : "+actbuyerPriceBookFilter_Item             +" Value Expected : "+expbuyerPriceBookFilter_Item);
		System.out.println("buyerPriceBookFilter_Currency         : "+actbuyerPriceBookFilter_Currency         +" Value Expected : "+expbuyerPriceBookFilter_Currency);
		System.out.println("buyerPriceBookFilter_Customer         : "+actbuyerPriceBookFilter_Customer         +" Value Expected : "+expbuyerPriceBookFilter_Customer);
		System.out.println("buyerPriceBookFilter_Department       : "+actbuyerPriceBookFilter_Department       +" Value Expected : "+expbuyerPriceBookFilter_Department);
		System.out.println("buyerPriceBookFilter_ClearBtn         : "+actbuyerPriceBookFilter_ClearBtn         +" Value Expected : "+expbuyerPriceBookFilter_ClearBtn);
		System.out.println("buyerPriceBookFilter_FilterBtn        : "+actbuyerPriceBookFilter_FilterBtn        +" Value Expected : "+expbuyerPriceBookFilter_FilterBtn);
		System.out.println("buyerPriceBookFilter_FilterAndLoadBtn : "+actbuyerPriceBookFilter_FilterAndLoadBtn +" Value Expected : "+expbuyerPriceBookFilter_FilterAndLoadBtn);
		System.out.println("buyerPriceBook_AdvanceExcelImportBtn  : "+actbuyerPriceBook_AdvanceExcelImportBtn  +" Value Expected : "+expbuyerPriceBook_AdvanceExcelImportBtn);
		System.out.println("buyerPriceBook_PrintRateHistoryBtn    : "+actbuyerPriceBook_PrintRateHistoryBtn    +" Value Expected : "+expbuyerPriceBook_PrintRateHistoryBtn);
		System.out.println("buyerPriceBook_ExportToXmlBtn         : "+actbuyerPriceBook_ExportToXmlBtn         +" Value Expected : "+expbuyerPriceBook_ExportToXmlBtn);
		System.out.println("buyerPriceBook_ImportFromXmlBtn       : "+actbuyerPriceBook_ImportFromXmlBtn       +" Value Expected : "+expbuyerPriceBook_ImportFromXmlBtn);
		System.out.println("buyerPriceBook_ExportToExcel          : "+actbuyerPriceBook_ExportToExcel          +" Value Expected : "+expbuyerPriceBook_ExportToExcel);
		System.out.println("buyerPriceBook_ImportFromExcel        : "+actbuyerPriceBook_ImportFromExcel        +" Value Expected : "+expbuyerPriceBook_ImportFromExcel);
		System.out.println("buyerPriceBook_ViewPrice              : "+actbuyerPriceBook_ViewPrice              +" Value Expected : "+expbuyerPriceBook_ViewPrice);
		System.out.println("buyerPriceBook_UpdatePriceList        : "+actbuyerPriceBook_UpdatePriceList        +" Value Expected : "+expbuyerPriceBook_UpdatePriceList);
		
	  	if (actbuyerPriceBook_PriceBook==expbuyerPriceBook_PriceBook && actbuyerPriceBook_AbbreviationBtn==expbuyerPriceBook_AbbreviationBtn
	  			&& actbuyerPriceBookLabel==expbuyerPriceBookLabel && actbuyerPriceBookClearBtn==expbuyerPriceBookClearBtn
	  			&& actbuyerPriceBookPasteBtn==expbuyerPriceBookPasteBtn && actbuyerPriceBookCopyBtn==expbuyerPriceBookCopyBtn
	  			&& actbuyerPriceBookSelectAllBtn==expbuyerPriceBookSelectAllBtn && actbuyerPriceBookCloseBtn==expbuyerPriceBookCloseBtn
	  			&& actbuyerPriceBookSaveBtn==expbuyerPriceBookSaveBtn && actbuyerPriceBookDeleteBtn==expbuyerPriceBookDeleteBtn
	  			&& actbuyerPriceBookStaringDate==expbuyerPriceBookStaringDate && actbuyerPriceBookENdingDate==expbuyerPriceBookENdingDate
	  			&& actbuyerPriceBookInActiveCheckBox==expbuyerPriceBookInActiveCheckBox && actbuyerPriceBookFilter_Item==expbuyerPriceBookFilter_Item
	  			&& actbuyerPriceBookFilter_Currency==expbuyerPriceBookFilter_Currency && actbuyerPriceBookFilter_Customer==expbuyerPriceBookFilter_Customer
	  			&& actbuyerPriceBookFilter_Department==expbuyerPriceBookFilter_Department && actbuyerPriceBookFilter_ClearBtn==expbuyerPriceBookFilter_ClearBtn
	  			&& actbuyerPriceBookFilter_FilterBtn==expbuyerPriceBookFilter_FilterBtn && actbuyerPriceBookFilter_FilterAndLoadBtn==expbuyerPriceBookFilter_FilterAndLoadBtn
	  			&& actbuyerPriceBook_AdvanceExcelImportBtn==expbuyerPriceBook_AdvanceExcelImportBtn && actbuyerPriceBook_PrintRateHistoryBtn==expbuyerPriceBook_PrintRateHistoryBtn
	  			&& actbuyerPriceBook_ExportToXmlBtn==expbuyerPriceBook_ExportToXmlBtn && actbuyerPriceBook_ImportFromXmlBtn==expbuyerPriceBook_ImportFromXmlBtn
	  			&& actbuyerPriceBook_ExportToExcel==expbuyerPriceBook_ExportToExcel && actbuyerPriceBook_ImportFromExcel==expbuyerPriceBook_ImportFromExcel
	  			&& actbuyerPriceBook_ViewPrice==expbuyerPriceBook_ViewPrice && actbuyerPriceBook_UpdatePriceList==expbuyerPriceBook_UpdatePriceList)
	  	{
	  		System.out.println("Pass: Dispalyed Seller Price Book Screen ");
	  		excelReader.setCellData(xlfile, "Sheet1", 837, 9, resPass);
	  		return true;
		} 
	  	else
	  	{
	  		System.out.println("Fail: Dispalyed Seller Price Book Screen ");
	  		excelReader.setCellData(xlfile, "Sheet1", 837, 9, resFail);
	  		return false;
		}
	}
	
	
	
	
	
	///---------------------------------------------------
	
	public boolean checkSellerPricebookSettingsOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{

		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	  	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
		buyerPriceBook_CustamizeSettingBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeDateRangeCheckBox));
		
		boolean actbuyerPriceBook_CustamizeDateRangeCheckBox   	=buyerPriceBook_CustamizeDateRangeCheckBox.isDisplayed();
		boolean actbuyerPriceBook_CustomizeVendorCheckBox   	=buyerPriceBook_CustomizeVendorCheckBox.isDisplayed();
		boolean actbuyerPriceBook_WarehouseCheckBox   			=buyerPriceBook_WarehouseCheckBox.isDisplayed();
		boolean actbuyerPriceBook_QualityRangeCheckBox   		=buyerPriceBook_QualityRangeCheckBox.isDisplayed();
		boolean actbuyerPriceBook_CurrencyCheckBox   			=buyerPriceBook_CurrencyCheckBox.isDisplayed();
		boolean actbuyerPriceBook_UnitCheckBox   				=buyerPriceBook_UnitCheckBox.isDisplayed();
		
		boolean expbuyerPriceBook_CustamizeDateRangeCheckBox   	=true;
		boolean expbuyerPriceBook_CustomizeVendorCheckBox   	=true;
		boolean expbuyerPriceBook_WarehouseCheckBox   			=true;
		boolean expbuyerPriceBook_QualityRangeCheckBox   		=true;
		boolean expbuyerPriceBook_CurrencyCheckBox   			=true;
		boolean expbuyerPriceBook_UnitCheckBox   				=true;
		
		System.out.println("**************************checkSellerPricebookSettingsOption**************************");
		
		System.out.println("buyerPriceBook_CustamizeDateRangeCheckBox  : "+actbuyerPriceBook_CustamizeDateRangeCheckBox +" Value Expected : "+expbuyerPriceBook_CustamizeDateRangeCheckBox);
		System.out.println("buyerPriceBook_CustomizeVendorCheckBox     : "+actbuyerPriceBook_CustomizeVendorCheckBox    +" Value Expected : "+expbuyerPriceBook_CustomizeVendorCheckBox);
		System.out.println("buyerPriceBook_WarehouseCheckBox           : "+actbuyerPriceBook_WarehouseCheckBox          +" Value Expected : "+expbuyerPriceBook_WarehouseCheckBox);
		System.out.println("buyerPriceBook_QualityRangeCheckBox        : "+actbuyerPriceBook_QualityRangeCheckBox       +" Value Expected : "+expbuyerPriceBook_QualityRangeCheckBox);
		System.out.println("buyerPriceBook_CurrencyCheckBox            : "+actbuyerPriceBook_CurrencyCheckBox           +" Value Expected : "+expbuyerPriceBook_CurrencyCheckBox);
		System.out.println("buyerPriceBook_UnitCheckBox                : "+actbuyerPriceBook_UnitCheckBox               +" Value Expected : "+expbuyerPriceBook_UnitCheckBox);
		
		if (actbuyerPriceBook_CustamizeDateRangeCheckBox==expbuyerPriceBook_CustamizeDateRangeCheckBox
				&& actbuyerPriceBook_CustomizeVendorCheckBox==expbuyerPriceBook_CustomizeVendorCheckBox
				&& actbuyerPriceBook_WarehouseCheckBox==expbuyerPriceBook_WarehouseCheckBox 
				&& actbuyerPriceBook_QualityRangeCheckBox==expbuyerPriceBook_QualityRangeCheckBox
				&& actbuyerPriceBook_CurrencyCheckBox==expbuyerPriceBook_CurrencyCheckBox
				&& actbuyerPriceBook_UnitCheckBox==expbuyerPriceBook_UnitCheckBox) 	
		{
			System.out.println("Pass: Customize Screen Displayed");
			excelReader.setCellData(xlfile, "Sheet1", 838, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: Customize Screen Displayed");
			excelReader.setCellData(xlfile, "Sheet1", 838, 9, resFail);
			return false;
		}
	}
	
	
	
	
	public boolean checkCustomizeColumnsInSellerPricebook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	  	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeDateRangeCheckBox));
		buyerPriceBook_CustamizeDateRangeCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustomizeVendorCheckBox));
		buyerPriceBook_CustomizeVendorCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_WarehouseCheckBox));
		buyerPriceBook_WarehouseCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_QualityRangeCheckBox));
		buyerPriceBook_QualityRangeCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CurrencyCheckBox));
		buyerPriceBook_CurrencyCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_UnitCheckBox));
		buyerPriceBook_UnitCheckBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
		buyerPriceBook_CustamizeSettingBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(beforeClickOnToggle));
		beforeClickOnToggle.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemColoumn));
		
		boolean actgridItemColoumn        =gridItemColoumn.isDisplayed();
		boolean actgridWareHouseColoumn   =gridWareHouseColoumn.isDisplayed();
		boolean actstartingdate   		  =startingdate.isDisplayed();
		boolean actendingdate             =endingdate.isDisplayed();
		boolean actcustomer               =customer.isDisplayed();
		boolean actcurrency               =currency.isDisplayed();
		
		boolean expgridItemColoumn        =true;
		boolean expgridWareHouseColoumn   =true;
		boolean expstartingdate   		  =true;
		boolean expendingdate             =true;
		boolean expcustomer               =true;
		boolean expcurrency               =true;
		
		System.out.println("*************************checkCustomizeColumnsInSellerPricebook*****************************");
		
		System.out.println("gridItemColoumn      : "+actgridItemColoumn      +" Value Expected : "+expgridItemColoumn);
		System.out.println("gridWareHouseColoumn : "+actgridWareHouseColoumn +" Value Expected : "+expgridWareHouseColoumn);
		System.out.println("startingdate         : "+actstartingdate         +" Value Expected : "+expstartingdate);
		System.out.println("endingdate           : "+actendingdate           +" Value Expected : "+expendingdate);
		System.out.println("customer             : "+actcustomer             +" Value Expected : "+expcustomer);
		System.out.println("currency             : "+actcurrency             +" Value Expected : "+expcurrency);
		
		
		if (actgridItemColoumn==expgridItemColoumn && actgridWareHouseColoumn==expgridWareHouseColoumn && actstartingdate==expstartingdate 
				&& actendingdate==expendingdate && actcustomer==expcustomer && actcurrency==expcurrency) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyFirstCellSelect));
			maxQtyFirstCellSelect.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
			maxQtyValue.click();
			maxQtyValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQty));
			
			boolean actminQty =minQty.isDisplayed();
			boolean actmaxQty =maxQty.isDisplayed();
			boolean actunit   =unit.isDisplayed();
			boolean actrate   =rate.isDisplayed();
			boolean actso     =so.isDisplayed();
			boolean actst     =st.isDisplayed();
			
			boolean expminQty =true;
			boolean expmaxQty =true;
			boolean expunit   =true;
			boolean exprate   =true;
			boolean expso     =true;
			boolean expst     =true;
			
			System.out.println("minQty  : "+actminQty+" Value Expected : "+expminQty);
			System.out.println("maxQty  : "+actmaxQty+" Value Expected : "+expmaxQty);
			System.out.println("unit    : "+actunit  +" Value Expected : "+expunit);
			System.out.println("rate    : "+actrate  +" Value Expected : "+exprate);
			System.out.println("so      : "+actso    +" Value Expected : "+expso);
			System.out.println("st      : "+actst    +" Value Expected : "+expst);
			
			if (actminQty==expminQty && actmaxQty==expmaxQty && actunit==expunit && actrate==exprate && actso==expso && actst==expst) 
			{	
				System.out.println("Pass: Customize Screen Displayed");
				excelReader.setCellData(xlfile, "Sheet1", 839, 9, resPass);
				return true;
			} 
			else 
			{
				System.out.println("Fail: Customize Screen Displayed");
				excelReader.setCellData(xlfile, "Sheet1", 839, 9, resFail);
				return false;
			}
		} 
		else 
		{
			System.out.println("Fail: Customize Screen Displayed");
			excelReader.setCellData(xlfile, "Sheet1", 839, 9, resFail);
			return false;
		}
	}
	
	
	
	public boolean checkCustomizeColumnsAfterInputofSellerPriceName() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	  	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitFirstCellSelect));
		unitFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys("Test Selling Price Hyd Wh");
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.click();
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemColoumn));
		
		boolean actgridItemColoumn        =gridItemColoumn.isDisplayed();
		boolean actgridWareHouseColoumn   =gridWareHouseColoumn.isDisplayed();
		boolean actstartingdate   		  =startingdate.isDisplayed();
		boolean actendingdate             =endingdate.isDisplayed();
		boolean actcustomer               =customer.isDisplayed();
		boolean actcurrency               =currency.isDisplayed();
		
		boolean expgridItemColoumn        =true;
		boolean expgridWareHouseColoumn   =true;
		boolean expstartingdate   		  =true;
		boolean expendingdate             =true;
		boolean expcustomer               =true;
		boolean expcurrency               =true;
		
		System.out.println("*************************checkCustomizeColumnsAfterInputofSellerPriceName*****************************");
		
		System.out.println("gridItemColoumn      : "+actgridItemColoumn      +" Value Expected : "+expgridItemColoumn);
		System.out.println("gridWareHouseColoumn : "+actgridWareHouseColoumn +" Value Expected : "+expgridWareHouseColoumn);
		System.out.println("startingdate         : "+actstartingdate         +" Value Expected : "+expstartingdate);
		System.out.println("endingdate           : "+actendingdate           +" Value Expected : "+expendingdate);
		System.out.println("customer             : "+actcustomer             +" Value Expected : "+expcustomer);
		System.out.println("currency             : "+actcurrency             +" Value Expected : "+expcurrency);
		
		
		if (actgridItemColoumn==expgridItemColoumn && actgridWareHouseColoumn==expgridWareHouseColoumn && actstartingdate==expstartingdate 
				&& actendingdate==expendingdate && actcustomer==expcustomer && actcurrency==expcurrency) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyFirstCellSelect));
			maxQtyFirstCellSelect.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
			maxQtyValue.click();
			maxQtyValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQty));
			
			boolean actminQty =minQty.isDisplayed();
			boolean actmaxQty =maxQty.isDisplayed();
			boolean actunit   =unit.isDisplayed();
			boolean actrate   =rate.isDisplayed();
			boolean actso     =so.isDisplayed();
			boolean actst     =st.isDisplayed();
			
			boolean expminQty =true;
			boolean expmaxQty =true;
			boolean expunit   =true;
			boolean exprate   =true;
			boolean expso     =true;
			boolean expst     =true;
			
			System.out.println("minQty  : "+actminQty+" Value Expected : "+expminQty);
			System.out.println("maxQty  : "+actmaxQty+" Value Expected : "+expmaxQty);
			System.out.println("unit    : "+actunit  +" Value Expected : "+expunit);
			System.out.println("rate    : "+actrate  +" Value Expected : "+exprate);
			System.out.println("so      : "+actso    +" Value Expected : "+expso);
			System.out.println("st      : "+actst    +" Value Expected : "+expst);
			
			if (actminQty==expminQty && actmaxQty==expmaxQty && actunit==expunit && actrate==exprate && actso==expso && actst==expst) 
			{
				System.out.println("Pass: Customize Screen Displayed");
				excelReader.setCellData(xlfile, "Sheet1", 840, 9, resPass);
				return true;
			} 
			else 
			{
				System.out.println("Fail: Customize Screen Displayed");
				excelReader.setCellData(xlfile, "Sheet1", 840, 9, resFail);
				return false;
			}
		} 
		else 
		{
			System.out.println("Fail: Customize Screen Displayed");
			excelReader.setCellData(xlfile, "Sheet1", 840, 9, resFail);
			return false;
		}
	}
	
	
	public boolean checkMandatoryFieldsInSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	  			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.clear();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		String expMessage="Enter price book name";
		
		System.out.println("*********************checkMandatoryFieldsInSellerPriceBook****************************");
		
		String actMessage=checkValidationMessage(expMessage);
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			System.out.println("Pass: displayed Mandatory Message as Enter Name of the Price Book");
			excelReader.setCellData(xlfile, "Sheet1", 841, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: displayed Mandatory Message as Enter Name of the Price Book");
			excelReader.setCellData(xlfile, "Sheet1", 841, 9, resFail);
			return false;
		}
	}
	
	
	public boolean checkProductMandatory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys("Test Selling Price Hyd Wh");
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		String expMessage="Select Item";
		
		System.out.println("***************************checkProductMandatory********************************");
		
		String actMessage=checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) 
	   {
			System.out.println("Pass: displayed message as Select Product");
            excelReader.setCellData(xlfile, "Sheet1", 842, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail:displayed message as Select Product");
            excelReader.setCellData(xlfile, "Sheet1", 842, 9, resFail);
			return false;
		}
	}
	
	
	public boolean checkSellerPriceBookRate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.END);
		itemValue.sendKeys(Keys.SHIFT,Keys.HOME);
		itemValue.sendKeys(Keys.SPACE);
		
		int count = itemComboBoxList.size();
		   
		for(int i=0;i<count;i++)
		{
			String data = itemComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("ITEMS GROUP"))
			{
				itemComboBoxList.get(i).click();
				break;
			}
		}
		   
		itemValue.sendKeys(Keys.TAB);
		
		//Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		String expMessage="Enter Rate In Row Number 1";
		
		System.out.println("**********************checkSellerPriceBookRate******************************");
		
		String actMessage=checkValidationMessage(expMessage);
	    
	    if (actMessage.equalsIgnoreCase(expMessage)) 
	    {
			System.out.println("Pass: displayed message as Select Product");
            excelReader.setCellData(xlfile, "Sheet1", 843, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail:displayed message as Select Product");
             excelReader.setCellData(xlfile, "Sheet1", 843, 9, resFail);
			return false;
		}
	}
	
	
	
	public boolean checkSaveSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseFirstCellSelect));
		wareHouseFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
		wareHouseValue.click();
		wareHouseValue.sendKeys(Keys.END);
		wareHouseValue.sendKeys(Keys.SHIFT,Keys.HOME);
		wareHouseValue.sendKeys(Keys.SPACE);
		
		int warehousecount = warehouseComboBoxList.size();
		   
		for(int i=0;i<warehousecount;i++)
		{
			String data = warehouseComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("HYDERABAD"))
			{
				warehouseComboBoxList.get(i).click();
				break;
			}
		}
		   
		wareHouseValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingDateValue));
		startingDateValue.click();
		startingDateValue.sendKeys(Keys.TAB);
		

		if(endingDateChkBox.isSelected()==true)
		{
			endingDateChkBox.click();
		}
		
		if(endingDateChkBox.isSelected()==false)
		{
			endingDateChkBox.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(endingDateValue));
		endingDateValue.click();
		endingDateValue.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
		vendorValue.click();
		vendorValue.sendKeys(Keys.END);
		vendorValue.sendKeys(Keys.SHIFT,Keys.HOME);
		vendorValue.sendKeys(Keys.SPACE);
		
		int vendorcount1 = vendorComboBoxList.size();
		   
		for(int i=0;i<vendorcount1;i++)
		{
			String data = vendorComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Customer B"))
			{
				vendorComboBoxList.get(i).click();
				break;
			}
		}
		   
		vendorValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyOptionValue));
		currencyOptionValue.click();
		currencyOptionValue.sendKeys(Keys.END);
		currencyOptionValue.sendKeys(Keys.SHIFT,Keys.HOME);
		currencyOptionValue.sendKeys("I");
		
		int currencycount1 = currencyComboBoxList.size();
		   
		for(int i=0;i<currencycount1;i++)
		{
			String data = currencyComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Indian Rupees"))
			{
				currencyComboBoxList.get(i).click();
				break;
			}
		}
		   
		currencyOptionValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
		minQtyValue.click();
		minQtyValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
		maxQtyValue.click();
		maxQtyValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateFirstCellSelect));
		rateFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("100");
		rateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PoValue));
		PoValue.sendKeys("101");
		PoValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();

		String expMessage="Price Book created successfully";
		
		System.out.println("**************************checkSaveSellerPriceBook*****************************");
		
		String actMessage=checkValidationMessage(expMessage);
	    
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			System.out.println("Pass: Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");
			excelReader.setCellData(xlfile, "Sheet1", 844, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail:  Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");
			excelReader.setCellData(xlfile, "Sheet1", 844, 9, resFail);
			return false;
		}	
	}
	

	public boolean checkEditSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(afterClickOnToggle));
		afterClickOnToggle.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.END);
		buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		
		int pricebookcount = pricebookComboBoxList.size();
		   
		for(int i=0;i<pricebookcount;i++)
		{
			String data = pricebookComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Test Selling Price Hyd Wh"))
			{
				pricebookComboBoxList.get(i).click();
				break;
			}
		}
		   
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.click();
		 
		String actbuyerPriceBook_PriceBook =buyerPriceBook_PriceBook.getAttribute("value");
		String expbuyerPriceBook_PriceBook ="Test Selling Price Hyd Wh";
		
		System.out.println("*****************************checkEditSellerPriceBook*******************************");
		
		System.out.println("buyerPriceBook_PriceBook : "+actbuyerPriceBook_PriceBook+" Value Expected : "+expbuyerPriceBook_PriceBook);
	    
		if(actbuyerPriceBook_PriceBook.equalsIgnoreCase(expbuyerPriceBook_PriceBook)) 
		{
			System.out.println("Pass: Edit Buyer PriceBook");
			excelReader.setCellData(xlfile, "Sheet1", 845, 9, resPass);
	        return true;
		 } 
		 else 
		 {
			System.out.println("Fail: Edit Buyer PriceBook");
			excelReader.setCellData(xlfile, "Sheet1", 845, 9, resFail);
	        return false;
		 }
	}	
	
	

	public boolean checkSavedSellerRatesInformation() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
	    buyerPriceBookFilter_FilterAndLoadBtn.click();
	 
	    //Thread.sleep(2000);
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(beforeClickOnToggle));
	    beforeClickOnToggle.click();
	    
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitFirstCellSelect));
	    unitFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		
		String actitemValue =itemValue.getAttribute("value");
	    String expitemValue ="ITEMS GROUP";    
	    
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseFirstCellSelect));
	    wareHouseFirstCellSelect.click();
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
		wareHouseValue.click();
	
		String actwareHouseValue =wareHouseValue.getAttribute("value");
	    String expwareHouseValue ="HYDERABAD";
		
	    System.out.println("*****************************checkSavedSellerRatesInformation*****************************");
	    
	    System.out.println("itemValue      : "+actitemValue      +" Value Expected : "+expitemValue);
	    System.out.println("wareHouseValue : "+actwareHouseValue +" Value Expected : "+expwareHouseValue);
	    
	    if (actitemValue.equalsIgnoreCase(expitemValue) && actwareHouseValue.equalsIgnoreCase(expwareHouseValue)) 
	    {
			System.out.println("Pass : Allow to select and HYD WH should be display in the Warehouse Combo box");
			excelReader.setCellData(xlfile, "Sheet1", 846, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: Allow to select and HYD WH should be display in the Warehouse Combo box");
			excelReader.setCellData(xlfile, "Sheet1", 846, 9, resFail);
			return false;
		}
	}
	
	
	public boolean checkUpdateSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.END);
		itemValue.sendKeys(Keys.SHIFT,Keys.HOME);
		itemValue.sendKeys(Keys.SPACE);
		
		int itemcount = itemComboBoxList.size();
		   
		for(int i=0;i<itemcount;i++)
		{
			String data = itemComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("BATCH BR ITEM"))
			{
				itemComboBoxList.get(i).click();
				break;
			}
		}
		   
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
		wareHouseValue.click();
		wareHouseValue.sendKeys(Keys.END);
		wareHouseValue.sendKeys(Keys.SHIFT,Keys.HOME);
		wareHouseValue.sendKeys(Keys.SPACE);
		
		int warehousecount = warehouseComboBoxList.size();
		   
		for(int i=0;i<warehousecount;i++)
		{
			String data = warehouseComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("HYDERABAD"))
			{
				warehouseComboBoxList.get(i).click();
				break;
			}
		}
		   
		wareHouseValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingDateValue));
		startingDateValue.click();
		startingDateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(endingDateValue));
		endingDateValue.click();
		endingDateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
		vendorValue.sendKeys(Keys.TAB);
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyOptionValue));
		currencyOptionValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
		minQtyValue.click();
		minQtyValue.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
		maxQtyValue.click();
		maxQtyValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitValue));
		unitValue.click();
		unitValue.sendKeys(Keys.END);
		unitValue.sendKeys(Keys.SHIFT,Keys.HOME);
		unitValue.sendKeys(Keys.SPACE);
		
		int unitcount1 = unitComboBoxList.size();
		   
		for(int i=0;i<unitcount1;i++)
		{
			String data = unitComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Pcs"))
			{
				unitComboBoxList.get(i).click();
				break;
			}
		}
		   
		unitValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("200");
		rateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PoValue));
		PoValue.sendKeys("202");
		PoValue.sendKeys(Keys.TAB);
		
		//Input in Third
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitSecondCellSelect));
		unitSecondCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
		itemThirdCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.END);
		itemValue.sendKeys(Keys.SHIFT,Keys.HOME);
		itemValue.sendKeys(Keys.SPACE);
		
		int itemcount1 = itemComboBoxList.size();
		   
		for(int i=0;i<itemcount1;i++)
		{
			String data = itemComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("BIN RAW MATERIAL ITEM"))
			{
				itemComboBoxList.get(i).click();
				break;
			}
		}
		   
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
		wareHouseValue.click();
		wareHouseValue.sendKeys(Keys.END);
		wareHouseValue.sendKeys(Keys.SHIFT,Keys.HOME);
		wareHouseValue.sendKeys(Keys.SPACE);
		
		int warehousecount1 = warehouseComboBoxList.size();
		   
		for(int i=0;i<warehousecount1;i++)
		{
			String data = warehouseComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("SECUNDERABAD"))
			{
				warehouseComboBoxList.get(i).click();
				break;
			}
		}
		   
		wareHouseValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingDateValue));
		startingDateValue.click();
		startingDateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(endingDateValue));
		endingDateValue.click();
		endingDateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
		vendorValue.click();
		vendorValue.sendKeys(Keys.END);
		vendorValue.sendKeys(Keys.SHIFT,Keys.HOME);
		vendorValue.sendKeys(Keys.SPACE);
		
		int vendorcount1 = vendorComboBoxList.size();
		   
		for(int i=0;i<vendorcount1;i++)
		{
			String data = vendorComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Customer C"))
			{
				vendorComboBoxList.get(i).click();
				break;
			}
		}
		   
		vendorValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyOptionValue));
		currencyOptionValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
		minQtyValue.click();
		minQtyValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
		maxQtyValue.click();
		maxQtyValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitValue));
		unitValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("300");
		rateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PoValue));
		PoValue.sendKeys("303");
		PoValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		String expMessage="Price Book updated successfully";
		
		System.out.println("*************************checkUpdateSellerPriceBook*****************************");
		
		String actMessage=checkValidationMessage(expMessage);
	    
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			System.out.println("Pass: Update Buyer PriceBook and Saved Successfully");
			excelReader.setCellData(xlfile, "Sheet1", 847, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail:   Update Buyer PriceBook and Saved Successfully");
			excelReader.setCellData(xlfile, "Sheet1", 847, 9, resFail);
			return false;
		}
	}
	
	
	public boolean checkCloseInSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCloseBtn));
		buyerPriceBookCloseBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
		
		boolean actlabelDashboard                  =labelDashboard.isDisplayed();
		boolean actselectDashboard                 =selectDashboard.isDisplayed();
		boolean actnewAddDashBoard                 =newAddDashBoard.isDisplayed();
		boolean actdashboardCustomizationSettings  =dashboardCustomizationSettings.isDisplayed();
		
		boolean explabelDashboard                  =true;
		boolean expselectDashboard                 =true;
		boolean expnewAddDashBoard                 =true;
		boolean expdashboardCustomizationSettings  =true;
		
		System.out.println("******************************checkCloseInSellerPriceBook************************");
		
		System.out.println("labelDashboard                  : "+actlabelDashboard                 +" Value Expected : "+explabelDashboard);
		System.out.println("selectDashboard                 : "+actselectDashboard                +" Value Expected : "+expselectDashboard);
		System.out.println("newAddDashBoard                 : "+actnewAddDashBoard                +" Value Expected : "+expnewAddDashBoard);
		System.out.println("dashboardCustomizationSettings  : "+actdashboardCustomizationSettings +" Value Expected : "+expdashboardCustomizationSettings);
		
		if(labelDashboard.isDisplayed() && selectDashboard.isDisplayed() && newAddDashBoard.isDisplayed() && dashboardCustomizationSettings.isDisplayed())
		{
			System.out.println("Pass: Close Button in Buyer Price Book Screen");
			excelReader.setCellData(xlfile, "Sheet1", 848, 9, resPass);
			return true;
		}
		else 
		{
			System.out.println("Fail: Close Button in Buyer Price Book Screen");
			excelReader.setCellData(xlfile, "Sheet1", 848, 9, resFail);
			return false;
		}
	}
	
	
	
	
	public boolean checkCopyClearPasteDeleteOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	  	homeMenu.click();
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterMenu));
	  	homeMasterMenu.click();
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	  	homeMasterItemMenu.click();
		
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemSellerPriceBookMenu));
	  	homeMasterItemSellerPriceBookMenu.click();
		
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.END);
		buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		
		int pricebookcount = pricebookComboBoxList.size();
		   
		for(int i=0;i<pricebookcount;i++)
		{
			String data = pricebookComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Test Selling Price Hyd Wh"))
			{
				pricebookComboBoxList.get(i).click();
				break;
			}
		}
		   
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	  	buyerPriceBook_AbbreviationBtn.click();
		 
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
	  	buyerPriceBookFilter_FilterAndLoadBtn.click();
	  	
	  	//Thread.sleep(2000);
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCopyBtn));
	  	buyerPriceBookCopyBtn.click();			 
			
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookClearBtn));
	  	buyerPriceBookClearBtn.click();
			
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	  	buyerPriceBook_PriceBook.click();
	  	buyerPriceBook_PriceBook.sendKeys("Test Pricebook to Delete");
	  	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		 
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookPasteBtn));
	  	buyerPriceBookPasteBtn.click();
		 
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	  	itemSecondCellSelect.click();

	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.END);
		itemValue.sendKeys(Keys.SHIFT,Keys.HOME);
		itemValue.sendKeys(Keys.SPACE);
		
		int count = itemComboBoxList.size();
		   
		for(int i=0;i<count;i++)
		{
			String data = itemComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("BATCH FIFO ITEM"))
			{
				itemComboBoxList.get(i).click();
				break;
			}
		}
		   
		itemValue.sendKeys(Keys.TAB);
			
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
	  	wareHouseValue.sendKeys(Keys.TAB);
					
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingDateValue));
	  	startingDateValue.click();
	  	startingDateValue.sendKeys(Keys.TAB);
			
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(endingDateValue));
	  	endingDateValue.click();
	  	endingDateValue.sendKeys(Keys.TAB);
			
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
	  	vendorValue.click();
	  	vendorValue.sendKeys(Keys.TAB);
		
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyOptionValue));
	  	currencyOptionValue.click();
	  	currencyOptionValue.sendKeys(Keys.TAB);
			
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
	  	minQtyValue.sendKeys("9");
	  	minQtyValue.sendKeys(Keys.TAB);
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
	  	maxQtyValue.sendKeys("39");
	  	maxQtyValue.sendKeys(Keys.TAB);
			
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitValue));
	  	unitValue.click();
	  	unitValue.sendKeys(Keys.TAB);
			
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	  	rateValue.sendKeys("20");
	  	rateValue.sendKeys(Keys.TAB);
			
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PoValue));
	  	PoValue.sendKeys("29");
	  	PoValue.sendKeys(Keys.TAB);

	  	getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
	
		String expMessage="Price Book created successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.END);
		buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		
		int pricebookcount1 = pricebookComboBoxList.size();
		   
		for(int i=0;i<pricebookcount1;i++)
		{
			String data = pricebookComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Test Pricebook to Delete"))
			{
				pricebookComboBoxList.get(i).click();
				break;
			}
		}
		   
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
		buyerPriceBookFilter_FilterAndLoadBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookDeleteBtn));
		buyerPriceBookDeleteBtn.click();
		
		getWaitForAlert();
		 
		String actAlertMessage =getAlert().getText();
		String expAlertMessage ="Do you want to delete Total Price book?";
		 
		System.out.println("*************************************checkCopyClearPasteDeleteOption**********************************");
		
		System.out.println("AlertMessage  : "+actAlertMessage+" Value Expected : "+expAlertMessage);
		
		if (actAlertMessage.equalsIgnoreCase(expAlertMessage)) 
		{
			System.out.println("Pass: Created price Book from Dropdown and Click on Delete button");
			excelReader.setCellData(xlfile, "Sheet1", 849, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: Created price Book from Dropdown and Click on Delete button");
			excelReader.setCellData(xlfile, "Sheet1", 849, 9, resFail);
			return false;
		}
	}
	
	
	
	public boolean checkSellerPricebookNoOptionOnClickOnDelete() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getAlert().dismiss();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(buyerPriceBook_PriceBook));
		String actbuyerPriceBook_PriceBook =buyerPriceBook_PriceBook.getAttribute("value");
	    String expbuyerPriceBook_PriceBook ="Test Pricebook to Delete";
	    
	    System.out.println("*****************************checkSellerPricebookNoOptionOnClickOnDelete****************************");
	    
	    System.out.println("buyerPriceBook_PriceBook : "+actbuyerPriceBook_PriceBook+" Value Expected : "+expbuyerPriceBook_PriceBook);
	    
	    if (actbuyerPriceBook_PriceBook.equalsIgnoreCase(expbuyerPriceBook_PriceBook)) 
	    {
			System.out.println("Pass:No Button in Delete option");
			 excelReader.setCellData(xlfile, "Sheet1", 850, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: No Button in Delete option");
			 excelReader.setCellData(xlfile, "Sheet1", 850, 9, resFail);
			return false;
		}
	}
	
	
	public boolean checkSellerPricebookYesOptionOnClickDelete() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookDeleteBtn));
		buyerPriceBookDeleteBtn.click();
	
		getAlert().accept();
	
		String expMessage="Pricebook details deleted successfully.";
		
		System.out.println("***********************checkSellerPricebookYesOptionOnClickDelete************************");
		
		String actMessage=checkValidationMessage(expMessage);
	    
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			System.out.println("Pass: Yes Button in Delete option ");
			excelReader.setCellData(xlfile, "Sheet1", 851, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail: Yes Button in Delete option ");
		    excelReader.setCellData(xlfile, "Sheet1", 851, 9, resFail);
			return false;
		}
	}
	
	
	
	public boolean checkSellerPricebookWithInactive() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.END);
		buyerPriceBook_PriceBook.sendKeys(Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		
		int pricebookcount = pricebookComboBoxList.size();
		   
		for(int i=0;i<pricebookcount;i++)
		{
			String data = pricebookComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("Test Selling Price Hyd Wh"))
			{
				pricebookComboBoxList.get(i).click();
				break;
			}
		}
		   
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.click();
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
		buyerPriceBookFilter_FilterAndLoadBtn.click();
		
		//Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCopyBtn));
		buyerPriceBookCopyBtn.click();			 
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookClearBtn));
		buyerPriceBookClearBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys("Test Pricebook to Inactive");
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookPasteBtn));
		buyerPriceBookPasteBtn.click();
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.END);
		itemValue.sendKeys(Keys.SHIFT,Keys.HOME);
		itemValue.sendKeys(Keys.SPACE);
		
		int count = itemComboBoxList.size();
		   
		for(int i=0;i<count;i++)
		{
			String data = itemComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("BATCH FIFO ITEM"))
			{
				itemComboBoxList.get(i).click();
				break;
			}
		}
		   
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
		wareHouseValue.click();
		wareHouseValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingDateValue));
		startingDateValue.click();
		startingDateValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(endingDateValue));
		endingDateValue.click();
		endingDateValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
		vendorValue.click();
		vendorValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyOptionValue));
		currencyOptionValue.click();
		currencyOptionValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
		minQtyValue.sendKeys("89");
		minQtyValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
		maxQtyValue.sendKeys("99");
		maxQtyValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitValue));
		unitValue.click();
		unitValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("90");
		rateValue.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PoValue));
		PoValue.sendKeys("99");
		PoValue.sendKeys(Keys.TAB);
		 
		buyerPriceBookInActiveCheckBox.click();
			 
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		String expMessage="Price Book created successfully";
		
		System.out.println("********************************checkSellerPricebookWithInactive*****************************");
		
		String actMessage=checkValidationMessage(expMessage);
	    
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			System.out.println("Pass: Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");
			excelReader.setCellData(xlfile, "Sheet1", 852, 9, resPass);
			return true;
		} 
		else 
		{
			System.out.println("Fail:  Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully");
			excelReader.setCellData(xlfile, "Sheet1", 852, 9, resFail);
			return false;
		}	
	}
	
	@FindBy (xpath="//*[@id='25']/span")
    private static WebElement dataMangementMenu;
	
		@FindBy (xpath="//*[@id='26']/span")
		private static WebElement backup;
		
		@FindBy (xpath="//input[@id='txtBackUpFile']")
		private static WebElement backupTXT;
		
		@FindBy (xpath="//div[@id='BackUpControls']//div[1]//div[1]")
		private static WebElement backupBtn;
		
		@FindBy (xpath="//div[contains(text(),'Cancel')]")
		private static WebElement cancelBtn;
		
	public boolean checkCloseOptionAfterSavingInactiveInSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCloseBtn));
		buyerPriceBookCloseBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
		
		boolean actlabelDashboard                  =labelDashboard.isDisplayed();
		boolean actselectDashboard                 =selectDashboard.isDisplayed();
		boolean actnewAddDashBoard                 =newAddDashBoard.isDisplayed();
		boolean actdashboardCustomizationSettings  =dashboardCustomizationSettings.isDisplayed();
		
		boolean explabelDashboard                  =true;
		boolean expselectDashboard                 =true;
		boolean expnewAddDashBoard                 =true;
		boolean expdashboardCustomizationSettings  =true;
		
		
		System.out.println("******************************checkCloseOptionAfterSavingInactiveInSellerPriceBook************************");
		
		System.out.println("labelDashboard                  : "+actlabelDashboard                 +" Value Expected : "+explabelDashboard);
		System.out.println("selectDashboard                 : "+actselectDashboard                +" Value Expected : "+expselectDashboard);
		System.out.println("newAddDashBoard                 : "+actnewAddDashBoard                +" Value Expected : "+expnewAddDashBoard);
		System.out.println("dashboardCustomizationSettings  : "+actdashboardCustomizationSettings +" Value Expected : "+expdashboardCustomizationSettings);
		
		if(labelDashboard.isDisplayed() && selectDashboard.isDisplayed() && newAddDashBoard.isDisplayed() && dashboardCustomizationSettings.isDisplayed())
		{
			System.out.println("Pass: Close Button in Buyer Price Book Screen");
			excelReader.setCellData(xlfile, "Sheet1", 853, 9, resPass);
			
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
			userNameDisplay.click();
			  
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			logoutOption.click();
			 
			return true;
		}
		else 
		{
			System.out.println("Fail: Close Button in Buyer Price Book Screen");
			excelReader.setCellData(xlfile, "Sheet1", 853, 9, resFail);
			
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
			userNameDisplay.click();
			  
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			logoutOption.click();
			 
			return false;
		}
	}
	
	//--------------------------------------------------------------------------------------------------------------------------
	
	// Import Excel Option in Advance option is not working and thats the reason there are some known pending in the Pricebook
	
	//--------------------------------------------------------------------------------------------------------------------------

	public boolean checkCustomizeDisplayInSelletPriceBookScreen() throws EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
		buyerPriceBook_CustamizeSettingBtn.click();
		
		if (buyerPriceBook_CustamizeDateRangeCheckBox.isDisplayed() && buyerPriceBook_CustomizeVendorCheckBox.isDisplayed()
				&& buyerPriceBook_WarehouseCheckBox.isDisplayed()&&buyerPriceBook_QualityRangeCheckBox.isDisplayed()
				&&buyerPriceBook_CurrencyCheckBox.isDisplayed()&& buyerPriceBook_UnitCheckBox.isDisplayed()) 
		{
			
			System.out.println("Pass: SellerPriceBook Screen Customize Screen Displayed");
			buyerPriceBook_WarehouseCheckBox.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
			buyerPriceBook_CustamizeSettingBtn.click();
			excelReader.setCellData(xlfile, "Sheet1", 838, 9, resPass);
			return true;
			
		} else {

			System.out.println("Fail: SellerPriceBook Screen Customize Screen Displayed");
			buyerPriceBook_WarehouseCheckBox.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
			buyerPriceBook_CustamizeSettingBtn.click();
			excelReader.setCellData(xlfile, "Sheet1", 838, 9, resFail);
			return false;

		}
		
		
	}
	
	
	

	
public boolean checkMandatoryFieldsInSelletPriceBookScreen() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	
	String expMessage="Enter price book name";
	
	String actMessage=checkValidationMessage(expMessage);
	
   if (actMessage.equalsIgnoreCase(expMessage)) 
   {
		System.out.println("Pass: displayed Mandatory Message as Enter Name of the Price Book in SellerPriceBook Screen");
		return true;
	} 
	else 
	{
		System.out.println("Fail: displayed Mandatory Message as Enter Name of the Price Book in SellerPriceBook Screen");
		return false;
	}
	
	
}	
	
	
	
public boolean CheckInputNameasTestBuyingPriceHydWhandclickonSavebuttonOFSellerScreen() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys("Test Buying Price Hyd Wh");
	//Thread.sleep(3000);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	
	String expMessage="Select Item";
	
	System.out.println("**************************CheckInputNameasTestBuyingPriceHydWhandclickonSavebuttonOFSellerScreen**********************");
	
	String actMessage=checkValidationMessage(expMessage);
    
   if (actMessage.equalsIgnoreCase(expMessage)) 
   {
		System.out.println("Pass: displayed message as Select Product in SellerPriceBook Screen");
		return true;
	} 
	else 
	{
		System.out.println("Fail:displayed message as Select Product in SellerPriceBook Screen");
		return false;
	}
}
	
	
	/*
	
public boolean checkbyinputdataintowarehouseAndSelectingProductsinCumboboxinSellerPriceScrenn() throws InterruptedException
{
   
    
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
	buyerPriceBook_CustamizeSettingBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_WarehouseCheckBox));
    buyerPriceBook_WarehouseCheckBox.click();
	  
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
	buyerPriceBook_CustamizeSettingBtn.click();
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRow));
	gridItemFirstRow.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRowafterSelecting));
	gridItemFirstRowafterSelecting.click();
	gridItemFirstRowafterSelecting.sendKeys("Group of RMA LIFO and FIFO");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseFirstRow));
	gridWarehouseFirstRow.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseFirstRowAfterSelecting));
	gridWarehouseFirstRowAfterSelecting.click();
	gridWarehouseFirstRowAfterSelecting.sendKeys("HYDERABAD");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateFirstRow));
	gridRateFirstRow.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateFirstRowAfterSelecting));
	gridRateFirstRowAfterSelecting.sendKeys("100");
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemSecondRow));
	gridItemSecondRow.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRowafterSelecting));
	gridItemFirstRowafterSelecting.click();
	gridItemFirstRowafterSelecting.sendKeys("Bin-BR");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseSecondRow));
	gridWarehouseSecondRow.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseFirstRowAfterSelecting));
	gridWarehouseFirstRowAfterSelecting.click();
	gridWarehouseFirstRowAfterSelecting.sendKeys("MUMBAI");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateSecondRow));
	gridRateSecondRow.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateFirstRowAfterSelecting));
	gridRateFirstRowAfterSelecting.sendKeys("101");
	
	
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemThirdRow));
	gridItemThirdRow.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRowafterSelecting));
	gridItemFirstRowafterSelecting.click();
	gridItemFirstRowafterSelecting.sendKeys("Batch-FIFO");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseThirdRow));
	gridWarehouseThirdRow.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseFirstRowAfterSelecting));
	gridWarehouseFirstRowAfterSelecting.click();
	gridWarehouseFirstRowAfterSelecting.sendKeys("SECUNDERABAD");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateThirdRow));
	gridRateThirdRow.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateFirstRowAfterSelecting));
	gridRateFirstRowAfterSelecting.sendKeys("102");
	
	

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFourthRow));
	gridItemFourthRow.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRowafterSelecting));
	gridItemFirstRowafterSelecting.click();
	gridItemFirstRowafterSelecting.sendKeys("AddGroup");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseFourthRow));
	gridWarehouseFourthRow.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseFirstRowAfterSelecting));
	gridWarehouseFirstRowAfterSelecting.click();
	gridWarehouseFirstRowAfterSelecting.sendKeys("WH1");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateFourthRow));
	gridRateFourthRow.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateFirstRowAfterSelecting));
	gridRateFirstRowAfterSelecting.sendKeys("103");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	

	String Actdata=newErrorMessage.getText();
    String Expdata="Price Book created successfully";
    
   if (Actdata.equalsIgnoreCase(Expdata)) 
   {
		System.out.println("Pass: Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully in SellerPriceBook Screen");
		newErrorMessageCloseBtn.click();
		return true;
	} 
	else 
	{
		System.out.println("Fail:  Input data in to warehouse And Selecting Products in Cumbobox and Saved Successfully in SellerPriceBook Screen");
		newErrorMessageCloseBtn.click();
		return false;
	}
	

}

	
	
	public boolean checkEditSellerPriceBook() throws InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	  	homeMenu.click();
	  	
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterMenu));
	  	homeMasterMenu.click();
	  
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	  	homeMasterItemMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemSellerPriceBookMenu));
		homeMasterItemSellerPriceBookMenu.click();
	  	
		//Thread.sleep(2000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		 buyerPriceBook_PriceBook.click();
		 buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		 buyerPriceBook_AbbreviationBtn.click();
		 
		 //Thread.sleep(2000);
		 
			String Actdata=buyerPriceBook_PriceBook.getAttribute("value");
		    String Expdata="Test Buying Price Hyd Wh";
		    
		   if (Actdata.equalsIgnoreCase(Expdata)) 
		   {
				System.out.println("Pass: Edit Buyer SellerPriceBook Screen");
		
				return true;
			} 
			else 
			{
				System.out.println("Fail: Edit Buyer SellerPriceBook Screen");
			
				return false;
			}
		 
	}
	
	
	
	
	
	
	
	
	

	public boolean SelectthePriceBookNameInSellerPriceBook() throws InterruptedException
	{
		//Thread.sleep(2000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
		 buyerPriceBookFilter_FilterAndLoadBtn.click();
		
		 //Thread.sleep(2000);
		 
		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRow));
			gridItemFirstRow.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRowafterSelecting));
			gridItemFirstRowafterSelecting.click();
			
	
	
			String Actdata=gridItemFirstRowafterSelecting.getAttribute("value");
		    String Expdata="Group of RMA LIFO and FIFO";
		    
		    
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseFirstRow));
		    gridWarehouseFirstRow.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseFirstRowAfterSelecting));
			gridWarehouseFirstRowAfterSelecting.click();
		
		
			String Actdata1=gridWarehouseFirstRowAfterSelecting.getAttribute("value");
		    String Expdata1="HYDERABAD";
			
		    
		   if (Actdata.equalsIgnoreCase(Expdata) &&Actdata1.equalsIgnoreCase(Expdata1)) 
		   {
				System.out.println("Pass : Allow to select and HYD WH should be display in the Warehouse Combo box and Product ");
		
				return true;
			} 
			else 
			{
				System.out.println("Fail: Allow to select and HYD WH should be display in the Warehouse Combo box and Product");
			
				return false;
			}
	}
	
	
	
	public boolean checkUpdateSellerPriceBook() throws InterruptedException
	{
		//Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemSecondRow));
		gridItemSecondRow.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRowafterSelecting));
		gridItemFirstRowafterSelecting.click();
		
		//Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		

		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(newErrorMessage));
		String Actdata=newErrorMessage.getText();
	    String Expdata="Price Book updated successfully";
	    
	   if (Actdata.equalsIgnoreCase(Expdata)) 
	   {
			System.out.println("Pass: Update Seller PriceBook and Saved Successfully");
			newErrorMessageCloseBtn.click();
			return true;
		} 
		else 
		{
			System.out.println("Fail:   Update Seller PriceBook and Saved Successfully");
			newErrorMessageCloseBtn.click();
			return false;
		}
		
	}
	
	
	
	
	
	
	
	public boolean CheckCloseinSellerPriceBook() throws InterruptedException
	{
		//Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCloseBtn));
		buyerPriceBookCloseBtn.click();
		
		if (dashBoardLabel.isDisplayed()) {
			System.out.println("Pass: Close Button in Seller Price Book Screen");
			return true;
		} else {
			System.out.println("Fail: Close Button in Seller Price Book Screen");
			return false;
		}
		
		
	}
	
	
	
	
	
	public boolean checkUpdatedProductINSellerPriceBook() throws InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	  	homeMenu.click();
	  	
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterMenu));
	  	homeMasterMenu.click();
	  
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	  	homeMasterItemMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemSellerPriceBookMenu));
		homeMasterItemSellerPriceBookMenu.click();
	  	
		//Thread.sleep(2000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		 buyerPriceBook_PriceBook.click();
		 buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		 buyerPriceBook_AbbreviationBtn.click();
		 //Thread.sleep(2000);
		 
			String Actdata=buyerPriceBook_PriceBook.getAttribute("value");
		    String Expdata="Test Buying Price Hyd Wh";
		    
		   if (Actdata.equalsIgnoreCase(Expdata)) 
		   {
				System.out.println("Pass: Edit Seller PriceBook");
		
				return true;
			} 
			else 
			{
				System.out.println("Fail: Edit Seller PriceBook");
			
				return false;
			}
		 
		 
	      }
	
	
	public boolean checktheWarehouseandProductNamewhileupdatingproductINSellerPriceBook()
	{
		
		buyerPriceBookFilter_FilterAndLoadBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemThirdRow));
		gridItemThirdRow.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRowafterSelecting));
		gridItemFirstRowafterSelecting.click();
	
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseThirdRow));
		gridWarehouseThirdRow.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseFirstRowAfterSelecting));
		gridWarehouseFirstRowAfterSelecting.click();
		
		String Actdata1=gridItemFirstRowafterSelecting.getAttribute("value");
	    String Expdata1="Batch-FIFO";
		
		
		String Actdata=gridWarehouseFirstRowAfterSelecting.getAttribute("value");
	    String Expdata="SECUNDERABAD";
	    
	   if (Actdata.equalsIgnoreCase(Expdata) && Actdata1.equalsIgnoreCase(Expdata1) ) 
	   {
			System.out.println("Pass: Warehouse and Product Name while updating product in Seller PriceBook");
	
			return true;
		} 
		else 
		{
			System.out.println("Fail:  Warehouse and Product Name while updating product in Seller PriceBook");
		
			return false;
		}
	 
	}
	
	
	public boolean checkbyaddingtheextraItemAndRateINSellerPriceBook()
	{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFifthRow));
		gridItemFifthRow.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRowafterSelecting));
		gridItemFirstRowafterSelecting.click();
		gridItemFirstRowafterSelecting.sendKeys("Batch-WA");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseFifthRow));
		gridWarehouseFifthRow.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseFirstRowAfterSelecting));
		gridWarehouseFirstRowAfterSelecting.click();
		gridWarehouseFirstRowAfterSelecting.sendKeys("WHG");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateFifthRow));
		gridRateFifthRow.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateFirstRowAfterSelecting));
		gridRateFirstRowAfterSelecting.sendKeys("104");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(newErrorMessage));
		String Actdata=newErrorMessage.getText();
	    String Expdata="Price Book updated successfully";
	    
	   if (Actdata.equalsIgnoreCase(Expdata)) 
	   {
			System.out.println("Pass: Update  Seller PriceBook and Saved Successfully");
			newErrorMessageCloseBtn.click();
			return true;
		} 
		else 
		{
			System.out.println("Fail:   Update  Seller PriceBook and Saved Successfully");
			newErrorMessageCloseBtn.click();
			return false;
		}
		
		
	}
	
	
	
	
	public boolean checkClearOptionINSellerPriceBook() throws InterruptedException
	{
		//Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	  	homeMenu.click();
	  	
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterMenu));
	  	homeMasterMenu.click();
	  
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	  	homeMasterItemMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemSellerPriceBookMenu));
		homeMasterItemSellerPriceBookMenu.click();
		
		
		//Thread.sleep(3000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		 buyerPriceBook_PriceBook.click();
		 buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		 buyerPriceBook_AbbreviationBtn.click();
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
		 buyerPriceBookFilter_FilterAndLoadBtn.click();
		
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookClearBtn));
		 buyerPriceBookClearBtn.click();
		 
		 getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(gridItemThirdRow));
		
		 String Actdata=gridItemThirdRow.getAttribute("value");
		 String Expdata="Batch-FIFO";
		 
		 try {
			 if (Actdata.equalsIgnoreCase(Expdata)) {
					System.out.println("Pass:  Clear Option Working Successfully in Seller PriceBook");
					return false;
				} 
				else 
				{
					System.out.println("Fail: Clear Option Working Successfully in Seller PriceBook");
					return true;
				}
		} catch (Exception e) {
			return true;
		}
	}
	
	
	
	
	
	

	  
	   public boolean checkCreationofPriceBookWithProductTypeisBininSellerPriceBook()
	   {
		   
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			 buyerPriceBook_PriceBook.click();
			 buyerPriceBook_PriceBook.sendKeys(" Testing Buying Rates Hyd WH");
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
			 buyerPriceBook_AbbreviationBtn.click();
		   
			   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
				buyerPriceBook_CustamizeSettingBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_WarehouseCheckBox));
			    buyerPriceBook_WarehouseCheckBox.click();
				  
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
				buyerPriceBook_CustamizeSettingBtn.click();
			 
			 
			 
			 
			 
			   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRow));
				gridItemFirstRow.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRowafterSelecting));
				gridItemFirstRowafterSelecting.click();
				gridItemFirstRowafterSelecting.sendKeys("Bin-FIFO");
				
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseFirstRow));
				gridWarehouseFirstRow.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseFirstRowAfterSelecting));
				gridWarehouseFirstRowAfterSelecting.click();
				gridWarehouseFirstRowAfterSelecting.sendKeys("HYDERABAD");
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateFirstRow));
				gridRateFirstRow.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateFirstRowAfterSelecting));
				gridRateFirstRowAfterSelecting.sendKeys("105");
				
			 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
				buyerPriceBookSaveBtn.click();

				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(newErrorMessage));
				String Actdata=newErrorMessage.getText();
			    String Expdata="Price Book created successfully";
			    
			   if (Actdata.equalsIgnoreCase(Expdata)) 
			   {
					System.out.println("Pass: Creation of  Seller PriceBook With Product Type is Bin");
					newErrorMessageCloseBtn.click();
					return true;
				} 
				else 
				{
					System.out.println("Fail:  Creation of  Seller PriceBook With Product Type is Bin");
					newErrorMessageCloseBtn.click();
					return false;
				}
	   }
	   
	   
	   
	   
	   
	   public boolean checkCopyandPasteinSellerPriceBook() throws InterruptedException
	   {

		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			 buyerPriceBook_PriceBook.click();
			 buyerPriceBook_PriceBook.sendKeys("Testing Buying Rates Hyd WH");
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
			 buyerPriceBook_AbbreviationBtn.click();
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
			 buyerPriceBookFilter_FilterAndLoadBtn.click();
		   
			 //Thread.sleep(2000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCopyBtn));
			 buyerPriceBookCopyBtn.click();
			 
			 //Thread.sleep(2000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookPasteBtn));
			 buyerPriceBookPasteBtn.click();
			 
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRow));
				gridItemFirstRow.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRowafterSelecting));
				gridItemFirstRowafterSelecting.click();
				
				
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(gridItemFirstRowafterSelecting));
				String Actdata=gridItemFirstRowafterSelecting.getAttribute("value");
			    String Expdata="Bin-FIFO";
			    
			   if (Actdata.equalsIgnoreCase(Expdata)) 
			   {
					System.out.println("Pass: Copy And Paste Option ");
					
					return true;
				} 
				else 
				{
					System.out.println("Fail: Copy And Paste Option");
					
					return false;
				}
	   }
	   
	
	public boolean checkSaveOptionAfterPerformingPasteOptioninSellerPriceBook()
	{
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(newErrorMessage));
		String Actdata=newErrorMessage.getText();
	    String Expdata="Price Book updated successfully";
	    
	   if (Actdata.equalsIgnoreCase(Expdata)) 
	   {
			System.out.println("Pass: Update in SellerPriceBook and Saved Successfully");
			newErrorMessageCloseBtn.click();
			return true;
		} 
		else 
		{
			System.out.println("Fail:   Update in SellerPriceBook and Saved Successfully");
			newErrorMessageCloseBtn.click();
			return false;
		}
	}
	
	
	
	
	
	
	public boolean checkPriceBookThroughCopyandPasteinSellerPriceBook() throws InterruptedException
	{
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		 buyerPriceBook_PriceBook.click();
		 buyerPriceBook_PriceBook.sendKeys(" Testing Buying Rates Hyd WH");
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		 buyerPriceBook_AbbreviationBtn.click();
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
		 buyerPriceBookFilter_FilterAndLoadBtn.click();
	   
		 //Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCopyBtn));
		 buyerPriceBookCopyBtn.click();
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCloseBtn));
		 buyerPriceBookCloseBtn.click();
		
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	     homeMenu.click();
	  	
	  	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterMenu));
	  	 homeMasterMenu.click();
	  
	  
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	  	 homeMasterItemMenu.click();
	 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemSellerPriceBookMenu));
	 	homeMasterItemSellerPriceBookMenu.click();
		
		
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		 buyerPriceBook_PriceBook.click();
		 buyerPriceBook_PriceBook.sendKeys("Test Two");
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		 buyerPriceBook_AbbreviationBtn.click();
		 
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookPasteBtn));
		buyerPriceBookPasteBtn.click();
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newErrorMessageCloseBtn));
		newErrorMessageCloseBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys("Test Two");
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(buyerPriceBook_PriceBook));
		String Actdata=buyerPriceBook_PriceBook.getAttribute("value");
	    String Expdata="Test Two";
	    
	   if (Actdata.equalsIgnoreCase(Expdata)) 
	   {
			System.out.println("Pass:Seller Price Book Through Copy and Paste");
			
			return true;
		} 
		else 
		{
			System.out.println("Fail:Seller Price Book Through Copy and Paste");
			
			return false;
		}
		
	}
	
	
	
	
	 
	public boolean ClickonLoadButtoninSellerPriceBook() throws InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookFilter_FilterAndLoadBtn));
		buyerPriceBookFilter_FilterAndLoadBtn.click();
		 
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRow));
			gridItemFirstRow.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRowafterSelecting));
			gridItemFirstRowafterSelecting.click();
			
			
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(gridItemFirstRowafterSelecting));
			String Actdata=gridItemFirstRowafterSelecting.getAttribute("value");
		    String Expdata="Bin-FIFO";
		    
		   if (Actdata.equalsIgnoreCase(Expdata)) 
		   {
				System.out.println("Pass: Copy And Paste Option in Seller Price Book");
				 //Thread.sleep(2000);
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCopyBtn));
				 buyerPriceBookCopyBtn.click();
				 buyerPriceBookCloseBtn.click();
				return true;
			} 
			else 
			{
				System.out.println("Fail: Copy And Paste Option in Seller Price Book");
				 //Thread.sleep(2000);
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookCopyBtn));
				 buyerPriceBookCopyBtn.click();
		      	 buyerPriceBookCloseBtn.click();
				return false;
			}
	}
	
	
	
	public boolean checkDeleteinSellerPriceBook() throws InterruptedException
	{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	     homeMenu.click();
	  	
	  	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterMenu));
	  	 homeMasterMenu.click();
	  
	  	
	   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	  	 homeMasterItemMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
		homeMasterItem_buyerPriceBookMenu.click();
		
		
		//Thread.sleep(2000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		 buyerPriceBook_PriceBook.click();
		 buyerPriceBook_PriceBook.sendKeys("Test Hyd Wh");
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		 buyerPriceBook_AbbreviationBtn.click();
		 
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
			buyerPriceBook_CustamizeSettingBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_WarehouseCheckBox));
		    buyerPriceBook_WarehouseCheckBox.click();
			  
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
			buyerPriceBook_CustamizeSettingBtn.click();
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRow));
			gridItemFirstRow.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridItemFirstRowafterSelecting));
			gridItemFirstRowafterSelecting.click();
			gridItemFirstRowafterSelecting.sendKeys("Group of RMA LIFO and FIFO");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseFirstRow));
			gridWarehouseFirstRow.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridWarehouseFirstRowAfterSelecting));
			gridWarehouseFirstRowAfterSelecting.click();
			gridWarehouseFirstRowAfterSelecting.sendKeys("HYDERABAD");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateFirstRow));
			gridRateFirstRow.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateFirstRowAfterSelecting));
			gridRateFirstRowAfterSelecting.sendKeys("100");
		 
		
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		    buyerPriceBookSaveBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(newErrorMessage));
		String Actdata=newErrorMessage.getText();
	    String Expdata="Price Book created successfully";
	    
	    System.out.println(Actdata);
	   if (Actdata.equalsIgnoreCase(Expdata)) 
	   {
			System.out.println("Pass: Creation of Seller Price Book For Delete Option ");
			newErrorMessageCloseBtn.click();
			return true;
		} 
		else 
		{
			System.out.println("Fail: Creation of Seller Price Book For Delete Option");
			newErrorMessageCloseBtn.click();
			return false;
		}
	}
	
	
	
	
	public boolean SelecttheCreatedpriceBookfromDropdownandClickonDeletebuttonofSellerPriceBook()
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		 buyerPriceBook_PriceBook.click();
		 buyerPriceBook_PriceBook.sendKeys("Test Hyd Wh");
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		 buyerPriceBook_AbbreviationBtn.click();
		 
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookDeleteBtn));
		 buyerPriceBookDeleteBtn.click();
		 
		 String Actdata=getDriver().switchTo().alert().getText();
		 String Expdata="Do you want to delete Total Price book?";
		 
		 System.out.println(Actdata);
		 if (Actdata.equalsIgnoreCase(Expdata)) {
			System.out.println("Pass: Created Seller price Book from Dropdown and Click on Delete button");
			return true;
		} else {
			System.out.println("Fail: Created Seller price Book from Dropdown and Click on Delete button");
			return false;
		}
		
		 
	}
	
	
	
	public boolean checkwithNoButtoninDeleteoptionSellerPriceBook() throws InterruptedException
	{
		//Thread.sleep(2000);
		getDriver().switchTo().alert().dismiss();
		
		//Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(buyerPriceBook_PriceBook));
		String Actdata=buyerPriceBook_PriceBook.getAttribute("value");
	    String Expdata="Test Hyd Wh";
	    
	   if (Actdata.equalsIgnoreCase(Expdata)) 
	   {
			System.out.println("Pass:No Button in Delete option in SellerPriceBook");
			
			return true;
		} 
		else 
		{
			System.out.println("Fail: No Button in Delete option in SellerPriceBook");
			
			return false;
		}
		
	}
	
	
	

	public boolean checkwithYesButtoninDeleteoptioninSellerPriceBook() throws InterruptedException
	{
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookDeleteBtn));
		 buyerPriceBookDeleteBtn.click();
	
		//Thread.sleep(2000);
		getDriver().switchTo().alert().accept();
	

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(newErrorMessage));
		String Actdata=newErrorMessage.getText();
	    String Expdata="Pricebook details deleted successfully.";
	    
	   if (Actdata.equalsIgnoreCase(Expdata)) 
	   {
			System.out.println("Pass: Yes Button in Delete option in SellerPriceBook ");
			newErrorMessageCloseBtn.click();
			return true;
		} 
		else 
		{
			System.out.println("Fail: Yes Button in Delete option in SellerPriceBook ");
			newErrorMessageCloseBtn.click();
			return false;
		}
	
	}
	
	 
	
	
	
	
	public boolean checkExportExcelinSellerPriceBook() throws InterruptedException
	{
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	     homeMenu.click();
	  	
	  	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterMenu));
	  	 homeMasterMenu.click();
	  
	  	
	   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	  	 homeMasterItemMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemSellerPriceBookMenu));
		homeMasterItemSellerPriceBookMenu.click();
		
		
		//Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(buyerPriceBook_PriceBook));
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		 buyerPriceBook_PriceBook.click();
		 buyerPriceBook_PriceBook.sendKeys("Test Two");
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		 buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_ExportToExcel));
		 buyerPriceBook_ExportToExcel.click();
		
		 getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(newErrorMessage));
			String Actdata=newErrorMessage.getText();
		    String Expdata="Pricebook exported successfully";
		    
		   if (Actdata.equalsIgnoreCase(Expdata)) 
		   {
				System.out.println("Pass: Yes Button in Delete option in SellerPriceBook ");
				newErrorMessageCloseBtn.click();
				return true;
			} 
			else 
			{
				System.out.println("Fail: Yes Button in Delete option in SellerPriceBook");
				newErrorMessageCloseBtn.click();
				return false;
			}
	}
	
  */
	
	public MasterBuyingAndSellingPage(WebDriver driver)
    {
    	
   	 PageFactory.initElements(driver, this);
   	
    }
	
	
	
}
