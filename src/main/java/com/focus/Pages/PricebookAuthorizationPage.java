package com.focus.Pages;

import java.awt.AWTException;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

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
import org.openqa.selenium.support.ui.WebDriverWait;

import com.focus.base.BaseEngine;

public class PricebookAuthorizationPage extends BaseEngine{
	
	
	 @FindBy(xpath="//*[@id='16']//*[contains(text(),'Settings')]")
	private static WebElement  settingsmenuBtn;
	
	@FindBy(xpath="//*[@id='235']/span")
	private static WebElement  pricebookAuthorizationBtn;
	
	@FindBy(xpath="//input[@id='txtDoc']")
	private static WebElement  pricebookDropdown;
	
	@FindBy(xpath="//li[normalize-space(text())='Buying Price Book']")
	private static WebElement  BuyingPriceBook;
	
	@FindBy(xpath="//li[normalize-space(text())='Selling Price Book']")
	private static WebElement  SellingPriceBook;
	
	  @FindBy(xpath="//*[@id='ctrlAuthorizationH']")
	  private static WebElement transAuthNameTxt; 
	  
	  
	  @FindBy(xpath="//*[@id='spnAddLevel']/i")
	  private static WebElement addLevelBtn; 
	  
	  
	  @FindBy(xpath="//*[@id='txtDescription']")
	  private static WebElement descriptionTxt; 
		
	  @FindBy(xpath="//*[@id='chkConditionNotRequired']")
	  private static WebElement conditionNotRequiredChkBox; 
	  
	  @FindBy(xpath="//*[@id='id_report_popup_body']/div[2]//ul[1]/li[2]/a")
	  private static WebElement userSelectionTab;
	
	  @FindBy(xpath="//*[@id='ddlUserSelection']")
	  private static WebElement userselectionDropdown;
	
	  @FindBy(xpath="//*[@id='ddlCriteria']")
	  private static WebElement criteriaDropdown;
	  
	  @FindBy(xpath="//*[@id='id_tblUserSelection_User']")
	  private static WebElement userDropdown;
		
	  @FindBy(xpath="//td[@id='id_tblUserSelection_col_1-1']")
	  private static WebElement userrow1;
		
	  @FindBy(xpath="//td[@id='id_tblUserSelection_col_2-1']")
	  private static WebElement userrow2;
	  
	  @FindBy(xpath="//*[@id='id_report_popup_ok']")
	  private static WebElement definitionOkBtn;					
				
	  @FindBy(xpath="//span[@id='id_report_popup_cancel']")
	  private static WebElement definitionCloseBtn;	
	  
	  
	  @FindBy(xpath="//*[@id='ulAuthBtns']/li[5]/span/i")
	  private static WebElement tranAuthCancelBtn; 
			
	  @FindBy(xpath="//*[@id='btnSaveAuthorization']")
	  private static WebElement tranAuthSaveBtn; 
	  
	  @FindBy(xpath="//*[@id='spndeleteProfile']")
	  private static WebElement tranAuthDeleteBtn; 
	  
		@FindBy(xpath="//*[@id='id_report_popup_body']/div[2]//ul[1]/li[3]/a")
		private static WebElement alertsTab;
	  
		@FindBy(xpath="//u[contains(text(),'Email Header Template')]")
		private static WebElement emailHeaderTemplateLink;
	  
		@FindBy(xpath="//input[@id='smsTemplateOptCtrl']")
		private static WebElement templateTxt;
	  
		@FindBy(xpath="//tbody[@id='smsTemplateOptCtrl_table_body']/tr/td[2]")
		private static List<WebElement> custTemplatesList;
	  
		@FindBy(xpath="//span[@id='btnCloseSMSModal']")
		private static WebElement custTemplateCloseBtn;
		
		
		@FindBy(xpath="//*[@id='chkActive']")
		private static WebElement activeChkBox;  
		
		@FindBy(xpath="//*[@id='chkAlwaysauthorizationstartsfromfirstlevel']")
		private static WebElement alwaysauthorizationstartsfromfirstlevelChkBox; 
		
		@FindBy(xpath="//*[@id='spnAddParallel']/i")
		private static WebElement addParallelBtn;    
		
		@FindBy(xpath="//*[@class='icon-clear hiconright2']")
		private static WebElement clearbtnInAuth; 
		
		@FindBy(xpath="//*[@class='icon-cancel icon-font7']")
		private static WebElement cancelBtnInAuth; 
			
		@FindBy(xpath="//*[@id='chkDateRange']")
		private static WebElement dateRangeChkBox; 
			
		@FindBy(xpath="//*[@id='dpFromDate']")
		private static WebElement fromDateTxt; 
			
		@FindBy(xpath="//*[@id='dpToDate']")
		private static WebElement toDateTxt; 
		
		@FindBy(xpath="//*[@id='chkTimeRange']")
		private static WebElement timeRangeChkBox;
			
		@FindBy(xpath="//*[@id='tpFromTime']")
		private static WebElement fromTimeTxt; 
			
		@FindBy(xpath="//*[@id='tpToTime']")
		private static WebElement totimeTxt; 
		
		@FindBy(xpath="//i[@id='id_StartNode']")
		private static WebElement startPointBtn; 
		
		@FindBy(xpath="//*[@id='tabcondition']")
		private static WebElement conditionTab; 
		
		
		@FindBy(xpath="//*[@id='235_0_AdvanceFilter_']/table/tbody/tr/td[1]/select")
		private static WebElement conjunctionDropdown; 
			
		@FindBy(xpath="//tbody//input[@placeholder='Select Field']")
		private static WebElement selectFieldTxt; 
		
		@FindBy(xpath="//*[@id='235_0_AdvanceFilter_']//tbody/tr/td[3]/select[@class='custom-select borderr']")
		private static WebElement selectOperatorDropdown; 
			
		@FindBy(xpath="//*[@id='235_0_AdvanceFilter_']//tbody/tr/td[4]/select[@class='custom-select borderr']")
		private static WebElement compareWithDropdown; 
		
		@FindBy(xpath="//*[@id='235_0_AdvanceFilter_']//tbody/tr/td[4]")
		private static WebElement compareWithDropdown1; 
			
			
		@FindBy(xpath="//*[@id='235_0_AdvanceFilter_']//tbody/tr/td[5]/input")
		private static WebElement valueTxt;
		
		@FindBy(xpath="//select[@id='ddlStatus']")
		private static WebElement selectStatus;
		
		@FindBy(xpath="//*[@id='chkFullAuthorization']")
		private static WebElement fullAuthorizationChkBox;
		
		@FindBy(xpath="//select[@id='ddlUndoAuthEditing']")
		private static WebElement selectUndoAuthEditing;
		
		@FindBy(xpath="//*[@id='id_report_popup_cancel']/i")
		private static WebElement definitionCancelBtn;	
		
		@FindBy(xpath="//*[@id='chkReceived']")
		private static WebElement receivedChkBox;	
			
		@FindBy(xpath="//*[@id='chkApproved']")
		private static WebElement approvedChkBox;	
			
		@FindBy(xpath="//*[@id='chkRejected']")
		private static WebElement rejectedChkBox;	
			
		@FindBy(xpath="//*[@id='chkStopped']")
		private static WebElement stoppedChkBox;	
		
		@FindBy(xpath="//*[@id='id_tblAlerts_control_heading_2']/div/u")
		private static WebElement emailHeaderTemplateBtn;	
		
		@FindBy(xpath="//*[@id='id_tblAlerts_control_heading_5']/div/u")
		private static WebElement alertsHeaderBtn;	
		
		@FindBy(xpath="//u[contains(text(),'Email Template')]")
		private static WebElement emailTemplateBtn;	
		
		@FindBy(xpath="//u[contains(text(),'Alerts')]")
		private static WebElement headerAlert;

		@FindBy(xpath="//div[@class='divStopNode']//i[@class='icon icon-end']")
		private static WebElement endpoint;
		
		@FindBy(xpath="//li[normalize-space(text())='Buying Price Book (1)']")
		private static WebElement  createBuyingPriceBook;
		
		@FindBy(xpath="//li[normalize-space(text())='Selling Price Book (1)']")
		private static WebElement  createSellingPriceBook;
		
	
		
		@FindBy(xpath="//*[@id='235_0_AdvanceFilter_']//tbody/tr/td[5]/input")
		//@FindBy(xpath="//*[@id='advancefilter_master_235_0_']")
		private static WebElement enter_valueTxt;
		
		@FindBy(xpath="//*[@id='advancefilter_master_235_0_']")
		private static WebElement enter_ProdvalueTxt;
		
		
		@FindBy(xpath="(//*[@id='1'])[2]")
		private static WebElement selectFieldpricebookName;
		
				
		@FindBy(xpath="(//*[@id='15'])[2]")
		private static WebElement selectFieldRate;
		
		@FindBy(xpath="(//*[@id='3'])[2]")
		private static WebElement selectFieldProductName;
		
		@FindBy(xpath="(//*[@id='6'])[2]")
		private static WebElement selectFieldAccountName;
		
		@FindBy(xpath="(//*[@id='13'])[2]")
		private static WebElement selectFieldMinimumQty;
		
		@FindBy(xpath="(//*[@id='10' and text()='Department'])[1]")
		private static WebElement selectFieldTag;
		
		
		@FindBy(xpath="//ul[@class='nav nav-tabs AuthTabs']//*[contains(text(),'Alerts')]")
		private static WebElement alertTabInEdit;
		
		@FindBy(xpath="//u[contains(text(),'SMS template')]")
		private static WebElement headerSMSTemplate;	
		

		@FindBy(xpath="//input[@id='smsTemplateOptCtrl']")
		private static WebElement emailTemplates;
		
		@FindBy(xpath="//i[@class='icon-open icon-font7']")
		private static WebElement companyMenu;
		
		@FindBy(xpath="//span[contains(text(),'Company Name')]")
		private static WebElement companyNameSource;
		
		@FindBy(xpath="//textarea[@id='txtTemplate']")
		private static WebElement companyNameTemplateArea;
		
		@FindBy(xpath="//span[@id='btnSaveSmsTemplate']")
		private static WebElement saveBtnInEmailTemplateHeader;
		
		@FindBy(xpath="//span[@id='btnNewSmsTemplate']")
		private static WebElement newButtonInCreateTemplate;
		
		@FindBy(xpath="//span[@id='btnCloseSMSModal']")
		private static WebElement closeButtonInCreateTemplate;
		
		@FindBy(xpath="//*[@id='ctrlAuthorizationH']")
		private static WebElement nameTxt;
		
		
		@FindBy(xpath="//*[@id='btnSaveAuthorization']")
		private static WebElement saveBtn;
		
		@FindBy(xpath="//*[@id='btnDelete']")
		public static WebElement deleteBtn;
		
		@FindBy(xpath="//*[@id='spndeleteProfile']")
		public static WebElement AuthdeleteBtn;
		
		@FindBy(xpath="//*[@id='135']")
		public static WebElement Inventorybut;
		
		@FindBy(xpath="//*[@id='137']")
		public static WebElement InventoryTransactionsbut;
		
		@FindBy(xpath="//*[@id='140']")
		public static WebElement salesbut;
		
		@FindBy(xpath="//*[@id='2017']")
		public static WebElement Salesorder;
		
		@FindBy(xpath="//*[@id='id_transaction_homescreen_new']/i")
		public static WebElement Salesorder_newbut;
		
		@FindBy(xpath="//*[@id='id_transactionentry_previous']/a/span")
		public static WebElement Salesorder_Previousbut;
		
		@FindBy(xpath="//*[@id='id_header_4']")
		public static WebElement CustomerAcc;
		
		@FindBy(xpath="//*[@id='id_body_536870915']")
		public static WebElement DepartmentTxt;
		
		@FindBy(xpath="//*[@id='id_body_536870916']")
		public static WebElement WarehouseTxt;
		
		@FindBy(xpath="//*[@id='id_body_23']")
		public static WebElement ItemTxt;
		
		@FindBy(xpath="//*[@id='id_body_24']")
		public static WebElement UnitTxt;
		
		@FindBy(xpath="//*[@id='id_body_26']")
		public static WebElement QuantityTxt;
		
		@FindBy(xpath="//*[@id='id_body_27']")
		public static WebElement RateTxt;
		
		@FindBy(xpath="//*[@id='id_body_28']")
		public static WebElement GrossTxt;
		
		@FindBy(xpath="//*[@id='id_body_33554573']")
		public static WebElement Val1Txt;
		
		@FindBy(xpath="//*[@id='id_body_33554440']")
		public static WebElement DiscountTxt;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[2]")
		public static WebElement select1strow_1stcolumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[5]")
		public static WebElement select1strow_4thcolumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[9]")
		public static WebElement select1strow_8thcolumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[12]")
		public static WebElement select1strow_11thcolumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[2]")
		public static WebElement select2ndrow_1stcolumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[5]")
		public static WebElement select2ndrow_4thcolumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[9]")
		public static WebElement select2ndrow_8thcolumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[12]")
		public static WebElement select2ndrow_11thcolumn;
		
		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td")
		private static List<WebElement> Row1List;
		
		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td")
		private static List<WebElement> Row2List;
		
		@FindBy(xpath="//*[@id='id_transactionentry_save']/a/span")
		public static WebElement SaveButton;
		
		@FindBy(xpath="//*[@id='id_transactionentry_close']/a/span")
		public static WebElement closeButton;
		
		
		
		
		
		public boolean 	checkOpenPricebookAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(4000);
			getDriver().navigate().refresh();
			Thread.sleep(4000);
			
			ClickUsingJs(settingsmenuBtn);
			
			Thread.sleep(2000);
		/*	
			((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
			Thread.sleep(2000);*/
			
			ClickUsingJs(pricebookAuthorizationBtn);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pricebookDropdown));
			
		    boolean actpricebookDropdown                       = pricebookDropdown.isDisplayed();
			boolean actnameTxt                              = nameTxt.isDisplayed();
			boolean actactiveChkBox                         = activeChkBox.isDisplayed();
			boolean actalwaysauthorstartsfirstlevelChkBox   = alwaysauthorizationstartsfromfirstlevelChkBox.isDisplayed();
			boolean actaddLevelBtn                          = addLevelBtn.isDisplayed();
			boolean actaddParallelBtn                       = addParallelBtn.isDisplayed();
			boolean actclearbtn               				= clearbtnInAuth.isDisplayed();
			boolean actsaveBtn  							= saveBtn.isDisplayed();
			boolean actcancelBtn 							= closeBtn.isDisplayed();
			boolean actdeleteBtn			                = AuthdeleteBtn.isDisplayed();
			
			boolean exppricebookDropdown                       = true;
			boolean expnameTxt            					= true;
			boolean expactiveChkBox                			= true;
			boolean expalwaysauthorstartsfirstlevelChkBox   = true;
			boolean expaddLevelBtn          				= true;
			boolean expaddParallelBtn              			= true;
			boolean expclearbtn               				= true;
			boolean expsaveBtn  							= true;
			boolean expcancelBtn							= true;
			boolean expdeleteBtn							= true;
			
			System.out.println("********************* checkTransactionAuthorizationOptions ******************");
		
			System.out.println("Home Page pricebookDropdown                     :"+actpricebookDropdown                      +"  Value Expected : "+exppricebookDropdown);
			System.out.println("Home Page nameTxt                            :"+actnameTxt                             +"  Value Expected : "+expnameTxt);
			System.out.println("Home Page activeChkBox                       :"+actactiveChkBox                        +"  Value Expected : "+expactiveChkBox);
			System.out.println("Home Page alwaysauthorstartsfirstlevelChkBox :"+actalwaysauthorstartsfirstlevelChkBox  +"  Value Expected : "+expalwaysauthorstartsfirstlevelChkBox);
	        System.out.println("Home Page addLevelBtn                        :"+actaddLevelBtn                         +"  Value Expected : "+expaddLevelBtn);
			System.out.println("Home Page addParallelBtn                     :"+actaddParallelBtn                      +"  Value Expected : "+expaddParallelBtn);
			System.out.println("Home Page clearbtn                           :"+actclearbtn                            +"  Value Expected : "+expclearbtn);
			System.out.println("Home Page saveBtn                            :"+actsaveBtn                             +"  Value Expected : "+expsaveBtn);
			System.out.println("Home Page cancelBtn 						 :"+actcancelBtn                           +"  Value Expected : "+expcancelBtn);
			System.out.println("Home Page deleteBtn                          :"+actdeleteBtn                           +"  Value Expected : "+expdeleteBtn);
			
			
			boolean res =  actpricebookDropdown==exppricebookDropdown  && actnameTxt==expnameTxt &&
					actactiveChkBox==expactiveChkBox && actalwaysauthorstartsfirstlevelChkBox==expalwaysauthorstartsfirstlevelChkBox && 
					actaddLevelBtn==expaddLevelBtn  && actaddParallelBtn==expaddParallelBtn && 
				    actclearbtn==expclearbtn && actsaveBtn==expsaveBtn && actcancelBtn==expcancelBtn && actdeleteBtn==expdeleteBtn;
			
			String actres = Boolean.toString(res);
			String expres = "true";
			
			
			
			System.err.println("Dispaly Result : "+actres+" Value Expected : "+expres);
			
			if(actres.equalsIgnoreCase(expres))
				
			{
				System.out.println("Test Pass : PriceBook Authorization is Displayed");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : PriceBook Authorization is Displayed");
				
				return false;
			}
		}
		
		
		public boolean 	checkPricebookDropdownOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			
			String actpricebookDropdown=pricebookDropdown.getAttribute("placeholder");
			String exppricebookDropdown="pricebook Type";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pricebookDropdown));
			pricebookDropdown.click();
			
			Thread.sleep(2000);
			pricebookDropdown.sendKeys(Keys.SPACE);
			
			Thread.sleep(2000);
			 boolean actbuyingpricebook        = BuyingPriceBook.isDisplayed();
			 boolean actsellingpricebook        = SellingPriceBook.isDisplayed();
			 
			 boolean expbuyingpricebook        = true;
			 boolean expsellingpricebook        = true;

				
			 System.out.println("****************************** checkpricebookDropdownOptions  ********************************************");
			 System.out.println("Buyingpricebook        : "+actbuyingpricebook         +"   Value Expected  : "+expbuyingpricebook);
			 System.out.println("Sellingpricebook               :"+actsellingpricebook                 +"   Value Expected : "+expsellingpricebook);
			 
			 boolean res =  actbuyingpricebook==expbuyingpricebook && actsellingpricebook==expsellingpricebook;
				
			 String actres = Boolean.toString(res);
			 String expres = "true";
				
			 
				
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Master Dropdown Options Are As expected");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Master Dropdown Options Are NOT As expected");
				
				return false;
			}
		}

		public boolean 	checkSaveWithoutSelectionOfPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			
			System.out.println("************************************checkSaveWithoutSelectionOfPriceBook**********************************");
			
			String expMessage="Give proper Authorization name.";
			
			String actMessage=checkValidationMessage(expMessage);	
			
			
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : PriceBook Authorization is Displayed");
				
				return false;
			}
		}
		
		
		
		public boolean 	checkSelectingBuyingPriceBookinPriceBookDropdown() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pricebookDropdown));
			pricebookDropdown.click();
			pricebookDropdown.sendKeys(Keys.SPACE);
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SellingPriceBook));
			SellingPriceBook.click();
			
			Thread.sleep(2000);
			
			
			String actPriceBookDropdown=pricebookDropdown.getAttribute("value");
			String expPriceBookDropdown="Selling Price Book";
			
			

			System.out.println("********************************checkSelectingBuyingPriceBookinPriceBookDropdown*********************"); 
			
			
			System.out.println("PriceBookDropdown   : "+actPriceBookDropdown+" Value Expected : "+expPriceBookDropdown);
			
			if(actPriceBookDropdown.equalsIgnoreCase(expPriceBookDropdown))
			{
				System.out.println("Test Pass : Buying PriceBook is Selected");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail :Buying PriceBook is Not Selected");
				
				return false;
			}
			
			
		}

		public boolean 	checkSaveWithoutName() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
					
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			System.out.println("********************************************checkSaveWithoutName***************************");
			
			String expMessage="Give proper Authorization name.";
			
			String actMessage=checkValidationMessage(expMessage);
					
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Pricebook Authorization is Displayed");
				
				return false;
			}
		}  
	  
	  
		public boolean 	checkActiveCheckBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(activeChkBox));
			
			boolean actactiveChkBox=activeChkBox.isSelected();
			boolean expactiveChkBox=true;
			
			System.out.println("*****************************************checkActiveCheckBox***********");
			
			System.out.println("activeChkBox   :"+actactiveChkBox +"  Value Expected : "+expactiveChkBox);
			
			boolean res = actactiveChkBox==expactiveChkBox;
			
			String actres = Boolean.toString(res);
			String expres = "true";
	
			if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Active ChkBox IS Enabled/Checked");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Active ChkBox IS NOT Enabled/Checked");
				
				return false;
			}
		}
		
		public boolean 	checkInputNameAndClickOnSaveWithoutAuthorizationSettings() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.sendKeys("Seller");
			//nameTxt.sendKeys(Keys.TAB);
			
			
			Thread.sleep(2000);
			
			String actnameTxt=nameTxt.getAttribute("value");
			String expnameTxt="Seller";
			
			
			
			System.out.println("**********************************************checkInputNameAndClickOnSaveWithoutAuthorizationSettings**********************");
			
			System.out.println("nameTxt   :"+actnameTxt +" Value Expected : "+expnameTxt);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			String expMessage="Create Authorization before Saving";
			
			String actMessage=checkValidationMessage(expMessage);
			
			
			
			if(actnameTxt.equalsIgnoreCase(expnameTxt) && actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Test Pass : Account inputed Successfully");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Account NOT inputed Successfully");
				return false;
			}
		}	
	
		
		public boolean 	checkStartPoint() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
			ClickUsingJs(startPointBtn);
			
			
			 Thread.sleep(2000);
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionTab));
			 boolean actconditionTab            = conditionTab.isDisplayed();
			 boolean actuserSelectionTab        = userSelectionTab.isDisplayed();
			 boolean actalertsTab               = alertsTab.isDisplayed();
			  
			 boolean expconditionTab            = true;
			 boolean expuserSelectionTab        = true;
			 boolean expalertsTab               = true;
			  
			 System.out.println("***********************************************checkStartPoint****************");
			  
			 System.out.println("conditionTab            "+actconditionTab         +"  Value Expected  :"+expconditionTab);
			 System.out.println("userSelectionTab        "+actuserSelectionTab     +"  Value Expected  :"+expuserSelectionTab);
			 System.out.println("alertsTab               "+actalertsTab            +"  Value Expected  :"+expalertsTab);
			 
			 boolean res = actconditionTab==expconditionTab && actuserSelectionTab==expuserSelectionTab && actalertsTab==expalertsTab;
				
			 String actres = Boolean.toString(res);
			 String expres = "true";
					
			
			  
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Start Point works As Expected");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Start Point NOT works As Expected");
				
				return false;
			}
		}
		
		
		public boolean 	checkConditionTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionTab));
			 conditionTab.click();
			
			 Thread.sleep(2000);

			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			 boolean actdescriptionTxt                    = descriptionTxt.isDisplayed();
			 boolean actconditionNotRequiredChkBox        = conditionNotRequiredChkBox.isDisplayed();
			 boolean actdateRangeChkBox                   = dateRangeChkBox.isDisplayed();
			 boolean acttimeRangeChkBox           		  = timeRangeChkBox.isDisplayed();
			 boolean actfromDateTxt           			  = fromDateTxt.isDisplayed();
			 boolean acttoDateTxt           			  = toDateTxt.isDisplayed();
			 boolean actfromTimeTxt            			  = fromTimeTxt.isDisplayed();
			 boolean acttotimeTxt                         = totimeTxt.isDisplayed();
			 boolean actconjunctionDropdown        		  = conjunctionDropdown.isDisplayed();
			 boolean actselectFieldTxt               	  = selectFieldTxt.isDisplayed();
			 boolean actselectOperatorDropdown            = selectOperatorDropdown.isDisplayed();
			 boolean actcompareWithDropdown               = compareWithDropdown.isDisplayed();
			 boolean actvalueTxt               			  = valueTxt.isDisplayed();
			 boolean actStatus							  = selectStatus.isDisplayed();
			 boolean actfullAuthorizationChkBox			  = fullAuthorizationChkBox.isDisplayed();
			 boolean actUndoAuthEditing					  =	selectUndoAuthEditing.isDisplayed();
			 
				
			 boolean expdescriptionTxt                    = true;
			 boolean expconditionNotRequiredChkBox        = true;
			 boolean expdateRangeChkBox                   = true;
			 boolean exptimeRangeChkBox           		  = true;
			 boolean expfromDateTxt           			  = true;
			 boolean exptoDateTxt                         = true;
			 boolean expfromTimeTxt                       = true;
			 boolean exptotimeTxt                         = true;
			 boolean expconjunctionDropdown               = true;
			 boolean expselectFieldTxt                    = true;
			 boolean expselectOperatorDropdown            = true;
			 boolean expcompareWithDropdown               = true;
			 boolean expvalueTxt               			  = true;
			 boolean expStatus							  = true;
			 boolean expfullAuthorizationChkBox			  = true;
			 boolean expUndoAuthEditing					  =	true;
			 
			 
			 
			
			 System.out.println("****************************** checkConditionTabOptions  ********************************************");
			 
			 System.out.println("descriptionTxt               :"+actdescriptionTxt             +"    Value Expected : "+expdescriptionTxt);
			 System.out.println("conditionNotRequiredChkBox   :"+actconditionNotRequiredChkBox +"    Value Expected : "+expconditionNotRequiredChkBox);
			 System.out.println("dateRangeChkBox              :"+actdateRangeChkBox            +"	 Value Expected : "+expdateRangeChkBox);
			 System.out.println("timeRangeChkBox              :"+acttimeRangeChkBox            +"	 Value Expected : "+exptimeRangeChkBox);
			 System.out.println("fromDateTxt                  :"+actfromDateTxt                +"	 Value Expected : "+expfromDateTxt);
			 System.out.println("toDateTxt                    :"+acttoDateTxt                  +"	 Value Expected : "+exptoDateTxt);
			 System.out.println("fromTimeTxt                  :"+actfromTimeTxt                +"	 Value Expected : "+expfromTimeTxt);
			 System.out.println("totimeTxt                    :"+acttotimeTxt                  +"	 Value Expected : "+exptotimeTxt);
			 System.out.println("conjunctionDropdown          :"+actconjunctionDropdown        +"	 Value Expected : "+expconjunctionDropdown);
			 System.out.println("selectFieldTxt               :"+actselectFieldTxt             +"	 Value Expected : "+expselectFieldTxt);
			 System.out.println("selectOperatorDropdown       :"+actselectOperatorDropdown     +"	 Value Expected : "+expselectOperatorDropdown);
			 System.out.println("compareWithDropdown          :"+actcompareWithDropdown        +"    Value Expected : "+expcompareWithDropdown);
			 System.out.println("valueTxt                     :"+actvalueTxt                   +"	 Value Expected : "+expvalueTxt);
			 System.out.println("status                       :"+actStatus                     +"	 Value Expected : "+expStatus);
			 System.out.println("fullAuthChkBox               :"+actfullAuthorizationChkBox    +"	 Value Expected : "+expfullAuthorizationChkBox);
			 System.out.println("undoAuthEditing              :"+actUndoAuthEditing            +"	 Value Expected : "+expUndoAuthEditing);
			 
			
			 boolean res = actdescriptionTxt==expdescriptionTxt  && actconditionNotRequiredChkBox==expconditionNotRequiredChkBox && actdateRangeChkBox==expdateRangeChkBox
					 && acttimeRangeChkBox==exptimeRangeChkBox && actfromDateTxt==expfromDateTxt && acttoDateTxt==exptoDateTxt&& actfromTimeTxt==expfromTimeTxt 
					 && acttotimeTxt==exptotimeTxt && actconjunctionDropdown==expconjunctionDropdown && actselectFieldTxt==expselectFieldTxt 
					 && actselectOperatorDropdown==expselectOperatorDropdown && actcompareWithDropdown==expcompareWithDropdown && actvalueTxt==expvalueTxt 
					&&actStatus==expStatus && actfullAuthorizationChkBox==expfullAuthorizationChkBox && actUndoAuthEditing==expUndoAuthEditing;
				
			 String actres = Boolean.toString(res);
			 String expres = "true";
					
			
			 
			 
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Condition Tab Options Are Displayed ");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Condition Tab Options Are NOT Displayed ");
				
				return false;
			}
			
		}
		
		public boolean 	checkUserSelectionTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
			userSelectionTab.click();
			
			Thread.sleep(2000);
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userselectionDropdown));
			 boolean actuserselectionDropdown             = userselectionDropdown.isDisplayed();
			 boolean actcriteriaDropdown                  = criteriaDropdown.isDisplayed();
			 boolean actuserrow1                          = userrow1.isDisplayed();
			 boolean actuserrow2                          = userrow2.isDisplayed();
			  
			 boolean expuserselectionDropdown             = true;
			 boolean expcriteriaDropdown                  = true;
			 boolean expuserrow1                          = true;
			 boolean expuserrow2                          = true;
			  
			 System.out.println("***********************************************checkUserSelectionTabOptions****************");
			  
			 System.out.println("UserselectionDropdown    :"+actuserselectionDropdown  +"  Value Expected  :"+expuserselectionDropdown);
			 System.out.println("CriteriaDropdown         :"+actcriteriaDropdown       +"  Value Expected  :"+expcriteriaDropdown);
			 System.out.println("Userrow1                 :"+actuserrow1               +"  Value Expected  :"+expuserrow1);
			 System.out.println("userrow2                 :"+actuserrow2               +"  Value Expected  :"+expuserrow2);
			 
			 boolean res = actuserselectionDropdown==expuserselectionDropdown && actcriteriaDropdown==expcriteriaDropdown && 
					  actuserrow1==expuserrow1 && actuserrow2==expuserrow2;
				
			 String actres = Boolean.toString(res);
			 String expres = "true";
					
			 
			  
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : User Selection TAb Options Are Displayed");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : User Selection TAb Options Are NOT Displayed");
				
				return false;
			}
		}
		
		
		
		
		/*@FindBy(xpath="//*[@id='id_tblUserSelection_User']")
		private static WebElement userDropdown;
		*/
		
		public boolean 	checkUserDropdownOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{		
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userselectionDropdown));
			userselectionDropdown.click();
			
			String userq=userselectionDropdown.getAttribute("value");
			
			System.out.println("userq"+userq);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(criteriaDropdown));
			criteriaDropdown.click();
			
			String userc=criteriaDropdown.getText();
			
			System.out.println("userc"+userc);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
			userrow1.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userDropdown));
			userDropdown.click();
			userDropdown.sendKeys("UserALOptionsST");
			userDropdown.sendKeys(Keys.ENTER);
			
			String Strlpl= userDropdown.getText();
		
			System.out.println(Strlpl);
			 
	        String[] testString=Strlpl.split("\n");
	        
	        List<String> listg = Arrays.asList(testString);
	        
	        System.out.println("********************************************checkUserDropdownOptions*************************************");
	        
	        System.out.println("Converting to string array"+listg);
	        
	        
	        String actlist = listg.toString();
	        String explist = "[SU, UserAllOptionsST, UserRestrictionST, UserExclusionST, UserForItemRestriction, Test, UserAccountRestrictions, PortalUser]";
	        					
	        System.out.println("The Users in User Dropdown Value Expected : " + explist + "Value Actual : " + actlist);
			
	        
	       if(actlist.equalsIgnoreCase(explist))
	       {
	    	  
	    	   return true;
	       }
	       else
	       {
	    	   
	    	   return false;
	       }
	}	
		
		public boolean 	checkAlertsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertsTab));
			 alertsTab.click();
			
			 Thread.sleep(2000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receivedChkBox));
			 boolean actreceivedChkBox                   = receivedChkBox.isDisplayed();
		     boolean actapprovedChkBox                   = approvedChkBox.isDisplayed();
			 boolean actrejectedChkBox                   = rejectedChkBox.isDisplayed();
			 boolean actemailHeaderTemplateBtn           = emailHeaderTemplateBtn.isDisplayed();
			 boolean actemailTemplateBtn                 = emailTemplateBtn.isDisplayed();
			 boolean actheaderAlert                      = headerAlert.isDisplayed();
			 
			 boolean expreceivedChkBox                   = true;
			 boolean expapprovedChkBox                   = true;
			 boolean exprejectedChkBox                   = true;
			 boolean expemailHeaderTemplateBtn           = true;
			 boolean expemailTemplateBtn                 = true;
			 boolean expheaderAlert                      = true;
			 	
			 System.out.println("****************************** checkAlertsTabOptions  ********************************************");
			 
			 System.out.println("ReceivedChkBox                :"+actreceivedChkBox         +"   Value Expected : "+expreceivedChkBox);
			 System.out.println("ApprovedChkBox                :"+actapprovedChkBox         +"   Value Expected : "+expapprovedChkBox);
			 System.out.println("RejectedChkBox                :"+actrejectedChkBox         +"	 Value Expected : "+exprejectedChkBox);
			 System.out.println("EmailHeaderTemplateBtn        :"+actemailHeaderTemplateBtn +"	 Value Expected : "+expemailHeaderTemplateBtn);
			 System.out.println("EmailTemplateBtn      	       :"+actemailTemplateBtn       +"	 Value Expected : "+expemailTemplateBtn);
			 System.out.println("HeaderAlert                   :"+actheaderAlert            +"	 Value Expected : "+expheaderAlert);
			 
			 boolean res = actreceivedChkBox==expreceivedChkBox&& actapprovedChkBox==expapprovedChkBox  && actrejectedChkBox==exprejectedChkBox 
					 && actemailHeaderTemplateBtn==expemailHeaderTemplateBtn 
					  && actemailTemplateBtn==expemailTemplateBtn && actheaderAlert==expheaderAlert ;
				
			 String actres = Boolean.toString(res);
			 String expres = "true";
					
			 
			  	
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Alerts TAb Options Are Displayed");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
				
				return false;
			}
		}
		
		
		
		
		public boolean 	checkCancelButtonInDefinitionScreen() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionCancelBtn));
			definitionCancelBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pricebookDropdown));
			
			boolean actbudgetDropdown                       = pricebookDropdown.isDisplayed();
			boolean actnameTxt                              = nameTxt.isDisplayed();
			boolean actactiveChkBox                         = activeChkBox.isDisplayed();
			boolean actalwaysauthorstartsfirstlevelChkBox   = alwaysauthorizationstartsfromfirstlevelChkBox.isDisplayed();
			boolean actaddLevelBtn                          = addLevelBtn.isDisplayed();
			boolean actaddParallelBtn                       = addParallelBtn.isDisplayed();
			boolean actclearbtn               				= clearbtnInAuth.isDisplayed();
			boolean actsaveBtn  							= saveBtn.isDisplayed();
			boolean actcancelBtn 							= closeBtn.isDisplayed();
			boolean actdeleteBtn			                = AuthdeleteBtn.isDisplayed();
			
			boolean expbudgetDropdown                       = true;
			boolean expnameTxt            					= true;
			boolean expactiveChkBox                			= true;
			boolean expalwaysauthorstartsfirstlevelChkBox   = true;
			boolean expaddLevelBtn          				= true;
			boolean expaddParallelBtn              			= true;
			boolean expclearbtn               				= true;
			boolean expsaveBtn  							= true;
			boolean expcancelBtn							= true;
			boolean expdeleteBtn							= true;
			
			System.out.println("********************************************** checkCancelButtonInDefinitionScreen ******************");
		
			System.out.println("Home Page budgetDropdown                     :"+actbudgetDropdown                      +"  Value Expected : "+expbudgetDropdown);
			System.out.println("Home Page nameTxt                            :"+actnameTxt                             +"  Value Expected : "+expnameTxt);
			System.out.println("Home Page activeChkBox                       :"+actactiveChkBox                        +"  Value Expected : "+expactiveChkBox);
			System.out.println("Home Page alwaysauthorstartsfirstlevelChkBox :"+actalwaysauthorstartsfirstlevelChkBox  +"  Value Expected : "+expalwaysauthorstartsfirstlevelChkBox);
	        System.out.println("Home Page addLevelBtn                        :"+actaddLevelBtn                         +"  Value Expected : "+expaddLevelBtn);
			System.out.println("Home Page addParallelBtn                     :"+actaddParallelBtn                      +"  Value Expected : "+expaddParallelBtn);
			System.out.println("Home Page clearbtn                           :"+actclearbtn                            +"  Value Expected : "+expclearbtn);
			System.out.println("Home Page saveBtn                            :"+actsaveBtn                             +"  Value Expected : "+expsaveBtn);
			System.out.println("Home Page cancelBtn 						 :"+actcancelBtn                           +"  Value Expected : "+expcancelBtn);
			System.out.println("Home Page deleteBtn                          :"+actdeleteBtn                           +"  Value Expected : "+expdeleteBtn);
			
			
			 boolean res =  actbudgetDropdown==expbudgetDropdown  && actnameTxt==expnameTxt &&
						actactiveChkBox==expactiveChkBox && actalwaysauthorstartsfirstlevelChkBox==expalwaysauthorstartsfirstlevelChkBox && 
						actaddLevelBtn==expaddLevelBtn  && actaddParallelBtn==expaddParallelBtn && 
					    actclearbtn==expclearbtn && actsaveBtn==expsaveBtn && actcancelBtn==expcancelBtn && actdeleteBtn==expdeleteBtn ;
				
			 String actres = Boolean.toString(res);
			 String expres = "true";
					
						  	
			if(actres.equalsIgnoreCase(expres))
				
			{
				System.out.println("Test Pass : Cancel Option Working ");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Cancel Option Not Working");
				return false;
			}
		}
		
		
		
		public boolean 	checkDescriptionAsMandatory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			Thread.sleep(2000);
			
		//	getAction().doubleClick(startPointBtn).build().perform();
			
			ClickUsingJs(startPointBtn);
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			descriptionTxt.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			System.out.println("****************************checkDescriptionAsMandatory************************************"); 
		
			
			String expMessage="Description is cannot be blank";
			
			String actMessage=checkValidationMessage(expMessage);
			
			
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				
				return true;
			}
			else
			{
				System.out.println(" NO Error Message");
				
				return false;
			}
		}
		
		
		
		
		
		public boolean 	checkConditionAsMandatory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
						
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			descriptionTxt.click();
			
			descriptionTxt.sendKeys("SellerPriceBook");
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			System.out.println("*****************************************************checkConditionAsMandatory************************************");
			
			String expMessage="In User Selection tab select user , User selection is mandatory";
			
			String actMessage=checkValidationMessage(expMessage);
			
			
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				
				return true;
			}
			else
			{
				System.out.println(" NO Error Message");
				
				return false;
			}
		}
		
		
		
		public boolean 	checkUserSelectionIsManadatory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
			conditionNotRequiredChkBox.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			System.out.println("*******************************************************checkUserSelectionIsManadatory*****************************************");
			
			String expMessage="In User Selection tab select user , User selection is mandatory";
			
			String actMessage=checkValidationMessage(expMessage);
			
			
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				
				return true;
			}
			else
			{
				System.out.println(" NO Error Message");
				
				return false;
			}
		}
		
				
		public boolean 	checkOkButtonOnSelectingUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
			userSelectionTab.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
			userrow1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userDropdown));
			userDropdown.click();
			
			userDropdown.sendKeys("UserAllOptionsST");
			
			/*Select s=new Select(userDropdown);
			s.selectByVisibleText(excelReader.getCellData("BudgetAuthorization", 45, 6));*/
			
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			Thread.sleep(2000);
			
			boolean Endpoint=endpoint.isDisplayed();


			String actendpoint = Boolean.toString(Endpoint);
			
			String expendpoint = "true";
			
			System.out.println("***************************************checkOkButtonOnSelectingUser*****************************");
			System.out.println("endpoint     : "+actendpoint +"  Value Expected : "+expendpoint);
			
			if(actendpoint.equalsIgnoreCase(expendpoint))
			{
				System.out.println("Test Pass : Image Displayed In Design Area");
			
				return true;
			}
			else
			{
				System.out.println("Test Fail : Image NOT Displayed In Design Area");
			
				return false;
			}
		}
		
		
		
		
		
		
		public boolean 	checkSavingSellingPriceBookAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{		
			
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(2000);
			System.out.println("**********************************checkSavingDefineBudgetAuthorization*****************************");
			
			String expMessage="Authorization Flow saved successfully.";
			
			String actMessage=checkValidationMessage(expMessage);
			
				
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				
				return true;
			}
			else
			{
				System.out.println(" NO Error Message");
				
				return false;
			}
		}
		
		
		
		
		public boolean 	checkCancelButton() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
			closeBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboard));
			
			boolean Dashboard=dashboard.isDisplayed();
			
			String actdashboard = Boolean.toString(Dashboard);
			
			String expdashboard = "true";	
			
			System.out.println("********************************checkCancelButton*************************");
			
			System.out.println("dashboard  :"+actdashboard +"  Value Expected  : "+expdashboard);
			
			if(actdashboard.equalsIgnoreCase(expdashboard))
				
			{
				System.out.println("Test Pass : Cancel Option Working ");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Cancel Option Not Working");
				
				return false;
			}
		}

	
		
		public boolean 	checkDisplayingMastersNameInListOfMasters() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(4000);
			getDriver().navigate().refresh();
			Thread.sleep(4000);
			
			getAction().moveToElement(settingsmenuBtn).build().perform();
			
		/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
			settingsmenuBtn.click();*/
			
			ClickUsingJs(settingsmenuBtn);
			
			Thread.sleep(2000);
			
			
		/*	((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
			Thread.sleep(2000);*/
			
			
			getAction().moveToElement(pricebookAuthorizationBtn).build().perform();
			Thread.sleep(2000);
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pricebookAuthorizationBtn));
			pricebookAuthorizationBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pricebookDropdown));
			pricebookDropdown.click();
			pricebookDropdown.sendKeys(Keys.SPACE);
			Thread.sleep(2000);
			boolean sellingPricebook=createSellingPriceBook.isDisplayed();


			String actbuyingPricebook = Boolean.toString(sellingPricebook);
			String expbuyingPricebook = "true";
			System.out.println("***********************************************checkDisplayingMastersNameInListOfMasters*******");
			
			System.out.println("Buying Price Book       : "+actbuyingPricebook+"  Value Expected : "+expbuyingPricebook);
			
			if(actbuyingPricebook.equalsIgnoreCase(expbuyingPricebook))
			{
				System.out.println("Test Pass : Selling Price Book(1) Is Displayed");
			
				return true;
			}
			else
			{
				System.out.println("Test Fail : Buying Price Book(1)) Is NOT Displayed");
				
				return false;
			}
		}
			
		
		
		public boolean 	checkSavedNameDisplayInNameTextBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createSellingPriceBook));
			createSellingPriceBook.click();
			Thread.sleep(6000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			nameTxt.sendKeys("Seller");
			
			Thread.sleep(1500);
			
			nameTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(10000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
			boolean actstartPointBtn     = startPointBtn.isDisplayed();
			boolean actendpoint          = endpoint.isDisplayed();
			
			boolean expstartPointBtn     = true;
			boolean expendpoint          = true;
			
			boolean res = actstartPointBtn==expstartPointBtn && actendpoint==expendpoint;
			
			String actres = Boolean.toString(res);
			
			String expres = "true";
			
			System.out.println("*************************************checkSavedNameDisplayInNameTextBox************************************");
			
			System.out.println("startPointBtn    :"+actstartPointBtn   +"  Value Expected : "+expstartPointBtn);
			System.out.println("endpoint         :"+actendpoint        +"  Value Expected : "+expendpoint);
			
			if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Saved Data is Displayed");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Saved Data is NOT Displayed");
				
				return false;
			}
			
		}
		

		@FindBy(xpath="//*[@id='divLevel']//*[@class='NodeIcon icon-report-1']")
		private static WebElement levelBtn;
		
		@FindBy(xpath="//div[@id='divLevel1']//i[@class='icon-edit']")
		private static WebElement Level1editBtn;
		

		public boolean 	checkEditPriceBookAuthorizationSettings() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			Thread.sleep(4500);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1Icon));
			getAction().doubleClick(Level1Icon).build().perform();

			Thread.sleep(2000);
			
			
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionTab));
	        boolean actconditionTab                 = conditionTab.isDisplayed();
			boolean actuserSelectionTab             = userSelectionTab.isDisplayed();
			boolean actalertsTab                    = alertsTab.isDisplayed();
			
			boolean expconditionTab                 = true;
			boolean expuserSelectionTab             = true;
			boolean expalertsTab                    = true;
			
			boolean res = actconditionTab==expconditionTab  && actuserSelectionTab==expuserSelectionTab && actalertsTab==expalertsTab;
			
			String actres = Boolean.toString(res);
			
			String expres = "true";
			
			
			
			System.out.println("********************* checkEditBudgetAuthorizationSettings ******************");
		
			System.out.println("ConditionTab             :"+actconditionTab     +"  Value Expected : "+expconditionTab);
			System.out.println("UserSelectionTab         :"+actuserSelectionTab +"  Value Expected : "+expuserSelectionTab);
			System.out.println("AlertsTab                :"+actalertsTab        +"	Value Expected : "+expalertsTab);
			
			if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Edit Btn Opened definition Screen");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Edit Btn NOT Opened definition Screen");
				
				return false;
			}
			
		}
		
			
		public boolean checkConditionCheckbox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			Thread.sleep(10000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1Icon));
			getAction().doubleClick(Level1Icon).build().perform();

			Thread.sleep(20000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Auth_dateRangeChkBox));
			if(Auth_dateRangeChkBox.isSelected()==true)
			{
				Auth_dateRangeChkBox.click();
			}
			
			getAction().moveToElement(conditionNotRequiredChkBox).build().perform();
			Thread.sleep(6000);
			
			if(conditionNotRequiredChkBox.isSelected()==true)
			{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
			conditionNotRequiredChkBox.click();
			}
			Thread.sleep(2000);
			
			
		/*	((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
			Thread.sleep(2000);*/
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
			conjunctionDropdown.click();
			Thread.sleep(1500);
			Select s=new Select(conjunctionDropdown);
			s.selectByValue("0");
			//conjunctionDropdown.sendKeys("where");
			//conjunctionDropdown.sendKeys(Keys.TAB);
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
			selectFieldTxt.click();
			Thread.sleep(2000);
			
			selectFieldTxt.click();
			Thread.sleep(2000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldRate));
			ClickUsingJs(selectFieldRate);
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
			selectOperatorDropdown.click();
			selectOperatorDropdown.sendKeys("Greater than");
			selectOperatorDropdown.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
			compareWithDropdown.click();
			compareWithDropdown.sendKeys("value");
			Thread.sleep(2000);
			//compareWithDropdown.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_valueTxt));
			enter_valueTxt.click();;
			enter_valueTxt.sendKeys("25.00");
			Thread.sleep(1500);
			/*enter_valueTxt.sendKeys(Keys.TAB);
			Thread.sleep(3000);*/
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
			Select s1=new Select(conjunctionDropdown);
			String actconjunctionDropdown=s1.getFirstSelectedOption().getText();
			String expconjunctionDropdown="Where";
			
			
			String actselectFieldTxt=selectFieldTxt.getAttribute("value");
			String expselectFieldTxt="Rate";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
			Select s2=new Select(selectOperatorDropdown);
			String actselectOperatorDropdown=s2.getFirstSelectedOption().getText();
			String expselectOperatorDropdown="Greater than";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
			Select s3=new Select(compareWithDropdown);
			String actcompareWithDropdown=s3.getFirstSelectedOption().getText();
			String expcompareWithDropdown="value";
			
			String actvalueTxt=enter_valueTxt.getAttribute("value");
			String expvalueTxt="25.00";
			
			Thread.sleep(2000);
			
/*//			getAction().moveToElement(definitionOkBtn).build().perform();
//			Thread.sleep(2000);
			ScrollIntoView(definitionOkBtn);
			Thread.sleep(2000);
			getAction().moveToElement(definitionOkBtn).build().perform();
			Thread.sleep(2000);*/
			
			getAction().moveToElement(definitionOkBtn).build().perform();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			Thread.sleep(8000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(4000);
			System.out.println("**********************************checkSavingAuthorization*****************************");
			
			String expMessage="Authorization Flow saved successfully.";
			
			String actMessage=checkValidationMessage(expMessage);
			
			
			
			System.out.println("************************************checkConditionCheckbox******************************************************************************");
			
			System.out.println("conjunctionDropdown     :"+actconjunctionDropdown        +" Value Expected  : "+expconjunctionDropdown); 
			System.out.println("selectFieldTxt          :"+actselectFieldTxt             +" Value Expected  : "+expselectFieldTxt);
			System.out.println("selectOperatorDropdown  :"+actselectOperatorDropdown     +"  Value Expected : "+expselectOperatorDropdown); 
			System.out.println("compareWithDropdown     :"+actcompareWithDropdown        +" Value Expected  : "+expcompareWithDropdown);
			System.out.println("valueTxt                :"+actvalueTxt                   +" Value Expected  : "+expvalueTxt);
			System.out.println("Saved                	:"+actMessage                    +" Value Expected  : "+expMessage);
			
			
			
			if(actconjunctionDropdown.equals(expconjunctionDropdown) && actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt) && 
					actselectOperatorDropdown.equals(expselectOperatorDropdown)
					&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueTxt.equalsIgnoreCase(expvalueTxt)
					&& actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Test Pass : Data entered");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT entered");
				
				return false;
			}
			
		}
		

		
		public boolean checkConditionCheckboxforBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			Thread.sleep(3000);

			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Auth_dateRangeChkBox));
			if(Auth_dateRangeChkBox.isSelected()==true)
			{
				Auth_dateRangeChkBox.click();
			}
			
			
			
			if(conditionNotRequiredChkBox.isSelected()==true)
			{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
			conditionNotRequiredChkBox.click();
			}
			Thread.sleep(2000);
			
			
			((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
			conjunctionDropdown.click();
			Thread.sleep(1500);
			Select s=new Select(conjunctionDropdown);
			s.selectByValue("0");
			//conjunctionDropdown.sendKeys("where");
			//conjunctionDropdown.sendKeys(Keys.TAB);
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
			selectFieldTxt.click();
			Thread.sleep(2000);
			selectFieldTxt.sendKeys(Keys.HOME,Keys.SHIFT);
			selectFieldTxt.sendKeys(Keys.DELETE);
			Thread.sleep(2000);
			selectFieldTxt.click();
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldRate));
			ClickUsingJs(selectFieldRate);
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
			selectOperatorDropdown.click();
			selectOperatorDropdown.sendKeys("Greater than");
			selectOperatorDropdown.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
			compareWithDropdown.click();
			compareWithDropdown.sendKeys("value");
			compareWithDropdown.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ProdvalueTxt));
			enter_ProdvalueTxt.clear();
			enter_ProdvalueTxt.sendKeys("25.00");
			Thread.sleep(1500);
			/*enter_valueTxt.sendKeys(Keys.TAB);
			Thread.sleep(3000);*/
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
			Select s1=new Select(conjunctionDropdown);
			String actconjunctionDropdown=s1.getFirstSelectedOption().getText();
			String expconjunctionDropdown="Where";
			
			
			String actselectFieldTxt=selectFieldTxt.getAttribute("value");
			String expselectFieldTxt="Rate";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
			Select s2=new Select(selectOperatorDropdown);
			String actselectOperatorDropdown=s2.getFirstSelectedOption().getText();
			String expselectOperatorDropdown="Greater than";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
			Select s3=new Select(compareWithDropdown);
			String actcompareWithDropdown=s3.getFirstSelectedOption().getText();
			String expcompareWithDropdown="value";
			
			String actvalueTxt=enter_valueTxt.getAttribute("value");
			String expvalueTxt="25.00";
			
			Thread.sleep(2000);
			
/*//			getAction().moveToElement(definitionOkBtn).build().perform();
//			Thread.sleep(2000);
			ScrollIntoView(definitionOkBtn);
			Thread.sleep(2000);
			getAction().moveToElement(definitionOkBtn).build().perform();
			Thread.sleep(2000);*/
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(2000);
			System.out.println("**********************************checkSavingAuthorization*****************************");
			
			String expMessage="Authorization Flow saved successfully.";
			
			String actMessage=checkValidationMessage(expMessage);
			
			
			
			System.out.println("************************************checkConditionCheckbox******************************************************************************");
			
			System.out.println("conjunctionDropdown     :"+actconjunctionDropdown        +" Value Expected  : "+expconjunctionDropdown); 
			System.out.println("selectFieldTxt          :"+actselectFieldTxt             +" Value Expected  : "+expselectFieldTxt);
			System.out.println("selectOperatorDropdown  :"+actselectOperatorDropdown     +"  Value Expected : "+expselectOperatorDropdown); 
			System.out.println("compareWithDropdown     :"+actcompareWithDropdown        +" Value Expected  : "+expcompareWithDropdown);
			System.out.println("valueTxt                :"+actvalueTxt                   +" Value Expected  : "+expvalueTxt);
			System.out.println("Saved                	:"+actMessage                    +" Value Expected  : "+expMessage);
			
			
			
			if(actconjunctionDropdown.equals(expconjunctionDropdown) && actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt) && 
					actselectOperatorDropdown.equals(expselectOperatorDropdown)
					&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueTxt.equalsIgnoreCase(expvalueTxt)
					&& actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Test Pass : Data entered");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT entered");
				
				return false;
			}
			
		}
		

		
		

		public boolean 	checkEditOnSelectAlertsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			 Thread.sleep(6000);
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertTabInEdit));
			 alertTabInEdit.click();
			
			 Thread.sleep(2000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receivedChkBox));
			 boolean actreceivedChkBox                   = receivedChkBox.isDisplayed();
		     boolean actapprovedChkBox                   = approvedChkBox.isDisplayed();
			 boolean actrejectedChkBox                   = rejectedChkBox.isDisplayed();
			 boolean actemailHeaderTemplateBtn           = emailHeaderTemplateBtn.isDisplayed();
			 boolean actemailTemplateBtn                 = emailTemplateBtn.isDisplayed();
			 boolean actheaderAlert                      = headerAlert.isDisplayed();
			 
			 boolean expreceivedChkBox                   = true;
			 boolean expapprovedChkBox                   = true;
			 boolean exprejectedChkBox                   = true;
			 boolean expemailHeaderTemplateBtn           = true;
			 boolean expemailTemplateBtn                 = true;
			 boolean expheaderAlert                      = true;
			 	
			 System.out.println("****************************** checkEditOnSelectAlertsTabOptions  ********************************************");
			 
			 System.out.println("ReceivedChkBox                :"+actreceivedChkBox         +"   Value Expected : "+expreceivedChkBox);
			 System.out.println("ApprovedChkBox                :"+actapprovedChkBox         +"   Value Expected : "+expapprovedChkBox);
			 System.out.println("RejectedChkBox                :"+actrejectedChkBox         +"	 Value Expected : "+exprejectedChkBox);
			 System.out.println("EmailHeaderTemplateBtn        :"+actemailHeaderTemplateBtn +"	 Value Expected : "+expemailHeaderTemplateBtn);
			 System.out.println("EmailTemplateBtn      	       :"+actemailTemplateBtn       +"	 Value Expected : "+expemailTemplateBtn);
			 System.out.println("HeaderAlert                   :"+actheaderAlert            +"	 Value Expected : "+expheaderAlert);
			  	
			 
			 boolean res = actreceivedChkBox==expreceivedChkBox&& actapprovedChkBox==expapprovedChkBox  && actrejectedChkBox==exprejectedChkBox 
					 && actemailHeaderTemplateBtn==expemailHeaderTemplateBtn 
					  && actemailTemplateBtn==expemailTemplateBtn && actheaderAlert==expheaderAlert;
				
			 String actres = Boolean.toString(res);
			 
			 String expres = "true";
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Alerts TAb Options Are Displayed");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
				
				return false;
			}
		}

		public boolean checkSaveWithAllMandatoryInputs() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(8000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailHeaderTemplateBtn));
			emailHeaderTemplateBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTemplates));
			emailTemplates.sendKeys("SellerPricebookTemplate");

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyNameTemplateArea));
			companyNameTemplateArea.click();
			Thread.sleep(2000);
			
			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\EmailHeaderTemplate1.exe");
			
			Thread.sleep(150000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtnInEmailTemplateHeader));
			saveBtnInEmailTemplateHeader.click();
			Thread.sleep(2000);
			
			System.out.println("**************************************checkSaveWithAllMandatoryInputs*******************************");
			
			String expMessage="Template saved successfully";
			 
			String actMessage=checkValidationMessage(expMessage);
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeButtonInCreateTemplate));
			//closeButtonInCreateTemplate.click();
			ClickUsingJs(closeButtonInCreateTemplate);
			Thread.sleep(2000);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Test Pass : Alerts TAb Options Are Displayed");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
				
				return false;
			}
		}
		
		
		public boolean checkSaveWithAllMandatoryInputsofBuyerAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailHeaderTemplateBtn));
			emailHeaderTemplateBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTemplates));
			emailTemplates.sendKeys("BuyerPricebookTemplate");

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyNameTemplateArea));
			companyNameTemplateArea.click();
			
			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\EmailPricebookAuth.exe");
			
			Thread.sleep(85000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtnInEmailTemplateHeader));
			saveBtnInEmailTemplateHeader.click();
			Thread.sleep(2000);
			
			System.out.println("**************************************checkSaveWithAllMandatoryInputs*******************************");
			
			String expMessage="Template saved successfully";
			 
			String actMessage=checkValidationMessage(expMessage);
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeButtonInCreateTemplate));
			//closeButtonInCreateTemplate.click();
			
			ClickUsingJs(closeButtonInCreateTemplate);
			Thread.sleep(2000);
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Test Pass : Alerts TAb Options Are Displayed");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
				
				return false;
			}
		}
		
		
		
		@FindBy(xpath="//td[@id='id_tblAlerts_col_1-2']")
		private static WebElement recievedEmailHeaderTemplate;
		
		@FindBy(xpath="//select[@id='id_tblAlerts_EmailHeaderTemplate']")
		private static WebElement receivedEmailHeaderTemplateInput;
		
		
		public boolean checkCreatedEmailInEmailHedaerTemplate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			if(receivedChkBox.isSelected()==false)
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receivedChkBox));
				receivedChkBox.click();
			}
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(recievedEmailHeaderTemplate));
			recievedEmailHeaderTemplate.click();
			Thread.sleep(2000);
			
			receivedEmailHeaderTemplateInput.click();
			Thread.sleep(2000);
			Select s=new Select(receivedEmailHeaderTemplateInput);
			s.selectByVisibleText("SellerPricebookTemplate");
			Thread.sleep(2000);
			receivedEmailHeaderTemplateInput.sendKeys(Keys.TAB);
			
			
			
			Thread.sleep(2000);
			
					
			 String actreceivedEmailHeaderTemplateInput=s.getFirstSelectedOption().getText();
			
			String expreceivedEmailHeaderTemplateInput="SellerPricebookTemplate";
			
		
			
			System.out.println("actreceivedEmailHeaderTemplateInput : "+actreceivedEmailHeaderTemplateInput);
			System.out.println("expreceivedEmailHeaderTemplateInput : "+expreceivedEmailHeaderTemplateInput);
			
			System.out.println("**************************************checkCreatedEmailInEmailHedaerTemplate*********************");
			
			System.out.println(" receivedEmailHeaderTemplateInput  :"+actreceivedEmailHeaderTemplateInput+"  Value Expected : "+expreceivedEmailHeaderTemplateInput);
			
			if(actreceivedEmailHeaderTemplateInput.equalsIgnoreCase(expreceivedEmailHeaderTemplateInput))
			{
				System.out.println("Test Pass : Alerts TAb Options Are Displayed");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
				
				return false;
			}
		}
		
		
		public boolean checkCreatedEmailInEmailHedaerTemplateofBuyerAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			if(receivedChkBox.isSelected()==false)
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receivedChkBox));
				receivedChkBox.click();
			}
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(recievedEmailHeaderTemplate));
			recievedEmailHeaderTemplate.click();
			Thread.sleep(2000);
			
			receivedEmailHeaderTemplateInput.click();
			Thread.sleep(2000);
			Select s=new Select(receivedEmailHeaderTemplateInput);
			s.selectByVisibleText("BuyerPricebookTemplate");
			Thread.sleep(2000);
			receivedEmailHeaderTemplateInput.sendKeys(Keys.TAB);
			
			
			
			Thread.sleep(2000);
			
					
			 String actreceivedEmailHeaderTemplateInput=s.getFirstSelectedOption().getText();
			
			String expreceivedEmailHeaderTemplateInput="BuyerPricebookTemplate";
			
		
			
			System.out.println("actreceivedEmailHeaderTemplateInput : "+actreceivedEmailHeaderTemplateInput);
			System.out.println("expreceivedEmailHeaderTemplateInput : "+expreceivedEmailHeaderTemplateInput);
			
			System.out.println("**************************************checkCreatedEmailInEmailHedaerTemplate*********************");
			
			System.out.println(" receivedEmailHeaderTemplateInput  :"+actreceivedEmailHeaderTemplateInput+"  Value Expected : "+expreceivedEmailHeaderTemplateInput);
			
			if(actreceivedEmailHeaderTemplateInput.equalsIgnoreCase(expreceivedEmailHeaderTemplateInput))
			{
				System.out.println("Test Pass : Alerts TAb Options Are Displayed");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
				
				return false;
			}
		}
		
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/i")
		private static WebElement companyExpandBtn;
		
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']//li[1]//li/a")
		private static List<WebElement> companyExpandBtnList;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/i")
		private static WebElement userDetailsExpandBtn;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']//li[3]//li/a")
		private static List<WebElement> userDetailsExpandBtnList;
		
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[4]/i")
		private static WebElement miscellaneousExpandBtn;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[4]//li/a")
		private static List<WebElement> miscellaneousExpandBtnList;

		
		
		
		@FindBy(xpath="//span[contains(text(),'Users')]//i[@class='icon-expand icon-font8']")
		private static WebElement userExpandBtn;


		@FindBy(xpath="//span[contains(text(),'Users')]//i[@class='icon-collepse icon-font8']")
		private static WebElement userMinimizeBtn;


		@FindBy(xpath="//span[contains(text(),'Logged in Username')]")
		private static WebElement  user_LoggedInUserNameDragandDrop;

		@FindBy(xpath="//span[contains(text(),'Logged in Login Abbreviation')]")
		private static WebElement  user_LoggedInLoginAbbreviationDragandDrop;

		@FindBy(xpath="//span[contains(text(),'End User')]")
		private static WebElement  user_EndUserDragandDrop;

		@FindBy(xpath="//i[@class='icon-new hiconright2']")
		private static WebElement  Newdropdown;

		public boolean checkEmailTemplateExpandInFieldsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTemplateBtn));
		 	 emailTemplateBtn.click();
		
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Newdropdown));
			
			
			
		
			 Thread.sleep(3000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyExpandBtn));
			 boolean actCompanyExpBtn             			= companyExpandBtn.isDisplayed();
		     boolean actPricebookFieldsExpBtn    			= pricebookFieldsExpdBtn.isDisplayed();
			 boolean actUserExpdBtn                     	 = userDetailsExpandBtn.isDisplayed();
			 boolean actMiscellExpdBtn                    	  = miscellaneousExpandBtn.isDisplayed();
			
			 
			 boolean expCompanyExpBtn             			= true;
		     boolean expPricebookFieldsExpBtn    			= true;
			 boolean expUserExpdBtn                     	= true;
			 boolean expMiscellExpdBtn                    	= true;
			
				
			 System.out.println("****************************** checkEmailTemplateUserExpandInFieldsTabOptions  ********************************************");
			 
			 System.out.println("Company Expd Btn          		:"+		actCompanyExpBtn           			+"    Value Expected : 		"+		expCompanyExpBtn);
			 System.out.println("Pricebook Fields Expd Btn		:"+		actPricebookFieldsExpBtn  			+"    Value Expected : 		"+		expPricebookFieldsExpBtn);
			 System.out.println("User Expd Btn                  :"+		actUserExpdBtn                   	+"	  Value Expected :		 "+		expUserExpdBtn);
			 System.out.println("Miscellanious Expd Btn         :"+		actMiscellExpdBtn                   +"	  Value Expected : 		"+		expMiscellExpdBtn);
			 
			 boolean res = actCompanyExpBtn==expCompanyExpBtn && actPricebookFieldsExpBtn==expPricebookFieldsExpBtn
					 && actUserExpdBtn==expUserExpdBtn && actMiscellExpdBtn== expMiscellExpdBtn;
				
			 String actres = Boolean.toString(res);
			 
			 String expres = "true";
			 
			
			 
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Fields Options Are Displayed");
				
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Fields Options Are NOT Displayed");
				
				
				return false;
			}
		}

		//@FindBy(xpath="//span[contains(text(),'Pricebook Fields')]//i[@class='icon-expand icon-font8']")
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/i")
		private static WebElement pricebookFieldsExpdBtn;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']//li[2]//li/a")
		private static List<WebElement> pricebookFieldsExpdBtnList;
		
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/a/span")
		private static WebElement  pricebook_pricebookName;

		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[2]/a/span")
		private static WebElement  pricebook_productCode;

		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[3]/a/span")
		private static WebElement  pricebook_productName;

		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[4]/a/span")
		private static WebElement  pricebook_accountCode;

		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[5]/a/span")
		private static WebElement  pricebook_AccountNameDragandDrop;

		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[6]/a/span")
		private static WebElement  pricebook_Unit;

		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[7]/a/span")
		private static WebElement  pricebook_currency;

		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[8]/a/span")
		private static WebElement  pricebook_Tag;

		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[9]/a/span")
		private static WebElement  pricebook_StartDate;

		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[10]/a/span")
		private static WebElement  pricebook_EndDate;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[11]/a/span")
		private static WebElement  pricebook_MinQty;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[12]/a/span")
		private static WebElement  pricebook_MaxQty;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[13]/a/span")
		private static WebElement  pricebook_Rate;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[14]/a/span")
		private static WebElement  pricebook_Val1;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[15]/a/span")
		private static WebElement  pricebook_Val2;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[16]/a/span")
		private static WebElement  pricebook_Val3;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[17]/a/span")
		private static WebElement  pricebook_Val4;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[18]/a/span")
		private static WebElement  pricebook_Val5;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[19]/a/span")
		private static WebElement  pricebook_Val6;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[20]/a/span")
		private static WebElement  pricebook_Val7;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[21]/a/span")
		private static WebElement  pricebook_Val8;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[22]/a/span")
		private static WebElement  pricebook_Val9;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[23]/a/span")
		private static WebElement  pricebook_Val10;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[24]/a/span")
		private static WebElement  pricebook_Val11;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[25]/a/span")
		private static WebElement  pricebook_Val12;
		
		@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[26]/a/span")
		private static WebElement  pricebook_Val13;
		
		
public boolean checkFieldExpansionDetails() throws InterruptedException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyExpandBtn));
	companyExpandBtn.click();
	Thread.sleep(2000);
	
	int c=companyExpandBtnList.size();
	ArrayList<String> companyList=new ArrayList<String>();
	
	for(int i=0;i<c;i++)
	{
		String data=companyExpandBtnList.get(i).getText();
		companyList.add(data);
	}
	
	String actCompanyList=companyList.toString();
	String expCompanyList="[Company Name, Company Code, Accounting Date, Address1, Address2, City, Zip Code, Country, Phone No1, Phone No2, Fax, Tax Registration Number, Excise Registration Number]";
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyExpandBtn));
	companyExpandBtn.click();
	Thread.sleep(2000);
		
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pricebookFieldsExpdBtn));
	pricebookFieldsExpdBtn.click();
	Thread.sleep(2000);
	
	int c1=pricebookFieldsExpdBtnList.size();
	ArrayList<String> pricebookFieldsList=new ArrayList<String>();
	
	for(int i=0;i<c1;i++)
	{
		String data=pricebookFieldsExpdBtnList.get(i).getText();
		pricebookFieldsList.add(data);
	}
	
	String actPricebookList=pricebookFieldsList.toString();
	String expPricebookList="[PriceBookName, ProductCode, ProductName, AccountCode, AccountName, Unit, Currency, Department, StartDate, EndDate, MinQty, MaxQty, Rate, Val 1, Val 2, Val3, Val4, Val5, Val6, Val7, Val8, Val9, Val10, Val11, Val12, Val13]";
	
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pricebookFieldsExpdBtn));
	pricebookFieldsExpdBtn.click();
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userDetailsExpandBtn));
	userDetailsExpandBtn.click();
	Thread.sleep(2000);
	
	int c2=userDetailsExpandBtnList.size();
	ArrayList<String> userDetailsList=new ArrayList<String>();
	
	for(int i=0;i<c2;i++)
	{
		String data=userDetailsExpandBtnList.get(i).getText();
		userDetailsList.add(data);
	}
	
	String actUserList=userDetailsList.toString();
	String expUserList="[User name, Mobile number, Phone number, E-Mail Id, Signature, Employee, Authorise User name, Authorise User Mobile number, Authorise User Phone number, Authorise User E-Mail Id, Authorise User Signature, Authorise by, Authorize remarks, Next Authorise User name, Next User Mobile number, Next User Phone number, Next User E-Mail Id, Next User Signature, Created by signature, Modified by signature]";
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userDetailsExpandBtn));
	userDetailsExpandBtn.click();
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousExpandBtn));
	miscellaneousExpandBtn.click();
	Thread.sleep(2000);
	
	int c3=miscellaneousExpandBtnList.size();
	ArrayList<String> miscList=new ArrayList<String>();
	
	for(int i=0;i<c3;i++)
	{
		String data=miscellaneousExpandBtnList.get(i).getText();
		miscList.add(data);
	}
	
	String actmiscList=miscList.toString();
	String expmiscList="[Login date, Page Number, Print Date, Print Time, Total Page, S.No, Page Type, Ageing details, Ageing amounts, Ageing headings, Ageing balance, Ageing slab1, Ageing slab2, Ageing slab3, Ageing slab4, Ageing slab5, Ageing slab6, Ageing month-wise]";
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousExpandBtn));
	miscellaneousExpandBtn.click();
	Thread.sleep(2000);
	
	
	System.out.println("Comapny Expanssion					:"	+	actCompanyList		);
	System.out.println("Value Expected						:"	+	expCompanyList);
	System.out.println("Pricebook Fields  Expanssion		:"	+	actPricebookList	);
	System.out.println("Value Expected						:"	+	expPricebookList);
	
	System.out.println("User Details  Expanssion			:"	+	actUserList);
	System.out.println("Value Expected						:"	+	expUserList);
	
	System.out.println("Miscellanious  Expanssion			:"	+	actmiscList);
	System.out.println("Value Expected						:"	+	expmiscList);
	
	if(actPricebookList.equalsIgnoreCase(expPricebookList) && actCompanyList.equalsIgnoreCase(expCompanyList)
			&& actUserList.equalsIgnoreCase(expUserList) && actmiscList.equalsIgnoreCase(expmiscList))
	{
		return true;
	}
	else
	{
		return false;
	}
		
		
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


@FindBy(xpath="//*[@id='PriceBookWebGrid_val1']")
private static WebElement val1Value;

@FindBy(xpath="//*[@id='EndingDateH_checkbox']")
private static WebElement buyerPriceBook_EndDateChkBox;

@FindBy(xpath="//*[@id='StartingDateH']")
private static WebElement buyerPriceBook_StartDate;


@FindBy(xpath="//*[@id='EndingDateH']")
private static WebElement buyerPriceBook_EndDate;

public boolean checkCreateaSellerPriceBookafterAuthorizationSaved() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
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
		buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization");
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_StartDate));
		buyerPriceBook_StartDate.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_EndDateChkBox));
		if(buyerPriceBook_EndDateChkBox.isSelected()==false)
		{
			buyerPriceBook_EndDateChkBox.click();
		}
		buyerPriceBook_EndDateChkBox.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys("Pricebook Item1");
		Thread.sleep(2000);
		
/*		int count = itemComboBoxList.size();
		
		for(int i=0;i<count;i++)
		{
		String data = itemComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("Pricebook Item1"))
		{
		itemComboBoxList.get(i).click();
		break;
		}
		}
*/		
		itemValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("103");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("103.25");
		val1Value.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		System.out.println("*********************************checkSavingSellerPriceBook********************************");
		
		String expMessage="Price Book created successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		
		
		
		if (actMessage.equalsIgnoreCase(expMessage)  ) 
		{
		
		
		return true;
		} 
		else 
		{
		
		
		return false;
		}
		
		
}


		public boolean checkStatusofSellerPricebookafterAuthorizationSaved() throws InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
			//buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			buyerPriceBook_PriceBook.click();
			buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
			buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization");
			Thread.sleep(1500);
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
//			buyerPriceBook_PriceBook.click();
//			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
			buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
			Thread.sleep(2000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
			priceBookFilterAndLoadBtn.click();
			Thread.sleep(8000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
			itemFirstCellSelect.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
			itemValue.click();
			itemValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
		//	getAction().moveToElement(StatusFirstCellSelect).build().perform();
		//	ScrollToElement(StatusFirstCellSelect);
			getAction().moveToElement(StatusFirstCellSelect).build().perform();
			Thread.sleep(2000);
			String test=StatusFirstCellSelect.getText();
			String exp="Pending";
			
			
			System.out.println("Status :"	+	test	+"Expected"	+	exp);
			if(test.equalsIgnoreCase(exp))
			{
				return true;
			}
			else
			{
				return false;
			}
		}

		
		
		
		public boolean checkRateValueloadingbeforeAuthorizationPricebookinSalesInvoiceVoucher() throws InterruptedException
		{
			
			checkOpenSalesInvoiceVoucher();
			Thread.sleep(3000);
			
			if(checkAddingFormulaforAccount())
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
				newBtn.click();
			}

			
			Thread.sleep(2000);
			
			
			
			
			String actRate1=checkRateValue("Pricebook Item1");
			String expRate1="0.00";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			String item_Txt1=enter_ItemTxt.getAttribute("value");
			
			
			System.out.println("Item Name	:	"	+item_Txt1);
			System.out.println("Actual Rate	:	"	+actRate1	+"Expected	:"	+expRate1);
			
			if(actRate1.equalsIgnoreCase(expRate1))
			{
				return true;
			}
			else
			{
				return false;
			}
			
		}
		
		
		
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
			getAction().moveToElement(settingsUpdateBtn).build().perform();
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
		
		
		public static boolean checkAddingFormulaforGroupAccountofPurchase() throws InterruptedException
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
			editScreenFormulaTxtBox.sendKeys("&purrtgrp(0)");
			Thread.sleep(1000);


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFormulaOkBtn));
			editScreenFormulaOkBtn.click();
			
			Thread.sleep(2000);
			getAction().moveToElement(settingsUpdateBtn).build().perform();
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

			getAction().moveToElement(settingsUpdateBtn).build().perform();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsUpdateBtn));
			settingsUpdateBtn.click();
			Thread.sleep(6000);

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
		
		
		
		public static boolean checkAddingFormulaforAccountofPurchase() throws InterruptedException
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
			editScreenFormulaTxtBox.sendKeys("&purrt(0)");
			Thread.sleep(1000);


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenFormulaOkBtn));
			editScreenFormulaOkBtn.click();
			Thread.sleep(2000);

			getAction().moveToElement(settingsUpdateBtn).build().perform();
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
		
		
		
		@FindBy(xpath="//i[@class='icon-settings hiconright2']")
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
		
		
		public static String checkRateValueforPurchase(String item) throws InterruptedException
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

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
			select1stRow_6thColumn.click();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.click();
			Thread.sleep(1000);

			String actRate=enter_Rate.getAttribute("value");
			return actRate;
			
			
		}
		
		
		
		
		
		
		public static boolean checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
		{
			Thread.sleep(2000);
			
			getDriver().navigate().refresh();
			Thread.sleep(6000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
			userNameDisplayLogo.click();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			logoutOption.click();

			Thread.sleep(4000);



			LoginPage lp=new LoginPage(getDriver()); 

			String unamelt="UserAllOptionsST";

			String pawslt="12345";

			lp.enterUserName(unamelt);

			lp.enterPassword(pawslt);
			
			  String compname = "Pricebook Validation and Authorization";

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

			

			Thread.sleep(15000);
			

			String actUserInfo=userNameUserDisplay.getText();

			System.out.println("User Info  : "+actUserInfo);

			System.out.println("User Info Capture Text  :  "+userNameUserDisplay.getText());

			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
			companyLogo.click();*/

			String getCompanyTxt=Company_Name.getText();
			String getLoginCompanyName=getCompanyTxt.substring(0, 38);
			System.out.println("company name  :  "+ getLoginCompanyName);
			//companyLogo.click();


			String expUserInfo           ="UserAllOptionsST";
			String expLoginCompanyName   ="Pricebook Validation and Authorization";

			boolean actDashboard = dashboard.isDisplayed();
			boolean expDashboard = true;

			getDriver().navigate().refresh();

			Thread.sleep(2000);

			System.out.println("**********************************************checkLogoutAndLoginWithUserAllOptionsSTToAuthorizePurchaseOrder*********************************************");

			System.out.println("User Info               : "+actUserInfo         +" Value Expected "+expUserInfo);
			System.out.println("Login Company           : "+getLoginCompanyName +" Value Expected "+expLoginCompanyName);
			System.out.println("Dashboard is Displaying : "+actDashboard        +" Value Expected "+expDashboard );

			if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName) && actDashboard==expDashboard) 
			{
				return true;
			}
			else
			{
				return false;
			}
		}


		@FindBy(xpath="//*[@id='10']/span/i")
		public static WebElement alertExtraMaster;
		
		@FindBy(xpath="//*[@id='7']/span/i")
		public static WebElement SellerPricebookAuthorization;
		
		@FindBy(xpath="//*[@id='13']/span/i")
		public static WebElement BuyerPricebookAuthorization;
		
		@FindBy(xpath="//*[@id='masterChild_7']/a/span")
		public static WebElement SellerPricebookAuthorizationTxt;
		
		@FindBy(xpath="//*[@id='masterChild_13']/a/span")
		public static WebElement BuyerPricebookAuthorizationTxt;
		
		
		@FindBy(xpath="//*[@id='AuthorizeLI']")
		public static WebElement SellerPricebookAuthorizationBtn;
		
		//@FindBy(xpath="//*[@id='GridPBAuthorize_col_1-1']/input")
		@FindBy(xpath="//*[@id='GridPBAuthorize_body']/tr/td[2]/input")
		public static WebElement firstcheckBoxforAuthrozation;
		
		@FindBy(xpath="//*[@id='btnAuthorize']")
		public static WebElement authorizeBtn;
		
		@FindBy(xpath="//*[@id='btnRejectAuth']")
		public static WebElement rejectionBtn;
		
		@FindBy(xpath="//*[@id='btnStopAuth']")
		public static WebElement stopBtn;
		
		@FindBy(xpath="//*[@id='btnCancelAuth']")
		public static WebElement cancelBtn;
		
		
		
		
		@FindBy(xpath="//a[@title='Alerts']//i")
		public static WebElement alertsBtn;
		
		
		
		
	public boolean checkAuthorizationofSellerPricebookwithUserAllOptionsST() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
		alertsBtn.click();
		Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
		alertExtraMaster.click();
		Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorization));
		SellerPricebookAuthorization.click();
		Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationTxt));
		SellerPricebookAuthorizationTxt.click();

		Thread.sleep(35000);
		
		//new WebDriverWait(getDriver(), 1000).until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationBtn));
		
		try
		{
		getAction().moveToElement(SellerPricebookAuthorizationBtn).build().perform();
		getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationBtn));
		SellerPricebookAuthorizationBtn.click();
		}
		catch(Exception E)
		{
			
		}
		Thread.sleep(5000);
		
		getWebDriverWait().until(ExpectedConditions.visibilityOf(firstcheckBoxforAuthrozation));
		firstcheckBoxforAuthrozation.click();
		
		getWebDriverWait().until(ExpectedConditions.visibilityOf(authorizeBtn));
		authorizeBtn.click();
		Thread.sleep(2000);
		
		String expMsg="Authorized Successfully";
		String actMsg=checkValidationMessage(expMsg);
		Thread.sleep(4000);
		
		getAction().moveToElement(cancelBtn).build().perform();
		getWebDriverWait().until(ExpectedConditions.visibilityOf(cancelBtn));
		ClickUsingJs(cancelBtn);
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
		//	getAction().moveToElement(StatusFirstCellSelect).build().perform();
		//	ScrollToElement(StatusFirstCellSelect);
			getAction().moveToElement(StatusFirstCellSelect).build().perform();
			Thread.sleep(2000);
			String test=StatusFirstCellSelect.getText();
			String exp="Authorized";
			
			
			System.out.println("Status :"	+	test	+"Expected"	+	exp);
			if(test.equalsIgnoreCase(exp) && actMsg.equalsIgnoreCase(expMsg))
			{
				return true;
			}
			else
			{
				return false;
			}
		
		
	}
	public boolean  checkLogoutandloginSU() throws InterruptedException
	{
		 Thread.sleep(2000);
		getDriver().navigate().refresh();
		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		
		 Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(3000);

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt="su";

		String pawslt="su";

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);

		 Thread.sleep(2000);

	        String compname = "Pricebook Validation and Authorization";

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

		//checkRefershPopOnlogin();

		//checkPopUpWindow();

		Thread.sleep(6000);

		String actUserInfo=userNameDisplay.getText();

		System.out.println("User Info  : "+actUserInfo);

		System.out.println("User Info Capture Text  :  "+userNameDisplayLogo.getText());

		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		//companyLogo.click();

		String getCompanyTxt=Company_Name.getText();
		String getLoginCompanyName=getCompanyTxt.substring(0, 38);
		System.out.println("company name  :  "+ getLoginCompanyName);
		//companyLogo.click();


		String expUserInfo           ="SU";
		String expLoginCompanyName   ="Pricebook Validation and Authorization";

		boolean actDashboard = dashboard.isDisplayed();
		boolean expDashboard = true;

		System.out.println("**********************************************checkLogoutAndLoginWithSUForPurchasesVoucherVAT*********************************************");

		System.out.println("User Info               : "+actUserInfo         +" Value Expected "+expUserInfo);
		System.out.println("Login Company           : "+getLoginCompanyName +" Value Expected "+expLoginCompanyName);
		System.out.println("Dashboard is Displaying : "+actDashboard        +" Value Expected "+expDashboard );

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName) && actDashboard==expDashboard) 
		{
			return true;
		}
		else
		{
			return false;
		}
	
	}
	
	
	public boolean checkRateLoadingafterAuthorizeSellerPricebook() throws InterruptedException
	{
		Thread.sleep(2000);
		checkOpenSalesInvoiceVoucher();
		Thread.sleep(2000);
		
		if(checkAddingFormulaforAccount())
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		}

		
		Thread.sleep(2000);
		
			
		
		String actRate1=checkRateValue("Pricebook Item1");
		String expRate1="103.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		String item_Txt1=enter_ItemTxt.getAttribute("value");
		
		
		System.out.println("Item Name	:	"	+item_Txt1);
		System.out.println("Actual Rate	:	"	+actRate1	+"Expected	:"	+expRate1);
		
		if(actRate1.equalsIgnoreCase(expRate1))
		{
			return true;
		}
		else
		{
			return false;
		}
	}
		

	@FindBy(xpath="//*[@id='btnSelectAllAuth']")
	public static WebElement Auth_selectAllBtn;
	
	@FindBy(xpath="//*[@id='chkDateRange']")
	public static WebElement Auth_dateRangeChkBox;
	
	
	public boolean checkDateRangeCheckBoxinAuthorizationforBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	
	{
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1Icon));
		getAction().doubleClick(Level1Icon).build().perform();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Auth_dateRangeChkBox));
		if(Auth_dateRangeChkBox.isSelected()==false)
		{
			Auth_dateRangeChkBox.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
		definitionOkBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
		Thread.sleep(2000);
		System.out.println("**********************************checkSavingAuthorization*****************************");
		
		String expMessage="Authorization Flow saved successfully.";
		
		String actMessage=checkValidationMessage(expMessage);
		
			
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			System.out.println("Error Message is As Expected");
			
			return true;
		}
		else
		{
			System.out.println(" NO Error Message");
			
			return false;
		}
		
		
	}
	
	
public boolean checkDateRangeCheckBoxinAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	
	{
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Auth_dateRangeChkBox));
		if(Auth_dateRangeChkBox.isSelected()==false)
		{
			Auth_dateRangeChkBox.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
		definitionOkBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
		Thread.sleep(2000);
		System.out.println("**********************************checkSavingAuthorization*****************************");
		
		String expMessage="Authorization Flow saved successfully.";
		
		String actMessage=checkValidationMessage(expMessage);
		
			
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			System.out.println("Error Message is As Expected");
			
			return true;
		}
		else
		{
			System.out.println(" NO Error Message");
			
			return false;
		}
		
		
	}
	
	
	
	public boolean checkEditingthePricebookAfterClickingDateRangeinAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		homeMasterItemMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
		homeMasterItem_sellerPriceBookMenu.click();
		
		Thread.sleep(5000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization");
		Thread.sleep(3000);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(6000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_StartDate));
		buyerPriceBook_StartDate.click();
		buyerPriceBook_StartDate.sendKeys(Keys.TAB);
		
		
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_EndDateChkBox));
		if(buyerPriceBook_EndDateChkBox.isSelected()==true)
		{
			buyerPriceBook_EndDateChkBox.click();
			Thread.sleep(1000);
		}*/
	//	buyerPriceBook_EndDateChkBox.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys("Pricebook Item2");
		Thread.sleep(2000);
		

		itemValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("203");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("203.25");
		val1Value.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization");
		Thread.sleep(2000);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(6000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_StartDate));
		buyerPriceBook_StartDate.click();
		buyerPriceBook_StartDate.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_EndDateChkBox));
		if(buyerPriceBook_EndDateChkBox.isSelected()==true)
		{
			buyerPriceBook_EndDateChkBox.click();
			Thread.sleep(1000);
		}
		
		//buyerPriceBook_EndDateChkBox.sendKeys(Keys.TAB);
		Thread.sleep(2000);
			
		
		//getAction().moveToElement(itemThirdCellSelect).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
		itemThirdCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys("RMA STITEM");
		Thread.sleep(2000);
		

		itemValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("303");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("403.25");
		val1Value.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		System.out.println("*********************************checkSavingSellerPriceBook********************************");
		
		String expMessage="Price Book updated successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		
		
		
		if (actMessage.equalsIgnoreCase(expMessage)  ) 
		{
		
		
		return true;
		} 
		else 
		{
		
		
		return false;
		}
		
		
		
	}
	
	
	public boolean checkStatusofeditPricebook() throws InterruptedException
	{

		
		Thread.sleep(3000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
	
		getAction().moveToElement(StatusSecondCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusSecondCellSelect.getText();
		String expStatus="Pending";
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
		itemThirdCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusThirdCellSelect));
	
		ScrollIntoView(StatusThirdCellSelect);
		Thread.sleep(2000);
		String actStatus1=StatusThirdCellSelect.getText();
		String expStatus1="Authorized";
		
		
		
		
		System.out.println("Status :"	+	actStatus	+"Expected"	+	expStatus);
		
		System.out.println("Status1 :"	+	actStatus1	+"Expected1"	+	expStatus1);
		if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1))
		{
			return true;
		}
		else
		{
			return false;
		}
	
	
	
	
	}
	
	
	
	public boolean checkRateLoadingforPendingandAuthorizedPricebookItemsinSalesInvoiceVoucher() throws InterruptedException
	{

		
		checkOpenSalesInvoiceVoucher();
		Thread.sleep(2000);
		
		if(checkAddingFormulaforAccount())
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		}

		
		Thread.sleep(2000);
		
		
		
		
		String actRate1=checkRateValue("Pricebook Item2");
		String expRate1="0.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		String item_Txt1=enter_ItemTxt.getAttribute("value");
		
		
		System.out.println("Item Name	:	"	+item_Txt1);
		System.out.println("Actual Rate	:	"	+actRate1	+"Expected	:"	+expRate1);
		
		

		String actRate2=checkRateValue("RMA STITEM");
		String expRate2="303.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		String item_Txt2=enter_ItemTxt.getAttribute("value");
		
		
		System.out.println("Item Name	:	"	+item_Txt2);
		System.out.println("Actual Rate	:	"	+actRate2	+"Expected	:"	+expRate2);
		
		
		
		
		if(actRate1.equalsIgnoreCase(expRate1) && actRate2.equalsIgnoreCase(expRate2))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	
	}
	
	
	
	public boolean checkAuthorizationforEditedRowsinSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(2000);
				
		getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
		alertsBtn.click();
		Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
		alertExtraMaster.click();
		Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorization));
		SellerPricebookAuthorization.click();
		Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationTxt));
		String actAlertTxt=SellerPricebookAuthorizationTxt.getText();
		String expAlertTxt="";
		
		SellerPricebookAuthorizationTxt.click();
		
		

		Thread.sleep(35000);
		try
		{
		
			getAction().moveToElement(SellerPricebookAuthorizationBtn).build().perform();
			Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationBtn));
		SellerPricebookAuthorizationBtn.click();
		Thread.sleep(2500);
		}
		catch(Exception e)
		{
			
		}
		
		getWebDriverWait().until(ExpectedConditions.visibilityOf(Auth_selectAllBtn));
		Auth_selectAllBtn.click();
		
		getWebDriverWait().until(ExpectedConditions.visibilityOf(authorizeBtn));
		authorizeBtn.click();
		Thread.sleep(2000);
		
		String expMsg="Authorized Successfully";
		String actMsg=checkValidationMessage(expMsg);
		Thread.sleep(6000);
		
		getAction().moveToElement(cancelBtn).build().perform();
		Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(cancelBtn));
		//cancelBtn.click();
		ClickUsingJs(cancelBtn);
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		String itemTxt=itemValue.getAttribute("value");
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
			getAction().moveToElement(StatusFirstCellSelect).build().perform();
			Thread.sleep(2000);
			String actStatus=StatusFirstCellSelect.getText();
			String expStatus="Authorized";
		
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
			itemSecondCellSelect.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
			itemValue.click();
			String itemTxt1=itemValue.getAttribute("value");
			itemValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
			getAction().moveToElement(StatusSecondCellSelect).build().perform();
			Thread.sleep(2000);
			String actStatus1=StatusSecondCellSelect.getText();
			String expStatus1="Authorized";
			
			
			
			
			System.out.println("Alert Text		:	"+	actAlertTxt 	+	"Expected	:"	+	expAlertTxt);
			
			System.out.println("Item 	:	"	+	itemTxt		+"Status :"	+	actStatus	+"Expected"	+	expStatus);
			System.out.println("Item 	:	"	+	itemTxt		+"Status :"	+	actStatus1	+"Expected"	+	expStatus1);
			
		
			
			if(actStatus.equalsIgnoreCase(expStatus) && actMsg.equalsIgnoreCase(expMsg)
					&& actStatus1.equalsIgnoreCase(expStatus1) )
			{
				return true;
			}
			else
			{
				return false;
			}

		
	}
	
	
	
	public boolean checkRateLoadingforPendingandAuthorizedPricebookItemsinSalesInvoiceVoucherafterAuthorization() throws InterruptedException
	{

		
		checkOpenSalesInvoiceVoucher();
		Thread.sleep(2000);
		
		if(checkAddingFormulaforAccount())
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		}

		
		Thread.sleep(2000);
		
		
		
		
		String actRate1=checkRateValue("Pricebook Item2");
		String expRate1="203.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		String item_Txt1=enter_ItemTxt.getAttribute("value");
		
		
		System.out.println("Item Name	:	"	+item_Txt1);
		System.out.println("Actual Rate	:	"	+actRate1	+"Expected	:"	+expRate1);
		
		

		String actRate2=checkRateValue("RMA STITEM");
		String expRate2="303.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		String item_Txt2=enter_ItemTxt.getAttribute("value");
		
		
		System.out.println("Item Name	:	"	+item_Txt2);
		System.out.println("Actual Rate	:	"	+actRate2	+"Expected	:"	+expRate2);
		
		
		
		
		if(actRate1.equalsIgnoreCase(expRate1) && actRate2.equalsIgnoreCase(expRate2))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	
	}
	
	public boolean checkSavingPricebookafterPricebookAuthorizewithCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
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
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with Condition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		
		
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys("FEDPAD SLURRY");
		Thread.sleep(2000);
		

		itemValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("20.00");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("100");
		val1Value.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		
		//Second row
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys("FEDPOUR SLURRY");
		Thread.sleep(2000);
		

		itemValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("25.00");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("22.32");
		val1Value.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		//Third row
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
		itemThirdCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys("BATCH BR ITEM");
		Thread.sleep(2000);
		

		itemValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("28.00");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("88.32");
		val1Value.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		System.out.println("*********************************checkSavingSellerPriceBook********************************");
		
		String expMessage="Price Book created successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		
		
		
		if (actMessage.equalsIgnoreCase(expMessage)  ) 
		{
		
		
		return true;
		} 
		else 
		{
		
		
		return false;
		}
		
			
		
		
	}
	
	
	
	
	
	public boolean checkStatusofPricebookforAuthorizationWithRateCondition() throws InterruptedException
	{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with Condition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
	
		getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Authorized";
		
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
	
		getAction().moveToElement(StatusSecondCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus1=StatusSecondCellSelect.getText();
		String expStatus1="Authorized";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
		itemThirdCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusThirdCellSelect));
	
		ScrollIntoView(StatusThirdCellSelect);
		Thread.sleep(2000);
		String actStatus2=StatusThirdCellSelect.getText();
		String expStatus2="Pending";
		
		
		System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
		
		System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
		
		System.out.println("Status 		:"	+		actStatus2	+	"Expected"	+		expStatus2);
		if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
				&& actStatus2.equalsIgnoreCase(expStatus2))
		{
			return true;
		}
		else
		{
			return false;
		}
	
	
	}
	
	
	public boolean checkStatusofPricebookAfterAuthorizationWithRateCondition() throws InterruptedException
	{
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
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
		homeMasterItem_sellerPriceBookMenu.click();
		
		Thread.sleep(8000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with Condition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(6000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
	
		getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Authorized";
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
	
		getAction().moveToElement(StatusSecondCellSelect).build().perform();;
		Thread.sleep(2000);
		String actStatus1=StatusSecondCellSelect.getText();
		String expStatus1="Authorized";
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
		itemThirdCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusThirdCellSelect));
	
		getAction().moveToElement(StatusThirdCellSelect).build().perform();;
		Thread.sleep(2000);
		String actStatus2=StatusThirdCellSelect.getText();
		String expStatus2="Authorized";
		
		
		System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
		
		System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
		
		System.out.println("Status 		:"	+		actStatus2	+	"Expected"	+		expStatus2);
		if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
				&& actStatus2.equalsIgnoreCase(expStatus2))
		{
			return true;
		}
		else
		{
			return false;
		}
	
	
	}
	

	
	public boolean checkDeletetheConditionPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{

	boolean result=checkDeletingSellerPriceBook("Pricebook With SellerAuthorization with Condition");

	if(result==true )
	{
	return true;
	}
	else
	{
	return false;
	}

	}
	
	
	
	
	public boolean checkEditingPricebookAuthorizationforProductNameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		

		
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1Icon));
		getAction().doubleClick(Level1Icon).build().perform();

		Thread.sleep(2000);
		
		
		if(conditionNotRequiredChkBox.isSelected()==true)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
		conditionNotRequiredChkBox.click();
		}
		Thread.sleep(2000);
		
		
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
		conjunctionDropdown.click();
		Thread.sleep(1500);
		Select s=new Select(conjunctionDropdown);
		s.selectByValue("0");
		//conjunctionDropdown.sendKeys("where");
		//conjunctionDropdown.sendKeys(Keys.TAB);
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
		selectFieldTxt.click();
		Thread.sleep(2000);
		selectFieldTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		selectFieldTxt.sendKeys(Keys.DELETE);
		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
		selectFieldTxt.click();
		Thread.sleep(2000);*/
		
		getAction().moveToElement(selectFieldProductName).build().perform();
		Thread.sleep(3000);
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldProductName));
		selectFieldProductName.click();
		Thread.sleep(2000);*/
		
		ClickUsingJs(selectFieldProductName);
		
		
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldProductName));
		selectFieldProductName.click();
		Thread.sleep(2000);*/
		
//		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
//		selectFieldTxt.click();
//		Thread.sleep(2000);
//		selectFieldProductName.click();
		
		selectFieldTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
		selectOperatorDropdown.click();
		selectOperatorDropdown.sendKeys("Equal to");
		Thread.sleep(2000);
		selectOperatorDropdown.sendKeys(Keys.TAB);
		//Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
		compareWithDropdown.click();
		compareWithDropdown.sendKeys("value");
		Thread.sleep(4000);
		//compareWithDropdown.sendKeys(Keys.TAB);
		//Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_valueTxt));
		enter_valueTxt.click();
		
		enter_ProdvalueTxt.click();
		//enter_valueTxt.clear();
		enter_ProdvalueTxt.sendKeys("RAW MATERIAL");
		Thread.sleep(4000);
		enter_ProdvalueTxt.sendKeys(Keys.TAB);
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
		Select s1=new Select(conjunctionDropdown);
		String actconjunctionDropdown=s1.getFirstSelectedOption().getText();
		String expconjunctionDropdown="Where";
		
		
		String actselectFieldTxt=selectFieldTxt.getAttribute("value");
		String expselectFieldTxt="ProductName";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
		Select s2=new Select(selectOperatorDropdown);
		String actselectOperatorDropdown=s2.getFirstSelectedOption().getText();
		String expselectOperatorDropdown="Equal to";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
		Select s3=new Select(compareWithDropdown);
		String actcompareWithDropdown=s3.getFirstSelectedOption().getText();
		String expcompareWithDropdown="value";
		
		String actvalueTxt=enter_ProdvalueTxt.getAttribute("value");
		String expvalueTxt="RAW MATERIAL";
		
		Thread.sleep(4000);
	
	//	((JavascriptExecutor)getDriver()).executeScript("window.scrollTo(0, 0)","");
		getAction().moveToElement(definitionOkBtn).build().perform();
		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
		definitionOkBtn.click();
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
		Thread.sleep(4000);
		System.out.println("**********************************checkSavingAuthorization*****************************");
		
		String expMessage="Authorization Flow saved successfully.";
		
		String actMessage=checkValidationMessage(expMessage);
		
		
		
		System.out.println("************************************checkConditionCheckbox******************************************************************************");
		
		System.out.println("conjunctionDropdown     :"+actconjunctionDropdown        +" Value Expected  : "+expconjunctionDropdown); 
		System.out.println("selectFieldTxt          :"+actselectFieldTxt             +" Value Expected  : "+expselectFieldTxt);
		System.out.println("selectOperatorDropdown  :"+actselectOperatorDropdown     +"  Value Expected : "+expselectOperatorDropdown); 
		System.out.println("compareWithDropdown     :"+actcompareWithDropdown        +" Value Expected  : "+expcompareWithDropdown);
		System.out.println("valueTxt                :"+actvalueTxt                   +" Value Expected  : "+expvalueTxt);
		System.out.println("Saved                	:"+actMessage                    +" Value Expected  : "+expMessage);
		
		
		
		if(actconjunctionDropdown.equals(expconjunctionDropdown) && actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt) && 
				actselectOperatorDropdown.equals(expselectOperatorDropdown)
				&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueTxt.equalsIgnoreCase(expvalueTxt)
				&& actMessage.equalsIgnoreCase(expMessage))
		{
			System.out.println("Test Pass : Data entered");
			
			return true;
		}
		else
		{
			System.out.println("Test Fail : Data NOT entered");
			
			return false;
		}
		
	
	}
	
	
	
	public boolean checkSavingPricebookforProductnameasCondition() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
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
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with ProductNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		
		
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys("RAW Material");
		Thread.sleep(2000);
		

		itemValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("20.00");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("100");
		val1Value.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		
		//Second row
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys("FEDPOUR SLURRY");
		Thread.sleep(2000);
		

		itemValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("25.00");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("22.32");
		val1Value.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		System.out.println("*********************************checkSavingSellerPriceBook********************************");
		
		String expMessage="Price Book created successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		
		
		
		if (actMessage.equalsIgnoreCase(expMessage)  ) 
		{
		
		
		return true;
		} 
		else 
		{
		
		
		return false;
		}
		

	}
	
	
	public boolean checkStatusofthePricebookforProductnameasCondition() throws InterruptedException
	{

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with ProductNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
	
		getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Pending";
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
	
		getAction().moveToElement(StatusSecondCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus1=StatusSecondCellSelect.getText();
		String expStatus1="Authorized";
		
		
		
		
		
		System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
		
		System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
		
		
		if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
				)
		{
			return true;
		}
		else
		{
			return false;
		}
	
	
		
	}
	
	
	public boolean checkStatusofthePricebookAfterAuthforProductnameasCondition() throws InterruptedException
	{
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
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with ProductNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
	
		getAction().moveToElement(StatusFirstCellSelect).build().perform();;
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Authorized";
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
	
		getAction().moveToElement(StatusSecondCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus1=StatusSecondCellSelect.getText();
		String expStatus1="Authorized";
		
		
		
		
		
		System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
		
		System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
		
		
		if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
				)
		{
			return true;
		}
		else
		{
			return false;
		}
	
	
		
	}
	
	
	public boolean checkDeletetheConditionProductNamePriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{

	boolean result=checkDeletingSellerPriceBook("Pricebook With SellerAuthorization with ProductNameCondition");

	if(result==true )
	{
	return true;
	}
	else
	{
	return false;
	}

	}
	
	
	
	
	
	
	@FindBy(xpath="//select[@id='ddlUndoAuthEditing']")
	private static WebElement select_undoAuthEditing;
	
	
	
	public boolean checkUndoAuthorizationEditinginPricebookAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1Icon));
		getAction().doubleClick(Level1Icon).build().perform();

		Thread.sleep(10000);
		
		
		getAction().moveToElement(select_undoAuthEditing).build().perform();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select_undoAuthEditing));
		
		Select s=new Select(select_undoAuthEditing);
		s.selectByIndex(1);
		Thread.sleep(2000);
		
		
		getAction().moveToElement(definitionOkBtn).build().perform();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
		definitionOkBtn.click();
		Thread.sleep(8000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
		Thread.sleep(4000);
		System.out.println("**********************************checkSavingAuthorization*****************************");
		
		String expMessage="Authorization Flow saved successfully.";
		
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("Saved Message	:"		+actMessage		+"Expected	:"		+expMessage);
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			return true;
		}
		
		else
		{
			return false;
		}
	}
	
	
	
	public boolean checkEditingaPricebookforUndoAuthEditingOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
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
		
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with ProductNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(6000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("156.23");
		rateValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		System.out.println("*********************************checkSavingSellerPriceBook********************************");
		
		String expMessage="Price Book updated successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with ProductNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		Thread.sleep(2500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
	
		getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Pending";
		
		System.out.println("Status	:	"+	actStatus	+	"Expected	:"	+	expStatus);
		
		if (actMessage.equalsIgnoreCase(expMessage)  && actStatus.equalsIgnoreCase(expStatus)) 
		{
		
		
		return true;
		} 
		else 
		{
		
		
		return false;
		}
		
		
	}
	
	
	
	public boolean checkEditingaBuyerPricebookforUndoAuthEditingOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
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
		
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with ProductNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("156.23");
		rateValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		System.out.println("*********************************checkSavingBuyerPriceBook********************************");
		
		String expMessage="Price Book updated successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with ProductNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		Thread.sleep(2500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
	
		getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Pending";
		
		System.out.println("Status	:	"+	actStatus	+	"Expected	:"	+	expStatus);
		
		if (actMessage.equalsIgnoreCase(expMessage)  && actStatus.equalsIgnoreCase(expStatus)) 
		{
		
		
		return true;
		} 
		else 
		{
		
		
		return false;
		}
		
		
	}
	
	
	
	
	public boolean checkEditingPricebookAuthorizationforAccountNameasCondition() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
	
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1Icon));
		getAction().doubleClick(Level1Icon).build().perform();

		Thread.sleep(2000);
		
		
		if(conditionNotRequiredChkBox.isSelected()==true)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
		conditionNotRequiredChkBox.click();
		}
		Thread.sleep(2000);
		
		
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
		conjunctionDropdown.click();
		Thread.sleep(1500);
		Select s=new Select(conjunctionDropdown);
		s.selectByValue("0");
		//conjunctionDropdown.sendKeys("where");
		//conjunctionDropdown.sendKeys(Keys.TAB);
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
		selectFieldTxt.click();
		Thread.sleep(2000);
		selectFieldTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		selectFieldTxt.sendKeys(Keys.DELETE);
		Thread.sleep(2000);
		selectFieldTxt.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldAccountName));
		ClickUsingJs(selectFieldAccountName);
		
		Thread.sleep(2000);
		
		selectFieldTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
		selectOperatorDropdown.click();
		selectOperatorDropdown.sendKeys("Equal to");
		Thread.sleep(2000);
		selectOperatorDropdown.sendKeys(Keys.TAB);
		//Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
		compareWithDropdown.click();
		compareWithDropdown.sendKeys("value");
		Thread.sleep(4000);
	/*	compareWithDropdown.sendKeys(Keys.TAB);
		Thread.sleep(2000);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_valueTxt));
		enter_valueTxt.click();
		
		enter_ProdvalueTxt.click();
		//enter_valueTxt.clear();
		enter_ProdvalueTxt.sendKeys("Customer B");
		Thread.sleep(4000);
		enter_ProdvalueTxt.sendKeys(Keys.TAB);
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
		Select s1=new Select(conjunctionDropdown);
		String actconjunctionDropdown=s1.getFirstSelectedOption().getText();
		String expconjunctionDropdown="Where";
		
		
		String actselectFieldTxt=selectFieldTxt.getAttribute("value");
		String expselectFieldTxt="AccountName";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
		Select s2=new Select(selectOperatorDropdown);
		String actselectOperatorDropdown=s2.getFirstSelectedOption().getText();
		String expselectOperatorDropdown="Equal to";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
		Select s3=new Select(compareWithDropdown);
		String actcompareWithDropdown=s3.getFirstSelectedOption().getText();
		String expcompareWithDropdown="value";
		
		String actvalueTxt=enter_ProdvalueTxt.getAttribute("value");
		String expvalueTxt="Customer B";
		
		Thread.sleep(2000);
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
		definitionOkBtn.click();
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
		Thread.sleep(2000);
		System.out.println("**********************************checkSavingAuthorization*****************************");
		
		String expMessage="Authorization Flow saved successfully.";
		
		String actMessage=checkValidationMessage(expMessage);
		
		
		
		System.out.println("************************************checkConditionCheckbox******************************************************************************");
		
		System.out.println("conjunctionDropdown     :"+actconjunctionDropdown        +" Value Expected  : "+expconjunctionDropdown); 
		System.out.println("selectFieldTxt          :"+actselectFieldTxt             +" Value Expected  : "+expselectFieldTxt);
		System.out.println("selectOperatorDropdown  :"+actselectOperatorDropdown     +"  Value Expected : "+expselectOperatorDropdown); 
		System.out.println("compareWithDropdown     :"+actcompareWithDropdown        +" Value Expected  : "+expcompareWithDropdown);
		System.out.println("valueTxt                :"+actvalueTxt                   +" Value Expected  : "+expvalueTxt);
		System.out.println("Saved                	:"+actMessage                    +" Value Expected  : "+expMessage);
		
		
		
		if(actconjunctionDropdown.equals(expconjunctionDropdown) && actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt) && 
				actselectOperatorDropdown.equals(expselectOperatorDropdown)
				&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueTxt.equalsIgnoreCase(expvalueTxt)
				&& actMessage.equalsIgnoreCase(expMessage))
		{
			System.out.println("Test Pass : Data entered");
			
			return true;
		}
		else
		{
			System.out.println("Test Fail : Data NOT entered");
			
			return false;
		}
		
		
		
	}
	
	
	public boolean checkEditingBuyerPricebookAuthorizationforAccountNameasCondition() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
	
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1Icon));
		getAction().doubleClick(Level1Icon).build().perform();

		Thread.sleep(2000);
		
		
		if(conditionNotRequiredChkBox.isSelected()==true)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
		conditionNotRequiredChkBox.click();
		}
		Thread.sleep(2000);
		
		
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
		conjunctionDropdown.click();
		Thread.sleep(1500);
		Select s=new Select(conjunctionDropdown);
		s.selectByValue("0");
		//conjunctionDropdown.sendKeys("where");
		//conjunctionDropdown.sendKeys(Keys.TAB);
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
		selectFieldTxt.click();
		Thread.sleep(2000);
		selectFieldTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		selectFieldTxt.sendKeys(Keys.DELETE);
		Thread.sleep(2000);
		selectFieldTxt.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldAccountName));
		ClickUsingJs(selectFieldAccountName);
		Thread.sleep(2000);
		
		selectFieldTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
		selectOperatorDropdown.click();
		selectOperatorDropdown.sendKeys("Equal to");
		Thread.sleep(2000);
		selectOperatorDropdown.sendKeys(Keys.TAB);
		//Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
		compareWithDropdown.click();
		compareWithDropdown.sendKeys("value");
		Thread.sleep(2000);
		/*compareWithDropdown.sendKeys(Keys.TAB);
		Thread.sleep(2000);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_valueTxt));
		enter_valueTxt.click();
		
		enter_ProdvalueTxt.click();
		//enter_valueTxt.clear();
		enter_ProdvalueTxt.sendKeys("Vendor B");
		Thread.sleep(1500);
		enter_ProdvalueTxt.sendKeys(Keys.TAB);
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
		Select s1=new Select(conjunctionDropdown);
		String actconjunctionDropdown=s1.getFirstSelectedOption().getText();
		String expconjunctionDropdown="Where";
		
		
		String actselectFieldTxt=selectFieldTxt.getAttribute("value");
		String expselectFieldTxt="AccountName";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
		Select s2=new Select(selectOperatorDropdown);
		String actselectOperatorDropdown=s2.getFirstSelectedOption().getText();
		String expselectOperatorDropdown="Equal to";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
		Select s3=new Select(compareWithDropdown);
		String actcompareWithDropdown=s3.getFirstSelectedOption().getText();
		String expcompareWithDropdown="value";
		
		String actvalueTxt=enter_ProdvalueTxt.getAttribute("value");
		String expvalueTxt="Vendor B";
		
		Thread.sleep(2000);
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
		definitionOkBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
		Thread.sleep(2000);
		System.out.println("**********************************checkSavingAuthorization*****************************");
		
		String expMessage="Authorization Flow saved successfully.";
		
		String actMessage=checkValidationMessage(expMessage);
		
		
		
		System.out.println("************************************checkConditionCheckbox******************************************************************************");
		
		System.out.println("conjunctionDropdown     :"+actconjunctionDropdown        +" Value Expected  : "+expconjunctionDropdown); 
		System.out.println("selectFieldTxt          :"+actselectFieldTxt             +" Value Expected  : "+expselectFieldTxt);
		System.out.println("selectOperatorDropdown  :"+actselectOperatorDropdown     +"  Value Expected : "+expselectOperatorDropdown); 
		System.out.println("compareWithDropdown     :"+actcompareWithDropdown        +" Value Expected  : "+expcompareWithDropdown);
		System.out.println("valueTxt                :"+actvalueTxt                   +" Value Expected  : "+expvalueTxt);
		System.out.println("Saved                	:"+actMessage                    +" Value Expected  : "+expMessage);
		
		
		
		if(actconjunctionDropdown.equals(expconjunctionDropdown) && actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt) && 
				actselectOperatorDropdown.equals(expselectOperatorDropdown)
				&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueTxt.equalsIgnoreCase(expvalueTxt)
				&& actMessage.equalsIgnoreCase(expMessage))
		{
			System.out.println("Test Pass : Data entered");
			
			return true;
		}
		else
		{
			System.out.println("Test Fail : Data NOT entered");
			
			return false;
		}
		
		
		
	}
	
	
	
	public boolean checkSavingPricebookforAccountNameasCondition() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Thread.sleep(4000);
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		homeMasterItemMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
		homeMasterItem_sellerPriceBookMenu.click();
		
		Thread.sleep(6000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with AccountNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
		buyerPriceBook_CustamizeSettingBtn.click();
		Thread.sleep(2000);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustomizeVendorCheckBoxSelected));
		buyerPriceBook_CustomizeVendorCheckBoxSelected.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
		buyerPriceBook_CustamizeSettingBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys("RAW Material");
		Thread.sleep(2000);
		

		itemValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
		vendorValue.click();
		vendorValue.sendKeys("Customer B");
		Thread.sleep(2000);
		vendorValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("20.00");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("100");
		val1Value.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		
		//Second row
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys("FEDPOUR SLURRY");
		Thread.sleep(2000);
		

		itemValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
		vendorValue.click();
		vendorValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		vendorValue.sendKeys("Customer A");
		Thread.sleep(2000);
		vendorValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("25.00");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("22.32");
		val1Value.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		System.out.println("*********************************checkSavingSellerPriceBook********************************");
		
		String expMessage="Price Book created successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		
		
		
		if (actMessage.equalsIgnoreCase(expMessage)  ) 
		{
		
		
		return true;
		} 
		else 
		{
		
		
		return false;
		}
		
	
	}
	
	
	
	
	public boolean checkStatusofthePricebookforAccountnameasCondition() throws InterruptedException
	{

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with AccountNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
	
		getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Pending";
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
	
		getAction().moveToElement(StatusSecondCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus1=StatusSecondCellSelect.getText();
		String expStatus1="Authorized";
		
		
		
		
		
		System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
		
		System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
		
		
		if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
				)
		{
			return true;
		}
		else
		{
			return false;
		}
	
	
		
	}
	
	
	public boolean checkDeletetheConditionAccountNamePriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{

	boolean result=checkDeletingSellerPriceBook("Pricebook With SellerAuthorization with AccountNameCondition");

	if(result==true )
	{
	return true;
	}
	else
	{
	return false;
	}

	}
	
	
	public boolean checkStatusofthePricebookAfterAuthforAccountnameasCondition() throws InterruptedException
	{

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
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with AccountNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
	
		getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Authorized";
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
	
		getAction().moveToElement(StatusSecondCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus1=StatusSecondCellSelect.getText();
		String expStatus1="Authorized";
		
		
		
		
		
		System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
		
		System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
		
		
		if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
				)
		{
			return true;
		}
		else
		{
			return false;
		}
	
	
		
	}
	
	
	public boolean checkSavingBuyerPricebookforAccountNameasCondition() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
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
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with AccountNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
		buyerPriceBook_CustamizeSettingBtn.click();
		Thread.sleep(2000);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustomizeVendorCheckBoxSelected));
		buyerPriceBook_CustomizeVendorCheckBoxSelected.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
		buyerPriceBook_CustamizeSettingBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys("RAW Material");
		Thread.sleep(2000);
		

		itemValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
		vendorValue.click();
		vendorValue.sendKeys("Vendor B");
		Thread.sleep(2000);
		vendorValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("20.00");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("100");
		val1Value.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		
		//Second row
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys("FEDPOUR SLURRY");
		Thread.sleep(2000);
		

		itemValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
		vendorValue.click();
		vendorValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		vendorValue.sendKeys("Vendor A");
		Thread.sleep(2000);
		vendorValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("25.00");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("22.32");
		val1Value.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		System.out.println("*********************************checkSavingBuyerPriceBook********************************");
		
		String expMessage="Price Book created successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		
		
		
		if (actMessage.equalsIgnoreCase(expMessage)  ) 
		{
		
		
		return true;
		} 
		else 
		{
		
		
		return false;
		}
		
	
	}
	
	
	
	
	public boolean checkStatusoftheBuyerPricebookforAccountnameasCondition() throws InterruptedException
	{

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with AccountNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
	
		getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Pending";
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
	
		getAction().moveToElement(StatusSecondCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus1=StatusSecondCellSelect.getText();
		String expStatus1="Authorized";
		
		
		
		
		
		System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
		
		System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
		
		
		if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
				)
		{
			return true;
		}
		else
		{
			return false;
		}
	
	
		
	}
	
	
	public boolean checkDeletetheConditionAccountNameBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{

	boolean result=checkDeletingBuyerPriceBook("Pricebook With Buyer Authorization with AccountNameCondition");

	if(result==true )
	{
	return true;
	}
	else
	{
	return false;
	}

	}
	
	
	public boolean checkStatusoftheBuyerPricebookAfterAuthforAccountnameasCondition() throws InterruptedException
	{

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
		
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with AccountNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
	
		getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Authorized";
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
	
		getAction().moveToElement(StatusSecondCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus1=StatusSecondCellSelect.getText();
		String expStatus1="Authorized";
		
		
		
		
		
		System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
		
		System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
		
		
		if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
				)
		{
			return true;
		}
		else
		{
			return false;
		}
	
	
		
	}
	
	
	
	
	
	
	
	
public boolean checkRateValueLoadingBeforeAuthofPricebookforAccountGroup() throws InterruptedException
{
	checkOpenSalesInvoiceVoucher();
	Thread.sleep(3000);
	
	if(checkAddingFormulaforGroupAccount())
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	}

	
	Thread.sleep(2000);
	
	
	
	
	String actRate1=checkRateValue("Group One Account One","PAPER");
	String expRate1="0.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt1=enter_ItemTxt.getAttribute("value");
	
	
	System.out.println("Item Name	:	"	+item_Txt1);
	System.out.println("Actual Rate	:	"	+actRate1	+"Expected	:"	+expRate1);
	
	Thread.sleep(4000);

	String actRate2=checkRateValue("Group One Account Two","PAPER");;
	String expRate2="0.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt2=enter_ItemTxt.getAttribute("value");
	
	
	System.out.println("Item Name	:	"	+item_Txt2);
	System.out.println("Actual Rate	:	"	+actRate2	+"Expected	:"	+expRate2);
	
	
	
	
	if(actRate1.equalsIgnoreCase(expRate1) && actRate2.equalsIgnoreCase(expRate2))
	{
		return true;
	}
	else
	{
		return false;
	}
}
	
	

public boolean checkRateValueLoadingBeforeAuthofBuyerPricebookforAccountGroup() throws InterruptedException
{
	checkOpenPurchaseVoucher();
	Thread.sleep(3000);
	
	if(checkAddingFormulaforGroupAccountofPurchase())
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	}

	
	Thread.sleep(2000);
	
	
	
	
	String actRate1=checkRateValueforPurchase("Group One Account One","PAPER");
	String expRate1="0.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt1=enter_ItemTxt.getAttribute("value");
	
	
	System.out.println("Item Name	:	"	+item_Txt1);
	System.out.println("Actual Rate	:	"	+actRate1	+"Expected	:"	+expRate1);
	
	

	String actRate2=checkRateValue("Group One Account Two","PAPER");;
	String expRate2="0.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt2=enter_ItemTxt.getAttribute("value");
	
	
	System.out.println("Item Name	:	"	+item_Txt2);
	System.out.println("Actual Rate	:	"	+actRate2	+"Expected	:"	+expRate2);
	
	
	
	
	if(actRate1.equalsIgnoreCase(expRate1) && actRate2.equalsIgnoreCase(expRate2))
	{
		return true;
	}
	else
	{
		return false;
	}
}
	
	



public boolean checkEditingPricebookAuthorizationforAccountNameasGroupAccCondition() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

	Thread.sleep(3000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1Icon));
	getAction().doubleClick(Level1Icon).build().perform();

	Thread.sleep(2000);
	
	
	if(conditionNotRequiredChkBox.isSelected()==true)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
	conditionNotRequiredChkBox.click();
	}
	Thread.sleep(2000);
	
	
	((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
	conjunctionDropdown.click();
	Thread.sleep(1500);
	Select s=new Select(conjunctionDropdown);
	s.selectByValue("0");
	//conjunctionDropdown.sendKeys("where");
	//conjunctionDropdown.sendKeys(Keys.TAB);
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
	selectFieldTxt.click();
	Thread.sleep(2000);
	selectFieldTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	selectFieldTxt.sendKeys(Keys.DELETE);
	Thread.sleep(2000);
	selectFieldTxt.click();
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldAccountName));
	ClickUsingJs(selectFieldAccountName);
	Thread.sleep(2000);
	
	selectFieldTxt.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
	selectOperatorDropdown.click();
	selectOperatorDropdown.sendKeys("Equal to");
	Thread.sleep(2000);
	selectOperatorDropdown.sendKeys(Keys.TAB);
	//Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
	compareWithDropdown.click();
	compareWithDropdown.sendKeys("value");
	Thread.sleep(4000);
	/*compareWithDropdown.sendKeys(Keys.TAB);
	Thread.sleep(2000);*/
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_valueTxt));
	enter_valueTxt.click();
	
	enter_ProdvalueTxt.click();
	
	enter_ProdvalueTxt.sendKeys("Pricebook Group One");
	Thread.sleep(4000);
	enter_ProdvalueTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
	Select s1=new Select(conjunctionDropdown);
	String actconjunctionDropdown=s1.getFirstSelectedOption().getText();
	String expconjunctionDropdown="Where";
	
	
	String actselectFieldTxt=selectFieldTxt.getAttribute("value");
	String expselectFieldTxt="AccountName";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
	Select s2=new Select(selectOperatorDropdown);
	String actselectOperatorDropdown=s2.getFirstSelectedOption().getText();
	String expselectOperatorDropdown="Equal to";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
	Select s3=new Select(compareWithDropdown);
	String actcompareWithDropdown=s3.getFirstSelectedOption().getText();
	String expcompareWithDropdown="value";
	
	String actvalueTxt=enter_ProdvalueTxt.getAttribute("value");
	String expvalueTxt="Pricebook Group One";
	
	Thread.sleep(2000);

	
	getAction().moveToElement(definitionOkBtn).build().perform();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
	definitionOkBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
	saveBtn.click();
	
	Thread.sleep(2000);
	System.out.println("**********************************checkSavingAuthorization*****************************");
	
	String expMessage="Authorization Flow saved successfully.";
	
	String actMessage=checkValidationMessage(expMessage);
	
	
	
	System.out.println("************************************checkConditionCheckbox******************************************************************************");
	
	System.out.println("conjunctionDropdown     :"+actconjunctionDropdown        +" Value Expected  : "+expconjunctionDropdown); 
	System.out.println("selectFieldTxt          :"+actselectFieldTxt             +" Value Expected  : "+expselectFieldTxt);
	System.out.println("selectOperatorDropdown  :"+actselectOperatorDropdown     +"  Value Expected : "+expselectOperatorDropdown); 
	System.out.println("compareWithDropdown     :"+actcompareWithDropdown        +" Value Expected  : "+expcompareWithDropdown);
	System.out.println("valueTxt                :"+actvalueTxt                   +" Value Expected  : "+expvalueTxt);
	System.out.println("Saved                	:"+actMessage                    +" Value Expected  : "+expMessage);
	
	
	
	if(actconjunctionDropdown.equals(expconjunctionDropdown) && actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt) && 
			actselectOperatorDropdown.equals(expselectOperatorDropdown)
			&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueTxt.equalsIgnoreCase(expvalueTxt)
			&& actMessage.equalsIgnoreCase(expMessage))
	{
		System.out.println("Test Pass : Data entered");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Data NOT entered");
		
		return false;
	}
	
		
}
	


public boolean checkSavingPricebookforGroupAccNameasCondition() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	

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
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with AccountGroupNameCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
	buyerPriceBook_CustamizeSettingBtn.click();
	Thread.sleep(2000);

	
	if(buyerPriceBook_CustomizeVendorCheckBox.isSelected()==false)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustomizeVendorCheckBoxSelected));
	buyerPriceBook_CustomizeVendorCheckBoxSelected.click();
	}
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
	buyerPriceBook_CustamizeSettingBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("PAPER");
	Thread.sleep(2000);
	

	itemValue.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
	vendorValue.click();
	vendorValue.sendKeys("Pricebook Group One");
	Thread.sleep(2000);
	vendorValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("120.00");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("800");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	//Second row
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("PREMIX A");
	Thread.sleep(2000);
	

	itemValue.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
	vendorValue.click();
	vendorValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	vendorValue.sendKeys("Pricebook Group Two");
	Thread.sleep(2000);
	vendorValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("205.00");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("212.32");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	
	System.out.println("*********************************checkSavingSellerPriceBook********************************");
	
	String expMessage="Price Book created successfully";
	
	String actMessage=checkValidationMessage(expMessage);
	
	
	
	
	if (actMessage.equalsIgnoreCase(expMessage)  ) 
	{
	
	
	return true;
	} 
	else 
	{
	
	
	return false;
	}
	

	
}


public boolean checkStatusofthePricebookforGroupAccNameasCondition() throws InterruptedException
{

	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with AccountGroupNameCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Pending";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}

}
	

public boolean checkDeletetheConditionGroupAccountNamePriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

boolean result=checkDeletingSellerPriceBook("Pricebook With SellerAuthorization with AccountGroupNameCondition");

if(result==true )
{
return true;
}
else
{
return false;
}

}



public boolean checkStatusofthePricebookAfterAuthforGroupAccNameasCondition() throws InterruptedException
{

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
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with AccountGroupNameCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Authorized";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}



}
	


public boolean checkSavingBuyerPricebookforGroupAccNameasCondition() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	

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
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with AccountGroupNameCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
	buyerPriceBook_CustamizeSettingBtn.click();
	Thread.sleep(2000);

	
	if(buyerPriceBook_CustomizeVendorCheckBox.isSelected()==false)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustomizeVendorCheckBoxSelected));
	buyerPriceBook_CustomizeVendorCheckBoxSelected.click();
	}
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
	buyerPriceBook_CustamizeSettingBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("PAPER");
	Thread.sleep(2000);
	

	itemValue.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
	vendorValue.click();
	vendorValue.sendKeys("Pricebook Group One");
	Thread.sleep(2000);
	vendorValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("120.00");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("800");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	//Second row
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("PREMIX A");
	Thread.sleep(2000);
	

	itemValue.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorValue));
	vendorValue.click();
	vendorValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	vendorValue.sendKeys("Pricebook Group Two");
	Thread.sleep(2000);
	vendorValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("205.00");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("212.32");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	
	System.out.println("*********************************checkSavingBuyerPriceBook********************************");
	
	String expMessage="Price Book created successfully";
	
	String actMessage=checkValidationMessage(expMessage);
	
	
	
	
	if (actMessage.equalsIgnoreCase(expMessage)  ) 
	{
	
	
	return true;
	} 
	else 
	{
	
	
	return false;
	}
	

	
}


public boolean checkStatusoftheBuyerPricebookforGroupAccNameasCondition() throws InterruptedException
{

	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with AccountGroupNameCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Pending";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}

}
	

public boolean checkDeletetheConditionGroupAccountNameBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

boolean result=checkDeletingBuyerPriceBook("Pricebook With Buyer Authorization with AccountGroupNameCondition");

if(result==true )
{
return true;
}
else
{
return false;
}

}



public boolean checkStatusoftheBuyerPricebookAfterAuthforGroupAccNameasCondition() throws InterruptedException
{

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
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with AccountGroupNameCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Authorized";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}



}
	




public static String checkRateValue(String Customer,String item) throws InterruptedException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	customerAccountTxt.click();

	customerAccountTxt.sendKeys(Keys.END);
	customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	customerAccountTxt.sendKeys(Customer);
	Thread.sleep(2000);
	customerAccountTxt.sendKeys(Keys.TAB);
	Thread.sleep(6000);
	
	
	if(getIsAlertPresent())
	{
		Thread.sleep(2000);
		getWaitForAlert();
		Thread.sleep(2000);
		getAlert().accept();
		Thread.sleep(4000);
	}
	
	
	Thread.sleep(10000);
	
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


	enter_ItemTxt.sendKeys(Keys.TAB);

	Thread.sleep(3000);
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
	enter_UnitTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	enter_UnitTxt.sendKeys(Keys.DELETE);
	enter_UnitTxt.sendKeys(Keys.ESCAPE);
	Thread.sleep(2000);
	enter_UnitTxt.sendKeys(Keys.TAB);*/

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	select1stRow_11thColumn.click();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	enter_Rate.click();
	Thread.sleep(1000);

	String actRate=enter_Rate.getAttribute("value");


	return actRate;
	
	
}


public static String checkRateValueforPurchase(String Customer,String item) throws InterruptedException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	customerAccountTxt.click();

	customerAccountTxt.sendKeys(Keys.END);
	customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	customerAccountTxt.sendKeys(Customer);
	
	customerAccountTxt.sendKeys(Keys.TAB);
	Thread.sleep(3000);
	
	
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
	
	enter_ItemTxt.sendKeys(item);
	Thread.sleep(2000);


	enter_ItemTxt.sendKeys(Keys.TAB);

	Thread.sleep(3000);
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
	enter_UnitTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	enter_UnitTxt.sendKeys(Keys.DELETE);
	enter_UnitTxt.sendKeys(Keys.ESCAPE);
	Thread.sleep(2000);
	enter_UnitTxt.sendKeys(Keys.TAB);*/

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	select1stRow_11thColumn.click();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	enter_Rate.click();
	Thread.sleep(1000);

	String actRate=enter_Rate.getAttribute("value");


	return actRate;
	
	
}



public boolean checkEditPricebookAuthorizationforMinimumQuantityasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	


	Thread.sleep(3000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1Icon));
	getAction().doubleClick(Level1Icon).build().perform();

	Thread.sleep(2000);
	
	
	if(conditionNotRequiredChkBox.isSelected()==true)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
	conditionNotRequiredChkBox.click();
	}
	Thread.sleep(2000);
	
	
	((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
	conjunctionDropdown.click();
	Thread.sleep(1500);
	Select s=new Select(conjunctionDropdown);
	s.selectByValue("0");
	//conjunctionDropdown.sendKeys("where");
	//conjunctionDropdown.sendKeys(Keys.TAB);
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
	selectFieldTxt.click();
	Thread.sleep(2000);
	selectFieldTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	selectFieldTxt.sendKeys(Keys.DELETE);
	Thread.sleep(2000);
	selectFieldTxt.click();
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldMinimumQty));
	ClickUsingJs(selectFieldMinimumQty);
	//selectFieldMinimumQty.click();
	Thread.sleep(2000);
	
	selectFieldTxt.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
	selectOperatorDropdown.click();
	selectOperatorDropdown.sendKeys("Less than or equal to");
	Thread.sleep(2000);
	selectOperatorDropdown.sendKeys(Keys.TAB);
	//Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
	compareWithDropdown.click();
	compareWithDropdown.sendKeys("value");
	Thread.sleep(4000);
	/*compareWithDropdown.sendKeys(Keys.TAB);
	Thread.sleep(2000);*/
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_valueTxt));
	enter_valueTxt.click();
	
//	enter_ProdvalueTxt.click();
	//enter_valueTxt.clear();
	enter_valueTxt.sendKeys("10");
	Thread.sleep(2000);
	enter_valueTxt.sendKeys(Keys.TAB);
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
	Select s1=new Select(conjunctionDropdown);
	String actconjunctionDropdown=s1.getFirstSelectedOption().getText();
	String expconjunctionDropdown="Where";
	
	
	String actselectFieldTxt=selectFieldTxt.getAttribute("value");
	String expselectFieldTxt="MinQty";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
	Select s2=new Select(selectOperatorDropdown);
	String actselectOperatorDropdown=s2.getFirstSelectedOption().getText();
	String expselectOperatorDropdown="Less than or equal to";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
	Select s3=new Select(compareWithDropdown);
	String actcompareWithDropdown=s3.getFirstSelectedOption().getText();
	String expcompareWithDropdown="value";
	
	String actvalueTxt=enter_valueTxt.getAttribute("value");
	String expvalueTxt="10";
	
	Thread.sleep(2000);

	getAction().moveToElement(definitionOkBtn).build().perform();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
	definitionOkBtn.click();
	Thread.sleep(6000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
	saveBtn.click();
	
	Thread.sleep(6000);
	System.out.println("**********************************checkSavingAuthorization*****************************");
	
	String expMessage="Authorization Flow saved successfully.";
	
	String actMessage=checkValidationMessage(expMessage);
	
	
	
	System.out.println("************************************checkConditionCheckbox******************************************************************************");
	
	System.out.println("conjunctionDropdown     :"+actconjunctionDropdown        +" Value Expected  : "+expconjunctionDropdown); 
	System.out.println("selectFieldTxt          :"+actselectFieldTxt             +" Value Expected  : "+expselectFieldTxt);
	System.out.println("selectOperatorDropdown  :"+actselectOperatorDropdown     +"  Value Expected : "+expselectOperatorDropdown); 
	System.out.println("compareWithDropdown     :"+actcompareWithDropdown        +" Value Expected  : "+expcompareWithDropdown);
	System.out.println("valueTxt                :"+actvalueTxt                   +" Value Expected  : "+expvalueTxt);
	System.out.println("Saved                	:"+actMessage                    +" Value Expected  : "+expMessage);
	
	
	
	if(actconjunctionDropdown.equals(expconjunctionDropdown) && actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt) && 
			actselectOperatorDropdown.equals(expselectOperatorDropdown)
			&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueTxt.equalsIgnoreCase(expvalueTxt)
			&& actMessage.equalsIgnoreCase(expMessage))
	{
		System.out.println("Test Pass : Data entered");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Data NOT entered");
		
		return false;
	}
	
		
}

public boolean checkSavingPricebookforMinimumQuantityasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	

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
			buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
			buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with MinQtyCondition");
			Thread.sleep(1500);
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
			buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
			buyerPriceBook_CustamizeSettingBtn.click();
			Thread.sleep(2000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_QualityRangeCheckBoxSelected));
			buyerPriceBook_QualityRangeCheckBoxSelected.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
			buyerPriceBook_CustamizeSettingBtn.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
			itemFirstCellSelect.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
			itemValue.click();
			itemValue.sendKeys("PAPER");
			Thread.sleep(2000);
			
			
			itemValue.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
			minQtyValue.click();
			minQtyValue.sendKeys("6");
			Thread.sleep(2000);
			minQtyValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
			maxQtyValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
			rateValue.sendKeys("120.00");
			rateValue.sendKeys(Keys.TAB);
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
			//val1Value.click();
			val1Value.sendKeys("800");
			val1Value.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			
			//Second row
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
			itemSecondCellSelect.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
			itemValue.click();
			itemValue.sendKeys("PREMIX A");
			Thread.sleep(2000);
			
			
			itemValue.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
			minQtyValue.click();
			minQtyValue.sendKeys("10");
			Thread.sleep(2000);
			minQtyValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
			maxQtyValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
			rateValue.sendKeys("205.00");
			rateValue.sendKeys(Keys.TAB);
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
			//val1Value.click();
			val1Value.sendKeys("212.32");
			val1Value.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
			itemThirdCellSelect.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
			itemValue.click();
			itemValue.sendKeys("RESIN");
			Thread.sleep(2000);
			
			
			itemValue.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
			minQtyValue.click();
			minQtyValue.sendKeys("11");
			Thread.sleep(2000);
			minQtyValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
			maxQtyValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
			rateValue.sendKeys("185.00");
			rateValue.sendKeys(Keys.TAB);
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
			//val1Value.click();
			val1Value.sendKeys("192.36");
			val1Value.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			
			
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
			buyerPriceBookSaveBtn.click();
			
			System.out.println("*********************************checkSavingSellerPriceBook********************************");
			
			String expMessage="Price Book created successfully";
			
			String actMessage=checkValidationMessage(expMessage);
			
			
			
			
			if (actMessage.equalsIgnoreCase(expMessage)  ) 
			{
			
			
			return true;
			} 
			else 
			{
			
			
			return false;
			}



}


public boolean checkStatusofthePricebookforMinimumQuantityasCondition() throws InterruptedException
{
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with MinQtyCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Pending";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Pending";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
	itemThirdCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusThirdCellSelect));

	getAction().moveToElement(StatusThirdCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus2=StatusThirdCellSelect.getText();
	String expStatus2="Authorized";
	
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	System.out.println("Status 		:"	+		actStatus2	+	"Expected"	+		expStatus2);
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
		&& actStatus2.equalsIgnoreCase(expStatus2)	)
	{
		return true;
	}
	else
	{
		return false;
	}



}
	

public boolean checkDeletetheConditionMinQtyPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

boolean result=checkDeletingSellerPriceBook("Pricebook With SellerAuthorization with MinQtyCondition");

if(result==true )
{
return true;
}
else
{
return false;
}

}




public boolean checkStatusofthePricebookAfterAuthforMinimumQuantityasCondition() throws InterruptedException
{
	
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
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with MinQtyCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Authorized";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
	itemThirdCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusThirdCellSelect));

	getAction().moveToElement(StatusThirdCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus2=StatusThirdCellSelect.getText();
	String expStatus2="Authorized";
	
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	System.out.println("Status 		:"	+		actStatus2	+	"Expected"	+		expStatus2);
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
		&& actStatus2.equalsIgnoreCase(expStatus2)	)
	{
		return true;
	}
	else
	{
		return false;
	}



}




public boolean checkSavingBuyerPricebookforMinimumQuantityasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	

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
			buyerPriceBook_PriceBook.click();
			buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
			buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with MinQtyCondition");
			Thread.sleep(1500);
			buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
			buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
			buyerPriceBook_CustamizeSettingBtn.click();
			Thread.sleep(2000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_QualityRangeCheckBoxSelected));
			buyerPriceBook_QualityRangeCheckBoxSelected.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
			buyerPriceBook_CustamizeSettingBtn.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
			itemFirstCellSelect.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
			itemValue.click();
			itemValue.sendKeys("PAPER");
			Thread.sleep(2000);
			
			
			itemValue.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
			minQtyValue.click();
			minQtyValue.sendKeys("6");
			Thread.sleep(2000);
			minQtyValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
			maxQtyValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
			rateValue.sendKeys("120.00");
			rateValue.sendKeys(Keys.TAB);
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
			//val1Value.click();
			val1Value.sendKeys("800");
			val1Value.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			
			//Second row
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
			itemSecondCellSelect.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
			itemValue.click();
			itemValue.sendKeys("PREMIX A");
			Thread.sleep(2000);
			
			
			itemValue.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
			minQtyValue.click();
			minQtyValue.sendKeys("10");
			Thread.sleep(2000);
			minQtyValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
			maxQtyValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
			rateValue.sendKeys("205.00");
			rateValue.sendKeys(Keys.TAB);
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
			//val1Value.click();
			val1Value.sendKeys("212.32");
			val1Value.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
			itemThirdCellSelect.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
			itemValue.click();
			itemValue.sendKeys("RESIN");
			Thread.sleep(2000);
			
			
			itemValue.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(minQtyValue));
			minQtyValue.click();
			minQtyValue.sendKeys("11");
			Thread.sleep(2000);
			minQtyValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxQtyValue));
			maxQtyValue.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
			rateValue.sendKeys("185.00");
			rateValue.sendKeys(Keys.TAB);
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
			//val1Value.click();
			val1Value.sendKeys("192.36");
			val1Value.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			
			
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
			buyerPriceBookSaveBtn.click();
			
			System.out.println("*********************************checkSavingBuyerPriceBook********************************");
			
			String expMessage="Price Book created successfully";
			
			String actMessage=checkValidationMessage(expMessage);
			
			
			
			
			if (actMessage.equalsIgnoreCase(expMessage)  ) 
			{
			
			
			return true;
			} 
			else 
			{
			
			
			return false;
			}



}


public boolean checkStatusoftheBuyerPricebookforMinimumQuantityasCondition() throws InterruptedException
{
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with MinQtyCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Pending";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Pending";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
	itemThirdCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusThirdCellSelect));

	getAction().moveToElement(StatusThirdCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus2=StatusThirdCellSelect.getText();
	String expStatus2="Authorized";
	
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	System.out.println("Status 		:"	+		actStatus2	+	"Expected"	+		expStatus2);
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
		&& actStatus2.equalsIgnoreCase(expStatus2)	)
	{
		return true;
	}
	else
	{
		return false;
	}



}
	

public boolean checkDeletetheConditionMinQtyBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

boolean result=checkDeletingBuyerPriceBook("Pricebook With Buyer Authorization with MinQtyCondition");

if(result==true )
{
return true;
}
else
{
return false;
}

}




public boolean checkStatusoftheBuyerPricebookAfterAuthforMinimumQuantityasCondition() throws InterruptedException
{
	
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
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with MinQtyCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Authorized";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
	itemThirdCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusThirdCellSelect));

	getAction().moveToElement(StatusThirdCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus2=StatusThirdCellSelect.getText();
	String expStatus2="Authorized";
	
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	System.out.println("Status 		:"	+		actStatus2	+	"Expected"	+		expStatus2);
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
		&& actStatus2.equalsIgnoreCase(expStatus2)	)
	{
		return true;
	}
	else
	{
		return false;
	}



}



public boolean  checkSettingsofSellerPriceBookforTag() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
		Thread.sleep(2000);
		getDriver().navigate().refresh();
		Thread.sleep(4000);
		
		getAction().moveToElement(settingsMenu).build().perform();
		Thread.sleep(2000);
		
		ClickUsingJs(settingsMenu);
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsMenu));
		settingsMenu.click();*/
		
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(preferencesBtn));
		preferencesBtn.click();
		Thread.sleep(2000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sellingRatesSelect));
		sellingRatesSelect.click();
		Thread.sleep(2000);
		
		Select s = new Select(sellingRatesSelect);
		s.selectByVisibleText("Department");
		Thread.sleep(2000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		
		Thread.sleep(1000);
		
		getWaitForAlert();
		Thread.sleep(2000);
		getAlert().accept();
		Thread.sleep(2000);
		
		String expMessage="Data saved Successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			return true;
		}
		
		else
		{
			return false;
		}
		
	
		
}




public boolean  checkSettingsofBuyerPriceBookforTag() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsMenu));
		settingsMenu.click();
		
		Thread.sleep(1500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(preferencesBtn));
		preferencesBtn.click();
		Thread.sleep(2000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyingRatesSelect));
		buyingRatesSelect.click();
		Thread.sleep(2000);
		
		Select s = new Select(buyingRatesSelect);
		s.selectByVisibleText("Warehouse");
		Thread.sleep(2000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		
		Thread.sleep(1000);
		
		getWaitForAlert();
		
		getAlert().accept();
		Thread.sleep(2000);
		
		String expMessage="Data saved Successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			return true;
		}
		
		else
		{
			return false;
		}
		
	
		
}




@FindBy(xpath="//a[@id='0']")
private static WebElement  TagsBtn;
@FindBy(xpath = "//*[@id='16']")
public static WebElement settingsMenu;

@FindBy(xpath="//select[@id='BuyingRates']")
public static WebElement buyingRatesSelect;

@FindBy(xpath="//select[@id='SellingRate']")
public static WebElement sellingRatesSelect;



public boolean checkEditingPricebookAuthorizationforTagasCondition() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException

{

	Thread.sleep(3000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1Icon));
	getAction().doubleClick(Level1Icon).build().perform();

	Thread.sleep(2000);
	
	
	if(conditionNotRequiredChkBox.isSelected()==true)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
	conditionNotRequiredChkBox.click();
	}
	Thread.sleep(2000);
	
	
	((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
	conjunctionDropdown.click();
	Thread.sleep(1500);
	Select s=new Select(conjunctionDropdown);
	s.selectByValue("0");
	//conjunctionDropdown.sendKeys("where");
	//conjunctionDropdown.sendKeys(Keys.TAB);
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
	selectFieldTxt.click();
	Thread.sleep(2000);
	selectFieldTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	selectFieldTxt.sendKeys(Keys.DELETE);
	Thread.sleep(2000);
	selectFieldTxt.click();
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTag));
	//selectFieldTag.click();
	ClickUsingJs(selectFieldTag);
	Thread.sleep(2000);
	
	selectFieldTxt.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
	selectOperatorDropdown.click();
	selectOperatorDropdown.sendKeys("Equal to");
	Thread.sleep(2000);
	selectOperatorDropdown.sendKeys(Keys.TAB);
	//Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
	compareWithDropdown.click();
	compareWithDropdown.sendKeys("value");
	Thread.sleep(6000);
	/*compareWithDropdown1.sendKeys(Keys.TAB);
	Thread.sleep(6000);
	*/
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_valueTxt));
	enter_valueTxt.click();
	
	enter_ProdvalueTxt.click();
	//enter_valueTxt.clear();
	enter_ProdvalueTxt.sendKeys("AMERICA");
	Thread.sleep(4000);
	enter_ProdvalueTxt.sendKeys(Keys.TAB);
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
	Select s1=new Select(conjunctionDropdown);
	String actconjunctionDropdown=s1.getFirstSelectedOption().getText();
	String expconjunctionDropdown="Where";
	
	
	String actselectFieldTxt=selectFieldTxt.getAttribute("value");
	String expselectFieldTxt="Department";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
	Select s2=new Select(selectOperatorDropdown);
	String actselectOperatorDropdown=s2.getFirstSelectedOption().getText();
	String expselectOperatorDropdown="Equal to";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
	Select s3=new Select(compareWithDropdown);
	String actcompareWithDropdown=s3.getFirstSelectedOption().getText();
	String expcompareWithDropdown="value";
	
	String actvalueTxt=enter_ProdvalueTxt.getAttribute("value");
	String expvalueTxt="AMERICA";
	
	Thread.sleep(2000);

	getAction().moveToElement(definitionOkBtn).build().perform();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
	definitionOkBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
	saveBtn.click();
	
	Thread.sleep(2000);
	System.out.println("**********************************checkSavingAuthorization*****************************");
	
	String expMessage="Authorization Flow saved successfully.";
	
	String actMessage=checkValidationMessage(expMessage);
	
	
	
	System.out.println("************************************checkConditionCheckbox******************************************************************************");
	
	System.out.println("conjunctionDropdown     :"+actconjunctionDropdown        +" Value Expected  : "+expconjunctionDropdown); 
	System.out.println("selectFieldTxt          :"+actselectFieldTxt             +" Value Expected  : "+expselectFieldTxt);
	System.out.println("selectOperatorDropdown  :"+actselectOperatorDropdown     +"  Value Expected : "+expselectOperatorDropdown); 
	System.out.println("compareWithDropdown     :"+actcompareWithDropdown        +" Value Expected  : "+expcompareWithDropdown);
	System.out.println("valueTxt                :"+actvalueTxt                   +" Value Expected  : "+expvalueTxt);
	System.out.println("Saved                	:"+actMessage                    +" Value Expected  : "+expMessage);
	
	
	
	if(actconjunctionDropdown.equals(expconjunctionDropdown) && actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt) && 
			actselectOperatorDropdown.equals(expselectOperatorDropdown)
			&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueTxt.equalsIgnoreCase(expvalueTxt)
			&& actMessage.equalsIgnoreCase(expMessage))
	{
		System.out.println("Test Pass : Data entered");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Data NOT entered");
		
		return false;
	}
	
		

}


public boolean checkEditingBuyerPricebookAuthorizationforTagasCondition() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException

{

	Thread.sleep(3000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1Icon));
	getAction().doubleClick(Level1Icon).build().perform();

	Thread.sleep(2000);
	
	
	if(conditionNotRequiredChkBox.isSelected()==true)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
	conditionNotRequiredChkBox.click();
	}
	Thread.sleep(2000);
	
	
	((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
	conjunctionDropdown.click();
	Thread.sleep(1500);
	Select s=new Select(conjunctionDropdown);
	s.selectByValue("0");
	//conjunctionDropdown.sendKeys("where");
	//conjunctionDropdown.sendKeys(Keys.TAB);
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
	selectFieldTxt.click();
	Thread.sleep(2000);
	selectFieldTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	selectFieldTxt.sendKeys(Keys.DELETE);
	Thread.sleep(2000);
	selectFieldTxt.click();
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTag));
	ClickUsingJs(selectFieldTag);
	Thread.sleep(2000);
	
	selectFieldTxt.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
	selectOperatorDropdown.click();
	selectOperatorDropdown.sendKeys("Equal to");
	Thread.sleep(2000);
	selectOperatorDropdown.sendKeys(Keys.TAB);
	//Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
	compareWithDropdown.click();
	compareWithDropdown.sendKeys("value");
	Thread.sleep(2000);
	/*compareWithDropdown.sendKeys(Keys.TAB);
	Thread.sleep(2000);*/
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_valueTxt));
	enter_valueTxt.click();
	
	enter_ProdvalueTxt.click();
	//enter_valueTxt.clear();
	enter_ProdvalueTxt.sendKeys("HYDERABAD");
	Thread.sleep(3000);
	enter_ProdvalueTxt.sendKeys(Keys.TAB);
	/*enter_valueTxt.sendKeys(Keys.TAB);
	Thread.sleep(3000);*/
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
	Select s1=new Select(conjunctionDropdown);
	String actconjunctionDropdown=s1.getFirstSelectedOption().getText();
	String expconjunctionDropdown="Where";
	
	
	String actselectFieldTxt=selectFieldTxt.getAttribute("value");
	String expselectFieldTxt="Tag";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
	Select s2=new Select(selectOperatorDropdown);
	String actselectOperatorDropdown=s2.getFirstSelectedOption().getText();
	String expselectOperatorDropdown="Equal to";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
	Select s3=new Select(compareWithDropdown);
	String actcompareWithDropdown=s3.getFirstSelectedOption().getText();
	String expcompareWithDropdown="value";
	
	String actvalueTxt=enter_ProdvalueTxt.getAttribute("value");
	String expvalueTxt="HYDERABAD";
	
	Thread.sleep(2000);

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
	definitionOkBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
	saveBtn.click();
	
	Thread.sleep(2000);
	System.out.println("**********************************checkSavingAuthorization*****************************");
	
	String expMessage="Authorization Flow saved successfully.";
	
	String actMessage=checkValidationMessage(expMessage);
	
	
	
	System.out.println("************************************checkConditionCheckbox******************************************************************************");
	
	System.out.println("conjunctionDropdown     :"+actconjunctionDropdown        +" Value Expected  : "+expconjunctionDropdown); 
	System.out.println("selectFieldTxt          :"+actselectFieldTxt             +" Value Expected  : "+expselectFieldTxt);
	System.out.println("selectOperatorDropdown  :"+actselectOperatorDropdown     +"  Value Expected : "+expselectOperatorDropdown); 
	System.out.println("compareWithDropdown     :"+actcompareWithDropdown        +" Value Expected  : "+expcompareWithDropdown);
	System.out.println("valueTxt                :"+actvalueTxt                   +" Value Expected  : "+expvalueTxt);
	System.out.println("Saved                	:"+actMessage                    +" Value Expected  : "+expMessage);
	
	
	
	if(actconjunctionDropdown.equals(expconjunctionDropdown) && actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt) && 
			actselectOperatorDropdown.equals(expselectOperatorDropdown)
			&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueTxt.equalsIgnoreCase(expvalueTxt)
			&& actMessage.equalsIgnoreCase(expMessage))
	{
		System.out.println("Test Pass : Data entered");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Data NOT entered");
		
		return false;
	}
	
		

}


public boolean checkSavingPricebookforTagasCondition() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	
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
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with TagCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
	buyerPriceBook_CustamizeSettingBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_WarehouseCheckBoxSelected));
	buyerPriceBook_WarehouseCheckBoxSelected.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
	buyerPriceBook_CustamizeSettingBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("PAPER");
	Thread.sleep(2000);
	
	
	itemValue.sendKeys(Keys.TAB);
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
	departmentValue.click();
	departmentValue.sendKeys("AMERICA");
	Thread.sleep(2000);
	departmentValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("120.00");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("800");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	//Second row
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("PREMIX A");
	Thread.sleep(2000);
	
	
	itemValue.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
	departmentValue.click();
	Thread.sleep(2000);
	departmentValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	departmentValue.sendKeys("DUBAI");
	Thread.sleep(2000);
	departmentValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("205.00");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("212.32");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	
	System.out.println("*********************************checkSavingSellerPriceBook********************************");
	
	String expMessage="Price Book created successfully";
	
	String actMessage=checkValidationMessage(expMessage);
	
	
	
	
	if (actMessage.equalsIgnoreCase(expMessage)  ) 
	{
	
	
	return true;
	} 
	else 
	{
	
	
	return false;
	}


}




public boolean checkStatusofthePricebookforTagasCondition() throws InterruptedException
{
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with TagCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Pending";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}



}
	


public boolean checkDeletetheConditionTagPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

boolean result=checkDeletingSellerPriceBook("Pricebook With SellerAuthorization with TagCondition");

if(result==true )
{
return true;
}
else
{
return false;
}

}


public boolean checkStatusofthePricebookAfterAuthforTagasCondition() throws InterruptedException
{
	
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
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With SellerAuthorization with TagCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Authorized";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	

	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}



}
	





public boolean checkSavingBuyerPricebookforTagasCondition() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	
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
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with TagCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
	buyerPriceBook_CustamizeSettingBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_WarehouseCheckBoxSelected));
	buyerPriceBook_WarehouseCheckBoxSelected.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
	buyerPriceBook_CustamizeSettingBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("PAPER");
	Thread.sleep(2000);
	
	
	itemValue.sendKeys(Keys.TAB);
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
	wareHouseValue.click();
	wareHouseValue.sendKeys("HYDERABAD");
	Thread.sleep(2000);
	wareHouseValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("120.00");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("800");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	//Second row
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("PREMIX A");
	Thread.sleep(2000);
	
	
	itemValue.sendKeys(Keys.TAB);
	
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
	wareHouseValue.click();
	wareHouseValue.sendKeys("SECUNDRABAD");
	Thread.sleep(2000);
	wareHouseValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("205.00");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("212.32");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	
	System.out.println("*********************************checkSavingBuyerPriceBook********************************");
	
	String expMessage="Price Book created successfully";
	
	String actMessage=checkValidationMessage(expMessage);
	
	
	
	
	if (actMessage.equalsIgnoreCase(expMessage)  ) 
	{
	
	
	return true;
	} 
	else 
	{
	
	
	return false;
	}


}




public boolean checkStatusoftheBuyerPricebookforTagasCondition() throws InterruptedException
{
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with TagCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Pending";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}



}
	


public boolean checkDeletetheConditionTagBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

boolean result=checkDeletingSellerPriceBook("Pricebook With Buyer Authorization with TagCondition");

if(result==true )
{
return true;
}
else
{
return false;
}

}


public boolean checkStatusoftheBuyerPricebookAfterAuthforTagasCondition() throws InterruptedException
{
	
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
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with TagCondition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Authorized";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	

	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}



}
	

@FindBy(xpath="(//*[@id='SliderMenu']/a[1]/i)[1]")
private static WebElement newLayout;


@FindBy(xpath="(//*[@id='drpdownSlider']/a/i)[3]")
private static WebElement saveDropdown;


@FindBy(xpath="(//*[@id='SliderMenu']/a[1])[3]")
private static WebElement saveLayout;


@FindBy(xpath="//*[@id='id_PopuplayoutName']")
private static WebElement saveTxt;

@FindBy(xpath="//*[@id='btn_PopUpOk']")
private static WebElement saveLayoutBtn;

@FindBy(xpath="//*[@id='li_InvoicePreview']")
private static WebElement previewLayout;

@FindBy(xpath="//*[@id='id_InvoiceDesingCancel']")
private static WebElement exitLayoutBtn;


@FindBy(xpath="(//i[@class='icon-close hiconright2'])[1]")
private static WebElement exit_LayoutBtn;


@FindBy(xpath="//td[@id='id_tblAlerts_col_1-3']")
private static WebElement recievedEmailTemplate;

@FindBy(xpath="//select[@id='id_tblAlerts_EmailTemplate']")
private static WebElement receivedEmailTemplateInput;

public boolean checkCreationofEmailTemplateinPricebookAuthorization() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
{
	//PricebookAuthEmailTemplate
	
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Newdropdown));
	  Newdropdown.click();
	  Thread.sleep(2000);
	  
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newLayout));
	  newLayout.click();
	  Thread.sleep(2000);
	  
	  
	  Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\EmailInvoiceLayout1.exe");
		
		Thread.sleep(150000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveDropdown));
		 saveDropdown.click();
		 Thread.sleep(2000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveLayout));
		 saveLayout.click();
		 Thread.sleep(2000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveTxt));
		 saveTxt.sendKeys("newAuthEmailLayout");
		 saveTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveLayoutBtn));
		 saveLayoutBtn.click();
		 Thread.sleep(2000);
		 String expMsg="Layout saved successfully";
		 String actMsg=checkValidationMessage(expMsg);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exit_LayoutBtn));
		 exit_LayoutBtn.click();
		 Thread.sleep(2000);
		 
		 if(actMsg.equalsIgnoreCase(expMsg))
		 {
			 return true;
		 }
		 else
		 {
			 return false;
		 }
	
}


@FindBy(xpath="//*[@id='id_tblAlerts_col_1-3']")
private static WebElement receivedEmailTxt;

public boolean checkCreatedEmailTemplate() throws InterruptedException
{
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(recievedEmailTemplate));
	recievedEmailTemplate.click();
	Thread.sleep(2000);
	
	
	receivedEmailTemplateInput.click();
	Thread.sleep(4000);
	Select s=new Select(receivedEmailTemplateInput);
	s.selectByVisibleText("newAuthEmailLayout");
	//s.selectByValue("10");
	Thread.sleep(2000);
	receivedEmailTemplateInput.sendKeys(Keys.TAB);
	
	
	
	Thread.sleep(2000);
	
	
	 String actreceivedEmailTemplateInput=receivedEmailTxt.getText();
	
	
	String expreceivedEmailTemplateInput="newAuthEmailLayout";
	

	
	System.out.println("actreceivedEmailTemplateInput 		: "		+		actreceivedEmailTemplateInput);
	System.out.println("expreceivedEmailTemplateInput 		: "		+		expreceivedEmailTemplateInput);
	
	System.out.println("**************************************checkCreatedEmailInEmailTemplate*********************");
	
	System.out.println(" receivedEmailHeaderTemplateInput  :"+actreceivedEmailTemplateInput+"  Value Expected : "+expreceivedEmailTemplateInput);
	
	if(actreceivedEmailTemplateInput.equalsIgnoreCase(expreceivedEmailTemplateInput))
	{
		System.out.println("Test Pass : Alerts TAb Options Are Displayed");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
		
		return false;
	}
	
}





public boolean checkSaveAlertsTemplate() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertsHeaderBtn));
	alertsHeaderBtn.click();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTemplates));
	emailTemplates.sendKeys("SellerPricebookAlertTemplate");

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyNameTemplateArea));
	companyNameTemplateArea.click();
	
	Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\EmailHeaderTemplate1.exe");
	
	Thread.sleep(150000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtnInEmailTemplateHeader));
	saveBtnInEmailTemplateHeader.click();
	Thread.sleep(2000);
	
	System.out.println("**************************************checkSaveWithAllMandatoryInputs*******************************");
	
	String expMessage="Template saved successfully";
	 
	String actMessage=checkValidationMessage(expMessage);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeButtonInCreateTemplate));
	//closeButtonInCreateTemplate.click();
	
	ClickUsingJs(closeButtonInCreateTemplate);
	Thread.sleep(2000);
	
	if(actMessage.equalsIgnoreCase(expMessage))
	{
		System.out.println("Test Pass : Alerts TAb Options Are Displayed");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
		
		return false;
	}

	
}


@FindBy(xpath="//span[@id='id_report_popup_ok']")
private static WebElement alertOkBtn;


public boolean checkCreatedAlertTemplate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(recievedAlertTemplate));
	recievedAlertTemplate.click();
	Thread.sleep(2000);

	receivedAlertTemplateInput.click();
	Thread.sleep(2000);
	Select s=new Select(receivedAlertTemplateInput);
	s.selectByVisibleText("SellerPricebookAlertTemplate");
	
	receivedAlertTemplateInput.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	 String actreceivedAlertTemplateInput=s.getFirstSelectedOption().getText();
	
	
	String expreceivedAlertTemplateInput="SellerPricebookAlertTemplate";
	

	
	System.out.println("actreceivedEmailAlertTemplateInput : "+actreceivedAlertTemplateInput);
	System.out.println("expreceivedEmailAlertTemplateInput : "+expreceivedAlertTemplateInput);
	
	System.out.println("**************************************checkCreatedAlertTemplate*********************");
	
	System.out.println(" receivedEmailAlertTemplateInput  :"+actreceivedAlertTemplateInput+"  Value Expected : "+expreceivedAlertTemplateInput);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertOkBtn));
	alertOkBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
	saveBtn.click();
	
	Thread.sleep(2000);
	System.out.println("**********************************checkSavingAuthorization*****************************");
	
	String expMessage="Authorization Flow saved successfully.";
	
	String actMessage=checkValidationMessage(expMessage);

	
	if(actreceivedAlertTemplateInput.equalsIgnoreCase(expreceivedAlertTemplateInput) && actMessage.equalsIgnoreCase(expMessage))
	{
		System.out.println("Test Pass : Alerts TAb Options Are Displayed");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
		
		return false;
	}

	
	
}


public boolean checkCreatedAlertTemplateofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(recievedAlertTemplate));
	recievedAlertTemplate.click();
	Thread.sleep(2000);

	receivedAlertTemplateInput.click();
	Thread.sleep(2000);
	Select s=new Select(receivedAlertTemplateInput);
	s.selectByVisibleText("BuyerPricebookTemplate");
	
	receivedAlertTemplateInput.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	 String actreceivedAlertTemplateInput=s.getFirstSelectedOption().getText();
	
	
	String expreceivedAlertTemplateInput="BuyerPricebookTemplate";
	

	
	System.out.println("actreceivedEmailAlertTemplateInput : "+actreceivedAlertTemplateInput);
	System.out.println("expreceivedEmailAlertTemplateInput : "+expreceivedAlertTemplateInput);
	
	System.out.println("**************************************checkCreatedAlertTemplate*********************");
	
	System.out.println(" receivedEmailAlertTemplateInput  :"+actreceivedAlertTemplateInput+"  Value Expected : "+expreceivedAlertTemplateInput);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertOkBtn));
	alertOkBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
	saveBtn.click();
	
	Thread.sleep(2000);
	System.out.println("**********************************checkSavingAuthorization*****************************");
	
	String expMessage="Authorization Flow saved successfully.";
	
	String actMessage=checkValidationMessage(expMessage);

	
	if(actreceivedAlertTemplateInput.equalsIgnoreCase(expreceivedAlertTemplateInput) && actMessage.equalsIgnoreCase(expMessage))
	{
		System.out.println("Test Pass : Alerts TAb Options Are Displayed");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
		
		return false;
	}

	
	
}








@FindBy(xpath="//td[@id='id_tblAlerts_col_1-5']")
private static WebElement recievedAlertTemplate;

@FindBy(xpath="//td[@id='id_tblAlerts_col_2-5']")
private static WebElement approvedAlertTemplate;

@FindBy(xpath="//td[@id='id_tblAlerts_col_2-6']")
private static WebElement approvedRecipient;


@FindBy(xpath="//td[@id='id_tblAlerts_col_3-5']")
private static WebElement rejectedAlertTemplate;

@FindBy(xpath="//td[@id='id_tblAlerts_col_3-6']")
private static WebElement rejectedRecipient;


@FindBy(xpath="//td[@id='id_tblAlerts_col_4-5']")
private static WebElement stoppedAlertTemplate;

@FindBy(xpath="//td[@id='id_tblAlerts_col_4-6']")
private static WebElement stoppedRecipient;




@FindBy(xpath="//select[@id='id_tblAlerts_Alerts']")
private static WebElement receivedAlertTemplateInput;

@FindBy(xpath="//select[@id='id_tblAlerts_Recepient']")
private static WebElement approvedRecipientInput;


@FindBy(xpath="//*[@id='id_tblAlerts_col_2-6]")
private static WebElement approvedRecipientInputTxt;


public boolean checkDeletingSellerPriceBook(String bname) throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		homeMasterItemMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
		homeMasterItem_sellerPriceBookMenu.click();
		
		Thread.sleep(8000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.sendKeys(bname);
		Thread.sleep(2000);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		
		Thread.sleep(6000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
		deleteBtn.click();
		Thread.sleep(2000);
		
		getWaitForAlert();
		Thread.sleep(2000);
		getAlert().accept();
		
		Thread.sleep(2000);
		
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



public boolean checkSavingAPricebookafterAlertsCreated() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
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
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
	homeMasterItem_sellerPriceBookMenu.click();
	
	Thread.sleep(4000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Authorization for Alerts");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
	buyerPriceBook_CustamizeSettingBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_WarehouseCheckBoxSelected));
	buyerPriceBook_WarehouseCheckBoxSelected.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
	buyerPriceBook_CustamizeSettingBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("PAPER");
	Thread.sleep(2000);
	
	
	itemValue.sendKeys(Keys.TAB);
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
	departmentValue.click();
	departmentValue.sendKeys("AMERICA");
	Thread.sleep(2000);
	departmentValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("120.00");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("800");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	//Second row
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("PREMIX A");
	Thread.sleep(2000);
	
	
	itemValue.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
	departmentValue.click();
	departmentValue.sendKeys(Keys.SHIFT,Keys.HOME);
	departmentValue.sendKeys("DUBAI");
	Thread.sleep(2000);
	departmentValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("205.00");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("212.32");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	
	System.out.println("*********************************checkSavingSellerPriceBook********************************");
	
	String expMessage="Price Book created successfully";
	
	String actMessage=checkValidationMessage(expMessage);
	
	
	
	
	if (actMessage.equalsIgnoreCase(expMessage)  ) 
	{
	
	
	return true;
	} 
	else 
	{
	
	
	return false;
	}

	
}


public boolean checkSavingABuyerPricebookafterAlertsCreated() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
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
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Buyer Pricebook With Authorization for Alerts");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
	buyerPriceBook_CustamizeSettingBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_WarehouseCheckBoxSelected));
	buyerPriceBook_WarehouseCheckBoxSelected.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
	buyerPriceBook_CustamizeSettingBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("PAPER");
	Thread.sleep(2000);
	
	
	itemValue.sendKeys(Keys.TAB);
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
	wareHouseValue.click();
	wareHouseValue.sendKeys("HYDERABAD");
	Thread.sleep(2000);
	wareHouseValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("120.00");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("800");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	//Second row
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("PREMIX A");
	Thread.sleep(2000);
	
	
	itemValue.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseValue));
	wareHouseValue.click();
	wareHouseValue.sendKeys(Keys.SHIFT,Keys.HOME);
	wareHouseValue.sendKeys("SECUNDRABAD");
	Thread.sleep(2000);
	wareHouseValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("205.00");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("212.32");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	
	System.out.println("*********************************checkSavingBuyerPriceBook********************************");
	
	String expMessage="Price Book created successfully";
	
	String actMessage=checkValidationMessage(expMessage);
	
	
	
	
	if (actMessage.equalsIgnoreCase(expMessage)  ) 
	{
	
	
	return true;
	} 
	else 
	{
	
	
	return false;
	}

	
}



public boolean checkStatusofPricebookforAlerts() throws InterruptedException
{
	

	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Authorization for Alerts");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Pending";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}
}


public boolean checkStatusofBuyerPricebookforAlerts() throws InterruptedException
{
	

	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Buyer Pricebook With Authorization for Alerts");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Pending";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}
}




public boolean checkStatusofPricebookforAlertsAfterAuthorize() throws InterruptedException
{
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
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Authorization for Alerts");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Authorized";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}

	
}



public boolean checkDeletetheAlertPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

boolean result=checkDeletingSellerPriceBook("Pricebook With Authorization for Alerts");

if(result==true )
{
return true;
}
else
{
return false;
}

}



public boolean checkStatusofBuyerPricebookforAlertsAfterAuthorize() throws InterruptedException
{
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
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Buyer Pricebook With Authorization for Alerts");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Authorized";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}

	
}



public boolean checkDeletetheAlertBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

boolean result=checkDeletingBuyerPriceBook("Buyer Pricebook With Authorization for Alerts");

if(result==true )
{
return true;
}
else
{
return false;
}

}



public boolean checkSelectCreatedAlertTemplateforApproved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
		
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertsTab));
	alertsTab.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(approvedChkBox));
	approvedChkBox.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(approvedAlertTemplate));
	approvedAlertTemplate.click();
	Thread.sleep(2000);

	receivedAlertTemplateInput.click();
	Thread.sleep(2000);
	Select s=new Select(receivedAlertTemplateInput);
	s.selectByVisibleText("SellerPricebookAlertTemplate");
	
	receivedAlertTemplateInput.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	Select s1=new Select(receivedAlertTemplateInput);
	String st=s1.getFirstSelectedOption().getText();
	Thread.sleep(2000);
	
	
	 String actreceivedAlertTemplateInput=receivedAlertTemplateInput.getText();
	
	
	String expreceivedAlertTemplateInput="SellerPricebookAlertTemplate";
	

	
	System.out.println("actreceivedEmailAlertTemplateInput : "+actreceivedAlertTemplateInput);
	System.out.println("expreceivedEmailAlertTemplateInput : "+expreceivedAlertTemplateInput);
	
	System.out.println("**************************************checkCreatedAlertTemplate*********************");
	
	System.out.println(" receivedEmailAlertTemplateInput  :"+actreceivedAlertTemplateInput+"  Value Expected : "+expreceivedAlertTemplateInput);
	
	
	
	if(actreceivedAlertTemplateInput.equalsIgnoreCase(expreceivedAlertTemplateInput) )
	{
		System.out.println("Test Pass : Alerts TAb Options Are Displayed");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
		
		return false;
	}

	
	
}



public boolean checkSelectCreatedAlertTemplateforApprovedofBuyerAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
		
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertsTab));
	alertsTab.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(approvedChkBox));
	approvedChkBox.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(approvedAlertTemplate));
	approvedAlertTemplate.click();
	Thread.sleep(2000);

	receivedAlertTemplateInput.click();
	Thread.sleep(2000);
	Select s=new Select(receivedAlertTemplateInput);
	s.selectByVisibleText("BuyerPricebookTemplate");
	
	receivedAlertTemplateInput.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	Select s1=new Select(receivedAlertTemplateInput);
	String st=s1.getFirstSelectedOption().getText();
	Thread.sleep(2000);
	
	
	 String actreceivedAlertTemplateInput=receivedAlertTemplateInput.getText();
	
	
	String expreceivedAlertTemplateInput="BuyerPricebookTemplate";
	

	
	System.out.println("actreceivedEmailAlertTemplateInput : "+actreceivedAlertTemplateInput);
	System.out.println("expreceivedEmailAlertTemplateInput : "+expreceivedAlertTemplateInput);
	
	System.out.println("**************************************checkCreatedAlertTemplate*********************");
	
	System.out.println(" receivedEmailAlertTemplateInput  :"+actreceivedAlertTemplateInput+"  Value Expected : "+expreceivedAlertTemplateInput);
	
	
	
	if(actreceivedAlertTemplateInput.equalsIgnoreCase(expreceivedAlertTemplateInput) )
	{
		System.out.println("Test Pass : Alerts TAb Options Are Displayed");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
		
		return false;
	}

	
	
}




public boolean checkRecipientforAlertsApproved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(approvedRecipient));
	approvedRecipient.click();
	Thread.sleep(2000);
	
	
	//approvedRecipientInput.click();
	Thread.sleep(2000);
	Select s=new Select(approvedRecipientInput);
	s.selectByVisibleText("Owner");
	Thread.sleep(2000);
	approvedRecipientInput.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	 String actapprovedAlertTemplateInput=approvedRecipientInput.getText();
	
	
	String expapprovedAlertTemplateInput="Owner";
	

	
	System.out.println("actapprovedEmailAlertTemplateInput : "+actapprovedAlertTemplateInput);
	System.out.println("expapprovedEmailAlertTemplateInput : "+expapprovedAlertTemplateInput);
	
	System.out.println("**************************************checkRecipientTemplate*********************");
	
	System.out.println(" approvedEmailAlertTemplateInput  :"+actapprovedAlertTemplateInput+"  Value Expected : "+expapprovedAlertTemplateInput);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertOkBtn));
	alertOkBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
	saveBtn.click();
	
	Thread.sleep(2000);
	System.out.println("**********************************checkSavingAuthorization*****************************");
	
	String expMessage="Authorization Flow saved successfully.";
	
	String actMessage=checkValidationMessage(expMessage);

	if(actapprovedAlertTemplateInput.equalsIgnoreCase(expapprovedAlertTemplateInput) && actMessage.equalsIgnoreCase(expMessage))
	{
		return true;
	}
	else
	{
		return false;
	}
	
	
}


public boolean checkAlertinOwnerAfterApprovedofAuthorization() throws InterruptedException
{
	Thread.sleep(2000);
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
	alertsBtn.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
	alertExtraMaster.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorization));
	SellerPricebookAuthorization.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationTxt));
	String actAlertTxt=SellerPricebookAuthorizationTxt.getText();
	String expAlertTxt="Pricebook Validation and AuthorizationUserAllOptionsSTPricebook With Authorization for Alerts PAPERAMERICA120.0000000 <<Val 1>>";
	
	
	SellerPricebookAuthorizationTxt.click();
	
	Thread.sleep(35000);
	
	System.out.println("Approved Message	:"	+	actAlertTxt		+"Value Expected	:"		+	expAlertTxt);
	
	if(actAlertTxt.equalsIgnoreCase(expAlertTxt))
	{
		return true;
	}
	else
	{
		return false;
	}
	
}



public boolean checkAlertinOwnerAfterApprovedofAuthorizationofBuyer() throws InterruptedException
{
	Thread.sleep(2000);
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
	alertsBtn.click();

	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
	alertExtraMaster.click();

	getWebDriverWait().until(ExpectedConditions.visibilityOf(BuyerPricebookAuthorization));
	BuyerPricebookAuthorization.click();

	getWebDriverWait().until(ExpectedConditions.visibilityOf(BuyerPricebookAuthorizationTxt));
	String actAlertTxt=BuyerPricebookAuthorizationTxt.getText();
	String expAlertTxt="";
	
	
	BuyerPricebookAuthorizationTxt.click();
	
	System.out.println("Approved Message	:"	+	actAlertTxt		+"Value Expected	:"		+	expAlertTxt);
	
	if(actAlertTxt.equalsIgnoreCase(expAlertTxt))
	{
		return true;
	}
	else
	{
		return false;
	}
	
}


public boolean checkSelectCreatedAlertTemplateforRejected() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
		
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertsTab));
	alertsTab.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rejectedChkBox));
	rejectedChkBox.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rejectedAlertTemplate));
	rejectedAlertTemplate.click();
	Thread.sleep(2000);

	receivedAlertTemplateInput.click();
	Thread.sleep(2000);
	Select s=new Select(receivedAlertTemplateInput);
	s.selectByVisibleText("SellerPricebookAlertTemplate");
	
	receivedAlertTemplateInput.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	Select s1=new Select(receivedAlertTemplateInput);
	String st=s1.getFirstSelectedOption().getText();
	Thread.sleep(2000);
	
	
	 String actreceivedAlertTemplateInput=receivedAlertTemplateInput.getText();
	
	
	String expreceivedAlertTemplateInput="SellerPricebookAlertTemplate";
	

	
	System.out.println("actrejectedEmailAlertTemplateInput : "+actreceivedAlertTemplateInput);
	System.out.println("exprejectedEmailAlertTemplateInput : "+expreceivedAlertTemplateInput);
	
	System.out.println("**************************************checkCreatedAlertTemplate*********************");
	
	System.out.println(" rejectedEmailAlertTemplateInput  :"+actreceivedAlertTemplateInput+"  Value Expected : "+expreceivedAlertTemplateInput);
	
	
	
	if(actreceivedAlertTemplateInput.equalsIgnoreCase(expreceivedAlertTemplateInput) )
	{
		System.out.println("Test Pass : Alerts TAb Options Are Displayed");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
		
		return false;
	}

	
	
}


public boolean checkSelectCreatedAlertTemplateforRejectedofBuyerAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
		
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertsTab));
	alertsTab.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rejectedChkBox));
	rejectedChkBox.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rejectedAlertTemplate));
	rejectedAlertTemplate.click();
	Thread.sleep(2000);

	receivedAlertTemplateInput.click();
	Thread.sleep(2000);
	Select s=new Select(receivedAlertTemplateInput);
	s.selectByVisibleText("BuyerPricebookTemplate");
	
	receivedAlertTemplateInput.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	Select s1=new Select(receivedAlertTemplateInput);
	String st=s1.getFirstSelectedOption().getText();
	Thread.sleep(2000);
	
	
	 String actreceivedAlertTemplateInput=receivedAlertTemplateInput.getText();
	
	
	String expreceivedAlertTemplateInput="BuyerPricebookTemplate";
	

	
	System.out.println("actrejectedEmailAlertTemplateInput : "+actreceivedAlertTemplateInput);
	System.out.println("exprejectedEmailAlertTemplateInput : "+expreceivedAlertTemplateInput);
	
	System.out.println("**************************************checkCreatedAlertTemplate*********************");
	
	System.out.println(" rejectedEmailAlertTemplateInput  :"+actreceivedAlertTemplateInput+"  Value Expected : "+expreceivedAlertTemplateInput);
	
	
	
	if(actreceivedAlertTemplateInput.equalsIgnoreCase(expreceivedAlertTemplateInput) )
	{
		System.out.println("Test Pass : Alerts TAb Options Are Displayed");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
		
		return false;
	}

	
	
}








public boolean checkRecipientforAlertsRejected() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rejectedRecipient));
	rejectedRecipient.click();
	Thread.sleep(2000);
	
	
	//approvedRecipientInput.click();
	Thread.sleep(2000);
	Select s=new Select(approvedRecipientInput);
	s.selectByVisibleText("Owner");
	
	approvedRecipientInput.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	 String actapprovedAlertTemplateInput=s.getFirstSelectedOption().getText();
	
	
	String expapprovedAlertTemplateInput="Owner";
	

	
	System.out.println("**************************************checkRecipientTemplate*********************");
	
	System.out.println(" approvedEmailAlertTemplateInput  :"+actapprovedAlertTemplateInput+"  Value Expected : "+expapprovedAlertTemplateInput);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertOkBtn));
	alertOkBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
	saveBtn.click();
	
	Thread.sleep(2000);
	System.out.println("**********************************checkSavingAuthorization*****************************");
	
	String expMessage="Authorization Flow saved successfully.";
	
	String actMessage=checkValidationMessage(expMessage);

	if(actapprovedAlertTemplateInput.equalsIgnoreCase(expapprovedAlertTemplateInput) && actMessage.equalsIgnoreCase(expMessage))
	{
		return true;
	}
	else
	{
		return false;
	}
	
	
}


public boolean checkRejectionforEditedRowsinSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	Thread.sleep(2000);
			
	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
	alertsBtn.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
	alertExtraMaster.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorization));
	SellerPricebookAuthorization.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationTxt));
	String actAlertTxt=SellerPricebookAuthorizationTxt.getText();
	String expAlertTxt="";
	
	SellerPricebookAuthorizationTxt.click();
	
	

	Thread.sleep(35000);
	
	try
	{
	getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationBtn));
	SellerPricebookAuthorizationBtn.click();
	Thread.sleep(2500);
	}
	
	catch(Exception e)
	{
		
	}
	
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(Auth_selectAllBtn));
	Auth_selectAllBtn.click();
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(rejectionBtn));
	rejectionBtn.click();
	Thread.sleep(2000);
	
	String expMsg="Rejected Successfully";
	String actMsg=checkValidationMessage(expMsg);
	Thread.sleep(4000);
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(cancelBtn));
	//cancelBtn.click();
	ClickUsingJs(cancelBtn);
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	String itemTxt=itemValue.getAttribute("value");
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
	getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Rejected";
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		String itemTxt1=itemValue.getAttribute("value");
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
		getAction().moveToElement(StatusSecondCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus1=StatusSecondCellSelect.getText();
		String expStatus1="Authorized";
		
		
		
		
		System.out.println("Alert Text		:	"+	actAlertTxt 	+	"Expected	:"	+	expAlertTxt);
		
		System.out.println("Item 	:	"	+	itemTxt		+"Status :"	+	actStatus	+"Expected"	+	expStatus);
		System.out.println("Item 	:	"	+	itemTxt1		+"Status :"	+	actStatus1	+"Expected"	+	expStatus1);
		
	
		
		if(actStatus.equalsIgnoreCase(expStatus) && actMsg.equalsIgnoreCase(expMsg)
				&& actStatus1.equalsIgnoreCase(expStatus1) )
		{
			return true;
		}
		else
		{
			return false;
		}

	
}




public boolean checkRejectionforEditedRowsinBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	Thread.sleep(2000);
			
	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
	alertsBtn.click();

	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
	alertExtraMaster.click();

	getWebDriverWait().until(ExpectedConditions.visibilityOf(BuyerPricebookAuthorization));
	BuyerPricebookAuthorization.click();

	getWebDriverWait().until(ExpectedConditions.visibilityOf(BuyerPricebookAuthorizationTxt));
	String actAlertTxt=BuyerPricebookAuthorizationTxt.getText();
	String expAlertTxt="";
	
	BuyerPricebookAuthorizationTxt.click();
	
	

	Thread.sleep(25000);
	
	try
	{
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationBtn));
	SellerPricebookAuthorizationBtn.click();
	Thread.sleep(2500);
	}
	
	catch(Exception e)
	{
		
	}
	
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(Auth_selectAllBtn));
	Auth_selectAllBtn.click();
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(rejectionBtn));
	rejectionBtn.click();
	Thread.sleep(2000);
	
	String expMsg="Rejected Successfully";
	String actMsg=checkValidationMessage(expMsg);
	Thread.sleep(4000);
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(cancelBtn));
	//cancelBtn.click();
	ClickUsingJs(cancelBtn);
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	String itemTxt=itemValue.getAttribute("value");
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
		getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Rejected";
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		String itemTxt1=itemValue.getAttribute("value");
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
		getAction().moveToElement(StatusSecondCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus1=StatusSecondCellSelect.getText();
		String expStatus1="Authorized";
		
		
		
		
		System.out.println("Alert Text		:	"+	actAlertTxt 	+	"Expected	:"	+	expAlertTxt);
		
		System.out.println("Item 	:	"	+	itemTxt		+"Status :"	+	actStatus	+"Expected"	+	expStatus);
		System.out.println("Item 	:	"	+	itemTxt1		+"Status :"	+	actStatus1	+"Expected"	+	expStatus1);
		
	
		
		if(actStatus.equalsIgnoreCase(expStatus) && actMsg.equalsIgnoreCase(expMsg)
				&& actStatus1.equalsIgnoreCase(expStatus1) )
		{
			return true;
		}
		else
		{
			return false;
		}

	
}

public boolean checkAlertinOwnerAfterRejectedofAuthorization() throws InterruptedException
{
	Thread.sleep(2000);
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
	alertsBtn.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
	alertExtraMaster.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorization));
	SellerPricebookAuthorization.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationTxt));
	String actAlertTxt=SellerPricebookAuthorizationTxt.getText();
	String expAlertTxt="Pricebook Validation and AuthorizationUserAllOptionsSTPricebook With Authorization for Alerts PAPERAMERICA120.0000000 <<Val 1>>";
	
	System.out.println("Approved Message	:"	+	actAlertTxt		+"Value Expected	:"		+	expAlertTxt);
	
	if(actAlertTxt.equalsIgnoreCase(expAlertTxt))
	{
		return true;
	}
	else
	{
		return false;
	}
	
}

public boolean checkStatusofthePricebbokafterRejection() throws InterruptedException
{
			
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
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Authorization for Alerts");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();	
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Rejected";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}

	
	
}



public boolean checkStatusoftheBuyerPricebbokafterRejection() throws InterruptedException
{
			
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
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Buyer Pricebook With Authorization for Alerts");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Rejected";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}

	
	
}




public boolean checkStopedforEditedRowsinSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	Thread.sleep(2000);
			
	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
	alertsBtn.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
	alertExtraMaster.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorization));
	SellerPricebookAuthorization.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationTxt));
	String actAlertTxt=SellerPricebookAuthorizationTxt.getText();
	String expAlertTxt="";
	
	SellerPricebookAuthorizationTxt.click();
	
	

	Thread.sleep(35000);
	
	try
	{
	getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationBtn));
	SellerPricebookAuthorizationBtn.click();
	Thread.sleep(2500);
	}
	
	catch(Exception e)
	{
		
	}
	
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(Auth_selectAllBtn));
	Auth_selectAllBtn.click();
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(stopBtn));
	stopBtn.click();
	Thread.sleep(2000);
	
	String expMsg="Stopped Successfully";
	String actMsg=checkValidationMessage(expMsg);
	Thread.sleep(4000);
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(cancelBtn));
	//cancelBtn.click();
	ClickUsingJs(cancelBtn);
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	String itemTxt=itemValue.getAttribute("value");
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
		getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Stopped";
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		String itemTxt1=itemValue.getAttribute("value");
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
		getAction().moveToElement(StatusSecondCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus1=StatusSecondCellSelect.getText();
		String expStatus1="Authorized";
		
		
		
		
		System.out.println("Alert Text		:	"+	actAlertTxt 	+	"Expected	:"	+	expAlertTxt);
		
		System.out.println("Item 	:	"	+	itemTxt		+"Status :"	+	actStatus	+"Expected"	+	expStatus);
		System.out.println("Item 	:	"	+	itemTxt		+"Status :"	+	actStatus1	+"Expected"	+	expStatus1);
		
	
		
		if(actStatus.equalsIgnoreCase(expStatus) && actMsg.equalsIgnoreCase(expMsg)
				&& actStatus1.equalsIgnoreCase(expStatus1) )
		{
			return true;
		}
		else
		{
			return false;
		}

	
}



public boolean checkStopedforEditedRowsinBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	Thread.sleep(2000);
			
	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
	alertsBtn.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
	alertExtraMaster.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(BuyerPricebookAuthorization));
	BuyerPricebookAuthorization.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(BuyerPricebookAuthorizationTxt));
	String actAlertTxt=BuyerPricebookAuthorizationTxt.getText();
	String expAlertTxt="";
	
	SellerPricebookAuthorizationTxt.click();
	
	

	Thread.sleep(35000);
	try
	{
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationBtn));
	SellerPricebookAuthorizationBtn.click();
	Thread.sleep(2500);
	}
	catch(Exception e)
	{
		
	}
	
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(Auth_selectAllBtn));
	Auth_selectAllBtn.click();
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(stopBtn));
	stopBtn.click();
	Thread.sleep(2000);
	
	String expMsg="Stopped Successfully";
	String actMsg=checkValidationMessage(expMsg);
	Thread.sleep(2000);
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(cancelBtn));
	//cancelBtn.click();
	ClickUsingJs(cancelBtn);
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	String itemTxt=itemValue.getAttribute("value");
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
	getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Stopped";
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		String itemTxt1=itemValue.getAttribute("value");
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
		getAction().moveToElement(StatusSecondCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus1=StatusSecondCellSelect.getText();
		String expStatus1="Authorized";
		
		
		
		
		System.out.println("Alert Text		:	"+	actAlertTxt 	+	"Expected	:"	+	expAlertTxt);
		
		System.out.println("Item 	:	"	+	itemTxt		+"Status :"	+	actStatus	+"Expected"	+	expStatus);
		System.out.println("Item 	:	"	+	itemTxt		+"Status :"	+	actStatus1	+"Expected"	+	expStatus1);
		
	
		
		if(actStatus.equalsIgnoreCase(expStatus) && actMsg.equalsIgnoreCase(expMsg)
				&& actStatus1.equalsIgnoreCase(expStatus1) )
		{
			return true;
		}
		else
		{
			return false;
		}

	
}






public boolean checkStatusofthePricebbokafterStopped() throws InterruptedException
{
			
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
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Authorization for Alerts");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Stopped";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}

	
	
}



public boolean checkStatusoftheBuyerPricebbokafterStopped() throws InterruptedException
{
			
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
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Buyer Pricebook With Authorization for Alerts");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Stopped";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	
	
	
	
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	
	
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			)
	{
		return true;
	}
	else
	{
		return false;
	}

	
	
}






public boolean checkSelectCreatedAlertTemplateforStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
		
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertsTab));
	alertsTab.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stoppedChkBox));
	stoppedChkBox.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stoppedAlertTemplate));
	stoppedAlertTemplate.click();
	Thread.sleep(2000);

	receivedAlertTemplateInput.click();
	Thread.sleep(2000);
	Select s=new Select(receivedAlertTemplateInput);
	s.selectByVisibleText("SellerPricebookAlertTemplate");
	
	receivedAlertTemplateInput.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	Select s1=new Select(receivedAlertTemplateInput);
	String st=s1.getFirstSelectedOption().getText();
	Thread.sleep(2000);
	
	
	 String actreceivedAlertTemplateInput=receivedAlertTemplateInput.getText();
	
	
	String expreceivedAlertTemplateInput="SellerPricebookAlertTemplate";
	

	
	System.out.println("actrejectedEmailAlertTemplateInput : "+actreceivedAlertTemplateInput);
	System.out.println("exprejectedEmailAlertTemplateInput : "+expreceivedAlertTemplateInput);
	
	System.out.println("**************************************checkCreatedAlertTemplate*********************");
	
	System.out.println(" rejectedEmailAlertTemplateInput  :"+actreceivedAlertTemplateInput+"  Value Expected : "+expreceivedAlertTemplateInput);
	
	
	
	if(actreceivedAlertTemplateInput.equalsIgnoreCase(expreceivedAlertTemplateInput) )
	{
		System.out.println("Test Pass : Alerts TAb Options Are Displayed");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
		
		return false;
	}

	
	
}



public boolean checkSelectCreatedAlertTemplateforStoppedofBuyerAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
		
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertsTab));
	alertsTab.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stoppedChkBox));
	stoppedChkBox.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stoppedAlertTemplate));
	stoppedAlertTemplate.click();
	Thread.sleep(2000);

	receivedAlertTemplateInput.click();
	Thread.sleep(2000);
	Select s=new Select(receivedAlertTemplateInput);
	s.selectByVisibleText("BuyerPricebookTemplate");
	
	receivedAlertTemplateInput.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	Select s1=new Select(receivedAlertTemplateInput);
	String st=s1.getFirstSelectedOption().getText();
	Thread.sleep(2000);
	
	
	 String actreceivedAlertTemplateInput=receivedAlertTemplateInput.getText();
	
	
	String expreceivedAlertTemplateInput="BuyerPricebookTemplate";
	

	
	System.out.println("actrejectedEmailAlertTemplateInput : "+actreceivedAlertTemplateInput);
	System.out.println("exprejectedEmailAlertTemplateInput : "+expreceivedAlertTemplateInput);
	
	System.out.println("**************************************checkCreatedAlertTemplate*********************");
	
	System.out.println(" rejectedEmailAlertTemplateInput  :"+actreceivedAlertTemplateInput+"  Value Expected : "+expreceivedAlertTemplateInput);
	
	
	
	if(actreceivedAlertTemplateInput.equalsIgnoreCase(expreceivedAlertTemplateInput) )
	{
		System.out.println("Test Pass : Alerts TAb Options Are Displayed");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
		
		return false;
	}

	
	
}



public boolean checkRecipientforAlertsStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stoppedRecipient));
	stoppedRecipient.click();
	Thread.sleep(2000);
	
	
	//approvedRecipientInput.click();
	Thread.sleep(2000);
	Select s=new Select(approvedRecipientInput);
	s.selectByVisibleText("Owner");
	
	approvedRecipientInput.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	 String actapprovedAlertTemplateInput=s.getFirstSelectedOption().getText();
	
	
	String expapprovedAlertTemplateInput="Owner";
	

	
	System.out.println("actapprovedEmailAlertTemplateInput : "+actapprovedAlertTemplateInput);
	System.out.println("expapprovedEmailAlertTemplateInput : "+expapprovedAlertTemplateInput);
	
	System.out.println("**************************************checkRecipientTemplate*********************");
	
	System.out.println(" approvedEmailAlertTemplateInput  :"+actapprovedAlertTemplateInput+"  Value Expected : "+expapprovedAlertTemplateInput);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertOkBtn));
	alertOkBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
	saveBtn.click();
	
	Thread.sleep(2000);
	System.out.println("**********************************checkSavingAuthorization*****************************");
	
	String expMessage="Authorization Flow saved successfully.";
	
	String actMessage=checkValidationMessage(expMessage);

	if(actapprovedAlertTemplateInput.equalsIgnoreCase(expapprovedAlertTemplateInput) && actMessage.equalsIgnoreCase(expMessage))
	{
		return true;
	}
	else
	{
		return false;
	}
	
	
}


public boolean checkAlertinOwnerAfterStoppedofAuthorization() throws InterruptedException
{
	Thread.sleep(2000);
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
	alertsBtn.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
	alertExtraMaster.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorization));
	SellerPricebookAuthorization.click();
	Thread.sleep(2000);
	getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationTxt));
	String actAlertTxt=SellerPricebookAuthorizationTxt.getText();
	String expAlertTxt="Pricebook Validation and AuthorizationUserAllOptionsSTPricebook With Authorization for Alerts PAPERAMERICA120.0000000 <<Val 1>>";
	
	SellerPricebookAuthorizationTxt.click();
	Thread.sleep(8000);
	System.out.println("Approved Message	:"	+	actAlertTxt		+"Value Expected	:"		+	expAlertTxt);
	
	if(actAlertTxt.equalsIgnoreCase(expAlertTxt))
	{
		return true;
	}
	else
	{
		return false;
	}
	
}

//////////////Buyer Pricebook Authorization

public boolean checkSavingBuyerPricebookAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	Thread.sleep(2000);
	getDriver().navigate().refresh();
	Thread.sleep(4000);
	
	
	getAction().moveToElement(settingsmenuBtn).build().perform();
//	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
//	settingsmenuBtn.click();
	ClickUsingJs(settingsmenuBtn);
	
	Thread.sleep(2000);
	
	((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
	Thread.sleep(2000);
	
	
	getAction().moveToElement(pricebookAuthorizationBtn).build().perform();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pricebookAuthorizationBtn));
	pricebookAuthorizationBtn.click();
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pricebookDropdown));
	
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pricebookDropdown));
	pricebookDropdown.click();
	pricebookDropdown.sendKeys(Keys.SPACE);
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BuyingPriceBook));
	BuyingPriceBook.click();
	
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
	nameTxt.sendKeys("Buyer");
	//nameTxt.sendKeys(Keys.TAB);
	
	
	Thread.sleep(2000);
	

	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
	ClickUsingJs(startPointBtn);
	
	
	 Thread.sleep(2000);
	 
	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
		descriptionTxt.click();
		descriptionTxt.sendKeys("BuyerPriceBook");
		Thread.sleep(3000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionTab));
		 conditionTab.click();
		
		 Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
		conditionNotRequiredChkBox.click();
		
		Thread.sleep(3000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
		userSelectionTab.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
		userrow1.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userDropdown));
		userDropdown.click();
		
		userDropdown.sendKeys("UserAllOptionsST");
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
		definitionOkBtn.click();
	 
	
	
	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
		Thread.sleep(2000);
		System.out.println("**********************************checkSavingDefineBudgetAuthorization*****************************");
		
		String expMessage="Authorization Flow saved successfully.";
		
		String actMessage=checkValidationMessage(expMessage);
		
			
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			System.out.println("Error Message is As Expected");
			
			return true;
		}
		else
		{
			System.out.println(" NO Error Message");
			
			return false;
		}
	 
	
	
	
}




public boolean checkCreateBuyerPricebookAfterAuthorizationSaved() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	
	getDriver().navigate().refresh();
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
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization");
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_StartDate));
	buyerPriceBook_StartDate.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_EndDateChkBox));
	if(buyerPriceBook_EndDateChkBox.isSelected()==false)
	{
		buyerPriceBook_EndDateChkBox.click();
	}
	buyerPriceBook_EndDateChkBox.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("Pricebook Item1");
	Thread.sleep(2000);
	
/*		int count = itemComboBoxList.size();
	
	for(int i=0;i<count;i++)
	{
	String data = itemComboBoxList.get(i).getText();
	
	if(data.equalsIgnoreCase("Pricebook Item1"))
	{
	itemComboBoxList.get(i).click();
	break;
	}
	}
*/		
	itemValue.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("103");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("103.25");
	val1Value.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	
	System.out.println("*********************************checkSavingSellerPriceBook********************************");
	
	String expMessage="Price Book created successfully";
	
	String actMessage=checkValidationMessage(expMessage);
	
	
	
	
	if (actMessage.equalsIgnoreCase(expMessage)  ) 
	{
	
	
	return true;
	} 
	else 
	{
	
	
	return false;
	}

	
}


public boolean checkStatusofBuyerPricebookafterAuthorizationSaved() throws InterruptedException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	//buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
//	buyerPriceBook_PriceBook.click();
//	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
//	getAction().moveToElement(StatusFirstCellSelect).build().perform();
//	ScrollToElement(StatusFirstCellSelect);
	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String test=StatusFirstCellSelect.getText();
	String exp="Pending";
	
	
	System.out.println("Status :"	+	test	+"Expected"	+	exp);
	if(test.equalsIgnoreCase(exp))
	{
		return true;
	}
	else
	{
		return false;
	}
}




public boolean checkRateValueloadingbeforeAuthorizationPricebookinPurchaseVoucher() throws InterruptedException
{
	
	checkOpenPurchaseVoucher();
	Thread.sleep(3000);
	
	if(checkAddingFormulaforAccountofPurchase())
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	}

	
	Thread.sleep(2000);
	
	
	
	
	String actRate1=checkRateValueforPurchase("Pricebook Item1");
	String expRate1="0.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt1=enter_ItemTxt.getAttribute("value");
	
	
	System.out.println("Item Name	:	"	+item_Txt1);
	System.out.println("Actual Rate	:	"	+actRate1	+"Expected	:"	+expRate1);
	
	if(actRate1.equalsIgnoreCase(expRate1))
	{
		return true;
	}
	else
	{
		return false;
	}
	
}



public static void checkOpenPurchaseVoucher() throws InterruptedException
{
	
	getDriver().navigate().refresh();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	financialsMenu.click();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesVoucherBtn));
	purchasesVoucherBtn.click();

	Thread.sleep(3000);
}



public boolean checkAuthorizationforEditedRowsinBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	Thread.sleep(2000);
			
	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
	alertsBtn.click();

	getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
	alertExtraMaster.click();

	getWebDriverWait().until(ExpectedConditions.visibilityOf(BuyerPricebookAuthorization));
	BuyerPricebookAuthorization.click();

	getWebDriverWait().until(ExpectedConditions.visibilityOf(BuyerPricebookAuthorizationTxt));
	String actAlertTxt=BuyerPricebookAuthorizationTxt.getText();
	String expAlertTxt="";
	
	BuyerPricebookAuthorizationTxt.click();
	
	

	Thread.sleep(20000);
	
	try
	{
	getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationBtn));
	SellerPricebookAuthorizationBtn.click();
	Thread.sleep(2500);
	}
	
	catch(Exception e)
	{
		
	}
	
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(Auth_selectAllBtn));
	Auth_selectAllBtn.click();
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(authorizeBtn));
	authorizeBtn.click();
	Thread.sleep(2000);
	
	String expMsg="Authorized Successfully";
	String actMsg=checkValidationMessage(expMsg);
	Thread.sleep(4000);
	
	getWebDriverWait().until(ExpectedConditions.visibilityOf(cancelBtn));
	//cancelBtn.click();
	ClickUsingJs(cancelBtn);
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	String itemTxt=itemValue.getAttribute("value");
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
	getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Authorized";
	
	/*	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		String itemTxt1=itemValue.getAttribute("value");
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
		ScrollIntoView(StatusSecondCellSelect);
		Thread.sleep(2000);
		String actStatus1=StatusSecondCellSelect.getText();
		String expStatus1="Authorized";
		
		*/
		
		
		System.out.println("Alert Text		:	"+	actAlertTxt 	+	"Expected	:"	+	expAlertTxt);
		
		System.out.println("Item 	:	"	+	itemTxt		+"Status :"	+	actStatus	+"Expected"	+	expStatus);
		//System.out.println("Item 	:	"	+	itemTxt1		+"Status :"	+	actStatus1	+"Expected"	+	expStatus1);
		
	
		
		if(actStatus.equalsIgnoreCase(expStatus) && actMsg.equalsIgnoreCase(expMsg))
		{
			return true;
		}
		else
		{
			return false;
		}

	
}


public boolean checkRateLoadingafterAuthorizeBuyerPricebook() throws InterruptedException
{
	Thread.sleep(2000);
	checkOpenPurchaseVoucher();
	Thread.sleep(2000);
	
	if(checkAddingFormulaforAccountofPurchase())
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	}

	
	Thread.sleep(2000);
	
		
	
	String actRate1=checkRateValueforPurchase("Pricebook Item1");
	String expRate1="103.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt1=enter_ItemTxt.getAttribute("value");
	
	
	System.out.println("Item Name	:	"	+item_Txt1);
	System.out.println("Actual Rate	:	"	+actRate1	+"Expected	:"	+expRate1);
	
	if(actRate1.equalsIgnoreCase(expRate1))
	{
		return true;
	}
	else
	{
		return false;
	}
}
	
public boolean 	checkEditPriceBookAuthorizationSettingsforBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	Thread.sleep(4000);

	
	getDriver().navigate().refresh();
	Thread.sleep(4000);
	
	getAction().moveToElement(settingsmenuBtn).build().perform();
	
/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
	settingsmenuBtn.click();*/
	
	ClickUsingJs(settingsmenuBtn);
	
	Thread.sleep(2000);
	
	((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
	Thread.sleep(2000);
	
	
	getAction().moveToElement(pricebookAuthorizationBtn).build().perform();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pricebookAuthorizationBtn));
	pricebookAuthorizationBtn.click();
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pricebookDropdown));
	
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pricebookDropdown));
	pricebookDropdown.click();
	pricebookDropdown.sendKeys(Keys.SPACE);
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createBuyingPriceBook));
	createBuyingPriceBook.click();
	
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
	nameTxt.sendKeys("Buyer");
	Thread.sleep(2000);
	nameTxt.sendKeys(Keys.TAB);
	
	
	Thread.sleep(3000);

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
	boolean actstartPointBtn     = startPointBtn.isDisplayed();
	boolean actendpoint          = endpoint.isDisplayed();
	
	boolean expstartPointBtn     = true;
	boolean expendpoint          = true;
	
	boolean res = actstartPointBtn==expstartPointBtn && actendpoint==expendpoint;
	
	String actres = Boolean.toString(res);
	
	String expres = "true";
	
	System.out.println("*************************************checkSavedNameDisplayInNameTextBox************************************");
	
	System.out.println("startPointBtn    :"+actstartPointBtn   +"  Value Expected : "+expstartPointBtn);
	System.out.println("endpoint         :"+actendpoint        +"  Value Expected : "+expendpoint);
	
	if(actres.equalsIgnoreCase(expres))
	{
		System.out.println("Test Pass : Saved Data is Displayed");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Saved Data is NOT Displayed");
		
		return false;
	}
	
	
}


public boolean checkEditingtheBuyerPricebookAfterClickingDateRangeinAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
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
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(3000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_StartDate));
	buyerPriceBook_StartDate.click();
	buyerPriceBook_StartDate.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_EndDateChkBox));
	if(buyerPriceBook_EndDateChkBox.isSelected()==true)
	{
		buyerPriceBook_EndDateChkBox.click();
		Thread.sleep(1000);
	}
	buyerPriceBook_EndDateChkBox.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("Pricebook Item2");
	Thread.sleep(2000);
	

	itemValue.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("203");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("203.25");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_StartDate));
	buyerPriceBook_StartDate.click();
	buyerPriceBook_StartDate.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_EndDateChkBox));
	if(buyerPriceBook_EndDateChkBox.isSelected()==false)
	{
		buyerPriceBook_EndDateChkBox.click();
	}
	
	buyerPriceBook_EndDateChkBox.sendKeys(Keys.TAB);
	Thread.sleep(2000);
		
	
	//getAction().moveToElement(itemThirdCellSelect).build().perform();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
	itemThirdCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("RMA STITEM");
	Thread.sleep(2000);
	

	itemValue.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("303");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("403.25");
	val1Value.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	
	System.out.println("*********************************checkSavingSellerPriceBook********************************");
	
	String expMessage="Price Book updated successfully";
	
	String actMessage=checkValidationMessage(expMessage);
	
	
	
	
	if (actMessage.equalsIgnoreCase(expMessage)  ) 
	{
	
	
	return true;
	} 
	else 
	{
	
	
	return false;
	}
	
	
	
}


public boolean checkStatusofeditBuyerPricebook() throws InterruptedException
{

	
	Thread.sleep(3000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusSecondCellSelect.getText();
	String expStatus="Authorized";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
	itemThirdCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusThirdCellSelect));

	getAction().moveToElement(StatusThirdCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusThirdCellSelect.getText();
	String expStatus1="Pending";
	
	
	
	
	System.out.println("Status :"	+	actStatus	+"Expected"	+	expStatus);
	
	System.out.println("Status :"	+	actStatus1	+"Expected"	+	expStatus1);
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1))
	{
		return true;
	}
	else
	{
		return false;
	}


}


public boolean checkRateLoadingforPendingandAuthorizedPricebookItemsinPurchaseVoucher() throws InterruptedException
{

	
	checkOpenPurchaseVoucher();
	Thread.sleep(2000);
	
	if(checkAddingFormulaforAccountofPurchase())
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	}

	
	Thread.sleep(2000);
	
	
	
	
	String actRate1=checkRateValueforPurchase("Pricebook Item2");
	String expRate1="203.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt1=enter_ItemTxt.getAttribute("value");
	
	
	System.out.println("Item Name	:	"	+item_Txt1);
	System.out.println("Actual Rate	:	"	+actRate1	+"Expected	:"	+expRate1);
	
	

	String actRate2=checkRateValueforPurchase("RMA STITEM");
	String expRate2="0.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt2=enter_ItemTxt.getAttribute("value");
	
	
	System.out.println("Item Name	:	"	+item_Txt2);
	System.out.println("Actual Rate	:	"	+actRate2	+"Expected	:"	+expRate2);
	
	
	
	
	if(actRate1.equalsIgnoreCase(expRate1) && actRate2.equalsIgnoreCase(expRate2))
	{
		return true;
	}
	else
	{
		return false;
	}
	

}



public boolean checkRateLoadingforPendingandAuthorizedPricebookItemsinPurchaseVoucherafterAuthorization() throws InterruptedException
{

	
	checkOpenPurchaseVoucher();
	Thread.sleep(2000);
	
	if(checkAddingFormulaforAccountofPurchase())
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
	}

	
	Thread.sleep(2000);
	
	
	
	
	String actRate1=checkRateValueforPurchase("Pricebook Item2");
	String expRate1="203.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt1=enter_ItemTxt.getAttribute("value");
	
	
	System.out.println("Item Name	:	"	+item_Txt1);
	System.out.println("Actual Rate	:	"	+actRate1	+"Expected	:"	+expRate1);
	
	

	String actRate2=checkRateValueforPurchase("RMA STITEM");
	String expRate2="303.00";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	String item_Txt2=enter_ItemTxt.getAttribute("value");
	
	
	System.out.println("Item Name	:	"	+item_Txt2);
	System.out.println("Actual Rate	:	"	+actRate2	+"Expected	:"	+expRate2);
	
	
	
	
	if(actRate1.equalsIgnoreCase(expRate1) && actRate2.equalsIgnoreCase(expRate2))
	{
		return true;
	}
	else
	{
		return false;
	}
	

}




public boolean checkSavingBuyerPricebookafterPricebookAuthorizewithCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
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
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with Condition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("FEDPAD SLURRY");
	Thread.sleep(2000);
	

	itemValue.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("20.00");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("100");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	//Second row
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("FEDPOUR SLURRY");
	Thread.sleep(2000);
	

	itemValue.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("25.00");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("22.32");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	//Third row
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
	itemThirdCellSelect.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys("BATCH BR ITEM");
	Thread.sleep(2000);
	

	itemValue.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
	rateValue.sendKeys("28.00");
	rateValue.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
	//val1Value.click();
	val1Value.sendKeys("88.32");
	val1Value.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
	buyerPriceBookSaveBtn.click();
	
	System.out.println("*********************************checkSavingSellerPriceBook********************************");
	
	String expMessage="Price Book created successfully";
	
	String actMessage=checkValidationMessage(expMessage);
	
	
	
	
	if (actMessage.equalsIgnoreCase(expMessage)  ) 
	{
	
	
	return true;
	} 
	else 
	{
	
	
	return false;
	}
	
		
	
	
}





public boolean checkStatusofBuyerPricebookforAuthorizationWithRateCondition() throws InterruptedException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with Condition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Authorized";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
	itemThirdCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusThirdCellSelect));

	getAction().moveToElement(StatusThirdCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus2=StatusThirdCellSelect.getText();
	String expStatus2="Pending";
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	System.out.println("Status 		:"	+		actStatus2	+	"Expected"	+		expStatus2);
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			&& actStatus2.equalsIgnoreCase(expStatus2))
	{
		return true;
	}
	else
	{
		return false;
	}


}


public boolean checkStatusofBuyerPricebookAfterAuthorizationWithRateCondition() throws InterruptedException
{
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
	
	buyerPriceBook_PriceBook.click();
	buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
	buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with Condition");
	Thread.sleep(1500);
	buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
	buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
	itemFirstCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

	getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Authorized";
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
	itemThirdCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusThirdCellSelect));

	getAction().moveToElement(StatusThirdCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus2=StatusThirdCellSelect.getText();
	String expStatus2="Authorized";
	
	
	System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
	
	System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
	
	System.out.println("Status 		:"	+		actStatus2	+	"Expected"	+		expStatus2);
	if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
			&& actStatus2.equalsIgnoreCase(expStatus2))
	{
		return true;
	}
	else
	{
		return false;
	}


}



	public boolean checkDeletetheConditionBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{

		boolean result=checkDeletingBuyerPriceBook("Pricebook With Buyer Authorization with Condition");
		
		if(result==true )
		{
		return true;
		}
		else
		{
		return false;
		}

	}


	
	
	public boolean checkSavingBuyerPricebookforProductnameasCondition() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
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
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with ProductNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		
		
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys("RAW Material");
		Thread.sleep(2000);
		

		itemValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("20.00");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("100");
		val1Value.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		
		//Second row
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys("FEDPOUR SLURRY");
		Thread.sleep(2000);
		

		itemValue.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("25.00");
		rateValue.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
		//val1Value.click();
		val1Value.sendKeys("22.32");
		val1Value.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		System.out.println("*********************************checkSavingSellerPriceBook********************************");
		
		String expMessage="Price Book created successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		
		
		
		if (actMessage.equalsIgnoreCase(expMessage)  ) 
		{
		
		
		return true;
		} 
		else 
		{
		
		
		return false;
		}
		

	}
	
	
	public boolean checkStatusoftheBuyerPricebookforProductnameasCondition() throws InterruptedException
	{

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with ProductNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
	
		getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Pending";
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
	
		getAction().moveToElement(StatusSecondCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus1=StatusSecondCellSelect.getText();
		String expStatus1="Authorized";
		
		
		
		
		
		System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
		
		System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
		
		
		if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
				)
		{
			return true;
		}
		else
		{
			return false;
		}
	
	
		
	}
	
	
	public boolean checkStatusoftheBuyerPricebookAfterAuthforProductnameasCondition() throws InterruptedException
	{
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
		
		buyerPriceBook_PriceBook.click();
		buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
		buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		buyerPriceBook_PriceBook.sendKeys("Pricebook With Buyer Authorization with ProductNameCondition");
		Thread.sleep(1500);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
		buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
	
		getAction().moveToElement(StatusFirstCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus=StatusFirstCellSelect.getText();
		String expStatus="Authorized";
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
		itemSecondCellSelect.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
	
		getAction().moveToElement(StatusSecondCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus1=StatusSecondCellSelect.getText();
		String expStatus1="Authorized";
		
		
		
		
		
		System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);
		
		System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);
		
		
		if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
				)
		{
			return true;
		}
		else
		{
			return false;
		}
	
	
		
	}
	
	
	public boolean checkDeletetheConditionProductNameBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{

	boolean result=checkDeletingBuyerPriceBook("Pricebook With Buyer Authorization with ProductNameCondition");

	if(result==true )
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

	        lp.enterPassword(pawslt);

	        Thread.sleep(2000);

	        String compname = "Pricebook Validation and Authorization";

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

	       /* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	        companyLogo.click();*/

	        String getCompanyTxt1=Company_Name.getText();
	        String getLoginCompanyName1=getCompanyTxt1.substring(0, 38);
	        System.out.println("company name  :  "+ getLoginCompanyName1);
	      //  companyLogo.click();

	        String expUserInfo1           ="SU";
	        String expLoginCompanyName1   ="Pricebook Validation and Authorization";

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


	
	
	public boolean checkLogoutPricebookAuthorizationPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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

	
	///Multiple Users and Roles
	

	
public boolean checkSavingPricebookAuthorizationforMultipleUsers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
	userSelectionTab.click();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userselectionDropdown));
	userselectionDropdown.click();
	
	Select s=new Select(userselectionDropdown);
	s.selectByVisibleText("Users");
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(criteriaDropdown));
	criteriaDropdown.click();
	
	String userc=criteriaDropdown.getText();
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
	userrow1.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userDropdown));
	userDropdown.click();
	userDropdown.sendKeys("UserALOptionsST");
	userDropdown.sendKeys(Keys.ENTER);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow2));
	userrow2.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userDropdown));
	userDropdown.click();
	userDropdown.sendKeys("Test");
	userDropdown.sendKeys(Keys.ENTER);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
	definitionOkBtn.click();
	Thread.sleep(4000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
	saveBtn.click();
	
	Thread.sleep(2000);
	System.out.println("**********************************checkSavingDefineBudgetAuthorization*****************************");
	
	String expMessage="Authorization Flow saved successfully.";
	
	String actMessage=checkValidationMessage(expMessage);
	
		
	if(actMessage.equalsIgnoreCase(expMessage))
	{
		System.out.println("Error Message is As Expected");
		
		return true;
	}
	else
	{
		System.out.println(" NO Error Message");
		
		return false;
	}
	
	
}

public boolean checkSavingPricebookforMultipleUsers() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{


Thread.sleep(4000);

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
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Users");
Thread.sleep(1500);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_WarehouseCheckBoxSelected));
if(buyerPriceBook_WarehouseCheckBox.isSelected()==false)
{
	buyerPriceBook_WarehouseCheckBoxSelected.click();
}
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
itemFirstCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("PAPER");
Thread.sleep(2000);


itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.click();
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("120.00");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("800");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);


//Second row
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
itemSecondCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("PREMIX A");
Thread.sleep(2000);


itemValue.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.click();
departmentValue.sendKeys(Keys.SHIFT,Keys.HOME);
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("205.00");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("212.32");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);

//Third Cell
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
itemThirdCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("RAW MATERIAL");
Thread.sleep(2000);


itemValue.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.click();
departmentValue.sendKeys(Keys.SHIFT,Keys.HOME);
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("165.32");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("289.14");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book created successfully";

String actMessage=checkValidationMessage(expMessage);




if (actMessage.equalsIgnoreCase(expMessage)  ) 
{


return true;
} 
else 
{


return false;
}

}


public boolean checkStatusofPricebookforMultipleUsersBeforeAuthorization() throws InterruptedException
{
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));

buyerPriceBook_PriceBook.click();
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Users");
Thread.sleep(1500);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();
Thread.sleep(6000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
itemFirstCellSelect.click();
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

getAction().moveToElement(StatusFirstCellSelect).build().perform();
Thread.sleep(2000);
String actStatus=StatusFirstCellSelect.getText();
String expStatus="Pending";



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
itemSecondCellSelect.click();
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

getAction().moveToElement(StatusSecondCellSelect).build().perform();
Thread.sleep(2000);
String actStatus1=StatusSecondCellSelect.getText();
String expStatus1="Pending";


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
itemThirdCellSelect.click();
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusThirdCellSelect));

getAction().moveToElement(StatusThirdCellSelect).build().perform();
Thread.sleep(2000);
String actStatus2=StatusThirdCellSelect.getText();
String expStatus2="Pending";



System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);

System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);

System.out.println("Status 		:"	+		actStatus2	+	"Expected"	+		expStatus2);

if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
		&& actStatus2.equalsIgnoreCase(expStatus2)	)
{
	return true;
}
else
{
	return false;
}


}

@FindBy(xpath="//*[@id='btnCancelAuth']")
public static WebElement cancelAuthBtn;

public boolean checkAuthorizeMultipleUsersPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{


Thread.sleep(2000);
		
getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
alertsBtn.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
alertExtraMaster.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorization));
SellerPricebookAuthorization.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationTxt));
String actAlertTxt=SellerPricebookAuthorizationTxt.getText();
String expAlertTxt="";

SellerPricebookAuthorizationTxt.click();



Thread.sleep(35000);

try
{

getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationBtn));
SellerPricebookAuthorizationBtn.click();
Thread.sleep(2500);
}

catch(Exception e)
{
	
}

getWebDriverWait().until(ExpectedConditions.visibilityOf(Auth_selectAllBtn));
Auth_selectAllBtn.click();

getWebDriverWait().until(ExpectedConditions.visibilityOf(authorizeBtn));
authorizeBtn.click();
Thread.sleep(2000);

String expMsg="Authorized Successfully";
String actMsg=checkValidationMessage(expMsg);
Thread.sleep(2000);

getWebDriverWait().until(ExpectedConditions.visibilityOf(cancelAuthBtn));
ClickUsingJs(cancelAuthBtn);
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
itemFirstCellSelect.click();
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
String itemTxt=itemValue.getAttribute("value");
itemValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Authorized";

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	String itemTxt1=itemValue.getAttribute("value");
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
	itemThirdCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	String itemTxt2=itemValue.getAttribute("value");
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusThirdCellSelect));
	getAction().moveToElement(StatusThirdCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus2=StatusThirdCellSelect.getText();
	String expStatus2="Authorized";
	
	
	System.out.println("Alert Text		:	"+	actAlertTxt 	+	"Expected	:"	+	expAlertTxt);
	
	System.out.println("Item 	:	"	+	itemTxt			+"Status :"	+	actStatus	+"Expected"	+	expStatus);
	System.out.println("Item 	:	"	+	itemTxt1		+"Status :"	+	actStatus1	+"Expected"	+	expStatus1);
	
	System.out.println("Item 	:	"	+	itemTxt2		+"Status :"	+	actStatus2	+"Expected"	+	expStatus2);
	
	if(actStatus.equalsIgnoreCase(expStatus) && actMsg.equalsIgnoreCase(expMsg)
			&& actStatus1.equalsIgnoreCase(expStatus1) )
	{
		return true;
	}
	else
	{
		return false;
	}

}

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[4]/td[2]")
public static WebElement itemFourthCellSelect; 	

@FindBy(xpath="(//tbody[@id='PriceBookWebGrid_body']//tr[4]//td[1])[1]")
public static WebElement sellerPricebook_fourthRow;

@FindBy(xpath="(//tbody[@id='PriceBookWebGrid_body']//tr[5]//td[1])[1]")
public static WebElement sellerPricebook_fifthRow;

@FindBy(xpath="(//tbody[@id='PriceBookWebGrid_body']//tr[6]//td[1])[1]")
public static WebElement sellerPricebook_sixthRow;

@FindBy(xpath="(//tbody[@id='PriceBookWebGrid_body']//tr[7]//td[1])[1]")
public static WebElement sellerPricebook_seventhRow;

@FindBy(xpath="(//tbody[@id='PriceBookWebGrid_body']//tr[8]//td[1])[1]")
public static WebElement sellerPricebook_eigthRow;

@FindBy(xpath="(//tbody[@id='PriceBookWebGrid_body']//tr[9]//td[1])[1]")
public static WebElement sellerPricebook_ninthRow;

public boolean checkEditingthePricebookAfterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{


Thread.sleep(4000);

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
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Users");
Thread.sleep(1500);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();
Thread.sleep(6000);

// 4th Row
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFourthCellSelect));
itemFourthCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("RESIN");
Thread.sleep(2000);


itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.click();
departmentValue.sendKeys(Keys.SHIFT,Keys.HOME);
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("120.00");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("800");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);
/*

//5th row
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFifthCellSelect));
itemFifthCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("RMA STITEM");
Thread.sleep(2000);


itemValue.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.click();
departmentValue.sendKeys(Keys.SHIFT,Keys.HOME);
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("205.00");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("212.32");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);

*/

Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book updated successfully";

String actMessage=checkValidationMessage(expMessage);


if (actMessage.equalsIgnoreCase(expMessage)  ) 
{


return true;
} 
else 
{


return false;
}



}

@FindBy(xpath="//label[contains(text(),'Delete Row')]")
public static WebElement deleteRow;

public boolean checkDeleteRowinPricebookatUserLevel() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
Thread.sleep(4000);

Thread.sleep(4000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
homeMasterItem_sellerPriceBookMenu.click();

Thread.sleep(6000);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));

buyerPriceBook_PriceBook.click();
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Users");
Thread.sleep(1500);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();
Thread.sleep(4000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sellerPricebook_fourthRow));
getAction().contextClick(sellerPricebook_fourthRow).build().perform();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteRow));
deleteRow.click();

getWaitForAlert();
Thread.sleep(2000);
getAlert().accept();
Thread.sleep(6000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book updated successfully";

String actMessage=checkValidationMessage(expMessage);


if (actMessage.equalsIgnoreCase(expMessage)  ) 
{


return true;
} 
else 
{


return false;
}

}


public boolean checkDeletePricebookforMultipleUsers() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
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
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Users");
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();

Thread.sleep(3000);

Thread.sleep(3000);
((JavascriptExecutor)getDriver()).executeScript("window.scrollTo(0, 0)","");

Thread.sleep(5000);
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


public boolean checkSavingPricebookAuthorizationforMultipleRoles() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
userSelectionTab.click();
Thread.sleep(2000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userselectionDropdown));
userselectionDropdown.click();

Select s=new Select(userselectionDropdown);
s.selectByVisibleText("Roles");
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(criteriaDropdown));
criteriaDropdown.click();

String userc=criteriaDropdown.getText();



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
userrow1.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userDropdown));
userDropdown.click();
userDropdown.sendKeys("Role ST");
userDropdown.sendKeys(Keys.ENTER);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow2));
userrow2.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userDropdown));
userDropdown.click();
userDropdown.sendKeys("Testing");
userDropdown.sendKeys(Keys.ENTER);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
definitionOkBtn.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
saveBtn.click();

Thread.sleep(2000);
System.out.println("**********************************checkSavingDefineBudgetAuthorization*****************************");

String expMessage="Authorization Flow saved successfully.";

String actMessage=checkValidationMessage(expMessage);

	
if(actMessage.equalsIgnoreCase(expMessage))
{
	System.out.println("Error Message is As Expected");
	
	return true;
}
else
{
	System.out.println(" NO Error Message");
	
	return false;
}

}

public boolean checkSavingPricebookforMultipleRoles() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{


Thread.sleep(4000);

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
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Roles");
Thread.sleep(1500);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_WarehouseCheckBoxSelected));
if(buyerPriceBook_WarehouseCheckBox.isSelected()==false)
{
	buyerPriceBook_WarehouseCheckBoxSelected.click();
}

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
itemFirstCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("PAPER");
Thread.sleep(2000);


itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.click();
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("120.00");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("800");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);


//Second row
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
itemSecondCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("PREMIX A");
Thread.sleep(2000);


itemValue.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.click();
departmentValue.sendKeys(Keys.SHIFT,Keys.HOME);
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("205.00");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("212.32");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);

//Third Cell
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
itemThirdCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("RAW MATERIAL");
Thread.sleep(2000);


itemValue.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.click();
departmentValue.sendKeys(Keys.SHIFT,Keys.HOME);
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("165.32");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("289.14");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book created successfully";

String actMessage=checkValidationMessage(expMessage);




if (actMessage.equalsIgnoreCase(expMessage)  ) 
{


return true;
} 
else 
{


return false;
}

}


public boolean checkStatusofPricebookforMultipleRolesBeforeAuthorization() throws InterruptedException
{
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));

buyerPriceBook_PriceBook.click();
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Roles");
Thread.sleep(1500);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
itemFirstCellSelect.click();
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));

getAction().moveToElement(StatusFirstCellSelect).build().perform();
Thread.sleep(2000);
String actStatus=StatusFirstCellSelect.getText();
String expStatus="Pending";



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
itemSecondCellSelect.click();
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));

getAction().moveToElement(StatusSecondCellSelect).build().perform();
Thread.sleep(2000);
String actStatus1=StatusSecondCellSelect.getText();
String expStatus1="Pending";


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
itemThirdCellSelect.click();
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusThirdCellSelect));

getAction().moveToElement(StatusThirdCellSelect).build().perform();
Thread.sleep(2000);
String actStatus2=StatusThirdCellSelect.getText();
String expStatus2="Pending";



System.out.println("Status 		:"	+		actStatus	+	"Expected"	+		expStatus);

System.out.println("Status 		:"	+		actStatus1	+	"Expected"	+		expStatus1);

System.out.println("Status 		:"	+		actStatus2	+	"Expected"	+		expStatus2);

if(actStatus.equalsIgnoreCase(expStatus) && actStatus1.equalsIgnoreCase(expStatus1)
		&& actStatus2.equalsIgnoreCase(expStatus2)	)
{
	return true;
}
else
{
	return false;
}


}

public boolean checkAuthorizeMultipleRolesPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

Thread.sleep(2000);
		
getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
alertsBtn.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
alertExtraMaster.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorization));
SellerPricebookAuthorization.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationTxt));
String actAlertTxt=SellerPricebookAuthorizationTxt.getText();
String expAlertTxt="";

SellerPricebookAuthorizationTxt.click();


Thread.sleep(35000);

try
{

getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationBtn));
SellerPricebookAuthorizationBtn.click();
Thread.sleep(2500);
}

catch(Exception e)
{
	
}

getWebDriverWait().until(ExpectedConditions.visibilityOf(Auth_selectAllBtn));
Auth_selectAllBtn.click();

getWebDriverWait().until(ExpectedConditions.visibilityOf(authorizeBtn));
authorizeBtn.click();
Thread.sleep(2000);

String expMsg="Authorized Successfully";
String actMsg=checkValidationMessage(expMsg);
Thread.sleep(2000);

getWebDriverWait().until(ExpectedConditions.visibilityOf(cancelAuthBtn));
ClickUsingJs(cancelAuthBtn);
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
itemFirstCellSelect.click();
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
String itemTxt=itemValue.getAttribute("value");
itemValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFirstCellSelect));
getAction().moveToElement(StatusFirstCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFirstCellSelect.getText();
	String expStatus="Authorized";

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSecondCellSelect));
	itemSecondCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	String itemTxt1=itemValue.getAttribute("value");
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusSecondCellSelect));
	getAction().moveToElement(StatusSecondCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus1=StatusSecondCellSelect.getText();
	String expStatus1="Authorized";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
	itemThirdCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	String itemTxt2=itemValue.getAttribute("value");
	itemValue.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusThirdCellSelect));
	getAction().moveToElement(StatusThirdCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus2=StatusThirdCellSelect.getText();
	String expStatus2="Authorized";
	
	
	System.out.println("Alert Text		:	"+	actAlertTxt 	+	"Expected	:"	+	expAlertTxt);
	
	System.out.println("Item 	:	"	+	itemTxt			+"Status :"	+	actStatus	+"Expected"	+	expStatus);
	System.out.println("Item 	:	"	+	itemTxt1		+"Status :"	+	actStatus1	+"Expected"	+	expStatus1);
	
	System.out.println("Item 	:	"	+	itemTxt2		+"Status :"	+	actStatus2	+"Expected"	+	expStatus2);
	
	if(actStatus.equalsIgnoreCase(expStatus) && actMsg.equalsIgnoreCase(expMsg)
			&& actStatus1.equalsIgnoreCase(expStatus1) )
	{
		return true;
	}
	else
	{
		return false;
	}

}

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[5]/td[2]")
public static WebElement itemFifthCellSelect; 	

public boolean checkEditingthePricebookRolesAfterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{


Thread.sleep(4000);

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
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Roles");
Thread.sleep(1500);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();
Thread.sleep(2000);

// 4th Row
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFourthCellSelect));
itemFourthCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("RESIN");
Thread.sleep(2000);


itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.click();
departmentValue.sendKeys(Keys.SHIFT,Keys.HOME);
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("120.00");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("800");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);


//5th row
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFifthCellSelect));
itemFifthCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("RMA STITEM");
Thread.sleep(2000);


itemValue.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.click();
departmentValue.sendKeys(Keys.SHIFT,Keys.HOME);
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("205.00");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("212.32");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);



Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book updated successfully";

String actMessage=checkValidationMessage(expMessage);


if (actMessage.equalsIgnoreCase(expMessage)  ) 
{


return true;
} 
else 
{


return false;
}



}



public boolean checkDeleteRowinPricebookofRolesatUserLevel() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
Thread.sleep(4000);



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
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Roles");
Thread.sleep(1500);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sellerPricebook_fifthRow));
getAction().contextClick(sellerPricebook_fifthRow).build().perform();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteRow));
deleteRow.click();

getWaitForAlert();
Thread.sleep(2000);
getAlert().accept();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book updated successfully";

String actMessage=checkValidationMessage(expMessage);


if (actMessage.equalsIgnoreCase(expMessage)  ) 
{


return true;
} 
else 
{


return false;
}

}


public boolean checkDeletePricebookforMultipleRoles() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
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
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Roles");
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();

Thread.sleep(3000);

Thread.sleep(3000);
((JavascriptExecutor)getDriver()).executeScript("window.scrollTo(0, 0)","");

Thread.sleep(5000);
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


public boolean checkLogoutandLoginWithTestwithMultipleUsers() throws InterruptedException
{
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
	userNameDisplayLogo.click();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
	logoutOption.click();

	Thread.sleep(4000);

	LoginPage lp=new LoginPage(getDriver()); 

	String unamelt="Test";

	String pawslt="su";

	lp.enterUserName(unamelt);

	lp.enterPassword(pawslt);
	
	  String compname = "Pricebook Validation and Authorization";

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

	

	Thread.sleep(8000);

	String actUserInfo=userNameTestDisplay.getText();

	System.out.println("User Info  : "+actUserInfo);

	System.out.println("User Info Capture Text  :  "+userNameTestDisplay.getText());

	//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	//companyLogo.click();

	String getCompanyTxt=Company_Name.getText();
	String getLoginCompanyName=getCompanyTxt.substring(0, 38);
	System.out.println("company name  :  "+ getLoginCompanyName);
	//companyLogo.click();


	String expUserInfo           ="Test";
	String expLoginCompanyName   ="Pricebook Validation and Authorization";

	boolean actDashboard = dashboard.isDisplayed();
	boolean expDashboard = true;

	getDriver().navigate().refresh();

	Thread.sleep(2000);

	System.out.println("**********************************************checkLogoutAndLoginWithUserAllOptionsSTToAuthorizePurchaseOrder*********************************************");

	System.out.println("User Info               : "+actUserInfo         +" Value Expected "+expUserInfo);
	System.out.println("Login Company           : "+getLoginCompanyName +" Value Expected "+expLoginCompanyName);
	System.out.println("Dashboard is Displaying : "+actDashboard        +" Value Expected "+expDashboard );

	if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName) && actDashboard==expDashboard) 
	{
		return true;
	}
	else
	{
		return false;
	}
}


public boolean checkAlertDisplayedinUserAllOptionsSTUser() throws InterruptedException
{
Thread.sleep(2000);

getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
alertsBtn.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
alertExtraMaster.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorization));
SellerPricebookAuthorization.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationTxt));
String actAlertTxt=SellerPricebookAuthorizationTxt.getText();
String expAlertTxt="Pricebook Validation and AuthorizationUserAllOptionsSTPricebook For Multiple Users RESINAMERICA120.0000000 <<Val 1>>";

System.out.println(actAlertTxt);
System.out.println(expAlertTxt);

if(actAlertTxt.equalsIgnoreCase(expAlertTxt))
{
	return true;
}
else
{
	return false;
}
}

@FindBy(xpath="//*[@id='alertsUL']")
public static WebElement alertEmpty;


public boolean checkAlertDisplayedAfterAuthorization() throws InterruptedException
{
Thread.sleep(2000);
getDriver().navigate().refresh();

Thread.sleep(2000);

getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
alertsBtn.click();

boolean actAlertTxt=extraMasterTxt.getText().isEmpty();

System.out.println(actAlertTxt);

if(actAlertTxt)
{
	return true;
}
else
{
	return false;
}

}


@FindBy(xpath="(//tbody[@id='PriceBookWebGrid_body']//tr[3]//td[1])[1]")
public static WebElement sellerPricebook_thirdRow;

@FindBy(xpath="//label[contains(text(),'Duplicate Row')]")
public static WebElement duplicateRow;

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[4]/td[6]")
public static WebElement endDate4thRow;

@FindBy(xpath="//*[@id='PriceBookWebGrid_option_EndingDate_checkbox']")
public static WebElement endDateChkBox4thRow;

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[4]/td[12]")
public static WebElement pricebook_rateValue;

@FindBy(xpath="//*[@id='chkDateRange']/following-sibling::span")
public static WebElement buyerPriceBook_CustamizeDateRangeCheckBoxSelected;


public boolean checkAddingDuplicateRowinPricebook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

Thread.sleep(4000);

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
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Users");
Thread.sleep(1500);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeDateRangeCheckBoxSelected));
if(buyerPriceBook_CustamizeDateRangeCheckBox.isSelected()==false)
{
	buyerPriceBook_CustamizeDateRangeCheckBoxSelected.click();
}
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sellerPricebook_thirdRow));
getAction().contextClick(sellerPricebook_thirdRow).build().perform();
Thread.sleep(2000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(duplicateRow));
duplicateRow.click();

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(endDate4thRow));
endDate4thRow.click();


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(endDateChkBox4thRow));
endDateChkBox4thRow.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pricebook_rateValue));
pricebook_rateValue.click();
//pricebook_rateValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));

rateValue.sendKeys("110.00");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("200.10");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);

Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book updated successfully";

String actMessage=checkValidationMessage(expMessage);


if (actMessage.equalsIgnoreCase(expMessage)  ) 
{

return true;
} 
else 
{
return false;
}

}


public boolean checkAuthorizingtheRowfromHomePageinPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

Thread.sleep(4000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
homeMasterItem_sellerPriceBookMenu.click();

Thread.sleep(5000);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));

buyerPriceBook_PriceBook.click();
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Users");
Thread.sleep(1500);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();
Thread.sleep(6000);

try
{

getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationBtn));
SellerPricebookAuthorizationBtn.click();
Thread.sleep(8000);
}

catch(Exception e)
{
	
}


getWebDriverWait().until(ExpectedConditions.visibilityOf(Auth_selectAllBtn));
Auth_selectAllBtn.click();

getWebDriverWait().until(ExpectedConditions.visibilityOf(authorizeBtn));
authorizeBtn.click();
Thread.sleep(2000);

String expMsg="Authorized Successfully";
String actMsg=checkValidationMessage(expMsg);
Thread.sleep(2000);

getWebDriverWait().until(ExpectedConditions.visibilityOf(cancelAuthBtn));
ClickUsingJs(cancelAuthBtn);
Thread.sleep(2000);

if(actMsg.equalsIgnoreCase(expMsg))
{
	return true;
}
else
{
	return false;
}

}

@FindBy(xpath = "//label[contains(text(),'Insert Row')]")
public static WebElement insertRowBtn;

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[3]/td[12]")
public static WebElement itemThirdCellRateField; 


public boolean checkInsertingRowinPricebook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{


Thread.sleep(4000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();
Thread.sleep(1500);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();
Thread.sleep(1500);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();
Thread.sleep(1500);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
homeMasterItem_sellerPriceBookMenu.click();

Thread.sleep(6000);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));

buyerPriceBook_PriceBook.click();
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Users");
Thread.sleep(1500);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();
Thread.sleep(6000);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sellerPricebook_thirdRow));
getAction().contextClick(sellerPricebook_thirdRow).build().perform();
Thread.sleep(4000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(insertRowBtn));
insertRowBtn.click();

Thread.sleep(3000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();
Thread.sleep(2000);

if(buyerPriceBook_CustamizeDateRangeCheckBox.isSelected()==false)
{
	buyerPriceBook_CustamizeDateRangeCheckBoxSelected.click();
}

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_WarehouseCheckBoxSelected));
if(buyerPriceBook_WarehouseCheckBox.isSelected()==false)
{
	buyerPriceBook_WarehouseCheckBoxSelected.click();
}
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellSelect));
itemThirdCellSelect.click();

Thread.sleep(2000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("ANTI PIPING COMPOUND");
Thread.sleep(4000);


itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemThirdCellRateField));
itemThirdCellRateField.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
//rateValue.click();
rateValue.sendKeys("101.00");
Thread.sleep(2000);
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("600");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book updated successfully";

String actMessage=checkValidationMessage(expMessage);


if (actMessage.equalsIgnoreCase(expMessage)  ) 
{

return true;
} 
else 
{
return false;
}


}


@FindBy(xpath="//*[@id='GridPBAuthorize_body']/tr/td[2]/input")
public static WebElement Auth_FirstRowCheckBox;

public boolean checkAuthorizingRowafterInsertRowinPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{



Thread.sleep(6000);
getDriver().navigate().refresh();
Thread.sleep(6000);
		
getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
alertsBtn.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
alertExtraMaster.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorization));
SellerPricebookAuthorization.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationTxt));
String actAlertTxt=SellerPricebookAuthorizationTxt.getText();
String expAlertTxt="";

SellerPricebookAuthorizationTxt.click();



Thread.sleep(35000);
try
{

getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationBtn));
SellerPricebookAuthorizationBtn.click();
Thread.sleep(2500);
}

catch(Exception e)
{
	
}


getWebDriverWait().until(ExpectedConditions.visibilityOf(Auth_FirstRowCheckBox));
Auth_FirstRowCheckBox.click();

getWebDriverWait().until(ExpectedConditions.visibilityOf(authorizeBtn));
authorizeBtn.click();
Thread.sleep(2000);

String expMsg="Authorized Successfully";
String actMsg=checkValidationMessage(expMsg);
Thread.sleep(2000);

getWebDriverWait().until(ExpectedConditions.visibilityOf(cancelAuthBtn));
ClickUsingJs(cancelAuthBtn);
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFifthCellSelect));
itemFifthCellSelect.click();
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
String itemTxt=itemValue.getAttribute("value");
itemValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(StatusFifthCellSelect));
getAction().moveToElement(StatusFifthCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusFifthCellSelect.getText();
	String expStatus="Authorized";

	
	
	
	System.out.println("Alert Text		:	"+	actAlertTxt 	+	"Expected	:"	+	expAlertTxt);
	
	System.out.println("Item 	:	"	+	itemTxt		+"Status :"	+	actStatus	+"Expected"	+	expStatus);

	if(actStatus.equalsIgnoreCase(expStatus) && actMsg.equalsIgnoreCase(expMsg) )
	{
		return true;
	}
	else
	{
		return false;
	}


}

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[5]/td[29]")
public static WebElement StatusFifthCellSelect; 

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[6]/td[29]")
public static WebElement StatusSixthCellSelect; 


@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[7]/td[29]")
public static WebElement StatusSeventhCellSelect; 


@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[8]/td[29]")
public static WebElement StatusEigthCellSelect; 

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[9]/td[29]")
public static WebElement StatusNinthCellSelect; 

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[10]/td[29]")
public static WebElement StatusTenthCellSelect; 


@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[11]/td[29]")
public static WebElement StatusEleventhCellSelect; 

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[6]/td[2]")
public static WebElement itemSixthCellSelect; 	

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[7]/td[2]")
public static WebElement itemSeventhCellSelect; 	

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[7]/td[12]")
public static WebElement itemSeventhCellRateField; 

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[8]/td[2]")
public static WebElement itemEigthCellSelect; 	

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[8]/td[12]")
public static WebElement itemEigthCellRateField; 

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[9]/td[2]")
public static WebElement itemNinthCellSelect; 	

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[9]/td[12]")
public static WebElement itemNinthCellRateField; 

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[10]/td[2]")
public static WebElement itemTenthCellSelect; 	

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[10]/td[12]")
public static WebElement itemTenthCellRateField; 

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[11]/td[2]")
public static WebElement itemEleventhCellSelect; 

@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[11]/td[12]")
public static WebElement itemEleventhCellRateField; 


public boolean checkSavingPricebookforDeletePendingRow() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{



Thread.sleep(4000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();
Thread.sleep(1500);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();
Thread.sleep(1500);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();
Thread.sleep(1500);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
homeMasterItem_sellerPriceBookMenu.click();

Thread.sleep(8000);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));

buyerPriceBook_PriceBook.click();
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Users");
Thread.sleep(4000);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

Thread.sleep(8000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();
Thread.sleep(6000);


/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();
Thread.sleep(4000);

if(buyerPriceBook_CustamizeDateRangeCheckBox.isSelected()==true)
{
	buyerPriceBook_CustamizeDateRangeCheckBoxSelected.click();
}

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_WarehouseCheckBoxSelected));
if(buyerPriceBook_WarehouseCheckBox.isSelected()==false)
{
	buyerPriceBook_WarehouseCheckBoxSelected.click();
}
Thread.sleep(2000);

getAction().moveToElement(buyerPriceBook_CustamizeSettingBtn).build().perform();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();
Thread.sleep(2000);
*/
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemEigthCellSelect));
itemEigthCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("CERAMIC CHOPPED FIBRE");
Thread.sleep(2000);


itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemEigthCellRateField));
itemEigthCellRateField.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("205.36");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("125.22");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemNinthCellSelect));
itemNinthCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("DIRECT POUR SYSTEMS");
Thread.sleep(2000);


itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemNinthCellRateField));
itemNinthCellRateField.click();


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("145.88");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("222.33");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book updated successfully";

String actMessage=checkValidationMessage(expMessage);


if (actMessage.equalsIgnoreCase(expMessage)  ) 
{

return true;
} 
else 
{
return false;
}



}


public boolean checkDeleteRowsinUserLevelforAlertDisappear() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{


Thread.sleep(2000);

getDriver().navigate().refresh();
Thread.sleep(6000);
		
getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
alertsBtn.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
alertExtraMaster.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorization));
SellerPricebookAuthorization.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationTxt));
String actAlertTxt=SellerPricebookAuthorizationTxt.getText();
String expAlertTxt="";

SellerPricebookAuthorizationTxt.click();

Thread.sleep(35000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sellerPricebook_eigthRow));
getAction().contextClick(sellerPricebook_eigthRow).build().perform();


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteRow));
deleteRow.click();

getWaitForAlert();
Thread.sleep(2000);
getAlert().accept();
Thread.sleep(6000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sellerPricebook_eigthRow));
getAction().contextClick(sellerPricebook_eigthRow).build().perform();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteRow));
deleteRow.click();

getWaitForAlert();
Thread.sleep(2000);
getAlert().accept();
Thread.sleep(6000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book updated successfully";

String actMessage=checkValidationMessage(expMessage);

if(actMessage.equalsIgnoreCase(expMessage))
{
	return true;
}
else
{
	return false;
}



}


public boolean checkAddingRowManuallyinPricebook() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

Thread.sleep(4000);

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
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Users");
Thread.sleep(1500);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSeventhCellSelect));
itemSeventhCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("BORIC ACID - INDO BORAX MAKE");
Thread.sleep(2000);


itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.click();
departmentValue.sendKeys(Keys.SHIFT,Keys.HOME);
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSeventhCellRateField));
itemSeventhCellRateField.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("120.00");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("480.22");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);


Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book updated successfully";

String actMessage=checkValidationMessage(expMessage);


if (actMessage.equalsIgnoreCase(expMessage)  ) 
{


return true;
} 
else 
{


return false;
}
}


public boolean checkAuthorizingRowAfterAddingRowManually() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
Thread.sleep(2000);
getDriver().navigate().refresh();
Thread.sleep(6000);
		
getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
alertsBtn.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
alertExtraMaster.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorization));
SellerPricebookAuthorization.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationTxt));
String actAlertTxt=SellerPricebookAuthorizationTxt.getText();
String expAlertTxt="";

SellerPricebookAuthorizationTxt.click();



Thread.sleep(35000);

try
{

getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationBtn));
SellerPricebookAuthorizationBtn.click();
Thread.sleep(2500);

}

catch(Exception e)
{
	
}


getWebDriverWait().until(ExpectedConditions.visibilityOf(Auth_FirstRowCheckBox));
Auth_FirstRowCheckBox.click();

getWebDriverWait().until(ExpectedConditions.visibilityOf(authorizeBtn));
authorizeBtn.click();
Thread.sleep(2000);

String expMsg="Authorized Successfully";
String actMsg=checkValidationMessage(expMsg);
Thread.sleep(2000);

getWebDriverWait().until(ExpectedConditions.visibilityOf(cancelAuthBtn));
ClickUsingJs(cancelAuthBtn);
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSeventhCellSelect));
itemSeventhCellSelect.click();
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
String itemTxt=itemValue.getAttribute("value");
itemValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSeventhCellSelect));
getAction().moveToElement(StatusSeventhCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusSeventhCellSelect.getText();
	String expStatus="Authorized";

	
	
	
	System.out.println("Alert Text		:	"+	actAlertTxt 	+	"Expected	:"	+	expAlertTxt);
	
	System.out.println("Item 	:	"	+	itemTxt		+"Status :"	+	actStatus	+"Expected"	+	expStatus);

	if(actStatus.equalsIgnoreCase(expStatus) && actMsg.equalsIgnoreCase(expMsg) )
	{
		return true;
	}
	else
	{
		return false;
	}

}


public boolean checkSavingPricebookfroApprovingOneRow() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
Thread.sleep(4000);

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
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Users");
Thread.sleep(1500);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();
Thread.sleep(2000);






getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();
Thread.sleep(2000);

if(buyerPriceBook_CustamizeDateRangeCheckBox.isSelected()==true)
{
	buyerPriceBook_CustamizeDateRangeCheckBox.click();
}

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_WarehouseCheckBox));
if(buyerPriceBook_WarehouseCheckBox.isSelected()==false)
{
	buyerPriceBook_WarehouseCheckBox.click();
}
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSeventhCellSelect));
itemSeventhCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("CERAMIC CHOPPED FIBRE");
Thread.sleep(2000);


itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("205.36");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("125.22");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemEigthCellSelect));
itemEigthCellSelect.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
itemValue.sendKeys("DIRECT POUR SYSTEMS");
Thread.sleep(2000);


itemValue.sendKeys(Keys.TAB);




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValue));
departmentValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
departmentValue.sendKeys("AMERICA");
Thread.sleep(2000);
departmentValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
rateValue.sendKeys("145.88");
rateValue.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(val1Value));
//val1Value.click();
val1Value.sendKeys("222.33");
val1Value.sendKeys(Keys.TAB);

Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book updated successfully";

String actMessage=checkValidationMessage(expMessage);


if (actMessage.equalsIgnoreCase(expMessage)  ) 
{

return true;
} 
else 
{
return false;
}

}


public boolean checkAuthorizingOneRowinSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
Thread.sleep(2000);
		
getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
alertsBtn.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
alertExtraMaster.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorization));
SellerPricebookAuthorization.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationTxt));
String actAlertTxt=SellerPricebookAuthorizationTxt.getText();
String expAlertTxt="";

SellerPricebookAuthorizationTxt.click();



Thread.sleep(35000);

try
{
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationBtn));
SellerPricebookAuthorizationBtn.click();
Thread.sleep(2500);
}

catch(Exception e)
{
	
}


getWebDriverWait().until(ExpectedConditions.visibilityOf(Auth_FirstRowCheckBox));
Auth_FirstRowCheckBox.click();

getWebDriverWait().until(ExpectedConditions.visibilityOf(authorizeBtn));
authorizeBtn.click();
Thread.sleep(2000);

String expMsg="Authorized Successfully";
String actMsg=checkValidationMessage(expMsg);
Thread.sleep(2000);

getWebDriverWait().until(ExpectedConditions.visibilityOf(cancelAuthBtn));
ClickUsingJs(cancelAuthBtn);
Thread.sleep(6000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemEigthCellSelect));
itemEigthCellSelect.click();
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
itemValue.click();
String itemTxt=itemValue.getAttribute("value");
itemValue.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemEigthCellSelect));
	getAction().moveToElement(itemEigthCellSelect).build().perform();
	Thread.sleep(2000);
	String actStatus=StatusEigthCellSelect.getText();
	String expStatus="Authorized";

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemNinthCellSelect));
	itemNinthCellSelect.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
	itemValue.click();
	String itemTxt1=itemValue.getAttribute("value");
	itemValue.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemNinthCellSelect));
	getAction().moveToElement(StatusNinthCellSelect).build().perform();
		Thread.sleep(2000);
		String actStatus1=StatusNinthCellSelect.getText();
		String expStatus1="Pending";
	
	
	System.out.println("Alert Text		:	"+	actAlertTxt 	+	"Expected	:"	+	expAlertTxt);
	
	System.out.println("Item 	:	"	+	itemTxt		+"Status :"	+	actStatus	+"Expected"	+	expStatus);

	if(actStatus.equalsIgnoreCase(expStatus) && actMsg.equalsIgnoreCase(expMsg) )
	{
		return true;
	}
	else
	{
		return false;
	}


}


@FindBy(xpath="(//*[@id='10']/span)[1]")
public static WebElement extraMasterTxt;

public boolean checkAlertTxtAfterApprovingOneRowinPricebook() throws InterruptedException
{

Thread.sleep(2000);

getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsBtn));
alertsBtn.click();
Thread.sleep(4000);


getWebDriverWait().until(ExpectedConditions.visibilityOf(alertExtraMaster));
System.out.println(extraMasterTxt.getText().isEmpty());
alertExtraMaster.click();
Thread.sleep(2000);

getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorization));
SellerPricebookAuthorization.click();
Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(SellerPricebookAuthorizationTxt));
String actAlertTxt=SellerPricebookAuthorizationTxt.getText();
String expAlertTxt="Pricebook Validation and AuthorizationUserAllOptionsSTPricebook For Multiple Users DIRECT POUR SYSTEMSAMERICA145.8800000 <<Val 1>>";

System.out.println("Actual Alert Txt	"	+	actAlertTxt);
System.out.println("Expected Alert Txt	"	+	expAlertTxt);

if(actAlertTxt.equalsIgnoreCase(expAlertTxt))
{
	System.out.println("Alert Displayed Text  matched");
	return true;
}
else {
	System.out.println("Alert Displayed Text not  matched");
	return false;
}

}

public boolean checkDeletingPendingRowinPricebookAfterAuthofOneRow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

Thread.sleep(4000);

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
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Users");
Thread.sleep(1500);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sellerPricebook_ninthRow));
getAction().contextClick(sellerPricebook_ninthRow).build().perform();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteRow));
deleteRow.click();

getWaitForAlert();
Thread.sleep(2000);
getAlert().accept();
Thread.sleep(2000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
buyerPriceBookSaveBtn.click();

System.out.println("*********************************checkSavingSellerPriceBook********************************");

String expMessage="Price Book updated successfully";

String actMessage=checkValidationMessage(expMessage);

if(actMessage.equalsIgnoreCase(expMessage))
{
	return true;
}
else
{
	return false;
}


}
//Update Price List

@FindBy(xpath="//*[@id='BasisOn']")
public static WebElement selectBasisOn;

@FindBy(xpath="//*[@id='chk1']")
public static WebElement FilterOn_PricebookChkBox;

@FindBy(xpath="//*[@id='chk1']//..//span")
public static WebElement FilterOn_PricebookChkBoxSelected;

@FindBy(xpath="//*[@id='ctrlPriceBookUPH']")
public static WebElement FilterOnPricebookTxtBox;

@FindBy(xpath="//*[@id='chk7']")
public static WebElement FilterOn_DateRangeChkBox;

@FindBy(xpath="//*[@id='chk7']//..//span")
public static WebElement FilterOn_DateRangeChkBoxSelected;

@FindBy(xpath="//*[@id='ctrlStartingDateUPH_checkbox']")
public static WebElement FilterOn_DateRangeStartDateChkBox;

@FindBy(xpath="//*[@id='ctrlStartingDateUPH']")
public static WebElement FilterOn_DateRangeStartDateTxtBox;

@FindBy(xpath="//*[@id='ctrlEndingDateUPH_checkbox']")
public static WebElement FilterOn_DateRangeEndDateChkBox;

@FindBy(xpath="//*[@id='ctrlEndingDateUPH']")
public static WebElement FilterOn_DateRangeEndDateTxtBox;

@FindBy(xpath="//*[@id='ddlIncreaseDecrease']")
public static WebElement update_SelectAction;

@FindBy(xpath="//*[@id='rdoOverride1']")
public static WebElement update_OverrideRadioBtn;

@FindBy(xpath="//*[@id='ctrlStartingDateUP_checkbox']")
public static WebElement UpdateValue_DateRangeStartDateChkBox;

@FindBy(xpath="//*[@id='ctrlStartingDateUP']")
public static WebElement UpdateValue_DateRangeStartDateTxtBox;

@FindBy(xpath="//*[@id='ctrlEndingDateUP_checkbox']")
public static WebElement UpdateValue_DateRangeEndDateChkBox;

@FindBy(xpath="//*[@id='ctrlEndingDateUP']")
public static WebElement UpdateValue_DateRangeEndDateTxtBox;


@FindBy(xpath="//*[@id='btnOkForUpdate']")
public static WebElement UpdatePrice_OkBtn;

@FindBy(xpath="//*[@id='btnCloseForUpdate']")
public static WebElement UpdatePrice_CloseBtn;

@FindBy(xpath="//tbody[@id='PriceBookWebGrid_body']//td[5]")
public static List<WebElement> sellerPricebook_StartDateColList;

@FindBy(xpath="//tbody[@id='PriceBookWebGrid_body']//td[6]")
public static List<WebElement> sellerPricebook_EndDateColList;

@FindBy(xpath="//*[@id='btnUpdatePriceList']")
public static WebElement buyerPriceBook_UpdatePriceList;


public boolean checkStartandEndDateinPricebbokUpdatePriceList() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

Thread.sleep(4000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();
Thread.sleep(2000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();
Thread.sleep(2000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
homeMasterItemMenu.click();
Thread.sleep(2000);
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
homeMasterItem_sellerPriceBookMenu.click();

Thread.sleep(5000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
buyerPriceBook_PriceBook.click();
buyerPriceBook_PriceBook.sendKeys(Keys.SPACE);
buyerPriceBook_PriceBook.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
buyerPriceBook_PriceBook.sendKeys("Pricebook For Multiple Users");
Thread.sleep(1500);
buyerPriceBook_PriceBook.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_AbbreviationBtn));
buyerPriceBook_AbbreviationBtn.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
priceBookFilterAndLoadBtn.click();

Thread.sleep(8000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();
Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeDateRangeCheckBoxSelected));
if(buyerPriceBook_CustamizeDateRangeCheckBox.isSelected()==false)
{
	buyerPriceBook_CustamizeDateRangeCheckBoxSelected.click();
}
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_CustamizeSettingBtn));
buyerPriceBook_CustamizeSettingBtn.click();
Thread.sleep(2000);


Date date = Calendar.getInstance().getTime();
DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
String currentDate=df.format(date);
Calendar cal=Calendar.getInstance();

ArrayList<String> startDateArray=new ArrayList<String>();
for(WebElement e:sellerPricebook_StartDateColList)
{
	startDateArray.add(e.getText());
	
}

String actStartDate=startDateArray.toString();
String expStartDate="["+currentDate+", "+currentDate+", "+currentDate+", "+currentDate+", "+currentDate+", "+currentDate+", "+currentDate+", "+currentDate+", , , , , , , ]";

System.out.println("Actual 	 Start Date Before Update 	:"		+	actStartDate);
System.out.println("Expected Start Date Before Update 	:"		+	expStartDate);

ArrayList<String> endDateArray=new ArrayList<String>();
for(WebElement e:sellerPricebook_EndDateColList)
{
	endDateArray.add(e.getText());
	
}
String actEndDate=endDateArray.toString();
String expEndDate="["+currentDate+", "+currentDate+", "+currentDate+", "+currentDate+", , "+currentDate+", "+currentDate+", "+currentDate+", , , , , , , ]";

System.out.println("Actual 	 Start Date Before Update 	:"		+	actEndDate);
System.out.println("Expected Start Date Before Update 	:"		+	expEndDate);



  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_UpdatePriceList)); 
  buyerPriceBook_UpdatePriceList.click();
 
Thread.sleep(4000);



 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectBasisOn)); 
 Select s=new Select(selectBasisOn);
 s.selectByVisibleText("Sales Price");
 Thread.sleep(2000);
 
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FilterOn_PricebookChkBoxSelected));
 if(FilterOn_PricebookChkBox.isSelected()==false)
 {
	 FilterOn_PricebookChkBoxSelected.click();
 }
 FilterOn_PricebookChkBox.sendKeys(Keys.TAB);
 FilterOnPricebookTxtBox.sendKeys("Pricebook For Multiple Users");
 Thread.sleep(2000);
 FilterOnPricebookTxtBox.sendKeys(Keys.TAB);
 Thread.sleep(2000);

 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FilterOn_DateRangeChkBoxSelected));
 if(FilterOn_DateRangeChkBox.isSelected()==false)
 {
	 FilterOn_DateRangeChkBoxSelected.click();
 }
 
 //FilterOn_DateRangeChkBoxSelected.sendKeys(Keys.TAB);
 
 
	//String currentDate=df.format(cal.getTime());
	cal.add(Calendar.DATE, 2); 

	String next2DaysDate=df.format(cal.getTime());
	
	cal.add(Calendar.DATE, 4); 

	String next4DaysDate=df.format(cal.getTime());
 
 
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FilterOn_DateRangeStartDateChkBox));
 if(FilterOn_DateRangeStartDateChkBox.isSelected()==false)
 {
	 FilterOn_DateRangeStartDateChkBox.click();
 }
 FilterOn_DateRangeStartDateChkBox.sendKeys(Keys.TAB);
// FilterOn_DateRangeStartDateTxtBox.sendKeys(currentDate);
 
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FilterOn_DateRangeEndDateChkBox));
 if(FilterOn_DateRangeEndDateChkBox.isSelected()==false)
 {
	 FilterOn_DateRangeEndDateChkBox.click();
 }
 FilterOn_DateRangeEndDateChkBox.sendKeys(Keys.TAB);
// FilterOn_DateRangeEndDateTxtBox.sendKeys(currentDate);
 
 Thread.sleep(2000);
 
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(update_SelectAction));
 Select s1=new Select(update_SelectAction);
 s1.selectByVisibleText("Same");

 
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(update_OverrideRadioBtn));
 if(update_OverrideRadioBtn.isSelected()==false)
 {
	 update_OverrideRadioBtn.click();
 }
 
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UpdateValue_DateRangeStartDateChkBox));
 if(UpdateValue_DateRangeStartDateChkBox.isSelected()==false)
 {
	 UpdateValue_DateRangeStartDateChkBox.click();
 }
 UpdateValue_DateRangeStartDateChkBox.sendKeys(Keys.TAB);
 UpdateValue_DateRangeStartDateTxtBox.sendKeys(next2DaysDate);
 
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UpdateValue_DateRangeEndDateChkBox));
 if(UpdateValue_DateRangeEndDateChkBox.isSelected()==false)
 {
	 UpdateValue_DateRangeEndDateChkBox.click();
 }
 UpdateValue_DateRangeEndDateChkBox.sendKeys(Keys.TAB);
 UpdateValue_DateRangeEndDateTxtBox.sendKeys(next4DaysDate);
 
 Thread.sleep(2000);
 
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UpdatePrice_OkBtn));
 UpdatePrice_OkBtn.click();
 
 checkValidationMessage("Price Book updated successfully");
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UpdatePrice_CloseBtn));
 UpdatePrice_CloseBtn.click();
 Thread.sleep(2000);
 
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
	priceBookFilterAndLoadBtn.click();
	
	Thread.sleep(5000);
	
	
	
	
	ArrayList<String> startDateArray1=new ArrayList<String>();
	for(WebElement e:sellerPricebook_StartDateColList)
	{
		startDateArray1.add(e.getText());
		
	}
	String actStartDate1=startDateArray1.toString();
	String expStartDate1="["+currentDate+", "+next2DaysDate+", "+next2DaysDate+", "+next2DaysDate+", "+next2DaysDate+", "+next2DaysDate+", "+next2DaysDate+", "+next2DaysDate+", , , , , , , ]";
	
	System.out.println("Actual 	 Start Date After Update 	:"		+	actStartDate1);
	System.out.println("Expected Start Date After Update 	:"		+	expStartDate1);
	
	
	
	
	ArrayList<String> endDateArray1=new ArrayList<String>();
	for(WebElement e:sellerPricebook_EndDateColList)
	{
		endDateArray1.add(e.getText());
		
	}
	String actEndDate1=endDateArray1.toString();
	String expEndDate1="[, "+next4DaysDate+", "+next4DaysDate+", "+next4DaysDate+", "+next4DaysDate+", "+next4DaysDate+", "+next4DaysDate+", "+next4DaysDate+", , , , , , , ]";
	
	System.out.println("Actual 	 Start Date After Update 	:"		+	actEndDate1);
	System.out.println("Expected Start Date After Update 	:"		+	expEndDate1);
	
	if(actStartDate.equalsIgnoreCase(expStartDate) && actEndDate.equalsIgnoreCase(expEndDate) && 
			actStartDate1.equalsIgnoreCase(expStartDate1) && actEndDate1.equalsIgnoreCase(expEndDate1) )
	{
		return true;
	}
	else
	{
		return false;
	}
	
}

public static boolean checkRestorePricebookwithunits() throws InterruptedException, IOException, AWTException
{
	
	String actMessage=BaseEngine.restoreCompany("Pricebook with Units","Pricebook Validation and Authorization");
	String expMessage="Restore company code : 0G0";
	
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
	
public boolean checkpricebookvaluesinsalesorderforBOXUnitofGroup1Items() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Inventorybut));
	ClickUsingJs(Inventorybut);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InventoryTransactionsbut));
	ClickUsingJs(InventoryTransactionsbut);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesbut));
	ClickUsingJs(salesbut);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Salesorder));
	ClickUsingJs(Salesorder);
	Thread.sleep(4500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Salesorder_newbut));
	ClickUsingJs(Salesorder_newbut);
	Thread.sleep(3000);
	
	checkValidationMessage("Screen opened");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CustomerAcc));
	CustomerAcc.click();
	Thread.sleep(1000);
	CustomerAcc.sendKeys("Customer A");
	Thread.sleep(1000);
	CustomerAcc.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_1stcolumn));
	ClickUsingJs(select1strow_1stcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepartmentTxt));
	DepartmentTxt.click();
	Thread.sleep(1000);
	DepartmentTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	DepartmentTxt.sendKeys("INDIA");
	Thread.sleep(1000);
	DepartmentTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(WarehouseTxt));
	WarehouseTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	WarehouseTxt.sendKeys("HYDERABAD");
	Thread.sleep(1000);
	WarehouseTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemTxt));
	ItemTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	ItemTxt.sendKeys("Test-A");
	Thread.sleep(1000);
	ItemTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UnitTxt));
	UnitTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	UnitTxt.sendKeys("Box");
	Thread.sleep(1000);
	UnitTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_8thcolumn));
	ClickUsingJs(select1strow_8thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(QuantityTxt));
	QuantityTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	QuantityTxt.sendKeys("1");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_11thcolumn));
	ClickUsingJs(select1strow_11thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RateTxt));
	RateTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GrossTxt));
	GrossTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Val1Txt));
	Val1Txt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DiscountTxt));
	DiscountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(1500);
	
	int count1 = Row1List.size();
	System.out.println("Count1 : " + count1);

	ArrayList<String> Row1ListArray = new ArrayList<String>();

	for (int i = 1; i < count1; i++) {
		String data1 = Row1List.get(i).getText();

		Row1ListArray.add(data1);
	}

	String actRowlList = Row1ListArray.toString();

	String expRowlList = "[INDIA, HYDERABAD, Test-A, Box, 0.00, 0.00, 0, 1.00, , , 40.00, 40.00, 60.00, 0.00]";

	System.out.println("actentryPagerowlList : " + actRowlList);
	System.out.println("ExpentryPagerowlList : " + expRowlList);
	
	Thread.sleep(5000);

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndrow_1stcolumn));
	ClickUsingJs(select2ndrow_1stcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepartmentTxt));
	DepartmentTxt.click();
	Thread.sleep(1000);
	DepartmentTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	DepartmentTxt.sendKeys("INDIA");
	Thread.sleep(1000);
	DepartmentTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(WarehouseTxt));
	WarehouseTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	WarehouseTxt.sendKeys("HYDERABAD");
	Thread.sleep(1000);
	WarehouseTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemTxt));
	ItemTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	ItemTxt.sendKeys("Test-B");
	Thread.sleep(1000);
	ItemTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UnitTxt));
	UnitTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	UnitTxt.sendKeys("Box");
	Thread.sleep(1000);
	UnitTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndrow_8thcolumn));
	ClickUsingJs(select2ndrow_8thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(QuantityTxt));
	QuantityTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	QuantityTxt.sendKeys("1");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndrow_11thcolumn));
	ClickUsingJs(select2ndrow_11thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RateTxt));
	RateTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GrossTxt));
	GrossTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Val1Txt));
	Val1Txt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DiscountTxt));
	DiscountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(1500);
	
	int count2 = Row2List.size();
	System.out.println("Count2 : " + count2);

	ArrayList<String> Row2ListArray = new ArrayList<String>();

	for (int i = 1; i < count2; i++) {
		String data2 = Row2List.get(i).getText();

		Row2ListArray.add(data2);
	}

	String actRow2List = Row2ListArray.toString();

	String expRow2List = "[INDIA, HYDERABAD, Test-B, Box, 0.00, 0.00, 0, 1.00, , , 40.00, 40.00, 60.00, 0.00]";

	System.out.println("actentryPagerow2List : " + actRow2List);
	System.out.println("ExpentryPagerow2List : " + expRow2List);
	
	Thread.sleep(5000);
	
	ClickUsingJs(SaveButton);
	
	String expmsg= "This Transaction will make the Stock Negative";
	String actmsg= checkValidationMessage(expmsg);

	if(actRowlList.equalsIgnoreCase(expRowlList) && actRow2List.equalsIgnoreCase(expRow2List)
		&&	actmsg.equalsIgnoreCase(expmsg))
	{
		return true;
	}
	
	else
	{
		return false;
	}
	

}
	
public boolean checkpricebookvaluesinsalesorderforPENUnitofGroup1Items() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Salesorder_Previousbut));
	ClickUsingJs(Salesorder_Previousbut);
	Thread.sleep(1500);
	
	checkValidationMessage("Voucher loaded successfully");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_4thcolumn));
	ClickUsingJs(select1strow_4thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UnitTxt));
	UnitTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	UnitTxt.sendKeys("PEN");
	Thread.sleep(1000);
	UnitTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_8thcolumn));
	ClickUsingJs(select1strow_8thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(QuantityTxt));
	QuantityTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	QuantityTxt.sendKeys("1");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_11thcolumn));
	ClickUsingJs(select1strow_11thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RateTxt));
	RateTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GrossTxt));
	GrossTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Val1Txt));
	Val1Txt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DiscountTxt));
	DiscountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(1500);
	
	int count1 = Row1List.size();
	System.out.println("Count1 : " + count1);

	ArrayList<String> Row1ListArray = new ArrayList<String>();

	for (int i = 1; i < count1; i++) {
		String data1 = Row1List.get(i).getText();

		Row1ListArray.add(data1);
	}

	String actRowlList = Row1ListArray.toString();

	String expRowlList = "[INDIA, HYDERABAD, Test-A, PEN, 0.00, 0.00, 0, 1.00, , , 1.60, 1.60, 2.40, 0.00]";

	System.out.println("actentryPagerowlList : " + actRowlList);
	System.out.println("ExpentryPagerowlList : " + expRowlList);
	
	Thread.sleep(5000);

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndrow_4thcolumn));
	ClickUsingJs(select2ndrow_4thcolumn);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UnitTxt));
	UnitTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	UnitTxt.sendKeys("PEN");
	Thread.sleep(1000);
	UnitTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndrow_8thcolumn));
	ClickUsingJs(select2ndrow_8thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(QuantityTxt));
	QuantityTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	QuantityTxt.sendKeys("1");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndrow_11thcolumn));
	ClickUsingJs(select2ndrow_11thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RateTxt));
	RateTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GrossTxt));
	GrossTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Val1Txt));
	Val1Txt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DiscountTxt));
	DiscountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(1500);
	
	int count2 = Row2List.size();
	System.out.println("Count2 : " + count2);

	ArrayList<String> Row2ListArray = new ArrayList<String>();

	for (int i = 1; i < count2; i++) {
		String data2 = Row2List.get(i).getText();

		Row2ListArray.add(data2);
	}

	String actRow2List = Row2ListArray.toString();

	String expRow2List = "[INDIA, HYDERABAD, Test-B, PEN, 0.00, 0.00, 0, 1.00, , , 1.60, 1.60, 2.40, 0.00]";

	System.out.println("actentryPagerow2List : " + actRow2List);
	System.out.println("ExpentryPagerow2List : " + expRow2List);
	
	Thread.sleep(5000);
	
	ClickUsingJs(SaveButton);
	
	String expmsg= "This Transaction will make the Stock Negative";
	String actmsg= checkValidationMessage(expmsg);

	if(actRowlList.equalsIgnoreCase(expRowlList) && actRow2List.equalsIgnoreCase(expRow2List)
		&&	actmsg.equalsIgnoreCase(expmsg))
	{
		return true;
	}
	
	else
	{
		return false;
	}
	

}
	
public boolean checkpricebookvaluesinsalesorderforSTPUnitofGroup1Items() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Salesorder_Previousbut));
	ClickUsingJs(Salesorder_Previousbut);
	
	checkValidationMessage("Voucher loaded successfully");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_4thcolumn));
	ClickUsingJs(select1strow_4thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UnitTxt));
	UnitTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	UnitTxt.sendKeys("STP");
	Thread.sleep(1000);
	UnitTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_8thcolumn));
	ClickUsingJs(select1strow_8thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(QuantityTxt));
	QuantityTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	QuantityTxt.sendKeys("1");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_11thcolumn));
	ClickUsingJs(select1strow_11thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RateTxt));
	RateTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GrossTxt));
	GrossTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Val1Txt));
	Val1Txt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DiscountTxt));
	DiscountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(1500);
	
	int count1 = Row1List.size();
	System.out.println("Count1 : " + count1);

	ArrayList<String> Row1ListArray = new ArrayList<String>();

	for (int i = 1; i < count1; i++) {
		String data1 = Row1List.get(i).getText();

		Row1ListArray.add(data1);
	}

	String actRowlList = Row1ListArray.toString();

	String expRowlList = "[INDIA, HYDERABAD, Test-A, STP, 0.00, 0.00, 0, 1.00, , , 2.86, 2.86, 4.29, 0.00]";

	System.out.println("actentryPagerowlList : " + actRowlList);
	System.out.println("ExpentryPagerowlList : " + expRowlList);
	
	Thread.sleep(5000);

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndrow_4thcolumn));
	ClickUsingJs(select2ndrow_4thcolumn);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UnitTxt));
	UnitTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	UnitTxt.sendKeys("STP");
	Thread.sleep(1000);
	UnitTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndrow_8thcolumn));
	ClickUsingJs(select2ndrow_8thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(QuantityTxt));
	QuantityTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	QuantityTxt.sendKeys("1");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndrow_11thcolumn));
	ClickUsingJs(select2ndrow_11thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RateTxt));
	RateTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GrossTxt));
	GrossTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Val1Txt));
	Val1Txt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DiscountTxt));
	DiscountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(1500);
	
	int count2 = Row2List.size();
	System.out.println("Count2 : " + count2);

	ArrayList<String> Row2ListArray = new ArrayList<String>();

	for (int i = 1; i < count2; i++) {
		String data2 = Row2List.get(i).getText();

		Row2ListArray.add(data2);
	}

	String actRow2List = Row2ListArray.toString();

	String expRow2List = "[INDIA, HYDERABAD, Test-B, STP, 0.00, 0.00, 0, 1.00, , , 2.86, 2.86, 4.29, 0.00]";

	System.out.println("actentryPagerow2List : " + actRow2List);
	System.out.println("ExpentryPagerow2List : " + expRow2List);
	
	Thread.sleep(5000);
	
	ClickUsingJs(SaveButton);
	
	String expmsg= "This Transaction will make the Stock Negative";
	String actmsg= checkValidationMessage(expmsg);

	if(actRowlList.equalsIgnoreCase(expRowlList) && actRow2List.equalsIgnoreCase(expRow2List)
		&&	actmsg.equalsIgnoreCase(expmsg))
	{
		return true;
	}
	
	else
	{
		return false;
	}
	

}

public boolean checkDifferentDepartmentwithanyitemshouldnotshowPricebookvalues() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
	
	
	getDriver().navigate().refresh();
	Thread.sleep(1500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Inventorybut));
	ClickUsingJs(Inventorybut);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InventoryTransactionsbut));
	ClickUsingJs(InventoryTransactionsbut);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesbut));
	ClickUsingJs(salesbut);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Salesorder));
	ClickUsingJs(Salesorder);
	Thread.sleep(4500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Salesorder_newbut));
	ClickUsingJs(Salesorder_newbut);
	Thread.sleep(3000);
	
	checkValidationMessage("Screen opened");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CustomerAcc));
	CustomerAcc.click();
	Thread.sleep(1000);
	CustomerAcc.sendKeys("Customer A");
	Thread.sleep(1000);
	CustomerAcc.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_1stcolumn));
	ClickUsingJs(select1strow_1stcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepartmentTxt));
	DepartmentTxt.click();
	Thread.sleep(1000);
	DepartmentTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	DepartmentTxt.sendKeys("DUBAI");
	Thread.sleep(1000);
	DepartmentTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(WarehouseTxt));
	WarehouseTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	WarehouseTxt.sendKeys("HYDERABAD");
	Thread.sleep(1000);
	WarehouseTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemTxt));
	ItemTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	ItemTxt.sendKeys("Test-A");
	Thread.sleep(1000);
	ItemTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UnitTxt));
	UnitTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	UnitTxt.sendKeys("Box");
	Thread.sleep(1000);
	UnitTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_8thcolumn));
	ClickUsingJs(select1strow_8thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(QuantityTxt));
	QuantityTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	QuantityTxt.sendKeys("1");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_11thcolumn));
	ClickUsingJs(select1strow_11thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RateTxt));
	RateTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GrossTxt));
	GrossTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Val1Txt));
	Val1Txt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DiscountTxt));
	DiscountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(1500);
	
	int count1 = Row1List.size();
	System.out.println("Count1 : " + count1);

	ArrayList<String> Row1ListArray = new ArrayList<String>();

	for (int i = 1; i < count1; i++) {
		String data1 = Row1List.get(i).getText();

		Row1ListArray.add(data1);
	}

	String actRowlList = Row1ListArray.toString();

	String expRowlList = "[DUBAI, HYDERABAD, Test-A, Box, 0.00, 0.00, 0, 1.00, , , 0.00, 0.00, 0.00, 0.00]";

	System.out.println("actentryPagerowlList : " + actRowlList);
	System.out.println("ExpentryPagerowlList : " + expRowlList);
	

	if(actRowlList.equalsIgnoreCase(expRowlList))
	{
		return true;
	}
	
	else
	{
		return false;
	}
	

}
	
	
public boolean checkDifferentCustomerAccountwithanyitemshouldnotshowPricebookvalues() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
	
	
	getDriver().navigate().refresh();
	Thread.sleep(1500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Inventorybut));
	ClickUsingJs(Inventorybut);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InventoryTransactionsbut));
	ClickUsingJs(InventoryTransactionsbut);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesbut));
	ClickUsingJs(salesbut);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Salesorder));
	ClickUsingJs(Salesorder);
	Thread.sleep(4500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Salesorder_newbut));
	ClickUsingJs(Salesorder_newbut);
	Thread.sleep(3000);
	
	checkValidationMessage("Screen opened");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CustomerAcc));
	CustomerAcc.click();
	Thread.sleep(1000);
	CustomerAcc.sendKeys("Customer B");
	Thread.sleep(1000);
	CustomerAcc.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_1stcolumn));
	ClickUsingJs(select1strow_1stcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepartmentTxt));
	DepartmentTxt.click();
	Thread.sleep(1000);
	DepartmentTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	DepartmentTxt.sendKeys("INDIA");
	Thread.sleep(1000);
	DepartmentTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(WarehouseTxt));
	WarehouseTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	WarehouseTxt.sendKeys("HYDERABAD");
	Thread.sleep(1000);
	WarehouseTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemTxt));
	ItemTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	ItemTxt.sendKeys("Test-A");
	Thread.sleep(1000);
	ItemTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UnitTxt));
	UnitTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	UnitTxt.sendKeys("Box");
	Thread.sleep(1000);
	UnitTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_8thcolumn));
	ClickUsingJs(select1strow_8thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(QuantityTxt));
	QuantityTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	QuantityTxt.sendKeys("1");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_11thcolumn));
	ClickUsingJs(select1strow_11thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RateTxt));
	RateTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GrossTxt));
	GrossTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Val1Txt));
	Val1Txt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DiscountTxt));
	DiscountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(1500);
	
	int count1 = Row1List.size();
	System.out.println("Count1 : " + count1);

	ArrayList<String> Row1ListArray = new ArrayList<String>();

	for (int i = 1; i < count1; i++) {
		String data1 = Row1List.get(i).getText();

		Row1ListArray.add(data1);
	}

	String actRowlList = Row1ListArray.toString();

	String expRowlList = "[INDIA, HYDERABAD, Test-A, Box, 0.00, 0.00, 0, 1.00, , , 0.00, 0.00, 0.00, 0.00]";

	System.out.println("actentryPagerowlList : " + actRowlList);
	System.out.println("ExpentryPagerowlList : " + expRowlList);
	

	if(actRowlList.equalsIgnoreCase(expRowlList))
	{
		return true;
	}
	
	else
	{
		return false;
	}
	

}
public boolean checkpricebookvaluesinsalesorderforDifferentUnitofTestCItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
	
	getDriver().navigate().refresh();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Inventorybut));
	ClickUsingJs(Inventorybut);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InventoryTransactionsbut));
	ClickUsingJs(InventoryTransactionsbut);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesbut));
	ClickUsingJs(salesbut);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Salesorder));
	ClickUsingJs(Salesorder);
	Thread.sleep(4500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Salesorder_newbut));
	ClickUsingJs(Salesorder_newbut);
	Thread.sleep(3000);
	
	checkValidationMessage("Screen opened");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CustomerAcc));
	CustomerAcc.click();
	Thread.sleep(1000);
	CustomerAcc.sendKeys("Customer A");
	Thread.sleep(1000);
	CustomerAcc.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_1stcolumn));
	ClickUsingJs(select1strow_1stcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepartmentTxt));
	DepartmentTxt.click();
	Thread.sleep(1000);
	DepartmentTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	DepartmentTxt.sendKeys("INDIA");
	Thread.sleep(1000);
	DepartmentTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(WarehouseTxt));
	WarehouseTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	WarehouseTxt.sendKeys("HYDERABAD");
	Thread.sleep(1000);
	WarehouseTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemTxt));
	ItemTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	ItemTxt.sendKeys("Test-C");
	Thread.sleep(1000);
	ItemTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UnitTxt));
	UnitTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	UnitTxt.sendKeys("STP");
	Thread.sleep(1000);
	UnitTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_8thcolumn));
	ClickUsingJs(select1strow_8thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(QuantityTxt));
	QuantityTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	QuantityTxt.sendKeys("1");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_11thcolumn));
	ClickUsingJs(select1strow_11thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RateTxt));
	RateTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GrossTxt));
	GrossTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Val1Txt));
	Val1Txt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DiscountTxt));
	DiscountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(1500);
	
	int count1 = Row1List.size();
	System.out.println("Count1 : " + count1);

	ArrayList<String> Row1ListArray = new ArrayList<String>();

	for (int i = 1; i < count1; i++) {
		String data1 = Row1List.get(i).getText();

		Row1ListArray.add(data1);
	}

	String actRow1List = Row1ListArray.toString();

	String expRow1List = "[INDIA, HYDERABAD, Test-C, STP, 0.00, 0.00, 0, 1.00, , , 100.00, 100.00, 200.00, 0.00]";

	System.out.println("actentryPagerowlList : " + actRow1List);
	System.out.println("ExpentryPagerowlList : " + expRow1List);
	
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_4thcolumn));
	ClickUsingJs(select1strow_4thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UnitTxt));
	UnitTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	UnitTxt.sendKeys("TAB");
	Thread.sleep(1000);
	UnitTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_8thcolumn));
	ClickUsingJs(select1strow_8thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(QuantityTxt));
	QuantityTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	QuantityTxt.sendKeys("1");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_11thcolumn));
	ClickUsingJs(select1strow_11thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RateTxt));
	RateTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GrossTxt));
	GrossTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Val1Txt));
	Val1Txt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DiscountTxt));
	DiscountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(1500);
	
	int count2 = Row1List.size();
	System.out.println("Count2 : " + count2);

	ArrayList<String> Row2ListArray = new ArrayList<String>();

	for (int i = 1; i < count2; i++) {
		String data2 = Row1List.get(i).getText();

		Row2ListArray.add(data2);
	}

	String actRow2List = Row2ListArray.toString();

	String expRow2List = "[INDIA, HYDERABAD, Test-C, TAB, 0.00, 0.00, 0, 1.00, , , 6.67, 6.67, 13.33, 0.00]";

	System.out.println("actentryPagerow2List : " + actRow2List);
	System.out.println("ExpentryPagerow2List : " + expRow2List);
	
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_4thcolumn));
	ClickUsingJs(select1strow_4thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UnitTxt));
	UnitTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	UnitTxt.sendKeys("Box");
	Thread.sleep(1000);
	UnitTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_8thcolumn));
	ClickUsingJs(select1strow_8thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(QuantityTxt));
	QuantityTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	QuantityTxt.sendKeys("1");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_11thcolumn));
	ClickUsingJs(select1strow_11thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RateTxt));
	RateTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GrossTxt));
	GrossTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Val1Txt));
	Val1Txt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DiscountTxt));
	DiscountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(1500);
	
	int count3 = Row1List.size();
	System.out.println("Count3 : " + count3);

	ArrayList<String> Row3ListArray = new ArrayList<String>();

	for (int i = 1; i < count3; i++) {
		String data3 = Row1List.get(i).getText();

		Row3ListArray.add(data3);
	}

	String actRow3List = Row3ListArray.toString();

	String expRow3List = "[INDIA, HYDERABAD, Test-C, Box, 0.00, 0.00, 0, 1.00, , , 1,400.00, 1,400.00, 2800.00, 0.00]";

	System.out.println("actentryPagerow3List : " + actRow3List);
	System.out.println("ExpentryPagerow3List : " + expRow3List);
	
	Thread.sleep(3000);
	
	if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List)
			&& actRow3List.equalsIgnoreCase(expRow3List))
	{
		return true;
	}
	
	else
	{
		return false;
	}
	
	
}

public boolean checkpricebookvaluesinsalesorderforDifferentUnitofTestDItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
	
	getDriver().navigate().refresh();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Inventorybut));
	ClickUsingJs(Inventorybut);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InventoryTransactionsbut));
	ClickUsingJs(InventoryTransactionsbut);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesbut));
	ClickUsingJs(salesbut);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Salesorder));
	ClickUsingJs(Salesorder);
	Thread.sleep(4500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Salesorder_newbut));
	ClickUsingJs(Salesorder_newbut);
	Thread.sleep(3000);
	
	checkValidationMessage("Screen opened");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CustomerAcc));
	CustomerAcc.click();
	Thread.sleep(1000);
	CustomerAcc.sendKeys("Customer A");
	Thread.sleep(1000);
	CustomerAcc.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_1stcolumn));
	ClickUsingJs(select1strow_1stcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepartmentTxt));
	DepartmentTxt.click();
	Thread.sleep(1000);
	DepartmentTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	DepartmentTxt.sendKeys("INDIA");
	Thread.sleep(1000);
	DepartmentTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(WarehouseTxt));
	WarehouseTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	WarehouseTxt.sendKeys("HYDERABAD");
	Thread.sleep(1000);
	WarehouseTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemTxt));
	ItemTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	ItemTxt.sendKeys("Test-D");
	Thread.sleep(1000);
	ItemTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UnitTxt));
	UnitTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	UnitTxt.sendKeys("Pcs");
	Thread.sleep(1000);
	UnitTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_8thcolumn));
	ClickUsingJs(select1strow_8thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(QuantityTxt));
	QuantityTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	QuantityTxt.sendKeys("1");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_11thcolumn));
	ClickUsingJs(select1strow_11thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RateTxt));
	RateTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GrossTxt));
	GrossTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Val1Txt));
	Val1Txt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DiscountTxt));
	DiscountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(1500);
	
	int count1 = Row1List.size();
	System.out.println("Count1 : " + count1);

	ArrayList<String> Row1ListArray = new ArrayList<String>();

	for (int i = 1; i < count1; i++) {
		String data1 = Row1List.get(i).getText();

		Row1ListArray.add(data1);
	}

	String actRow1List = Row1ListArray.toString();

	String expRow1List = "[INDIA, HYDERABAD, Test-D, Pcs, 0.00, 0.00, 0, 1.00, , , 25.00, 25.00, 35.00, 0.00]";

	System.out.println("actentryPagerowlList : " + actRow1List);
	System.out.println("ExpentryPagerowlList : " + expRow1List);
	
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_4thcolumn));
	ClickUsingJs(select1strow_4thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UnitTxt));
	UnitTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	UnitTxt.sendKeys("Dozs");
	Thread.sleep(1000);
	UnitTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_8thcolumn));
	ClickUsingJs(select1strow_8thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(QuantityTxt));
	QuantityTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	QuantityTxt.sendKeys("1");
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1strow_11thcolumn));
	ClickUsingJs(select1strow_11thcolumn);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RateTxt));
	RateTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GrossTxt));
	GrossTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Val1Txt));
	Val1Txt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DiscountTxt));
	DiscountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(1500);
	
	int count2 = Row1List.size();
	System.out.println("Count2 : " + count2);

	ArrayList<String> Row2ListArray = new ArrayList<String>();

	for (int i = 1; i < count2; i++) {
		String data2 = Row1List.get(i).getText();

		Row2ListArray.add(data2);
	}

	String actRow2List = Row2ListArray.toString();

	String expRow2List = "[INDIA, HYDERABAD, Test-D, Dozs, 0.00, 0.00, 0, 1.00, , , 300.00, 300.00, 420.00, 0.00]";

	System.out.println("actentryPagerow2List : " + actRow2List);
	System.out.println("ExpentryPagerow2List : " + expRow2List);
	
	Thread.sleep(3000);
	
	
	if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List))
	{
		return true;
	}
	
	else
	{
		return false;
	}
	
	
}



	
	
	
	

		


	
	
	
	

public  PricebookAuthorizationPage(WebDriver driver) 
{
	PageFactory.initElements(driver, this);
}












}




