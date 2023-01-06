package com.focus.Pages;

import java.awt.AWTException;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Year;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.NoSuchElementException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
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

public class SanityFixedAssetsPage extends BaseEngine
{

	
	@FindBy(xpath="//*[@id='16']/div/span")
	private static WebElement  SeingsmenuBtn;
	
	@FindBy(xpath="//span[contains(text(),'Configure Transactions')]")
	private static WebElement  ConfigureTransactionBtn;
	
	@FindBy(xpath="//label[contains(text(),'Preferences')]")
	private static WebElement  PreferencesBtn;
	
	@FindBy(xpath="//div[@id='docCustomization']")
	private static WebElement  DocumentCustomizationBtn;
	
	@FindBy(xpath="//a[@id='0']")
	private static WebElement  TagsBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//a[@id='1']")
	private static WebElement  AccountsBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//ul//li//a[@id='2']")
	private static WebElement  BudgetBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//ul//li//a[@id='3']")
	private static WebElement  ARAPBtn;
	
	@FindBy(xpath="//a[@id='4']")
	private static WebElement  MiselleanousBtn;
	
	@FindBy(xpath="//a[@id='5']")
	private static WebElement  PDCBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//a[@id='6']")
	private static WebElement  InventoryBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//ul//li//a[@id='7']")
	private static WebElement  BatchBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//ul//li//a[@id='30']")
	private static WebElement  RMABtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//ul//li//a[@id='8']")
	private static WebElement  BinsBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//ul//li//a[@id='9']")
	private static WebElement  HirePurchaseBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//ul//li//a[@id='10']")
	private static WebElement  QuotationAnalysisBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//ul//li//a[@id='34']")
	private static WebElement  RFIDBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//ul//li//a[@id='29']")
	private static WebElement  AutoIndentBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//ul//li//a[@id='14']")
	private static WebElement  ABCAnalysisBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//a[@id='11']")
	private static WebElement  ReportBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//a[@id='12']")
	private static WebElement  InternetBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//ul//li//a[@id='13']")
	private static WebElement  MailSettingsBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//a[@id='16']")
	private static WebElement  MastersBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//a[@id='17']")
	private static WebElement  ExternalModulesBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//a[@id='18']")
	private static WebElement  InfoBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//a[@id='19']")
	private static WebElement  ProductionBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//a[@id='20']")
	private static WebElement  MrpBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//a[@id='21']")
	private static WebElement  QualityControlBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//a[@id='22']")
	private static WebElement  FixedAssetsBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//a[@id='25']")
	private static WebElement  VatBtn;
	
	@FindBy(xpath="//a[@id='28']")
	private static WebElement  PosBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//a[@id='35']")
	private static WebElement  WareHouseManagementBtn;
	
	@FindBy(xpath="//a[@id='36']")
	private static WebElement  LetterForCreditBtn;
	
	@FindBy(xpath="//ul[@id='preferenceUL']//a[@id='37']")
	private static WebElement  PronghornBnt;
	
	@FindBy(xpath="//input[@id='txtVoucherWizard']")
	private static WebElement  SearchTxt;
	
	@FindBy(xpath="//span[@id='updateButton']")
	private static WebElement  UpdateBtn;
	
	@FindBy(xpath="//i[@class='icon-close icon-font6']")
	private static WebElement  CloseBtn;
	
	@FindBy(xpath="//select[@id='Accounts']")
	private static WebElement  AccountingDropdown;
	
	@FindBy(xpath="//select[@id='Inventory']")
	private static WebElement  InventoryDropdown;
	
	@FindBy(xpath="//select[@id='Payroll']")
	private static WebElement  PayRollCostCenterDropdown;
	
	@FindBy(xpath="//select[@id='SellingRate']")
	private static WebElement  SellingsRateDropdown;
	
	@FindBy(xpath="//select[@id='BuyingRates']")
	private static WebElement  BuyingRatesdropdown;
	
	@FindBy(xpath="//select[@id='PayrolDept']")
	private static WebElement  PayRollDepartmentDropdown;
	
	@FindBy(xpath="//select[@id='PayrollSite']")
	private static WebElement  PayRollSiteDropdown;
	
	@FindBy(xpath="//select[@id='DeptwiseAppropriation']")
	private static WebElement  FinancialAppropriationByDropdown;
	
	@FindBy(xpath="//select[@id='VAT']")
	private static WebElement  VatDropdown;
	
	@FindBy(xpath="//select[@id='Schemes']")
	private static WebElement  SchemesDropdown;
	
	@FindBy(xpath="//select[@id='Inventoryallocationtag']")
	private static WebElement  InventoryAloocationTagDropdown;


	@FindBy(xpath="//div[@class='theme_color font-6']")
	public static WebElement errorMessage;
	 
	@FindBy(xpath="//span[@class='icon-reject2 theme_color']")
	public static WebElement errorMessageCloseBtn;
	



	@FindBy(xpath = "//select[@id='FixedAssetsLocation']")
	private static WebElement FixedAssets_FixedAssetsLocationDrrpdwn;
	
	@FindBy(xpath = "//input[@id='IsTag1']")
	private static WebElement FixedAssets_FixesAssetsTagChkbox1;
	
	@FindBy(xpath = "//input[@id='IsTag2']")
	private static WebElement FixedAssets_FixesAssetsTagChkbox2;
	
	@FindBy(xpath = "//input[@id='CreationOfFixedAssetInMandatoryFromPurchase']")
	private static WebElement FixedAssets_CreationOfFixedAssetInMandatoryFromPurchaseChkbox;
	
	@FindBy(xpath = "//input[@id='AlwaysDefineDepreciationRateAsPercentage']")
	private static WebElement FixedAssets_AlwaysDefineDepreciationRateAsPercentageChkbox;
	
	@FindBy(xpath = "//input[@id='IncludeAssetsInPurchaseAccountSelection']")
	private static WebElement FixedAssets_IncludeAssetsInPurchaseAccountSelectionChkbox;
	
	@FindBy(xpath = "//input[@id='PostDepreciation_0']")
	private static WebElement FixedAssetsPostDepreciationDefaultRadioBtn;
	
	@FindBy(xpath = "//input[@id='PostDepreciation_1']")
	private static WebElement PostDepreciationByMOnthRadioBtn;
	
	@FindBy(xpath = "//input[@id='PostDepreciation_2']")
	private static WebElement postFixedMonthlyDepresionRadioBtn;
	
	@FindBy(xpath = "//input[@id='PostDepreciation_3']")
	private static WebElement PostDepreciationByYearRadioBtn;
	
	@FindBy(xpath = "//input[@id='ConsolidatedPosting']")
	private static WebElement FixedAssetsConsolidatingPostingChkbox;
	
	@FindBy(xpath = "//input[@id='PostentryduringTransfer']")
	private static WebElement PostentryduringTransferChkbox;
	
	@FindBy(xpath = "//input[@id='PostentryduringDisposal']")
	private static WebElement PostentryduringDisposalChkbox;
	
	@FindBy(xpath = "//input[@id='Additionextends']")
	private static WebElement FixedAssets_AdditionExtendsChkbox;
	
	@FindBy(xpath = "//input[@id='AccelerateDepreciation']")
	private static WebElement AccelerateDepreciationChkbox;
	
	@FindBy(xpath = "//input[@id='AutoCalculatesRateOfDepreciation']")
	private static WebElement AutoCalculatesRateOfDepreciationChkbox;
	
	@FindBy(xpath = "//input[@id='CalculateDepriciation']")
	private static WebElement calculateDepresionBasedOnAccountingYearChkbox;
	
  
	 
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

	
	
	

	
	private static boolean methodReturnStatus;
	private static String xlSheetName = "FixedAssets";
	
				
		
	@FindBy(xpath="//span[contains(text(),'Configure Transactions')]")
	private static WebElement  settingsConfigureTransactions;   
	
	@FindBy(xpath="//*[@id='68']/span")
	private static WebElement  settingsTransactionAuthorization;
	
	@FindBy(xpath="//*[@id='235']/span")
	private static WebElement  settingsPricebookAuthorization;
	
	@FindBy(xpath="//*[@id='19']/span")
	private static WebElement  settingsConfigureMasters;
	
	@FindBy(xpath="//*[@id='67']/span")
	private static WebElement  settingsMasterAuthorization;

	@FindBy(xpath="//*[@id='69']/span")
	private static WebElement  settingsCreditLimitAuthorization;
	
	@FindBy(xpath="//*[@id='74']/span")
	private static WebElement  settingsDesignWorkflow;
	
	@FindBy(xpath="//*[@id='105']/span")
	private static WebElement  settingsSettingsWizard;
	
	@FindBy(xpath="//*[@id='232']/span")
	private static WebElement  settingsBatchCodeGeneration;
	
	@FindBy(xpath="//*[@id='34']/span")
	private static WebElement  settingsMapCreditLimit;
	
	@FindBy(xpath="//*[@id='257']/span")
	private static WebElement  settingsVATTaxCode;
	
	@FindBy(xpath="//*[@id='258']/span")
	private static WebElement  settingsCustomerPortalAdmin;
	
	@FindBy(xpath="//*[@id='3308']/span")
	private static WebElement  settingsBudgetAuthorization;
	
	@FindBy(xpath="//*[@id='3320']/span")
	private static WebElement  settingsPaymentGatewayIntegration;	
	
	@FindBy(xpath="//*[@id='3327']/span")
	private static WebElement  settingsFixedAssetsAuthorization;	

	@FindBy(xpath="//*[@id='navigation_menu']/li[8]/ul/li")
	private static List<WebElement>  settingMenuList;
	
	@FindBy(xpath = "//div[@id='idGlobalError']")
	public static WebElement validationConfirmationMessage;

	@FindBy(xpath = "//div[@class='theme_color font-6']")
	public static WebElement validationUpdatingConfirmationMessage;

	@FindBy(xpath = "//*[@id='idGlobalError']/div/table/tbody/tr/td[3]/span")
	public static WebElement closeValidationConfirmationMessage;
	
	@FindBy(xpath = "//span[@class='icon-reject2 theme_color']")
	public static WebElement newErrorMessageCloseBtn;
	
	@FindBy(xpath = "//div[contains(text(),'Settings')]")
	private static WebElement settingsmenuBtn;
	
	// Accounts Creation Page

		// Header tab Fields
		@FindBy(xpath = "//*[@id='divBtnGroup0']/div/a")
		public static WebElement headerGeneralTab;
		
		@FindBy(xpath = "//*[@id='divBtnGroup1']/div/a")
		public static WebElement headerAccountSettingTab;

		@FindBy(xpath = "//*[@id='divBtnGroup2']/div/a")
		public static WebElement headerDetailsTab;

		@FindBy(xpath = "//*[@id='divBtnGroup3']/div/a")
		public static WebElement headerPrintLayoutTab;

		@FindBy(xpath = "//*[@id='divBtnGroup1']/div/a")
		public static WebElement headerUnitsTab;

		@FindBy(xpath = "//*[@id='divBtnGroup2']/div/a")
		public static WebElement headerItemSettingsTab;

		@FindBy(xpath = "//*[@id='divBtnGroup3']/div/a")
		public static WebElement headerClassificationTab;

		@FindBy(xpath = "//*[@id='divBtnGroup4']/div/a")
		public static WebElement headerOtherDetailsTab;

		@FindBy(xpath = "//*[@id='divBtnGroup5']/div/a")
		public static WebElement headerReplenishmentTab;

		@FindBy(xpath = "//*[@id='divBtnGroup6']/div/a")
		public static WebElement headerOutletTab;

		// Header Section Fields
		@FindBy(xpath = "//*[@id='btnMasterSaveClick']")
		public static WebElement saveBtn;

		@FindBy(xpath = "//*[@id='GenerateNewMasterModel']/div/div[1]/div[2]/div/ul/li[2]/button[5]")
		public static WebElement closeBtn;

		
		@FindBy(xpath="//i[@class='icon-close icon-font6']")
		private static WebElement settingCloseIcon;
		  
		
		@FindBy(xpath="//i[@class='icon-close icon-font6']")
		private static WebElement  closeBtnInDocumentCustmization;
		
		// Edit option Fields
		@FindBy(xpath = "//*[@id='GenerateNewMasterModel']/div/div[1]/div[2]/div/ul/li[2]/button[6]")
		public static WebElement editcloseBtn;

		@FindBy(xpath = "//i[@class='icon-font6 icon-paste-clipboard']")
		public static WebElement pasteBtn;

		@FindBy(xpath = "//button[@id='btnMasterClone']")
		public static WebElement editcloneBtn;

		@FindBy(xpath = "//*[@id='btnMasterCopy']")
		public static WebElement copyBtn;

		// New General Elements
		// Name
		@FindBy(xpath = "//input[@id='sName']")
		public static WebElement nameTxt;

		// Code
		@FindBy(xpath = "//input[@id='sCode']")
		public static WebElement codeTxt;

		
		//Assert Type
		@FindBy(xpath = "//select[@id='iAssetType']")
		public static WebElement assertTypeDrpdwn;
		
		//Vendor Drpdwn
		@FindBy(xpath = "//input[@id='iVendor']")
		public static WebElement vendorDrpdwn;
		
		//Item Drpdwn
		@FindBy(xpath = "//input[@id='iProduct']")
		public static WebElement itemDrpdwn;
		
		
		//Purchase Drpdwn
		@FindBy(xpath = "//input[@id='iPurchaseDate']")
		public static WebElement purchaseDateTxt;
		
		//Price Txt
		@FindBy(xpath = "//input[@id='fPrice']")
		public static WebElement priceTxt;
				
		//Quantity Drpdwn
		@FindBy(xpath = "//input[@id='iQuantity']")
		public static WebElement quantityTxt;
		
		//Department Type
		@FindBy(xpath = "//input[@id='iTag0']")
		public static WebElement departmentdrpdwn;
		
		// AccountType
		@FindBy(xpath = "//select[@id='iAccountType']")
		public static WebElement accountTypeDropdown;

		// CreditLimit
		@FindBy(xpath = "//input[@id='fCreditLimit']")
		public static WebElement creditLimitTxt;

		// CreditDays
		@FindBy(xpath = "//input[@id='iCreditDays']")
		public static WebElement creditdaysTxt;

		// ChequeDiscountLimit
		@FindBy(xpath = "//input[@id='fChequeDiscountLimit']")
		public static WebElement chequeDiscountLimitTxt;

		// Rate Of Interest
		@FindBy(xpath = "//input[@id='fRateofinterest']")
		public static WebElement rateofinterestTxt;

		// BankAccount
		@FindBy(xpath = "//input[@id='iBankAc']")
		public static WebElement bankAccountTxt;

		@FindBy(xpath = "//*[@id='iBankAc_input_image']/span")
		public static WebElement bankAccount_ExpansionBtn;

		@FindBy(xpath = "//*[@id='iBankAc_input_settings']/span")
		public static WebElement bankAccount_SettingBtn;

		// Setting Elements

		// DebitCreditProposal
		@FindBy(xpath = "//select[@id='iDebitCreditProposal']")
		public static WebElement debitCreditProposalDropdown;

		// DebitCreditRequired
		@FindBy(xpath = "//select[@id='iDebitCreditRequired']")
		public static WebElement debitCreditRequiredDropdown;

		// ExchangeAdjustmentGainAC
		@FindBy(xpath = "//input[@id='iExchangeAdjustmentGainAC']")
		public static WebElement exchangeAdjustmentGainACTxt;

		@FindBy(xpath = "//*[@id='iExchangeAdjustmentGainAC_input_image']/span")
		public static WebElement exchangeAdjustmentGainAC_ExpansionBtn;

		@FindBy(xpath = "//*[@id='iExchangeAdjustmentGainAC_input_settings']/span")
		public static WebElement exchangeAdjustmentGainAC_SettingBtn;

		// ExchangeAdjustmentLossAC
		@FindBy(xpath = "//input[@id='iExchangeAdjustmentLossAC']")
		public static WebElement exchangeAdjustmentLossACTxt;

		@FindBy(xpath = "//*[@id='iExchangeAdjustmentLossAC_input_image']/span")
		public static WebElement exchangeAdjustmentLossAC_ExpansionBtn;

		@FindBy(xpath = "//*[@id='iExchangeAdjustmentLossAC_input_settings']/span")
		public static WebElement exchangeAdjustmentLossAC_SettingBtn;

		// PrimaryAccount
		@FindBy(xpath = "//input[@id='iPrimaryAccount']")
		public static WebElement primaryAccountTxt;

		@FindBy(xpath = "//*[@id='iPrimaryAccount_input_image']/span")
		public static WebElement primaryAccount_ExpansionBtn;

		@FindBy(xpath = "//*[@id='iPrimaryAccount_input_settings']/span")
		public static WebElement primaryAccount_SettingBtn;

		// DefaultCurrency
		@FindBy(xpath = "//input[@id='optCurrency']")
		public static WebElement defaultCurrencyTxt;

		@FindBy(xpath = "//*[@id='iDefaultCurrency_input_image']/span")
		public static WebElement defaultCurrency_ExpansionBtn;

		@FindBy(xpath = "//*[@id='iDefaultCurrency_input_settings']/span")
		public static WebElement defaultCurrency_SettingBtn;

		// ConsolidationMethod
		@FindBy(xpath = "//select[@id='iConsolidationMethod']")
		public static WebElement consolidationMethodDropdown;

		// PaymentTerms
		@FindBy(xpath = "//input[@id='iPaymentTerms']")
		public static WebElement paymentTermstxt;

		@FindBy(xpath = "//*[@id='iPaymentTerms_input_image']/span")
		public static WebElement paymentTerms_ExpansionBtn;

		@FindBy(xpath = "//*[@id='iPaymentTerms_input_settings']/span")
		public static WebElement paymentTerms_SettingBtn;

		// ReminderTerms
		@FindBy(xpath = "//*[@id='iReminderTerms']")
		public static WebElement reminderTermsTxt;

		@FindBy(xpath = "//*[@id='iReminderTerms_input_image']/span")
		public static WebElement reminderTerms_ExpansionBtn;

		@FindBy(xpath = "//*[@id='iReminderTerms_input_settings']/span")
		public static WebElement reminderTerms_SettingBtn;

		// FinanceChargeTerms
		@FindBy(xpath = "//*[@id='iFinanceChargeTerms']")
		public static WebElement financeChargeTermsTxt;

		@FindBy(xpath = "//*[@id='iFinanceChargeTerms_input_image']/span")
		public static WebElement financeChargeTerms_Expansion;

		@FindBy(xpath = "//*[@id='iFinanceChargeTerms_input_settings']/span")
		public static WebElement financeChargeTerms_SettingBtn;

		// Details Elements

		// Address
		@FindBy(xpath = "//textarea[@id='sAddress']")
		public static WebElement addressTxt;

		// City
		@FindBy(xpath = "//input[@id='iCity']")
		public static WebElement cityTxt;

		@FindBy(xpath = "//*[@id='iCity_input_image']/span")
		public static WebElement city_ExpansionBtn;

		@FindBy(xpath = "//*[@id='iCity_input_settings']/span")
		public static WebElement city_SettingBtn;

		// DeliveryAddress
		@FindBy(xpath = "//textarea[@id='sDeliveryAddress']")
		public static WebElement deliveryAddressTxt;

		// Pin
		@FindBy(xpath = "//input[@id='sPin']")
		public static WebElement pinTxt;

		// City2 DeliveryCity (City As Second TextBox)
		@FindBy(xpath = "//input[@id='iDeliveryCity']")
		public static WebElement city2Text;

		@FindBy(xpath = "//*[@id='iDeliveryCity_input_image']/span")
		public static WebElement city2_ExpansionBtn;

		@FindBy(xpath = "//*[@id='iDeliveryCity_input_settings']/span")
		public static WebElement city2_SettingBtn;

		//// Pin Delivery pin Second Pin TxtBox
		@FindBy(xpath = "//input[@id='sDeliveryPin']")
		public static WebElement pin2Txt;

		// Send Email
		@FindBy(xpath = "//input[@id='bSendEmailtocustomer']")
		public static WebElement sendEmailCheckBox;

		// AllowCustomerPortal
		@FindBy(xpath = "//input[@id='bAllowCustomerPortal']")
		public static WebElement allowCustomerPortalCheckBox;

		// SendEmail
		@FindBy(xpath = "//input[@id='sEMail']")
		public static WebElement sendEmailTxt;

		// Password
		@FindBy(xpath = "//input[@id='sPassword']")
		public static WebElement passwordTxt;

		// TelphoneNumber
		@FindBy(xpath = "//input[@id='sTelNo']")
		public static WebElement telphoneNumberTxt;

		// FaxNo
		@FindBy(xpath = "//input[@id='sFaxNo']")
		public static WebElement faxNoTxt;

		// PortalEmail
		@FindBy(xpath = "//input[@id='sPortalEmail']")
		public static WebElement portalEmailTxt;

		// Print Layout Elements
		// Voucher Type
		@FindBy(xpath = "//label[contains(text(),'Voucher Type')]")
		public static WebElement gridVoucherTypeTxt;

		@FindBy(xpath = "//*[@id='iVoucherType_input_image']/span")
		public static WebElement gridVoucher_ExpansionBtn;

		@FindBy(xpath = "//*[@id='iVoucherType_input_settings']/span")
		public static WebElement gridVoucher_SettingBtn;

		// Print layout
		@FindBy(xpath = "//label[contains(text(),'Print Layout')]")
		public static WebElement gridPrintLayoutTxt;

		@FindBy(xpath = "//*[@id='iPrintLayout_input_image']/span")
		public static WebElement gridPrintLayout_ExpansionBtn;

		@FindBy(xpath = "//*[@id='iPrintLayout_input_settings']/span")
		public static WebElement gridPrintLayout_SettingBtn;

		// Item
		@FindBy(xpath = "//*[@id='spnHeaderText']")
		private static WebElement itemLabel;

		@FindBy(xpath = "//*[@id='btnNew']")
		private static WebElement itemNewBtn;

		// Item General Tab
		@FindBy(xpath = "//*[@id='divBtnGroup0']/div/a")
		private static WebElement ItemNewnewGeneralBtn;

		@FindBy(xpath = "//input[@id='sName']")
		private static WebElement newGeneralName;

		@FindBy(xpath = "//input[@id='sCode']")
		private static WebElement newGeneralCode;

		@FindBy(xpath = "//select[@id='iProductType']")
		private static WebElement newGeneralItemTypeDropDown;

		@FindBy(xpath = "//input[@id='fReorderLevel']")
		private static WebElement newGeneralReorderLevel;

		@FindBy(xpath = "//input[@id='iBinCapacity']")
		private static WebElement newGeneralBinCapacity;

		@FindBy(xpath = "//select[@id='iValuationMethod']")
		private static WebElement newGeneralValuationMethodDropDown;

		@FindBy(xpath = "//input[@id='iCategory']")
		private static WebElement newGeneralCategoryDropDown;

		@FindBy(xpath = "//td[@id='iCategory_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement newGeneralCategorySettingBtn;

		@FindBy(xpath = "//select[@id='iProductMake']")
		private static WebElement newGeneralItemMakeDropDown;

		@FindBy(xpath = "//input[@id='Printer']")
		private static WebElement newGeneralPrint;

		@FindBy(xpath = "//input[@id='iAlternateCategory']")
		private static WebElement newGeneralAlternativeCategoryDropdown;

		@FindBy(xpath = "//td[@id='iAlternateCategory_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement newGeneralAlternativeCategoryDropdownSettingBtn;

		@FindBy(xpath = "//textarea[@id='sDescription']")
		private static WebElement newGeneralDescription;

		@FindBy(xpath = "//input[@id='bPerishableItem']")
		private static WebElement newGeneralPerishableItemCheckBox;

		@FindBy(xpath = "//input[@id='iBin']")
		private static WebElement newGeneralBinDropdown;

		@FindBy(xpath = "//td[@id='iBin_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement newGeneralBinDropdownSettingBtn;

		@FindBy(xpath = "//input[@id='pImage']")
		private static WebElement newGeneralUploadImageOrFileBtn;

		@FindBy(xpath = "//*[@id='btnMasterSaveClick']/i")
		private static WebElement newGeneralSaveBtn;

		@FindBy(xpath = "//*[@id='GenerateNewMasterModel']/div/div[1]/div[2]/div/ul/li[2]/button[5]")
		private static WebElement newGeneralCLoseBtn;

		// Item Units Tab
		@FindBy(xpath = "//*[@id='divBtnGroup1']/div/a")
		private static WebElement ItemNewUnitsBtn;

		@FindBy(xpath = "//input[@id='iDefaultBaseUnit']")
		private static WebElement itemUnitsDefaultBaseUnitDropdown;

		@FindBy(xpath = "//*[@id='iDefaultBaseUnit_table_data_body']/tr/td")
		private static List<WebElement> itemUnitsDefaultBaseUnitDropdownList;
		
		
		@FindBy(xpath = "//td[@id='iDefaultBaseUnit_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement itemUnitsDefaultBaseUnitSetting;

		@FindBy(xpath = "//input[@id='iDefaultSalesUnit']")
		private static WebElement itemUnitsDefaultSalesUnitDropdown;
		
		
		@FindBy(xpath = "//*[@id='iDefaultSalesUnit_table_data_body']/tr/td")
		private static List<WebElement> itemUnitsDefaultSalesUnitDropdownList;

		@FindBy(xpath = "//td[@id='iDefaultSalesUnit_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement itemUnitsDefaultSalesUnitSettingsBtn;

		@FindBy(xpath = "//input[@id='iDefaultPurchaseUnit']")
		private static WebElement itemUnitsDefaultPurchaseUnitDropDown;

		@FindBy(xpath = "//*[@id='iDefaultPurchaseUnit_table_data_body']/tr/td")
		private static List<WebElement> itemUnitsDefaultPurchaseUnitDropDownList;
		
		
		@FindBy(xpath = "//td[@id='iDefaultPurchaseUnit_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement itemUnitsDefaultPurchaseUnitSettingsBtn;

		@FindBy(xpath = "//input[@id='fLength']")
		private static WebElement itemUnitsLength;

		@FindBy(xpath = "//input[@id='fWidth']")
		private static WebElement itemUnitswidth;

		@FindBy(xpath = "//input[@id='fHeight']")
		private static WebElement itemUnitsHeight;

		@FindBy(xpath = "//*[@id='fCBM']")
		private static WebElement itemUnitsCBM;

		@FindBy(xpath = "//input[@id='fWeight']")
		private static WebElement itemUnitsWeight;

		@FindBy(xpath = "//input[@id='iPalletUnit']")
		private static WebElement itemUnitsPalletUnitTxt;

		@FindBy(xpath = "//input[@id='fPalletQty']")
		private static WebElement itemUnitsPalletQtyTxt;

		@FindBy(xpath = "//input[@id='fPalletLength']")
		private static WebElement itemUnitsPalletLengthTXt;

		@FindBy(xpath = "//input[@id='fPalletWidth']")
		private static WebElement itemUnitsPalletWidthTxt;

		@FindBy(xpath = "//input[@id='fPalletHeight']")
		private static WebElement itemUnitsPalletHeightTxt;

		// Master Item Settings
		@FindBy(xpath = "//*[@id='divBtnGroup2']/div/a")
		private static WebElement ItemNewSettingsBtn;

		@FindBy(xpath = "//*[@id='fStandardCost']")
		private static WebElement NewSettingsStandardCost;

		@FindBy(xpath = "//*[@id='fOverheadCost']")
		private static WebElement NewSettingsOverheadCost;

		@FindBy(xpath = "//*[@id='fIndirectCost']")
		private static WebElement NewSettingsIndirectCost;

		@FindBy(xpath = "//*[@id='fProfit']")
		private static WebElement NewSettingsProfit;

		@FindBy(xpath = "//*[@id='iDontshowproductexpireddays']")
		private static WebElement NewSettingsDontshowproductexpireddays;

		@FindBy(xpath = "//*[@id='iTaxCode']")
		private static WebElement NewSettingsTaxCode;

		@FindBy(xpath = "//*[@id='iTaxCode_input_settings']/span")
		private static WebElement NewSettingsTaxCodeSttingsBtn;

		@FindBy(xpath = "//*[@id='btnMasterSaveClick']")
		private static WebElement NewSettingsSaveBtn;

		@FindBy(xpath = "//*[@id='GenerateNewMasterModel']/div/div[1]/div[2]/div/ul/li[2]/button[5]")
		private static WebElement NewSettingsCloseBtn;

		// Master Item Classification
		@FindBy(xpath = "//*[@id='divBtnGroup3']/div/a")
		private static WebElement ItemNewClassificationBtn;

		@FindBy(xpath = "//*[@id='iABCCodeValue']")
		private static WebElement newItemClassificationABC_CodeValueDropdown;

		@FindBy(xpath = "//*[@id='iABCCodeMargin']")
		private static WebElement newItemClassificationABC_CodeMarginDropdOwn;

		@FindBy(xpath = "//*[@id='iABCCodeRevenue']")
		private static WebElement newItemClassificationABCCodeRevenueDropdown;

		@FindBy(xpath = "//*[@id='iABCCodeCarryingCost']")
		private static WebElement newItemClassificationABCCodeCarryingCostDropdown;

		@FindBy(xpath = "//*[@id='sBarcode1']")
		private static WebElement newItemClassificationBarcode1;

		@FindBy(xpath = "//*[@id='sBarcode2']")
		private static WebElement newItemClassificationBarcode2;

		@FindBy(xpath = "//*[@id='sBarcode3']")
		private static WebElement newItemClassificationBarcode3;

		@FindBy(xpath = "//*[@id='iUnit1']")
		private static WebElement newItemgridFirstrowUnit;

		@FindBy(xpath = "//*[@id='sBarcode1']")
		private static WebElement newItemgridFirstrowBarcode;

		// Grid element
		@FindBy(xpath = "//*[@id='sBatch1']")
		private static WebElement newItemClassificationgridFirstrowBatch;

		@FindBy(xpath = "//*[@id='iUnit2']")
		private static WebElement newItemClassificationgridSecondrowUnit;

		@FindBy(xpath = "//*[@id='sBarcode2']")
		private static WebElement newItemClassificationgridSecondrowBarcode;

		@FindBy(xpath = "//*[@id='sBatch2']")
		private static WebElement newItemClassificationgridSecondrowBatch;

		@FindBy(xpath = "//*[@id='iUnit3']")
		private static WebElement newItemClassificationgridThridrowUnit;

		@FindBy(xpath = "//*[@id='sBarcode3']")
		private static WebElement newItemClassificationgridThridrowBarcode;

		@FindBy(xpath = "//*[@id='sBatch3']")
		private static WebElement newItemClassificationgridThridrowBatch;

		@FindBy(xpath = "//*[@id='iUnit_input_settings']")
		private static WebElement newItemClassificationgridUnitSettings;

		@FindBy(xpath = "//div[@id='divTreeMenu']")
		private static WebElement treeUnitsDisplay;

		// Master Item Other Details
		@FindBy(xpath = "//*[@id='divBtnGroup4']/div/a")
		private static WebElement ItemNewOtherDetailsBtn;

		@FindBy(xpath = "//input[@id='iCostOfIssueAccount']")
		private static WebElement newOtherDetailsCostOfIssueitem;

		@FindBy(xpath = "//*[@id='iCostOfIssueAccount_table_data_body']/tr/td")
		private static List<WebElement> newOtherDetailsCostOfIssueitemList;
		
		
		@FindBy(xpath = "//td[@id='iCostOfIssueAccount_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement newOtherDetailsCostOfIssueitemsettingsBtn;

		@FindBy(xpath = "//input[@id='iStocksAccount']")
		private static WebElement newOtherDetailsStocksitem;

		@FindBy(xpath = "//*[@id='iStocksAccount_table_data_body']/tr/td")
		private static List<WebElement> newOtherDetailsStocksitemList;
		
		
		@FindBy(xpath = "//td[@id='iStocksAccount_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement newOtherDetailsiStocksitemsettingsBtn;

		@FindBy(xpath = "//input[@id='iSalesAccount']")
		private static WebElement newOtherDetailsSalesAccount;

		@FindBy(xpath = "//*[@id='iSalesAccount_table_data_body']/tr/td")
		private static List<WebElement> newOtherDetailsSalesAccountList;
		
		
		
		@FindBy(xpath = "//td[@id='iSalesAccount_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement newOtherDetailsSalesitemsettingsBtn;

		@FindBy(xpath = "//input[@id='iWIPAccount']")
		private static WebElement newOtherDetailsWIPitem;

		@FindBy(xpath = "//td[@id='iWIPAccount_input_settings']//span[contains(@class,'icon-settings optioncontrol_settings_margin')]")
		private static WebElement newOtherDetailsWIPitemsettingsBtn;

		@FindBy(xpath = "//input[@id='iCostofShortageStockAC']")
		private static WebElement newOtherDetailsCostofShortageStockAC;

		@FindBy(xpath = "//*[@id='iCostofShortageStockAC_table_data_body']/tr/td")
		private static List<WebElement> newOtherDetailsCostofShortageStockACList;
		
		
		@FindBy(xpath = "//td[@id='iCostofShortageStockAC_input_settings']//span[contains(@class,'icon-settings optioncontrol_settings_margin')]")
		private static WebElement newOtherDetailsCostofShortageStockACsettingsBtn;

		@FindBy(xpath = "//input[@id='iCostofExcessStockAC']")
		private static WebElement newOtherDetailsCostofExcessStockAC;

		@FindBy(xpath = "//*[@id='iCostofExcessStockAC_table_data_body']/tr/td")
		private static List<WebElement> newOtherDetailsCostofExcessStockACList;
		
		@FindBy(xpath = "//td[@id='iCostofExcessStockAC_input_settings']//span[contains(@class,'icon-settings optioncontrol_settings_margin')]")
		private static WebElement OtherDetailsCostofExcessStockACsettingsBtn;

		@FindBy(xpath = "//input[@id='iCostofSaleReturnAC']")
		private static WebElement OtherDetailsCostofSaleReturnAC;

		
		@FindBy(xpath = "//*[@id='iCostofSaleReturnAC_table_data_body']/tr/td")
		private static List<WebElement> OtherDetailsCostofSaleReturnACList;
		
		
		@FindBy(xpath = "//td[@id='iCostofSaleReturnAC_input_settings']//span[contains(@class,'icon-settings optioncontrol_settings_margin')]")
		private static WebElement OtherDetailsCostofSaleReturnACsettingsBtn;

		@FindBy(xpath = "//input[@id='iPurchaseVarianceAC']")
		private static WebElement OtherDetailsPurchaseVarianceAC;

		@FindBy(xpath = "//td[@id='iPurchaseVarianceAC_input_settings']//span[contains(@class,'icon-settings optioncontrol_settings_margin')]")
		private static WebElement OtherDetailsPurchaseVarianceACsettingsBtn;

		@FindBy(xpath = "//*[@id='btnMasterSaveClick']")
		private static WebElement newOtherDetailsSaveBtn;

		@FindBy(xpath = "//*[@id='GenerateNewMasterModel']/div/div[1]/div[2]/div/ul/li[2]/button[5]")
		private static WebElement newOtherDetailsCloseBtn;

		// Grid Elements
		@FindBy(xpath = "//*[@id='ioWarehouse1']")
		private static WebElement newOtherDetailsgridFirstrowWareHouse;

		@FindBy(xpath = "//*[@id='ioWarehouse_input_settings']/span")
		private static WebElement newGridWarehouseSettingBtn;

		@FindBy(xpath = "//*[@id='1']")
		private static WebElement newOtherDetailsgridFirstrowSupplier;

		@FindBy(xpath = "//*[@id='iSupplier_input_settings']/span")
		private static WebElement newGridSupplierSettingBtn;

		@FindBy(xpath = "//*[@id='sSupCode1']")
		private static WebElement newOtherDetailsgridFirstrowCode;

		@FindBy(xpath = "//*[@id='iStartDate1']")
		private static WebElement newOtherDetailsgridFirstrowStartDate;

		@FindBy(xpath = "//*[@id='iEndDate1']")
		private static WebElement newOtherDetailsgridFirstrowEnddate;

		@FindBy(xpath = "//*[@id='fRate1']")
		private static WebElement newOtherDetailsgridFirstrowRate;

		@FindBy(xpath = "//*[@id='iTolerance1']")
		private static WebElement newOtherDetailsgridFirstTolerance;

		@FindBy(xpath = "//*[@id='iLeadTime1']")
		private static WebElement newOtherDetailsgridFirstrowLeadTime;

		@FindBy(xpath = "//*[@id='iCreditDays1']")
		private static WebElement newOtherDetailsgridFirstrowUnitCreditDays;

		@FindBy(xpath = "//*[@id='iCreditLimit1']")
		private static WebElement newOtherDetailsgridFirstrowCreditLimit;

		@FindBy(xpath = "//*[@id='LeastDeliveryDate1']")
		private static WebElement newOtherDetailsgridFirstrowLastDeliveryDate;

		// Master Item Replenishment
		@FindBy(xpath = "//*[@id='divBtnGroup5']/div/a")
		private static WebElement ItemNewReplenishmentBtn;

		@FindBy(xpath = "//*[@id='iDefaultReplenishment']")
		private static WebElement newreplenishmentDefaultReplenishment;

		@FindBy(xpath = "//*[@id='iManufacturePolicy']")
		private static WebElement newreplenishmentManufacturePolicy;

		@FindBy(xpath = "//*[@id='iPackingBOM']")
		private static WebElement newreplenishmentPackingBOM;

		@FindBy(xpath = "//*[@id='iPackingBOM_input_settings']/span")
		private static WebElement newreplenishmentPackingBOMsettingsBtn;

		@FindBy(xpath = "//*[@id='iBOM']")
		private static WebElement newreplenishmentBOM;

		@FindBy(xpath = "//*[@id='iFlushing']")
		private static WebElement newreplenishmentFlushing;

		@FindBy(xpath = "//*[@id='btnMasterSaveClick']")
		private static WebElement newreplenishmentSaveBtn;

		@FindBy(xpath = "//*[@id='GenerateNewMasterModel']/div/div[1]/div[2]/div/ul/li[2]/button[5]")
		private static WebElement newreplenishmentCloseBtn;

		// Grid
		@FindBy(xpath = "//*[@id='irWarehouse1']")
		private static WebElement newreplenishmentGridFirstRowWarehouse;

		@FindBy(xpath = "//*[@id='irStartdate1']")
		private static WebElement newreplenishmentGridFirstRowStartDate;

		@FindBy(xpath = "//*[@id='irEnddate1']")
		private static WebElement newreplenishmentGridFirstRowEndDate;

		@FindBy(xpath = "//*[@id='iReplenishqty1']")
		private static WebElement newreplenishmentGridFirstRowReplenishqty;

		@FindBy(xpath = "//*[@id='iReorderPolicy1']")
		private static WebElement newreplenishmentGridFirstRowReorderPolicy;

		@FindBy(xpath = "//*[@id='fSafetyStock1']")
		private static WebElement newreplenishmentGridFirstRowSafetyStock;

		@FindBy(xpath = "//*[@id='iSafetyLeadTime1']")
		private static WebElement newreplenishmentFirstRowSafetyLeadTime;

		@FindBy(xpath = "//*[@id='fReordercycle1']")
		private static WebElement newreplenishmentFirstRowReordercycle;

		@FindBy(xpath = "//*[@id='fReorderPoint1']")
		private static WebElement newreplenishmentFirstRowReorderPoint;

		@FindBy(xpath = "//*[@id='fReorderQuantity1']")
		private static WebElement newreplenishmentFirstRowReorderQuantity;

		@FindBy(xpath = "//*[@id='fMaxInventoryLevel1']")
		private static WebElement newreplenishmentFirstRowMaxInventoryLevel;

		@FindBy(xpath = "//*[@id='iDefaultVendor1']")
		private static WebElement newreplenishmentFirstRowDefaultVendor;

		@FindBy(xpath = "//*[@id='fMinOrderQuantity1']")
		private static WebElement newreplenishmentFirstRowMinOrderQuantity;

		@FindBy(xpath = "//*[@id='fMaxOrderQuantity1']")
		private static WebElement newreplenishmentFirstRowMaxOrderQuantity1;

		@FindBy(xpath = "//*[@id='fOrderMultiple1']")
		private static WebElement newreplenishmentFirstRowOrderMultiple;

		// OutLet Elements
		@FindBy(xpath = "//*[@id='divBtnGroup7']/div/a")
		private static WebElement ItemNewOutletBtn;

		@FindBy(xpath = "//*[@id='txtsrch-term0']")
		private static WebElement newOutletSearchOutlet;

		@FindBy(xpath = "//*[@id='cmbUserTypeMaster']")
		private static WebElement newOutletAdvanceSearch;

		@FindBy(xpath = "//*[@id='chkResizeGrid0']/span")
		private static WebElement newOutletResizeBtn;

		@FindBy(xpath = "//*[@id='chkRetainSelection']")
		private static WebElement newOutletRetainSelctionCheckBox;

		@FindBy(xpath = "//a[contains(text(),'Tariff')]")
		private static WebElement tariffTab;

		@FindBy(xpath = "//input[@id='iTariffCode1']")
		private static WebElement tariffGridTariffCode;

		@FindBy(xpath = "//input[@id='iAffectedDate1']")
		private static WebElement tariffGridAffecteddate;

		@FindBy(xpath = "//*[@id='btnMasterSaveClick']/i")
		private static WebElement itemCreationScreenSaveBtn;

		@FindBy(xpath = "//*[@id='GenerateNewMasterModel']/div/div[1]/div[2]/div/ul/li[2]/button[5]")
		private static WebElement itemCreationScreenCloseBtn;

		@FindBy(xpath = "//*[@id='doc_Searchbox']")
		public static WebElement searchBoxInDocumentCustomization;

		@FindBy(xpath = "//*[@id='liSearchOn']/a")
		public static WebElement search_SearchOnBtn;

		@FindBy(xpath = "//*[@id='liAdvanceSearch']/a")
		public static WebElement search_AdvanceSearch;

		@FindBy(xpath = "//input[@id='cmbUserTypeMaster']")
		public static WebElement cmbSearchTxt;
		
		@FindBy(xpath = "//span[@id='btnProperties']")
		public static WebElement propertiesBtn;
		
		
		

		@FindBy(xpath = "//*[@id='cmbUserTypeMaster_input_image']/span")
		public static WebElement cmbSearchExpansionBtn;

		@FindBy(xpath = "//*[@id='cmbUserTypeMaster_input_settings']/span")
		public static WebElement cmbSearchSettingBtn;

		@FindBy(xpath = "//a[@id='chkResizeGrid0']")
		public static WebElement resizeGridBtn;

		@FindBy(xpath = "//input[@id='chkRetainSelection']")
		public static WebElement retainSelectionCheckbox;

		@FindBy(xpath = "//label[@id='oncheckaccorderan0']")
		public static WebElement checkAccorderanBtn;

		@FindBy(xpath = "//*[@id='iTreeId0']")
		public static WebElement treeIdDropdown;

		@FindBy(xpath = "//*[@id='iTreeViewId0']")
		public static WebElement treeViewIdDropdown;

		// WareHouse Master

		@FindBy(xpath = "//input[@id='iBins']")
		public static WebElement binsText;

		@FindBy(xpath = "//*[@id='iBins_input_image']/span")
		public static WebElement bins_ExpansionBtn;

		@FindBy(xpath = "//*[@id='iBins_input_settings']/span")
		public static WebElement bins_SettingBtn;

		@FindBy(xpath = "//input[@id='iPhysInventoryCountingFrequency']")
		public static WebElement physInventoryCountingFrequencyTxt;

		@FindBy(xpath = "//select[@id='iWarehouseType']")
		public static WebElement warehouseTypeDropdown;

		@FindBy(xpath = "//input[@id='bDontMaintainBin']")
		public static WebElement dontMaintainBinCheckBox;

		// State Master

		@FindBy(xpath = "//input[@id='iCountry']")
		public static WebElement countryTxt;

		@FindBy(xpath = "//*[@id='iCountry_input_image']/span")
		public static WebElement country_ExpansionBtn;

		@FindBy(xpath = "//*[@id='iCountry_input_settings']/span")
		public static WebElement country_SettingBtn;

		// City Master

		@FindBy(xpath = "//input[@id='iState']")
		public static WebElement stateTxt;

		@FindBy(xpath = "//*[@id='iState_input_image']/span")
		public static WebElement state_ExpansionBtn;

		@FindBy(xpath = "//*[@id='iState_input_settings']/span")
		public static WebElement state_SettingBtn;

		// Master Delete Field

		@FindBy(xpath = "//*[@id='btnDelete']/i")
		public static WebElement deleteBtn;

		@FindBy(xpath = "//*[@id='LandingGridBody']/tr[1]/td[11]")
		public static WebElement accountNewCreationNameOfLane1;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[2]/td[11]")
		public static WebElement accountNewCreationNameOfLane2;

		@FindBy(xpath = "//*[@id='LandingGridBody']/tr[1]/td[11]")
		public static WebElement masterCreationNameOfLane1;

		@FindBy(xpath = "//*[@id='LandingGridBody']/tr[2]/td[11]")
		public static WebElement masterCreationNameOfLane2;

		@FindBy(xpath = "//*[@id='LandingGridBody']/tr[1]/td[12]")
		public static WebElement masterCreationCodeOfLane1;

		@FindBy(xpath = "//*[@id='LandingGridBody']/tr[2]/td[12]")
		public static WebElement masterCreationCodeOfLane2;

		@FindBy(xpath = "//*[@id='LandingGridBody']/tr[1]/td[13]")
		public static WebElement masterCreationTypeOfLane1;

		@FindBy(xpath = "//*[@id='LandingGridBody']/tr[2]/td[13]")
		public static WebElement masterCreationTypeOfLane2;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[8]/div[1]/label[1]/input[1]")
		public static WebElement accountFirstCheckBoxToSelection;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/ol[1]/li[5]/span[1]")
		public static WebElement accountGroupTitleDisplay;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/ol[1]/li[2]/span[1]")
		public static WebElement productGroupTitleDisplay;

		// Master Customization

		@FindBy(xpath = "//i[@class='icon-expand']")
		private static WebElement plusBtn;

		@FindBy(xpath = "//i[@class='icon-collepse']")
		private static WebElement minusBtn;

		@FindBy(xpath = "//span[contains(text(),'Master Fields')]")
		private static WebElement masterFieldsOption;

		@FindBy(xpath = "//*[@id='masterFields_list']/li[1]/span")
		private static WebElement generalTab;

		@FindBy(xpath = "//span[@class='font-5'][contains(text(),'Main')]")
		private static WebElement generalMainTab;

		@FindBy(xpath = "//div[@id='mainTab_0']//tbody//tr[1]//td[4]")
		private static WebElement accountCustomizeName;

		@FindBy(xpath = "//div[@id='mainTab_0']//tr[2]//td[4]")
		private static WebElement accountCustomizeCode;

		@FindBy(xpath = "//div[@id='mainTab_0']//tr[3]//td[4]")
		private static WebElement accountCustomizeAccountType;

		@FindBy(xpath = "//div[@id='mainTab_0']//tr[4]//td[4]")
		private static WebElement accountCustomizeCreditLimit;

		@FindBy(xpath = "//div[@id='mainTab_0']//tr[5]//td[4]")
		private static WebElement accountCustomizeCreditDays;

		@FindBy(xpath = "//div[@id='tabId_1']//span[@class='font-5'][contains(text(),'Header Details')]")
		private static WebElement generalHeaderDetailsTab;

		@FindBy(xpath = "//div[@id='tabId_1']//tbody[contains(@class,'ui-sortable')]//tr[1]//td[4]")
		private static WebElement accountCustomizeChequeDiscountLimit;

		@FindBy(xpath = "//div[@id='HeaderDetails_0']//tr[2]//td[4]")
		private static WebElement accountCustomizeRateOfInterest;

		@FindBy(xpath = "//*[@id='HeaderDetails_0']/div/div/table/tbody/tr[3]/td[4]")
		private static WebElement accountCustomizeBankAc;

		@FindBy(xpath = "//*[@id='HeaderDetails_0']/div/div/table/tbody/tr[4]/td[4]")
		private static WebElement accountCustomizePDCDiscountedAccount;

		@FindBy(xpath = "//*[@id='btnMove_Field_MasterCust']/div[1]/span")
		private static WebElement accountHeaderMoveFiledBtn;

		@FindBy(xpath = "//div[contains(text(),'Preview')]")
		private static WebElement accountHeaderPreview;

		@FindBy(xpath = "//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Save')]")
		private static WebElement accountSaveBtn;

		@FindBy(xpath = "//*[@id='btnAdd_Field_MasterCust']/div[1]/span")
		private static WebElement accountAddBtn;

		@FindBy(xpath = "//*[@id='btnCloseMasterCutomizationFromMasterScreen']/div[1]/span")
		private static WebElement accountCloseBtn;

		@FindBy(xpath = "//span[@class='Flabel theme_icon-color'][contains(text(),'Settings')]")
		private static WebElement settingOption;

		@FindBy(xpath = "//div[@id='tabId_68']//span[@class='font-5'][contains(text(),'Header Details')]")
		private static WebElement settingsHeaderDetailsTab;

		@FindBy(xpath = "//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[1]/td[1]/i[1]")
		private static WebElement accountSettingsHeaderEditFirstRowBtn;

		@FindBy(xpath = "//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[2]/td[1]/i[1]")
		private static WebElement accountSettingsHeaderEditSecondRowBtn;

		@FindBy(xpath = "//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[3]/td[1]/i[1]")
		private static WebElement accountSettingsHeaderEditThirdRowBtn;

		@FindBy(xpath = "//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[4]/td[1]/i[1]")
		private static WebElement accountSettingsHeaderEditFourthRowBtn;

		@FindBy(xpath = "//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[5]/td[1]/i[1]")
		private static WebElement accountSettingsHeaderEditFifthRowBtn;

		@FindBy(xpath = "//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[6]/td[1]/i[1]")
		private static WebElement accountSettingsHeaderEditSixthRowBtn;

		@FindBy(xpath = "//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[7]/td[1]/i[1]")
		private static WebElement accountSettingsHeaderEditSeventhRowBtn;

		@FindBy(xpath = "//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[8]/td[1]/i[1]")
		private static WebElement accountSettingsHeaderEditEighthRowBtn;

		@FindBy(xpath = "//div[contains(text(),'Move Field')]")
		private static WebElement accountSettingsHeaderMoveFiledBtn;

		@FindBy(xpath = "//div[contains(text(),'Preview')]")
		private static WebElement accountSettingsHeaderPreview;

		@FindBy(xpath = "//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Save')]")
		private static WebElement accountSettingsSaveBtn;

		@FindBy(xpath = "//div[@id='btnAdd_Field_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Add')]")
		private static WebElement accountSettingsAddBtn;

		@FindBy(xpath = "//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
		private static WebElement accountSettingsCloseBtn;

		@FindBy(xpath = "//div[@id='tabId_68']//span[@class='font-5'][contains(text(),'Body Details')]")
		private static WebElement settingsBodyDetailsTab;

		@FindBy(xpath = "//div[@id='BodyDetails_1']//div[@class='scrollable tbl-emptyfields-mastercust']")
		private static WebElement settingsBodyDetailsEmptyTab;

		@FindBy(xpath = "//span[@class='Flabel theme_icon-color'][contains(text(),'Details')]")
		private static WebElement accountMasterFieldDetailsOption;

		@FindBy(xpath = "//div[@id='tabId_69']//span[@class='font-5'][contains(text(),'Header Details')]")
		private static WebElement detailsHeaderDetailsTab;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[1]/td[4]")
		private static WebElement detailsHeaderAddress;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[2]/td[4]")
		private static WebElement detailsHeaderCity;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[3]/td[4]")
		private static WebElement detailsHeaderPin;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[4]/td[4]")
		private static WebElement detailsHeaderDeliveryAddress;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[5]/td[4]")
		private static WebElement detailsHeaderCityOne;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[6]/td[4]")
		private static WebElement detailsHeaderPinOne;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[7]/td[4]")
		private static WebElement detailsHeaderSendEmailToCustomer;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[8]/td[4]")
		private static WebElement detailsHeaderAllowCustomerPortal;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[9]/td[4]")
		private static WebElement detailsHeaderEmail;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[10]/td[4]")
		private static WebElement detailsHeaderPassword;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[11]/td[4]")
		private static WebElement detailsHeaderTelNo;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[12]/td[4]")
		private static WebElement detailsHeaderFaxNo;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[13]/td[4]")
		private static WebElement detailsHeaderBankAccountName;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[14]/td[4]")
		private static WebElement detailsHeaderBankAccountNumber;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[15]/td[4]")
		private static WebElement detailsHeaderIFSCCode;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[16]/td[4]")
		private static WebElement detailsHeaderPaymentType;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[17]/td[4]")
		private static WebElement detailsHeaderFinanceEmail;

		@FindBy(xpath = "//*[@id='HeaderDetails_2']/div/div/table/tbody/tr[18]/td[4]")
		private static WebElement detailsHeaderPortalEmail;

		@FindBy(xpath = "//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Save')]")
		private static WebElement detailsHeaderBankSaveBtn;

		@FindBy(xpath = "//div[contains(text(),'Move Field')]")
		private static WebElement detailsHeaderMoveFieldBtn;

		@FindBy(xpath = "//div[@id='btnAdd_Field_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Add')]")
		private static WebElement detailsHeaderAddBtn;

		@FindBy(xpath = "//div[contains(text(),'Preview')]")
		private static WebElement detailsHeaderPreviewBtn;

		@FindBy(xpath = "//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
		private static WebElement detailsHeaderCloseBtn;

		@FindBy(xpath = "//div[@id='tabId_69']//span[@class='font-5'][contains(text(),'Body Details')]")
		private static WebElement detailsBodyDetailsTab;

		@FindBy(xpath = "//div[@id='BodyDetails_2']//div[@class='scrollable tbl-emptyfields-mastercust']")
		private static WebElement detailsBodyDetailsEmptyTab;

		@FindBy(xpath = "//*[@id='masterFields_list']/li[4]/span")
		private static WebElement accountMasterFieldPrintLayoutOption;

		@FindBy(xpath = "//div[@id='tabId_70']//span[@class='font-5'][contains(text(),'Header Details')]")
		private static WebElement printlayoutHeaderDetailsTab;

		@FindBy(xpath = "//div[@id='HeaderDetails_3']//div[@class='scrollable tbl-emptyfields-mastercust']")
		private static WebElement printLayoutHeaderDetailsEmpty;

		@FindBy(xpath = "//div[@id='tabId_70']//li[2]//a[1]")
		private static WebElement printLayouBodyDetailsTab;

		@FindBy(xpath = "//*[@id='BodyDetails_3']/div/div/table/tbody/tr[1]/td[4]")
		private static WebElement printlayoutBodyVoucherType;

		@FindBy(xpath = "//*[@id='BodyDetails_3']/div/div/table/tbody/tr[2]/td[4]")
		private static WebElement printLayoutBodyPrintlayout;

		@FindBy(xpath = "//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Save')]")
		private static WebElement printLayoutSaveBtn;

		@FindBy(xpath = "//div[@id='btnAdd_Field_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Add')]")
		private static WebElement printLayoutAddBtn;

		@FindBy(xpath = "//div[contains(text(),'Preview')]")
		private static WebElement printLayoutPreviewBtn;

		@FindBy(xpath = "//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
		private static WebElement printLayoutCloseBtn;

		@FindBy(xpath = "//span[contains(text(),'Unique Constraints')]")
		private static WebElement accountMasterFieldsUniqueConstraintsOption;

		@FindBy(xpath = "//div[@id='tabContent_UConstraints_MasterCust']//div[@id='btnAddOrEdit_UniqueConstraints']")
		private static WebElement UniqueConstraintsAddBtn;

		@FindBy(xpath = "//div[@id='tabContent_UConstraints_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Delete')]")
		private static WebElement UniqueConstraintsDeleteBtn;

		@FindBy(xpath = "//div[@id='tabContent_UConstraints_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
		private static WebElement UniqueConstraintsCloseBtn;

		@FindBy(xpath = "//input[@id='txtConstraintName']")
		private static WebElement ConstraintNameTxt;

		@FindBy(xpath = "//select[@id='ddlConstraintTabList']")
		private static WebElement UniqueConstraints_TabsDropdown;

		@FindBy(xpath = "//*[@id='tab_Main_UcFields']/i")
		private static WebElement UniqueConstraints_MainExpansionBtn;

		@FindBy(xpath = "//*[@id='chkList_MainFields_1']/div[1]")
		private static WebElement accountUniqueConstraints_Main_NameChekbox;

		@FindBy(xpath = "//div[@id='chkList_MainFields_1']//div[2]")
		private static WebElement accountUniqueConstraints_Main_CodeChekbox;

		@FindBy(xpath = "//div[@id='chkList_MainFields_1']//div[3]")
		private static WebElement accountUniqueConstraints_Main_AccountTypeCheckbox;

		@FindBy(xpath = "//div[@id='chkList_MainFields_1']//div[4]")
		private static WebElement accountUniqueConstraints_Main_CreditLimitCheckBox;

		@FindBy(xpath = "//div[@id='chkList_MainFields_1']//div[5]")
		private static WebElement accountUniqueConstraints_Main_CreditDays;

		@FindBy(xpath = "//div[@id='tab_Header_UcFields']//i[@class='pull-right icon-expand theme_icon-color']")
		private static WebElement UniqueConstraints_HeaderExpansionBtn;

		@FindBy(xpath = "//*[@id='chkList_HeaderFields_1']/div[1]")
		private static WebElement UniqueConstraints_Header_ChequeDiscountLimitChekbox;

		@FindBy(xpath = "//*[@id='chkList_HeaderFields_1']/div[2]")
		private static WebElement UniqueConstraints_Header_RateOfInterestChekbox;

		@FindBy(xpath = "//*[@id='chkList_HeaderFields_1']/div[3]")
		private static WebElement UniqueConstraints_Header_BankAcChekbox;

		@FindBy(xpath = "//*[@id='chkList_HeaderFields_1']/div[4]")
		private static WebElement UniqueConstraints_Header_PDCDiscountedChekbox;

		@FindBy(xpath = "//*[@id='tab_Body_UcFields']/i")
		private static WebElement UniqueConstraints_BodyExpansionBtn;

		@FindBy(xpath = "//span[contains(text(),'Rules')]")
		private static WebElement RulesOption;

		@FindBy(xpath = "//*[@id='newRuleTab_div_MasterRules']/div[1]/div/div[1]/div[1]")
		private static WebElement accountRules_SaveBtn;

		@FindBy(xpath = "//div[contains(text(),'New')]")
		private static WebElement accountRules_NewBtn;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[3]/div[1]/div[1]/div[1]/div[1]/div[3]/div[2]")
		private static WebElement accountRules_DeleteBtn;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[3]/div[1]/div[1]/div[1]/div[1]/div[4]/div[2]")
		private static WebElement accountRules_CloseBtn;

		@FindBy(xpath = "//*[@id='txtMasterRuleName_MasterRules']")
		private static WebElement Rules_RuleNameText;

		@FindBy(xpath = "//*[@id='chkCreatingGroup_Rule_MasterRules']")
		private static WebElement Rules_CreatingGroupChekbox;

		@FindBy(xpath = "//*[@id='chkNewRecord_Rule_MasterRules']")
		private static WebElement Rules_NewRecordChekbox;

		@FindBy(xpath = "//*[@id='chkEdit_Rule_MasterRules']")
		private static WebElement Rules_EditChekbox;

		@FindBy(xpath = "//*[@id='chkLoad_Rule_MasterRules']")
		private static WebElement Rules_LoadChekbox;

		@FindBy(xpath = "//*[@id='chkBeforeSave_Rule_MasterRules']")
		private static WebElement Rules_BeforeSaveChekbox;

		@FindBy(xpath = "//*[@id='chkOnLeave_Rule_MasterRules']")
		private static WebElement Rules_OnLeaveChekbox;

		@FindBy(xpath = "//*[@id='chkBeforeDelete_Rule_MasterRules']")
		private static WebElement Rules_BeforeDeleteChekbox;

		@FindBy(xpath = "//*[@id='chkOnEnter_Rule_MasterRules']")
		private static WebElement Rules_OnEnterChekbox;

		@FindBy(xpath = "//label[contains(text(),'Active')]")
		private static WebElement Rules_ActiveChekbox;

		@FindBy(xpath = "//*[@id='Conditions-Context_Menu']/li[3]")
		private static WebElement Rules_NoConditionChekbox;

		@FindBy(xpath = "//*[@id='allMasterCustTabs']/li[4]/span")
		private static WebElement rules_ExternalModules;

		@FindBy(xpath = "//select[@id='ddlOnEvent_MasterExternalModule']")
		private static WebElement rules_ExternalModulesOnEventDropdown;

		@FindBy(xpath = "//input[@id='txtbuttonCaption_MasterExternalModule']")
		private static WebElement rules_ExternalModulesButtonCaption;

		@FindBy(xpath = "//select[@id='ddlModuleType_MasterExternalModule']")
		private static WebElement rules_ExternalModules_ModuleTypeBtn;

		@FindBy(xpath = "//select[@id='ddlAllDllFiles_MasterExternalModule']")
		private static WebElement rules_ExternalModules_ModuleNameDropdown;

		@FindBy(xpath = "//select[@id='ddlClassNames_MasterExternalModule']")
		private static WebElement rules_ExternalModulesClassNameDropdown;

		@FindBy(xpath = "//select[@id='ddlClassMethods_MasterExternalModule']")
		private static WebElement rules_ExternalModulesFuctionalNameDropdown;

		@FindBy(xpath = "//div[contains(text(),'Update')]")
		private static WebElement rules_ExternalModulesUpdateBtn;

		@FindBy(xpath = "//div[@id='ExternalModulesDiv_MasterExternalModule']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Save')]")
		private static WebElement rules_ExternalModulesSaveBtn;

		@FindBy(xpath = "//div[contains(text(),'Edit')]")
		private static WebElement rules_ExternalModulesEditBtn;

		@FindBy(xpath = "//div[@id='ExternalModulesDiv_MasterExternalModule']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Delete')]")
		private static WebElement rules_ExternalModulesDeleteBtn;

		@FindBy(xpath = "//div[@id='ExternalModulesDiv_MasterExternalModule']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
		private static WebElement rules_ExternalModulesCloseBtn;

		@FindBy(xpath = "//span[contains(text(),'Info Panel Customization')]")
		private static WebElement itemInfoPanelCustomizationTab;

		@FindBy(xpath = "//body[@class='custom_scrollbar']/section[@id='mainDiv']/div[@id='id_focus8_wrapper_default']/div[@class='content-wrapper']/section[@id='page_Content']/div[@class='container']/div[@id='mainLandingDiv']/div[@id='masterFirstDiv']/div[@class='tab-content']/div[@id='divLanding']/div[@id='divSecondsub']/div/div[@class='col-xs-12 col-sm-8 col-md-9 col-lg-9 padding-left-right5 mainDivCustomWidth']/div[@id='tabContent_IPCust_MasterCust']/div[@class='col-xs-12 col-sm-12 padding0']/div[@class='pull-right']/div[1]")
		private static WebElement iteminfoPanelCust_AddBtn;

		@FindBy(xpath = "//div[@id='tabContent_IPCust_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Delete')]")
		private static WebElement infoPanelCust_DeleteBtn;

		@FindBy(xpath = "//div[@id='tabContent_IPCust_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
		private static WebElement infoPanelCust_CloseBtn;

		@FindBy(xpath = "//strong[contains(text(),'Credit History')]")
		private static WebElement infoPanelCustCreditHistory;

		@FindBy(xpath = "//div[@id='divPanel_2']//strong[contains(text(),'Ageing Analysis')]")
		private static WebElement infoPanelCustAgeingAnalysis;

		@FindBy(xpath = "//strong[contains(text(),'Ageing Analysis Base')]")
		private static WebElement infoPanelCustAgeingAnalysisBase;

		@FindBy(xpath = "//strong[contains(text(),'Ageing Analysis Local')]")
		private static WebElement infoPanelCustAgeingAnalysisLocal;

		@FindBy(xpath = "//strong[contains(text(),'Authorization')]")
		private static WebElement infoPanelAuthorization;

		@FindBy(xpath = "//span[@class='Flabel'][contains(text(),'Reports')]")
		private static WebElement accountReportsTab;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Save')]")
		private static WebElement accountReports_SaveBtn;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
		private static WebElement accountReports_CloseBtn;

		@FindBy(xpath = "//*[@id='AllReportsDiv_MasterCust']/div[1]/label/input")
		private static WebElement reports_AdvancedbudgetReportChkBox;

		@FindBy(xpath = "//div[@id='AllReportsDiv_MasterCust']//div[1]//label[1]")
		private static WebElement reports_MasterInfoChkBox;

		@FindBy(xpath = "//div[@class='col-xs-12 col-sm-8 col-md-8 col-lg-9 padding0']//div[2]//label[1]")
		private static WebElement reports_ABCAnalysisAmountChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[3]//label[1]")
		private static WebElement reports_BankBookChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[4]//label[1]")
		private static WebElement reports_BankreconciliationstatementChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[5]//label[1]")
		private static WebElement reports_CashBookChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[6]//label[1]")
		private static WebElement reports_ComparativeAnalysisChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[7]//label[1]")
		private static WebElement reports_CustomerAgeingDetailsAnalysisChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[8]//label[1]")
		private static WebElement reports_CustomerageingdetailbyduedateChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[9]//label[1]")
		private static WebElement reports_CustomerageingsummaryanalysisChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[10]//label[1]")
		private static WebElement reports_CustomerbillwisesummaryChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[11]//label[1]")
		private static WebElement reports_CustomerduedateanalysisChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[12]//label[1]")
		private static WebElement reports_CustomerlistingofoutstandingbillsChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[13]//label[1]")
		private static WebElement reports_CustomeroverdueanalysisChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[14]//label[1]")
		private static WebElement reports_CustomeroverduesummaryChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[15]//label[1]")
		private static WebElement reports_CustomerstatementsChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[16]//label[1]")
		private static WebElement reports_CustomersummaryageingbyduedateChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[17]//label[1]")
		private static WebElement reports_IncomeExpenseTrendChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[18]//label[1]")
		private static WebElement reports_InterestcalculationChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[19]//label[1]")
		private static WebElement reports_LedgerChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[20]//label[1]")
		private static WebElement reports_LedgerdetailChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[21]//label[1]")
		private static WebElement reports_MonthlySalesChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[22]//label[1]")
		private static WebElement reports_PeakandlowbalanceamountChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[23]//label[1]")
		private static WebElement reports_PettycashbookChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[24]//label[1]")
		private static WebElement reports_PurchaseregisterChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[25]//label[1]")
		private static WebElement reports_PurchasereturnregisterChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[26]//label[1]")
		private static WebElement reports_PurchasesgroupedbydepartmentChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[27]//label[1]")
		private static WebElement reports_PurchasesgroupedbyItemChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[28]//label[1]")
		private static WebElement reports_PurchasesgroupedbyvendorChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[29]//label[1]")
		private static WebElement reports_SalesdaybookChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[30]//label[1]")
		private static WebElement reports_SalesgroupedbycustomerChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[31]//label[1]")
		private static WebElement reports_salesgroupedbydepartmentrChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[32]//label[1]")
		private static WebElement reports_SalesgroupedbyItemChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[33]//label[1]")
		private static WebElement reports_SalesreturnregisterChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[34]//label[1]")
		private static WebElement reports_SchedulesChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[35]//label[1]")
		private static WebElement reports_SubledgerChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[36]//label[1]")
		private static WebElement reports_SummarypurchaseregisterChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[37]//label[1]")
		private static WebElement reports_SummarysalesregisterChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[38]//label[1]")
		private static WebElement reports_TopCustomersChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[39]//label[1]")
		private static WebElement reports_TradingaccountrChkBox;

		@FindBy(xpath = "//div[40]//label[1]")
		private static WebElement reports_TransactionstypeanalysisChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[41]//label[1]")
		private static WebElement reports_VendorageingdetailanalysisChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[42]//label[1]")
		private static WebElement reports_VendorageingdetailbyduedateChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[43]//label[1]")
		private static WebElement reports_VendorageingsummaryanalysisChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[44]//label[1]")
		private static WebElement reports_VendorbillwisesummaryChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[45]//label[1]")
		private static WebElement reports_VendorduedateanalysisChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[46]//label[1]")
		private static WebElement reports_VendorlistingofoutstandingbillsChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[47]//label[1]")
		private static WebElement reports_VendoroverdueanalysisChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[48]//label[1]")
		private static WebElement reports_VendoroverduesummaryChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[49]//label[1]")
		private static WebElement reports_VendorstatementsChkBox;

		@FindBy(xpath = "//div[@id='tabContent_Reports_MasterCust']//div[50]//label[1]")
		private static WebElement reports_VendorsummaryageingbyduedateChkBox;

		// Create Tabs
		@FindBy(xpath = "//button[@id='btnCreateTab_MCust']")
		private static WebElement createTab;

		@FindBy(xpath = "//a[contains(text(),'Tree Tab')]")
		private static WebElement createTabTreeTab;

		@FindBy(xpath = "//input[@id='txtNewTreeTabCaption']")
		private static WebElement treeTabCaptionTxt;

		@FindBy(xpath = "//input[@id='txtNewTreeTabName']")
		private static WebElement treeTabTabNameTxt;

		@FindBy(xpath = "//input[@id='fOption_tab_masterName']")
		private static WebElement treeTabNewMastertxt;

		@FindBy(xpath = "//*[@id='fOption_tab_masterName_input_image']/span")
		private static WebElement treeTabNewMasterdropdown;

		@FindBy(xpath = "//*[@id='fOption_tab_masterName_input_settings']/span")
		private static WebElement treeTab_NewMasterSettingsBtn;

		@FindBy(xpath = "//input[@id='txtNewTreeTabFieldCaption']")
		private static WebElement treeTab_FieldCaptiontxt;

		@FindBy(xpath = "//input[@id='txtNewTreeTabFieldName']")
		private static WebElement treeTab_FieldNameTxt;

		@FindBy(xpath = "//input[@id='bIsTreeHiddenTab']")
		private static WebElement treeTab_HiddenChkBox;

		@FindBy(xpath = "//div[@id='fOption_tab_masterName_customize_popup_container']")
		private static WebElement treeTab_NewMaster_ContainerTxt;

		@FindBy(xpath = "//div[@id='fOption_tab_masterName_customize_popup_footer']//input[1]")
		private static WebElement treeTab_NewMaster_StanadrdfieldsBtn;

		@FindBy(xpath = "//select[@id='fOption_tab_masterName_customize_popup_standardfields_list']")
		private static WebElement treeTab_NewMaster_Stanadrdfields_Fielddropdown;

		@FindBy(xpath = "//input[@id='fOption_tab_masterName_customize_popup_standardfields_header']")
		private static WebElement treeTab_NewMaster_Stanadrdfields_headerTxt;

		@FindBy(xpath = "//select[@id='fOption_tab_masterName_customize_popup_standardfields_alignment']")
		private static WebElement treeTab_NewMaster_Stanadrdfields_Allignmentdropdown;

		@FindBy(xpath = "//input[@id='fOption_tab_masterName_customize_popup_standardfields_width']")
		private static WebElement treeTab_NewMaster_Stanadrdfields_WidthTxt;

		@FindBy(xpath = "//div[@id='divSecondsub']//input[3]")
		private static WebElement treeTab_NewMaster_Stanadrdfields_OkBtn;

		@FindBy(xpath = "//div[@id='divSecondsub']//input[4]")
		private static WebElement treeTab_NewMaster_Stanadrdfields_CancelBtn;

		@FindBy(xpath = "//div[@id='fOption_tab_masterName_search_container']//input[2]")
		private static WebElement treeTab_NewMaster_deleteColumnBtn;

		@FindBy(xpath = "//div[@id='divSecondsub']//input[3]")
		private static WebElement treeTab_NewMaster_OkBtn;

		@FindBy(xpath = "//div[@id='divSecondsub']//input[4]")
		private static WebElement treeTab_NewMaster_CancelBtn;

		@FindBy(xpath = "//*[@id='MCustomize_CreateTreeTab_Modal']/div/div/div[3]/button[1]")
		private static WebElement treeTab_OkBtn;

		@FindBy(xpath = "//div[@id='MCustomize_EditTab_Modal']//button[@class='Fbutton font-5'][contains(text(),'OK')]")
		private static WebElement treeTab_UpdateOkBtn;

		@FindBy(xpath = "//*[@id='MCustomize_CreateTreeTab_Modal']/div/div/div[3]/button[2]")
		private static WebElement treeTab_CancelBtn;

		@FindBy(xpath = "//a[contains(text(),'Document Tab')]")
		private static WebElement createTab_DocumentTab;

		@FindBy(xpath = "//input[@id='txtNewDocTabCaption']")
		private static WebElement documentTab_TabCaption;

		@FindBy(xpath = "//input[@id='txtNewDocTabName']")
		private static WebElement documentTab_TabName;

		@FindBy(xpath = "//input[@id='bIsDocHiddenTab']")
		private static WebElement documentTab_HiddenChkbox;

		@FindBy(xpath = "//div[@id='MCustomize_CreateDocTab_Modal']//button[@class='Fbutton font-5'][contains(text(),'OK')]")
		private static WebElement documentTab_OkBtn;

		@FindBy(xpath = "//div[@id='MCustomize_CreateDocTab_Modal']//button[@class='Fbutton font-5'][contains(text(),'Cancel')]")
		private static WebElement documentTab_CancelBtn;

		@FindBy(xpath = "//label[@id='lblNameError_DocTab']")
		private static WebElement customizeMasterCreateTabDocumentTabMessage;

		@FindBy(xpath = "//span[contains(text(),'Document Tab')]")
		private static WebElement customizeMasterCreateDocumentTabNewlyCreate;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[6]/ul[1]/li[1]/a[1]/span[1]")
		private static WebElement customizeMasterCreateDocumentTabNewlyHeaderDetails;

		@FindBy(xpath = "//div[@id='HeaderDetails_5']//div[@class='scrollable tbl-emptyfields-mastercust']")
		private static WebElement customizeMasterCreateDocumentTabNewlyBlankInHeaderDetails;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[6]/ul[1]/li[2]/a[1]/span[1]")
		private static WebElement customizeMasterCreateDocumentTabNewlyBodyDetails;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[6]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[4]")
		private static WebElement customizeMasterCreateDocumentTabFieldName;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[6]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[4]")
		private static WebElement customizeMasterCreateDocumentTabDocument;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[6]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[3]/td[4]")
		private static WebElement customizeMasterCreateDocumentTabCreatedDate;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[6]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[4]/td[4]")
		private static WebElement customizeMasterCreateDocumentTabLastModified;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[6]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[5]/td[4]")
		private static WebElement customizeMasterCreateDocumentFileSize;

		@FindBy(xpath = "//a[contains(text(),'Create Tab')]")
		private static WebElement createTab_CreateTab;

		@FindBy(xpath = "//div[@id='MCustomize_CreateDocTab_Modal']//h4[@class='modal-title'][contains(text(),'Tab Details')]")
		private static WebElement createTab_TabDetailsLabel;

		@FindBy(xpath = "//input[@id='txtNewTabCaption']")
		private static WebElement createTab_TabCaptionTxt;

		@FindBy(xpath = "//input[@id='txtNewTabName']")
		private static WebElement createTab_TabNameTxt;

		@FindBy(xpath = "//input[@id='bIsHiddenTab']")
		private static WebElement createTab_PopHiddenChkbox;

		@FindBy(xpath = "//*[@id='MCustomize_CreateTab_Modal']/div/div/div[3]/button[1]")
		private static WebElement createTab_PopOkBtn;

		@FindBy(xpath = "//*[@id='MCustomize_CreateTab_Modal']/div/div/div[3]/button[2]")
		private static WebElement createTab_PopCancelBtn;

		@FindBy(xpath = "//label[@id='lblNameError_NewTab']")
		private static WebElement customizeMasterCreateTabCreateTabMessage;

		// Newly Created Create Tab
		@FindBy(xpath = "//span[contains(text(),'Create Tab')]")
		private static WebElement customizeMasterCreateCreateTabNewlyCreateTabCaption;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/ul[1]/li[1]/a[1]/span[1]")
		private static WebElement customizeMasterCreateCreateTabHeaderDetails;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/div[1]/div[1]/div[1]/div[1]")
		private static WebElement customizeMasterCreateCreateTabBlankHeaderDetails;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/ul[1]/li[2]/a[1]/span[1]")
		private static WebElement customizeMasterCreateCreateTabBodyDetails;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/div[1]/div[2]/div[1]/div[1]")
		private static WebElement customizeMasterCreateCreateTabBlankBodyDetails;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[8]/div[1]/div[2]")
		private static WebElement customizeMasterCreateCreateTabNewlyCreateTabSaveButton;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[8]/div[3]/div[2]")
		private static WebElement customizeMasterCreateCreateTabAddButton;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[8]/div[4]/div[2]")
		private static WebElement customizeMasterCreateCreateTabPreviewButton;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[8]/div[5]/div[2]")
		private static WebElement customizeMasterCreateCreateTabCloseButton;

		@FindBy(xpath = "//button[@id='btnDeleteTab_MCust']")
		private static WebElement tabs_DeleteBtn;

		@FindBy(xpath = "//label[@id='lblNameError_TreeTab']")
		private static WebElement customizeMasterCreateTabsMessage;

		@FindBy(xpath = "//label[@id='lblMasterNameError_TreeTab']")
		private static WebElement customizeMasterCreateTabsMasterNameMessage;

		@FindBy(xpath = "//span[contains(text(),'Tree Tab')]")
		private static WebElement customizeMasterCreateTreeTabNewlyCreate;

		@FindBy(xpath = "//input[@id='txtEditTabCaption']")
		private static WebElement customizeMasterCreateTreeTabNewlyCreateTabCaption;

		@FindBy(xpath = "//div[@id='MCustomize_EditTab_Modal']//label[@class='col-sm-12 col-xs-12 col-md-5 col-lg-5 form-control-static Flabel'][contains(text(),'Hidden')]")
		private static WebElement customizeMasterCreateTreeTabNewlyCreateHidden;

		@FindBy(xpath = "//div[@id='MCustomize_EditTab_Modal']//button[@class='Fbutton font-5'][contains(text(),'OK')]")
		private static WebElement customizeMasterCreateTreeTabNewlyCreateOkButton;

		@FindBy(xpath = "//div[@id='MCustomize_EditTab_Modal']//button[@class='Fbutton font-5'][contains(text(),'Cancel')]")
		private static WebElement customizeMasterCreateTreeTabNewlyCreateCancelButton;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[5]/ul[1]/li[1]/a[1]/span[1]")
		private static WebElement customizeMasterCreateTreeTabNewlyHeaderDetails;

		@FindBy(xpath = "//div[@id='HeaderDetails_4']//div[@class='scrollable tbl-emptyfields-mastercust']")
		private static WebElement customizeMasterCreateTreeTabNewlyBlankInHeaderDetails;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[5]/ul[1]/li[2]/a[1]/span[1]")
		private static WebElement customizeMasterCreateTreeTabNewlyBodyDetails;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[5]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[4]")
		private static WebElement customizeMasterCreateTreeTabNewlyFieldInBodyDetails;

		@FindBy(xpath = "//span[contains(text(),'Tree Tab Delete')]")
		private static WebElement CustomizeCreateTreeTab;

		@FindBy(xpath = "//*[@id='lblWarnMessage_TabModal']")
		private static WebElement CustomizeDeleteTabMessage;

		@FindBy(xpath = "//div[@id='MCustomize_DeleteTab_Modal']//button[@class='Fbutton font-5'][contains(text(),'Yes')]")
		private static WebElement CustomizeDeleteTabMessageYesBtn;

		@FindBy(xpath = "//div[@id='MCustomize_DeleteTab_Modal']//button[@class='Fbutton font-5'][contains(text(),'No')]")
		private static WebElement CustomizeDeleteTabMessageNoBtn;

		// Master Customization Extra Fields Creation On Click Add button
		@FindBy(xpath = "//input[@id='ExtraField_FieldCaption']")
		private static WebElement extraFields_FieldDetailsTab;

		
		 public static void checkUserFriendlyMessage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		  {
			try
			{
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(errorMessage));
				String actErrorMessage=errorMessage.getText();
				
				System.out.println("Open Page then Message Display  :  " + actErrorMessage);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
				errorMessageCloseBtn.click();
				
				/*try
				{	
					System.out.println("In Try Block Validation Message  :  " + actErrorMessage);
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
					errorMessageCloseBtn.click();
				}
				catch(Exception ee)
				{
					System.out.println("In Catch Block Validation Message  :  " + actErrorMessage);	
				}*/
			}
			catch(Exception e)
			{
				System.err.println("Error Message NOT Found or NOT Clickable");
				System.err.println(e.getMessage());
				
				String Exception=e.getMessage();
			}
		  }
			
		 public boolean checkVoucherSavingMessage(String docno) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		  {
			try
			{
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(errorMessage));
				String actErrorMessage=errorMessage.getText();
				String expErrorMessage="Voucher saved successfully";
				String expErrorMessage1=": "+docno;
				
				System.out.println("SavingMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage+" "+expErrorMessage1);
				
				if(actErrorMessage.startsWith(expErrorMessage) && actErrorMessage.endsWith(expErrorMessage1))
				{
					try
					{
						getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
						errorMessageCloseBtn.click();
						
						return  true;
					}
					catch(Exception ee)
					{
						return true;
					}	
				}
				else
				{
					return false;
				}	
			}
			catch(Exception e)
			{
				System.err.println("UNABLE TO COMPARE");
				return false;
			}
		  }
			
			

		  public boolean checkLoadingMessage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		  {
			try
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
				String actVoucherLoadingMessage=errorMessage.getText();
				String expVoucherLoadingMessage="Voucher loaded successfully";
				
				System.out.println("VoucherLoadingMessage  : " + actVoucherLoadingMessage + " Value Expected : " + expVoucherLoadingMessage);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
				errorMessageCloseBtn.click();

				if(actVoucherLoadingMessage.startsWith(expVoucherLoadingMessage))
				{
					return true;
				}
				else
				{
					return false;
				}
			}
			catch(Exception e)
			{
				System.err.println("UNABLE TO COMPARE");
				return false;
			}
		  }
			
			
		
		@FindBy(xpath = "//span[contains(text(),'Field Details')]")
		private static WebElement extraFields_ClickOnFieldDetailsTab;

		@FindBy(xpath = "//*[@id='ExtraField_FieldCaption']")
		private static WebElement extraFields_FieldDetailsCaption;

		@FindBy(xpath = "//div[@id='tab_FieldDetails_ExtraField']//div[1]//div[2]//div[1]//input[1]")
		private static WebElement extraFields_FieldDetailsName;

		@FindBy(xpath = "//*[@id='ddlDatatype_ExtraField']")
		private static WebElement extraFields_FieldDetailsDatTypeDropdown;

		@FindBy(xpath = "//*[@id='ExtraField_DefaultValue']")
		private static WebElement extraFields_FieldDetailsDefaultValue;

		@FindBy(xpath = "//*[@id='ExtraField_BannerText']")
		private static WebElement extraFields_FieldDetailsBannerText;

		@FindBy(xpath = "//*[@id='ExtraField_MaxSize']")
		private static WebElement extraFields_FieldDetails_MaxSize;

		@FindBy(xpath = "//*[@id='ExtraField_ControlType']")
		private static WebElement extraFields_FieldDetailsControlTypeDropdown;

		@FindBy(xpath = "//*[@id='ExtraField_ToolTipText']")
		private static WebElement extraFields_FieldDetailsToolTipText;

		@FindBy(xpath = "//*[@id='ExtraField_RegularExpression']")
		private static WebElement extraFields_FieldDetailsRegularExpressionText;

		@FindBy(xpath = " //*[@id='btnLoad_RExp']")
		private static WebElement extraFields_RegularExperssionLoadExpansionBtn;

		@FindBy(xpath = "//*[@id='extraFieldContainer_div']/ul/li[2]/a/span")
		private static WebElement extraFields_PropertiesTab;

		@FindBy(xpath = "//*[@id='tab_Properties_ExtraField']/div[1]/div[1]")
		private static WebElement extraFields_PropertiesTabPartOfDeliveryAddressChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_PartOfBillingAddress']")
		private static WebElement extraFields_PropertiesPartOfBillingAddressChekbox;

		@FindBy(xpath = "//input[@id='FieldsDetails_IsMandatory']")
		private static WebElement extraFields_PropertiesMandatoryChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_AllowAsParameter']")
		private static WebElement extraFields_PropertiesAllowAsParameterChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_IsHidden']")
		private static WebElement extraFields_PropertiesHiddenChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_HiddenInGroup']")
		private static WebElement extraFields_PropertiesHiddenInGroupChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_SpellCheck']")
		private static WebElement extraFields_PropertiesSpellCheckChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_MandatoryInGroup']")
		private static WebElement extraFields_PropertiesMandatoryInGroupChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_AuditTrial']")
		private static WebElement extraFields_PropertiesAuditTrialChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_NotAvailableForReports']")
		private static WebElement extraFields_PropertiesNotAvailableForReportsChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_MergeField']")
		private static WebElement extraFields_PropertiesMergeFieldChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_MassUpdate']")
		private static WebElement extraFields_PropertiesMassUpdateChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_IncludeInQuickCreate']")
		private static WebElement extraFields_PropertiesIncludeInQuickCreateChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_CannotBeExported']")
		private static WebElement extraFields_PropertiesCannotBeExportedChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_CannotBeImported']")
		private static WebElement extraFields_PropertiesCannotBeImportedChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_CopyFromParent']")
		private static WebElement extraFields_PropertiesCopyFromParentChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_ReadOnly']")
		private static WebElement extraFields_PropertiesReadOnlyChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_InformationField']")
		private static WebElement extraFields_PropertiesInformationFieldChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_IsMandatoryInRevision']")
		private static WebElement extraFields_PropertiesIsMandatoryInRevisionChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_AvailableinCustomerPortal']")
		private static WebElement extraFields_PropertiesAvailableinCustomerPortalChekbox;

		@FindBy(xpath = "//*[@id='FieldsDetails_EditableinCustomerPortal']")
		private static WebElement extraFields_PropertiesEditableinCustomerPortalChekbox;

		@FindBy(xpath = "//*[@id='Behaviour']")
		private static WebElement extraFields_PropertiesBehaviourDropdown;

		@FindBy(xpath = "//*[@id='extraFieldContainer_div']/ul/li[3]/a/span")
		private static WebElement extraFields_FormattingTab;

		@FindBy(xpath = "//*[@id='FieldsDetails_ColumnSpan']")
		private static WebElement extraFields_FormattingColumnSpan;

		@FindBy(xpath = "//*[@id='FieldsDetails_RowSpan']")
		private static WebElement extraFields_FormattingRowSpan;

		@FindBy(xpath = "//*[@id='CharacterCasing']")
		private static WebElement extraFields_FormattingCharacterCasingDropdown;

		@FindBy(xpath = "//*[@id='TextAlign']")
		private static WebElement extraFields_FormattingTextAlignDropdown;

		@FindBy(xpath = "//*[@id='FieldsDetails_Font']")
		private static WebElement extraFields_FormattingFieldsDetailsFont;

		@FindBy(xpath = "//*[@id='FieldsDetails_BackColor']")
		private static WebElement extraFields_FormattingBackColor;

		@FindBy(xpath = "//*[@id='FieldsDetails_Font_Button']")
		private static WebElement extraFields_FormattingFieldsDetailsFont_Btn;

		// Elements are used for font edit option pop in field details and rules
		@FindBy(xpath = "//*[@id='fontFamily_ExtraField_FontCtrl']")
		private static WebElement FontAreialDropdown;

		@FindBy(xpath = "//*[@id='fontStyle_ExtraField_FontCtrl']")
		private static WebElement fontStyleDropdown;

		@FindBy(xpath = "//*[@id='fontWeight_ExtraField_FontCtrl']")
		private static WebElement fontWeightDropdown;

		@FindBy(xpath = "//*[@id='fontSizes_ExtraField_FontCtrl']")
		private static WebElement fontSizeDropdown;

		@FindBy(xpath = "//*[@id='fontForeColor_ExtraField_FontCtrl']")
		private static WebElement foreColourDropdown;

		@FindBy(xpath = "//*[@id='fontBackColor_ExtraField_FontCtrl']")
		private static WebElement backColorDropdown;

		@FindBy(xpath = "//*[@id='chkBaseline_ExtraField_FontCtrl']")
		private static WebElement baselineChekbox;

		@FindBy(xpath = "//*[@id='chkOverLine_ExtraField_FontCtrl']")
		private static WebElement overLineChekbox;

		@FindBy(xpath = "//*[@id='chkStrikeThrough_ExtraField_FontCtrl']")
		private static WebElement strikeThroughChekbox;

		@FindBy(xpath = "//*[@id='chkUnderline_ExtraField_FontCtrl']")
		private static WebElement underlineChekbox;

		@FindBy(xpath = "//*[@id='chkEnableFont_ExtraField_FontCtrl']")
		private static WebElement enableFontChekbox;

		@FindBy(xpath = "//*[@id='previewfont_ExtraField_FontCtrl']")
		private static WebElement frontPriveiw;

		@FindBy(xpath = "//*[@id='myFontControl_IF_Save']/i")
		private static WebElement frontSaveBtn;

		@FindBy(xpath = "//*[@id='myFontControl_IF_Close']")
		private static WebElement frontCloseBtn;

		@FindBy(xpath = "//*[@id='extraFieldContainer_div']/ul/li[4]/a/span")
		private static WebElement extraFields_RulesTab;

		@FindBy(xpath = "//*[@id='newRuleTab_div_FieldRules']/div[1]/div/div[1]/div[2]")
		private static WebElement extraFields_RulesSaveBtn;

		@FindBy(xpath = "//*[@id='newRuleTab_div_FieldRules']/div[1]/div/div[2]/div[2]")
		private static WebElement extraFields_RulesNewBtn;

		@FindBy(xpath = "//*[@id='newRuleTab_div_FieldRules']/div[1]/div/div[3]/div[2]")
		private static WebElement extraFields_RulesDeleteBtn;

		@FindBy(xpath = "//div[@id='newRuleTab_div_FieldRules']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
		private static WebElement extraFields_RulesCloseBtn;

		@FindBy(xpath = "//*[@id='lnkRule_div_FieldRules']/input")
		private static WebElement extraFields_RulesClickHereAddRuleButtonBtn;

		@FindBy(xpath = "//*[@id='txtMasterRuleName_FieldRules']")
		private static WebElement extraFields_RuleNameDropdown;

		@FindBy(xpath = "//*[@id='txtMasterRuleName_FieldRules_input_settings']/span")
		private static WebElement extraFields_RuleNameSetingBtn;

		@FindBy(xpath = "//*[@id='extraFieldContainer_div']/ul/li[5]/a/span")
		private static WebElement extraFields_ExternalModulesTab;

		@FindBy(xpath = "//*[@id='ddlOnEvent_mExtraFieldExternalModule']")
		private static WebElement extraFields_ExternalModulesOnEventDropdown;

		@FindBy(xpath = "//*[@id='txtbuttonCaption_mExtraFieldExternalModule']")
		private static WebElement extraFields_ExternalModulesButtonCaption;

		@FindBy(xpath = "//*[@id='ddlModuleType_mExtraFieldExternalModule']")
		private static WebElement extraFields_ExternalModules_ModuleTypeBtn;

		@FindBy(xpath = "//*[@id='ddlAllDllFiles_mExtraFieldExternalModule']")
		private static WebElement extraFields_ExternalModules_ModuleNameDropdown;

		@FindBy(xpath = "//*[@id='ddlClassNames_mExtraFieldExternalModule']")
		private static WebElement extraFields_ExternalModulesClassNameDropdown;

		@FindBy(xpath = "//*[@id='ddlClassMethods_mExtraFieldExternalModule']")
		private static WebElement extraFields_ExternalModulesFuctionalNameDropdown;

		@FindBy(xpath = "//*[@id='ddlAvailability_mExtraFieldExternalModule']")
		private static WebElement extraFields_ExternalModulesAvabileInDropdown;

		@FindBy(xpath = "//*[@id='MCustomize_FontDialog_Modal']/div/div/div[3]/button[1]")
		private static WebElement extraFields_ExternalModulesForMatFieldDeatailsOkBtn;

		@FindBy(xpath = "//*[@id='MCustomize_FontDialog_Modal']/div/div/div[3]/button[2]")
		private static WebElement extraFields_externalModulesForMatFieldDeatailsCanceBtn;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]")
		private static WebElement extraFields_OkBtn;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[2]/div[1]/div[2]/div[1]/div[2]/div[2]")
		private static WebElement extraFields_CloseBtn;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[4]")
		private static WebElement extraFields_CreatedByUserInHeaderDetailsTab;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/i[1]")
		private static WebElement extraFields_EditOptionInHeaderDetailsTab;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/i[2]")
		private static WebElement extraFields_DeleteOptionInHeaderDetailsTab;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[4]")
		private static WebElement extraFields_SecondFieldInHeaderDetailsTab;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[8]/div[2]/div[2]")
		private static WebElement extraFields_MoveFieldOption;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[8]/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/select[1]")
		private static WebElement extraFields_MoveFieldTabName;

		@FindBy(xpath = "//div[@id='tableType_div']//div[1]//label[1]")
		private static WebElement extraFields_MoveFieldHeader;

		@FindBy(xpath = "//div[@id='rdbBodyDiv_MoveField_Modal']//label[1]")
		private static WebElement extraFields_MoveFieldBody;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[8]/div[1]/div[1]/div[3]/button[1]")
		private static WebElement extraFields_MoveFieldOkButton;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[8]/div[1]/div[1]/div[3]/button[2]")
		private static WebElement extraFields_MoveFieldCancelButton;

		// Settings Tab Move Field
		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[3]/td[1]/i[1]")
		private static WebElement extraFields_SettingsTab_MoveField;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[4]")
		private static WebElement extraFields_SettingsTab_First;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/i[2]")
		private static WebElement extraFields_SettingsTab_FirstDelete;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[3]/td[1]/i[2]")
		private static WebElement extraFields_SettingsTab_MoveFieldOptsDelete;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[9]/div[1]/div[1]/div[2]/div[1]/div[1]/label[1]")
		private static WebElement extraFields_SettingsTab_DeleteMessage;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[9]/div[1]/div[1]/div[3]/button[2]")
		private static WebElement extraFields_SettingsTab_DeleteMessageNoOption;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[9]/div[1]/div[1]/div[3]/button[1]")
		private static WebElement extraFields_SettingsTab_DeleteMessageYesOption;

		@FindBy(xpath = "//a[@class='icon-font7 theme_color-inverse']")
		private static WebElement extraFields_GeneralTab;

		@FindBy(xpath = "//a[contains(text(),'Settings')]")
		private static WebElement extraFields_SettingsTab;

		@FindBy(xpath = "//a[contains(text(),'Details')]")
		private static WebElement extraFields_DetailsTab;

		@FindBy(xpath = "//a[contains(text(),'Print Layout')]")
		private static WebElement extraFields_PrintLayout;

		@FindBy(xpath = "//li[5]/div/div/a")
		private static WebElement extraFields_TreeTabUpdate;

		@FindBy(xpath = "//a[@class='icon-font7 theme_color'][contains(text(),'Document Tab')]")
		private static WebElement extraFields_DocumentTab;

		@FindBy(xpath = "//a[@class='icon-font7 theme_color'][contains(text(),'Create Tab')]")
		private static WebElement extraFields_CreateTab;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[8]/div[4]/div[2]")
		private static WebElement extraFields_PreviewTab;

		@FindBy(xpath = "//input[@id='sFileName1']")
		private static WebElement extraFields_PreviewTab_DocumentTab_FileName;

		@FindBy(xpath = "//input[@id='biDocument1']")
		private static WebElement extraFields_PreviewTab_DocumentTab_Document;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[6]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[5]/div[1]/div[1]/div[1]")
		private static WebElement extraFields_PreviewTab_DocumentTab_UploadDocument;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[6]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[5]/div[1]/div[1]/div[2]")
		private static WebElement extraFields_PreviewTab_DocumentTab_SaveDocument;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[6]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[5]/div[1]/div[1]/div[3]")
		private static WebElement extraFields_PreviewTab_DocumentTab_DeleteDocument;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[6]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[7]/table[1]/tbody[1]/tr[1]/td[2]/input[2]")
		private static WebElement extraFields_PreviewTab_DocumentTab_CreateDate;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[6]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[9]/div[1]/input[1]")
		private static WebElement extraFields_PreviewTab_DocumentTab_FileSize;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[7]/div[1]/div[2]/div[2]/div[1]/div[1]/input[1]")
		private static WebElement extraFields_PreviewTab_CreateCreateTab_ExtraFieldUpdate;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[3]/div[2]/div[1]/ul[1]/li[2]/button[2]/i[1]")
		private static WebElement extraFields_PreviewTab_SaveBtn;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[3]/div[2]/div[1]/ul[1]/li[2]/button[5]")
		private static WebElement extraFields_PreviewTab_CloseBtn;

		@FindBy(xpath = "/html/body/section/div[2]/div/section[1]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td[8]/div[1]/label/input")
		private static WebElement selectVendorB;

		@FindBy(xpath = "//input[@id='iPDCDiscountedAC']")
		private static WebElement pdcDicountAccount;

		@FindBy(xpath = "//input[@id='uploadBtnbiDocument1']")
		private static WebElement addGroupDocumentTab_UploadDocument;

		@FindBy(xpath = "//tr[1]//td[5]//div[1]//div[1]//div[2]")
		private static WebElement addGroupDocumentTab_SaveDocument;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/form[1]/div[1]/div[6]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[5]/div[1]/div[1]/div[3]/i[1]")
		private static WebElement addGroupDocumentTab_DeleteDocument;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/form[1]/div[1]/div[7]/div[1]/div[2]/div[2]/div[1]/div[1]/input[1]")
		private static WebElement addGroupCreateCreateTab_ExtraFieldUpdate;

		@FindBy(xpath = "//td[contains(@class,'tdName padding0')][contains(text(),'Vendor Group')]")
		public static WebElement accountGroupCreation;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[8]/div[1]/label[1]/input[1]")
		private static WebElement accSelect;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[11]")
		public static WebElement accSelectBySingleClick;

		@FindBy(xpath = "//span[contains(text(),'Vendor Group')]")
		public static WebElement clickOnAccountGroupToDisplayTitle;

		@FindBy(xpath = "//span[contains(text(),'VG')]")
		public static WebElement checkAccountGroupTitle;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/ol[1]/li[4]/span[1]")
		public static WebElement clickOnAccountGroup;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/form[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/input[1]")
		public static WebElement getCpationOfGroup;

		@FindBy(xpath = "//label[@class='msgInfo']")
		public static WebElement clickOnEmptyGroup;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[3]/ul[1]/li[1]/ul[6]/li[1]/ul[3]/li[1]/ul[1]/li[1]/a[2]")
		public static WebElement selectVendorGroupCheckbox;

		@FindBy(xpath = "/html/body/section/div[2]/div/section[1]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table/tbody/tr[1]/td/div/table/tbody/tr[1]/td[8]/div[1]/label/input")
		public static WebElement selectVendorGroup;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[9]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]")
		public static WebElement getMsgOnDelete;

		@FindBy(xpath = "//button[@id='btnOkForDelete']")
		public static WebElement clickOnOkInDelete;

		@FindBy(xpath = "//button[@id='btnCancelForDelete']")
		public static WebElement clickOnCancelInDelete;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/ol[1]/li[4]/span[1]")
		public static WebElement clickOnVendorGroup;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/form[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/input[1]")
		public static WebElement getVendorGroupName;

		@FindBy(xpath = "//button[@class='btn-link theme_icon-color btnGenerateMasterclose closeMasterPartial clsMasterSaveClose']//i[@class='icon-font6 icon-close']")
		public static WebElement closeAccountCreationScreen;

		@FindBy(xpath = "//input[@id='chkRetainSelection']")
		public static WebElement accountMasterRetain;

		@FindBy(xpath = "//input[@id='liSelectAllMasters']")
		public static WebElement accountMasterSelect;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[8]/div[1]/label[1]/input[1]")
		public static WebElement accountFirstCheckBox;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[2]/td[8]/div[1]/label[1]/input[1]")
		public static WebElement accountSecondCheckBox;

		@FindBy(xpath = "//button[@class='btn-link theme_icon-color btnGenerateMasterclose closeMasterPartial clsMasterSaveClose']")
		public static WebElement closeAccMasterCreationScreen;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/ul[1]/li[2]/button[5]/i[1]")
		public static WebElement closeMaster;

		// Dash board Page
		@FindBy(xpath = "//*[@id='dashName']")
		private static WebElement labelDashboard;

		@FindBy(xpath = "//*[@id='Select_dash']")
		private static WebElement selectDashboard;

		@FindBy(xpath = "//*[@id='Dashboard_AddDash']")
		private static WebElement newAddDashBoard;

		@FindBy(xpath = "//*[@id='Dashboard_Dash_Config']")
		private static WebElement dashboardCustomizationSettings;

		// Properties Page
		@FindBy(xpath = "//span[contains(text(),'Account Properties')]")
		public static WebElement accountPropertiesLabel;

		@FindBy(xpath = "//input[@id='rbnNormal']")
		public static WebElement normalRadioBtn;

		@FindBy(xpath = "//input[@id='rbnDisplayDebitCreditTotalForEachAccountUnderGroup']")
		private static WebElement displayDebitCreditTotalForEachAccountUnderGroup;

		@FindBy(xpath = "//input[@id='rbnClubTranOfAllAcctsUnderGrpWhenDispInLedger']")
		private static WebElement ClubTheTransactionsofAllTheAccountsUnderGroup;

		@FindBy(xpath = "//input[@id='rbnShowSummaryOfAccount']")
		private static WebElement showSummaryOfTheAccountRadioBtn;

		@FindBy(xpath = "//input[@id='chkSuspendCreditLimit']")
		private static WebElement suspendCreditLimitChkbox;

		@FindBy(xpath = "//input[@id='rbnDisplayDebitCreditTotalForEachMonth']")
		private static WebElement displayDebitCreditTotalForEachMonthRadioBtn;

		@FindBy(xpath = "//input[@id='rbnrbnDisplayDebitCreditTotalForEachDay']")
		private static WebElement displayDebitCreditTotalForEachDayRadioBtn;

		@FindBy(xpath = "//input[@id='rbnNone']")
		public static WebElement noneRadioBtn;

		@FindBy(xpath = "//input[@id='rbnConsolidateWhenBothAccountsAreSame']")
		public static WebElement consolidateWhenBothAccountsAreSameRadioBtn;

		@FindBy(xpath = "//input[@id='rbnConsolidateAlways']")
		public static WebElement consolidateAlwaysRadioBtn;

		@FindBy(xpath = "//select[@id='ddlStatus']")
		public static WebElement statusDropdown;

		@FindBy(xpath = "//input[@id='chkDoNotRestrictSelectionEvenIfRightsNotAllotted']")
		public static WebElement doNotRestrictSelectionEvenIfRightsNotAllottedCheckbox;

		@FindBy(xpath = "//input[@id='chkAllowOtherCompaniesToViewRecords']")
		public static WebElement allowOtherCompaniesToViewRecordsCheckbox;

		@FindBy(xpath = "//input[@id='chkGenerateLedgerByCurrency']")
		public static WebElement generateLedgerByCurrencyCheckbox;

		@FindBy(xpath = "//input[@id='chkSendSmsEmailWhereverAcDebitedCredited']")
		public static WebElement sendSmsEmailWhereverAcDebitedCreditedCheckbox;

		@FindBy(xpath = "//span[@id='btnPropOk']")
		public static WebElement properties_okBtn;

		@FindBy(xpath = "//*[@id='divMasterProperty']/ul/li/span[2]")
		public static WebElement properties_cancelBtn;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/ul[1]/li[1]/ul[1]/li[5]/span[1]")
		public static WebElement customizeTabWithTreeTabUpdate;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/ul[1]/li[1]/ul[1]/li[6]/span[1]")
		public static WebElement customizeTabWithDocumentTab;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/ul[1]/li[1]/ul[1]/li[7]/span[1]")
		public static WebElement customizeTabWithCreateTreeTab;

		@FindBy(xpath = "//div[@id='Layout_Popup']//input[2]")
		private static WebElement noIn;

		@FindBy(xpath = "//*[@id='dashName']")
		private static WebElement dashboard;

		@FindBy(xpath = "//div[@id='chartdiv0']")
		public static WebElement dashboardGraph;

		@FindBy(xpath = "//span[@id='btnPropOk']")
		public static WebElement dashboardLedger;

		@FindBy(xpath = "//span[@id='btnPropOk']")
		public static WebElement dashboardInfoPanel;

		// General tab
		@FindBy(xpath = "//tbody[@id='iBankAc_table_data_body']/tr/td")
		private static List<WebElement> bankAccountListCount;

		@FindBy(xpath = "//tbody[@id='iPDCDiscountedAC_table_data_body']/tr/td")
		private static List<WebElement> pdcAccountListCount;

		// Settings Tab
		@FindBy(xpath = "//tbody[@id='iExchangeAdjustmentGainAC_table_data_body']/tr/td")
		private static List<WebElement> exgGainListCount;

		@FindBy(xpath = "//tbody[@id='iExchangeAdjustmentLossAC_table_data_body']/tr/td")
		private static List<WebElement> exgLossListCount;

		@FindBy(xpath = "//tbody[@id='iPrimaryAccount_table_data_body']/tr/td")
		private static List<WebElement> primaryAccountListCount;

		@FindBy(xpath = "//tbody[@id='iReminderTerms_table_data_body']")
		private static WebElement reminderTermsEmptyList;

		@FindBy(xpath = "//tbody[@id='iDefaultCurrency_table_data_body']/tr/td")
		private static List<WebElement> defaultCurrencyListCount;

		@FindBy(xpath = "//tbody[@id='iFinanceChargeTerms_table_data_body']")
		private static WebElement financeChargeTermsEmptyList;

		@FindBy(xpath = "//tbody[@id='iPaymentTerms_table_data_body']")
		private static WebElement paymentTermsEmptyList;

		// Detail tab
		@FindBy(xpath = "//tbody[@id='iCity_table_data_body']")
		private static WebElement city1EmptyList;

		@FindBy(xpath = "//tbody[@id='iDeliveryCity_table_data_body']")
		private static WebElement city2EmptyList;

		// Tree tab
		@FindBy(xpath = "//tbody[@id='cmbUserTypeMaster_table_data_body']/tr/td")
		private static List<WebElement> mastercmbMasterListCount;

		// Currency Page
		@FindBy(xpath = "//span[@class='navText']//span[contains(text(),'Currency Master')]")
		private static WebElement currencyMasterlabel;

		@FindBy(xpath = "//*[@id='ISOCurrencyCode']")
		private static WebElement ISOCurrencyCodeTxt;

		@FindBy(xpath = "//td[@id='ISOCurrencyCode_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement ISOCurrencyCodeSettingsBtn;

		@FindBy(xpath = "//h4[@id='ISOCurrencyCode_customize_popup_heading']")
		private static WebElement ISOCurrencyCode_CustomizeDisplayColumnsLabel;

		@FindBy(xpath = "//div[@id='ISOCurrencyCode_customize_popup_container']")
		private static WebElement ISOCurrencysCodeContainerTxt;

		@FindBy(xpath = "//div[@id='ISOCurrencyCode_customize_popup_footer']//input[1]")
		private static WebElement ISOCurrencyCodeStanadrdfieldsBtn;

		@FindBy(xpath = "//select[@id='ISOCurrencyCode_customize_popup_standardfields_list']")
		private static WebElement ISOCurrencyCodeStanadrdfields_Fielddropdown;

		@FindBy(xpath = "//input[@id='ISOCurrencyCode_customize_popup_standardfields_header']")
		private static WebElement ISOCurrencyCodeStanadrdfields_headerTxt;

		@FindBy(xpath = "//select[@id='ISOCurrencyCode_customize_popup_standardfields_alignment']")
		private static WebElement ISOCurrencyCodeStanadrdfields_Allignmentdropdown;

		@FindBy(xpath = "//input[@id='ISOCurrencyCode_customize_popup_standardfields_width']")
		private static WebElement ISOCurrencyCodeStanadrdfields_WidthTxt;

		@FindBy(xpath = "//div[@id='currencyMasterWidth']//input[3]")
		private static WebElement ISOCurrencyCodeStanadrdfields_OkBtn;

		@FindBy(xpath = "//section[@id='page_Content']//input[4]")
		private static WebElement ISOCurrencyCodeStanadrdfields_CancelBtn;

		@FindBy(xpath = "//div[@id='ISOCurrencyCode_search_container']//input[2]")
		private static WebElement ISOCurrencyCodedeleteColumnBtn;

		@FindBy(xpath = "//div[@id='currencyMasterWidth']//input[3]")
		private static WebElement ISOCurrencyCodeOkBtn;

		@FindBy(xpath = "//section[@id='page_Content']//input[4]")
		private static WebElement ISOCurrencyCodeCancelBtn;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[2]/fieldset[1]/div[1]/div[2]/div[2]/input[1]")
		private static WebElement coinsNameTxt;

		@FindBy(xpath = "//input[@id='currencyName']")
		private static WebElement currencyNameTxt;

		@FindBy(xpath = "//input[@id='GeneralRoundOff']")
		private static WebElement generalRoundOffTxt;

		@FindBy(xpath = "//select[@id='RoundingType']")
		private static WebElement roundingTypeDropdown;
		
		@FindBy(xpath = "//input[@id='CurrencyUnit']")
		private static WebElement currencyUnitTxt;

		@FindBy(xpath = "//input[@id='CurrencySubUnit']")
		private static WebElement currencySubUnitTxt;

		@FindBy(xpath = "//input[@id='connector']")
		private static WebElement connectorTxt;

		@FindBy(xpath = "//input[@id='controlSymbol']")
		private static WebElement controlSymbolDropdown;

		// Control SymbolPop Up
		@FindBy(xpath = "//td[@id='controlSymbol_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement symbolSettingsBtn;

		@FindBy(xpath = "//h4[@id='controlSymbol_customize_popup_heading']")
		private static WebElement symbol_CustomizeDisplayColumnsLabel;

		@FindBy(xpath = "//div[@id='controlSymbol_customize_popup_container']")
		private static WebElement ISOCurrencysSymbolContainerTxt;

		@FindBy(xpath = "//div[@id='controlSymbol_customize_popup_footer']//input[1]")
		private static WebElement symbolStanadrdfieldsBtn;

		@FindBy(xpath = "//select[@id='controlSymbol_customize_popup_standardfields_list']")
		private static WebElement symbolStanadrdfields_Fielddropdown;

		@FindBy(xpath = "//input[@id='controlSymbol_customize_popup_standardfields_header']")
		private static WebElement symbolStanadrdfields_headerTxt;

		@FindBy(xpath = "//select[@id='controlSymbol_customize_popup_standardfields_alignment']")
		private static WebElement symbolStanadrdfields_Allignmentdropdown;

		@FindBy(xpath = "//input[@id='controlSymbol_customize_popup_standardfields_width']")
		private static WebElement symbolStanadrdfields_WidthTxt;

		@FindBy(xpath = "//div[@id='currencyMasterWidth']//input[3]")
		private static WebElement symbolStanadrdfields_OkBtn;

		@FindBy(xpath = "//section[@id='page_Content']//input[4]")
		private static WebElement symbolStanadrdfields_CancelBtn;

		@FindBy(xpath = "//div[@id='controlSymbol_search_container']//input[2]")
		private static WebElement symboldeleteColumnBtn;

		@FindBy(xpath = "//div[@id='currencyMasterWidth']//input[3]")
		private static WebElement symbolOkBtn;

		@FindBy(xpath = "//section[@id='page_Content']//input[4]")
		private static WebElement symbolCancelBtn;

		@FindBy(xpath = "//input[@id='iNoOfDecimals']")
		private static WebElement noOfDecimalsTxt;

		@FindBy(xpath = "//input[@id='NoOfDesimals']")
		private static WebElement noOfDecimalsTxtInCurrency;
		
		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[2]/fieldset[1]/div[1]/div[2]/div[4]/table[1]/tbody[1]/tr[2]")
		private static WebElement tableRow;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[2]/fieldset[1]/div[1]/div[2]/div[4]/table[1]/tbody[1]/tr[2]/td[1]")
		private static WebElement tableCol;

		// Grid Fields in Denomination Details
		@FindBy(xpath = "//td[@id='CurrencyMasterGrid_col_1-1']")
		private static WebElement denominationCode1;

		@FindBy(xpath = "//input[@id='CurrencyMasterGrid_control_heading_ctrl_2']")
		private static WebElement denominationValue1;

		@FindBy(xpath = "//td[@id='CurrencyMasterGrid_col_2-1']")
		private static WebElement denominationCode2;

		@FindBy(xpath = "//td[@id='CurrencyMasterGrid_col_2-2']")
		private static WebElement denominationValue2;

		// Pop Up Field
		@FindBy(xpath = "//span[contains(text(),'Exchange Rate Definition')]")
		private static WebElement exchangeRateDefinitionLabel;

		@FindBy(xpath = "//i[@class='icon-font7 icon-import-from-excel']")
		private static WebElement importFromExcelIcon;

		@FindBy(xpath = "//i[@class='icon-font7 icon-clear']")
		private static WebElement exchangeRateDefinitionClearIcon;

		@FindBy(xpath = "//i[@class='icon-font7 icon-save']")
		private static WebElement exchangeRateDefinitionSaveIcon;

		@FindBy(xpath = "//span[@id='btnCurrencyMasterCancel']//i[@class='icon-font7 icon-close']")
		private static WebElement exchangeRateDefinitionPopCloseIcon;

		@FindBy(xpath = "//span[@id='closeExchangeRate']")
		private static WebElement exchangeRateDefinitionPageCloseIcon;

		@FindBy(xpath = "//input[@id='id_option']")
		private static WebElement baseCurrencyTxt;

		@FindBy(xpath = "//td[@id='id_option_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement baseCurrencySettingsBtn;
		
		@FindBy(xpath = "//h4[@id='id_option_customize_popup_heading']")
		private static WebElement baseCurrency_CustomizeDisplayColumnsLabel;

		@FindBy(xpath = "//div[@id='id_option_customize_popup_container']")
		private static WebElement baseCurrencyscodeContainerTxt;

		@FindBy(xpath = "//div[@id='id_option_customize_popup_footer']//input[1]")
		private static WebElement baseCurrencyStanadrdfieldsBtn;

		@FindBy(xpath = "//select[@id='id_option_customize_popup_standardfields_list']")
		private static WebElement baseCurrencyStanadrdfields_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_option_customize_popup_standardfields_header']")
		private static WebElement baseCurrencyStanadrdfields_headerTxt;

		@FindBy(xpath = "//select[@id='id_option_customize_popup_standardfields_alignment']")
		private static WebElement baseCurrencyStanadrdfields_Allignmentdropdown;

		@FindBy(xpath = "//input[@id='id_option_customize_popup_standardfields_width']")
		private static WebElement baseCurrencyStanadrdfields_WidthTxt;

		@FindBy(xpath = "//div[@class='row']//input[3]")
		private static WebElement baseCurrencyStanadrdfields_OkBtn;

		@FindBy(xpath = "//section[@id='page_Content']//input[4]")
		private static WebElement baseCurrencyStanadrdfields_CancelBtn;

		@FindBy(xpath = "//div[@id='id_option_search_container']//input[2]")
		private static WebElement baseCurrencydeleteColumnBtn;

		@FindBy(xpath = "//div[@class='row']//input[3]")
		private static WebElement baseCurrencyOkBtn;

		@FindBy(xpath = "//section[@id='page_Content']//input[4]")
		private static WebElement baseCurrencyCancelBtn;

		// Grid Fields
		@FindBy(xpath = "//input[@id='ExchangeTable_CurrencyName']")
		private static WebElement gridcurrencyName1Txt;

		@FindBy(xpath = "//td[@id='ExchangeTable_CurrencyName_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement currencyNameSettingsBtn;

		// Pop Up
		@FindBy(xpath = "//h4[@id='ExchangeTable_CurrencyName_customize_popup_heading']")
		private static WebElement currencyName_CustomizeDisplayColumnsLabel;

		@FindBy(xpath = "//div[@id='ExchangeTable_CurrencyName_customize_popup_container']")
		private static WebElement currencyNameContainerTxt;

		@FindBy(xpath = "//div[@id='ExchangeTable_CurrencyName_customize_popup_footer']//input[1]")
		private static WebElement currencyNameStanadrdfieldsBtn;

		@FindBy(xpath = "//select[@id='ExchangeTable_CurrencyName_customize_popup_standardfields_list']")
		private static WebElement currencyNameStanadrdfields_Fielddropdown;

		@FindBy(xpath = "//input[@id='ExchangeTable_CurrencyName_customize_popup_standardfields_header']")
		private static WebElement currencyNameStanadrdfields_headerTxt;

		@FindBy(xpath = "//select[@id='ExchangeTable_CurrencyName_customize_popup_standardfields_alignment']")
		private static WebElement currencyNameStanadrdfields_Allignmentdropdown;

		@FindBy(xpath = "//input[@id='ExchangeTable_CurrencyName_customize_popup_standardfields_width']")
		private static WebElement currencyNameStanadrdfields_WidthTxt;

		@FindBy(xpath = "//table[@id='ExchangeTable']//input[3]")
		private static WebElement currencyNameStanadrdfields_OkBtn;

		@FindBy(xpath = "//section[@id='page_Content']//input[4]")
		private static WebElement currencyNameStanadrdfields_CancelBtn;

		@FindBy(xpath = "//div[@id='ExchangeTable_CurrencyName_search_container']//input[2]")
		private static WebElement currencyNamedeleteColumnBtn;

		@FindBy(xpath = "//table[@id='ExchangeTable']//input[3]")
		private static WebElement currencyNameOkBtn;

		@FindBy(xpath = "//section[@id='page_Content']//input[4]")
		private static WebElement currencyNameCancelBtn;

		@FindBy(xpath = "//input[@id='EffectiveDate']")
		private static WebElement withEffectiveDateTxt;

		@FindBy(xpath = "//*[@id='EffectiveDate_input_image']/span")
		private static WebElement dOne;

		@FindBy(xpath = "//*[@id='EffectiveDate_day_today']/td/span[1]")
		private static WebElement dTwo;

		@FindBy(xpath = "//input[@id='ExchangeTable_DefinedAs']")
		private static WebElement gridDefineAsTxt;

		@FindBy(xpath = "//input[@id='ExchangeTable_Rate']")
		private static WebElement gridRateTxt;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_1-2']")
		private static WebElement gridDefineAs1Txt;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_1-3']")
		private static WebElement gridRate1Txt;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_1-4']")
		private static WebElement gridDescription1Txt;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_2-1']")
		private static WebElement gridcurrencyName2Txt;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_2-2']")
		private static WebElement gridDefineAs2Txt;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_2-3']")
		private static WebElement gridRate2Txt;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_2-4']")
		private static WebElement gridDescription2Txt;
		
		@FindBy(xpath = "//td[@id='ExchangeTable_col_3-1']")
		private static WebElement gridcurrencyName3Txt;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_3-2']")
		private static WebElement gridDefineAs3Txt;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_3-3']")
		private static WebElement gridRate3Txt;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_3-4']")
		private static WebElement gridDescription3Txt;

		@FindBy(xpath = "//span[@id='exchange_currency']")
		private static WebElement exchangeRateIcon;

		@FindBy(xpath = "//div[contains(text(),'Clear')]")
		private static WebElement clearIcon;

		@FindBy(xpath = "//span[@id='btnSave']")
		private static WebElement saveIcon;

		@FindBy(xpath = "//span[@id='btnClose']")
		private static WebElement cancelIcon;

		// Exchange Rate Page

		// Header Fields
		@FindBy(xpath = "//i[@class='icon-font7 icon-delete']")
		private static WebElement exchangeRateDefinitionDeleteIcon;

		@FindBy(xpath = " //span[@id='closeExchangeRate']//i[@class='icon-font7 icon-close']")
		private static WebElement exchangeRateDefinitionCloseIcon;

		// Exchange Rate History
		@FindBy(xpath = "//div[@class='navText']//span[contains(text(),'Exchange Rate History')]")
		private static WebElement exchangeRateHistoryLabel;

		@FindBy(xpath = "//input[@id='id_option']")
		private static WebElement baseCurrencyInExchangeRateHistory;

		@FindBy(xpath = "//*[@id='DateOption']")
		private static WebElement dateOptionDropdown;

		@FindBy(xpath = "//*[@id='FromDate']")
		private static WebElement fromDateTxt;

		@FindBy(xpath = "//*[@id='ToDate']")
		private static WebElement toDateTxt;

		@FindBy(xpath = "//i[@class='icon-font7 icon-refresh']")
		private static WebElement loadIcon;

		@FindBy(xpath = "//i[@class='icon-font7 icon-clear']")
		private static WebElement exchangeRateHistoryclearIcon;

		// Grid
		@FindBy(xpath = "//input[@id='chkSelectAll']")
		private static WebElement toSelectAllChkBox;

		@FindBy(xpath = "//input[@id='ExchangeTable_chk0']")
		private static WebElement row1Chkbox;

		@FindBy(xpath = "//input[@id='//td[@id='ExchangeTable_col_1-2']")
		private static WebElement selectCurrency1;

		// PopUp
		@FindBy(xpath = "//td[@id='ExchangeTable_CurrencyName_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement selectCurrencySettingsBtn;

		@FindBy(xpath = "//h4[@id='ExchangeTable_CurrencyName_customize_popup_heading']")
		private static WebElement selectCurrency_CustomizeDisplayColumnsLabel;

		@FindBy(xpath = "//div[@id='ExchangeTable_CurrencyName_customize_popup_container']")
		private static WebElement isoCurrencysselectCurrencysCodeContainerTxt;

		@FindBy(xpath = "//div[@id='ExchangeTable_CurrencyName_customize_popup_footer']//input[1]")
		private static WebElement selectCurrencyStanadrdfieldsBtn;

		@FindBy(xpath = "//select[@id='ExchangeTable_CurrencyName_customize_popup_standardfields_list']")
		private static WebElement selectCurrencyStanadrdfields_Fielddropdown;

		@FindBy(xpath = "//input[@id='ExchangeTable_CurrencyName_customize_popup_standardfields_header']")
		private static WebElement selectCurrencyStanadrdfields_headerTxt;

		@FindBy(xpath = "//select[@id='ExchangeTable_CurrencyName_customize_popup_standardfields_alignment']")
		private static WebElement selectCurrencyStanadrdfields_Allignmentdropdown;

		@FindBy(xpath = "//input[@id='ExchangeTable_CurrencyName_customize_popup_standardfields_width']")
		private static WebElement selectCurrencyStanadrdfields_WidthTxt;

		@FindBy(xpath = "//section[@id='page_Content']//input[3]")
		private static WebElement selectCurrencyStanadrdfields_OkBtn;

		@FindBy(xpath = "//section[@id='page_Content']//input[4]")
		private static WebElement selectCurrencyStanadrdfields_CancelBtn;

		@FindBy(xpath = "//div[@id='ExchangeTable_CurrencyName_search_container']//input[2]")
		private static WebElement selectCurrencydeleteColumnBtn;

		@FindBy(xpath = "//section[@id='page_Content']//input[3]")
		private static WebElement selectCurrencyOkBtn;

		@FindBy(xpath = "//section[@id='page_Content']//input[4]")
		private static WebElement selectCurrencyCancelBtn;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_1-3']")
		private static WebElement DefineAs1Txt;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[2]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[3]")
		private static WebElement valueUSD;

		@FindBy(xpath = "/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[2]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[4]")
		private static WebElement valueAED;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_1-2']")
		private static WebElement selCurrencyOne;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_2-2']")
		private static WebElement selCurrecnyTwo;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_1-3']")
		private static WebElement defCurrecnyOne;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_2-3']")
		private static WebElement defCurrecnyTwo;

		// Grid Second Row
		@FindBy(xpath = "//input[@id='ExchangeTable_chk1']")
		private static WebElement Row2Chkbox;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_2-2']")
		private static WebElement SelectCurrency2;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_1-3']")
		private static WebElement DefineAs2Txt;

		@FindBy(xpath = "//i[@class='icon-font6 icon-close']")
		private static WebElement exchangeRateHistoryCloseIcon;

		@FindBy(xpath = "//span[contains(text(),'Graph')]")
		private static WebElement graphTab;

		@FindBy(xpath = "//span[@class='font-5'][contains(text(),'Table')]")
		private static WebElement tableTab;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_1-0']")
		private static WebElement selectFirstRowIndex;

		@FindBy(xpath = "//td[@id='ExchangeTable_col_2-0']")
		private static WebElement selectSecondRowIndex;

		@FindBy(xpath = "//div[contains(text(),'Delete Row')]")
		private static WebElement selectDeleteOption;

		@FindBy(xpath = "//span[@id='DeleteData']")
		private static WebElement exchnageRateDeleteOption;

		// Vouchers Home Page Ribbon Control

		@FindBy(xpath = "//div[@id='myNavbar']//span[@class='icon-new icon-font6']")
		private static WebElement newBtn;

		@FindBy(xpath = "//div[@id='id_transaction_homescreen_edit']//span[@class='icon-edit icon-font6 ImagesinArabic']")
		private static WebElement editBtn;

		@FindBy(xpath = "//span[@class='icon-print icon-font6']")
		private static WebElement printBtn;

		@FindBy(xpath = "//div[@id='id_transaction_homescreen_Delete']//span[@class='icon-delete icon-font6']")
		private static WebElement transHomePageDeleteBtn;

		@FindBy(xpath = "//span[@class='icon-authorize icon-font6']")
		private static WebElement authorizeBtn;

		@FindBy(xpath = "//span[@class='icon-suspend icon-font6']")
		private static WebElement suspendBtn;

		@FindBy(xpath = "//span[@class='icon-reject2 icon-font6']")
		private static WebElement rejectBtn;

		@FindBy(xpath = "//span[@class='icon-cheque icon-font6']")
		private static WebElement chequeReturnBtn;

		@FindBy(xpath = "//span[@class='icon-printbarcode icon-font6']")
		private static WebElement printBarCodeBtn;

		@FindBy(xpath = "//span[@class='icon-convert icon-font6']")
		private static WebElement convertBtn;

		@FindBy(xpath = "//span[@class='icon-export icon-font6']")
		private static WebElement exportToXMLBtn;

		@FindBy(xpath = "//div[@id='myNavbar']//span[@class='icon-settings icon-font6']")
		private static WebElement settingsBtn;

		@FindBy(xpath = "//div[@id='dvHomeTransClose']//span[@class='icon-close icon-font6']")
		private static WebElement homeCloseBtn;

		// Vouchers Home Page Options
		@FindBy(xpath = "//a[@class='transaction_viewname_anchor theme_color-inverse font-5']")
		private static WebElement allVouchersOption;

		@FindBy(xpath = "//a[@class='transaction_viewname_anchor font-5']")
		private static WebElement pendingBillsOption;

		@FindBy(xpath = "//input[@id='SelectDefaultView']")
		private static WebElement selectDefaultViewChkBox;

		@FindBy(xpath = "//span[@class='icon-sorting icon-font6 toolbar_button_image theme_button_color dropdown-toggle']")
		private static WebElement sortingBtn;

		@FindBy(xpath = "//div[@id='btnCreateView']//span[@class='icon-new icon-font6']")
		private static WebElement createViewBtn;

		@FindBy(xpath = "//div[@id='btnCustomize']//span[@class='icon-settings icon-font6']")
		private static WebElement customizeBtn;

		@FindBy(xpath = "//span[@class='icon-filter icon-font6']")
		private static WebElement filterBtn;

		@FindBy(xpath = "//span[@id='reportRefresh']")
		private static WebElement refreshBtn;

		@FindBy(xpath = "//i[@class='icon-font6 icon-export']")
		private static WebElement exportBtn;

		// Options Button (Display Option)
		@FindBy(xpath = "//span[@id='transOptions']")
		private static WebElement optionsBtn;

		@FindBy(xpath = "//a[contains(text(),'Auto Adjust Width')]")
		private static WebElement autoAdjustWidthBtn;

		@FindBy(xpath = "//a[contains(text(),'Fit To Screen')]")
		private static WebElement fitToScreenBtn;

		@FindBy(xpath = "//a[contains(text(),'Normal')]")
		private static WebElement normalBtn;

		// Voucher Home Page Grid Header Columns
		@FindBy(xpath = "//input[@id='HeaderChkBox']")
		private static WebElement grid_HeaderChkBox;

		@FindBy(xpath = "//*[@id='tblHeadTransRenderRow']/th[3]")
		private static WebElement grid_Header_Date;

		@FindBy(xpath = "//*[@id='tblHeadTransRenderRow']/th[4]")
		private static WebElement grid_Header_VoucherNumber;

		@FindBy(xpath = "//*[@id='tblHeadTransRenderRow']/th[5]")
		private static WebElement grid_Header_CreatedBy;

		@FindBy(xpath = "//*[@id='tblHeadTransRenderRow']/th[6]")
		private static WebElement grid_Header_ModifiedBy;

		@FindBy(xpath = "//*[@id='tblHeadTransRenderRow']/th[7]")
		private static WebElement grid_Header_CreatedDate;

		@FindBy(xpath = "//*[@id='tblHeadTransRenderRow']/th[8]")
		private static WebElement grid_Header_ModifiedDate;

		@FindBy(xpath = "//*[@id='tblHeadTransRenderRow']/th[9]")
		private static WebElement grid_Header_CreatedTime;

		@FindBy(xpath = "//*[@id='tblHeadTransRenderRow']/th[10]")
		private static WebElement grid_Header_ModifiedTime;

		@FindBy(xpath = "//*[@id='tblHeadTransRenderRow']/th[11]")
		private static WebElement grid_Header_Suspended;

		@FindBy(xpath = "//*[@id='tblHeadTransRenderRow']/th[12]")
		private static WebElement grid_Header_Cancelled;

		@FindBy(xpath = "//*[@id='tblHeadTransRenderRow']/th[13]")
		private static WebElement grid_Header_AuthorizationStatus;

		// Grid Voucher DetailsBtn
		@FindBy(xpath = "//tr[@id='trRender_1']//i[@class='icon-info icon-font7']")
		private static WebElement grid_VoucherDetailsBtn1;

		@FindBy(xpath = "//tr[@id='trRender_2']//i[@class='icon-info icon-font7']")
		private static WebElement grid_VoucherDetailsBtn2;

		// Grid CheckBox 1 And 2

		@FindBy(xpath = "//tr[@id='trRender_1']//td//input")
		private static WebElement grid_ChkBox1;

		@FindBy(xpath = "//tr[@id='trRender_2']//td//input")
		private static WebElement grid_ChkBox2;

		@FindBy(xpath = "//tr[@id='trRender_3']//td//input")
		private static WebElement grid_ChkBox3;

		@FindBy(xpath = "//tr[@id='trRender_4']//td//input")
		private static WebElement grid_ChkBox4;

		// Vouchers Home Page Footer
		@FindBy(xpath = "//div[@id='tblFooterReportRender']//input[@id='txtSearch']")
		private static WebElement searchField;

		@FindBy(xpath = "//button[@id='frstPage']")
		private static WebElement firstPageBtn;

		@FindBy(xpath = "//button[@id='id_previousPage']")
		private static WebElement previousPageBtn;

		@FindBy(xpath = "//button[@id='id_pagenumber']")
		private static WebElement currentPageNo;

		@FindBy(xpath = "//button[@id='id_nextPage']")
		private static WebElement nextPageBtn;

		@FindBy(xpath = "//button[@id='lstPage']")
		private static WebElement lastPageBtn;

		// Vouchers Entry Page Ribbon Control Options
		@FindBy(xpath = "//span[@class='icon-scroll icon-font6']")
		private static WebElement freeFlowBtn;

		@FindBy(xpath = "//span[@class='icon-header icon-font6']")
		private static WebElement jumpToHeaderSectionBtn;

		@FindBy(xpath = "//span[@class='icon-panel icon-font6']")
		private static WebElement jumpToBodySectionBtn;

		@FindBy(xpath = "//span[@class='icon-footer icon-font6']")
		private static WebElement jumpToFooterBtn;

		@FindBy(xpath = "//*[@id='id_transactionentry_new']")
		private static WebElement new_newBtn;

		@FindBy(xpath = "//*[@id='id_transactionentry_save']")
		private static WebElement transSaveBtn;

		@FindBy(xpath = "//*[@id='id_transactionentry_previous']")
		private static WebElement previousBtn;

		@FindBy(xpath = "//*[@id='id_transactionentry_next']")
		private static WebElement nextBtn;

		@FindBy(xpath = "//*[@id='id_transactionentry_print']")
		private static WebElement new_PrintBtn;

		@FindBy(xpath = "//*[@id='id_transactionentry_suspend']")
		private static WebElement new_SuspendBtn;

		@FindBy(xpath = "//div[@id='id_transactionentry_close']//span[@class='icon-close icon-font6']")
		private static WebElement new_CloseBtn;

		@FindBy(xpath = "//div[@id='id_transactionentry_delete']//span[contains(@class,'icon-delete icon-font6')]")
		private static WebElement new_DeleteBtn;

		@FindBy(xpath = "//span[@class='icon-menu icon-font4']")
		private static WebElement toggleBtn;

		@FindBy(xpath = "//*[@id='id_transactionentry_copydocument']")
		private static WebElement copyDocumentOption;

		@FindBy(xpath = "//*[@id='id_transactionentry_copytoclipboard']")
		private static WebElement copyToClipBoardOption;

		@FindBy(xpath = "//*[@id='id_transactionentry_pastefromclipboard']")
		private static WebElement pasteFromClipBoardOption;

		@FindBy(xpath = "//*[@id='id_transactionentry_raisechequereturn']")
		private static WebElement raiseCheckReturnOption;

		@FindBy(xpath = "//*[@id='id_transactionentry_addtostock']")
		private static WebElement addToStockOption;

		@FindBy(xpath = "//*[@id='id_transactionentry_markconvert']")
		private static WebElement markConvertedOption;

		@FindBy(xpath = "//*[@id='id_transactionentry_bominput']")
		private static WebElement bOMinputOption;

		@FindBy(xpath = "//*[@id='id_transactionentry_autoallocate']")
		private static WebElement autoAllocateOption;

		@FindBy(xpath = "//*[@id='id_transactionentry_reverseentry']")
		private static WebElement reverseEntryOption;

		@FindBy(xpath = "//*[@id='id_transactionentry_printbarcode']")
		private static WebElement printBarCodeOption;

		@FindBy(xpath = "//*[@id='id_transactionentry_postingdetails']")
		private static WebElement postingDetailsOption;

		@FindBy(xpath = "//*[@id='id_transactionentry_exporttoxml']")
		private static WebElement exportToXMLOption;

		@FindBy(xpath = "//*[@id='id_transactionentry_settings']")
		private static WebElement settingsOption;

		// Voucher Entry Page Header Fields
		@FindBy(xpath = "//input[@id='id_header_1']")
		private static WebElement documentNumberTxt;

		@FindBy(xpath = "//*[@id='id_header_1_input_image']/span")
		private static WebElement documentNumberdropdown_ExpansionBtn;

		@FindBy(xpath = "//input[@id='id_header_2']")
		private static WebElement dateTxt;

		@FindBy(xpath = "//td[@id='id_header_2_input_image']//span[@class='icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow']")
		private static WebElement dateTxt_CalenderBtn;

		@FindBy(xpath = "//tr[@id='id_header_2_day_today']//span[@class='theme_color-inverse'][contains(text(),'Today')]")
		private static WebElement calender_TodayBtn;

		@FindBy(xpath = "//input[@id='id_header_4']")
		private static WebElement caskBankAccountTxt;

		@FindBy(xpath = "//*[@id='id_header_4_input_image']/span")
		private static WebElement caskBankAccount_ExpansionBtn;

		@FindBy(xpath = "//*[@id='id_header_4_input_settings']/span")
		private static WebElement cashBankAccount_SettingsBtn;

		@FindBy(xpath = "//div[@id='id_header_4_customize_popup_container']")
		private static WebElement cashBankAccount_Settings_Container;

		@FindBy(xpath = "//div[@id='id_header_4_customize_popup_footer']//input[1]")
		private static WebElement cashBankAccount_Settings_StandardFieldsBtn;

		@FindBy(xpath = "//select[@id='id_header_4_customize_popup_standardfields_list']")
		private static WebElement cashBankAccount_Settings_StandardFieldsBtn_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_header_4_customize_popup_standardfields_header']")
		private static WebElement cashBankAccount_Settings_StandardFieldsBtn_HeaderTxt;

		@FindBy(xpath = "//select[@id='id_header_4_customize_popup_standardfields_alignment']")
		private static WebElement cashBankAccount_Settings_StandardFieldsBtn_Alignmentdropdown;

		@FindBy(xpath = "//input[@id='id_header_4_customize_popup_standardfields_width']")
		private static WebElement cashBankAccount_Settings_StandardFieldsBtn_WidthTxt;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement cashBankAccount_Settings_StandardFieldsBtn_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement cashBankAccount_Settings_StandardFieldsBtn_CancelBtn;

		@FindBy(xpath = "//div[@id='id_header_4_search_container']//input[2]")
		private static WebElement cashBankAccount_Settings_DeleteColumnBtn;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement cashBankAccount_Settings_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement cashBankAccount_Settings_CancelBtn;

		@FindBy(xpath = "//input[@id='id_header_268435459']")
		private static WebElement departmentTxt;

		
		@FindBy(xpath = "//input[@id='id_header_268436057']")
		private static WebElement fixedAssetTxt;
		
		@FindBy(xpath = "//*[@id='id_header_268435459_input_image']/span")
		private static WebElement department_ExpansionBtn;

		@FindBy(xpath = "//*[@id='id_header_268435459_input_settings']/span")
		private static WebElement department_SettingsBtn;

		@FindBy(xpath = "//div[@id='id_header_268435459_customize_popup_container']")
		private static WebElement department_Settings_Container;

		@FindBy(xpath = "//div[@id='id_header_268435459_customize_popup_footer']//input[1]")
		private static WebElement department_Settings_StandardFieldsBtn;

		@FindBy(xpath = "//select[@id='id_header_268435459_customize_popup_standardfields_list']")
		private static WebElement department_Settings_StandardFieldsBtn_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_header_268435459_customize_popup_standardfields_header']")
		private static WebElement department_Settings_StandardFieldsBtn_HeaderTxt;

		@FindBy(xpath = "//select[@id='id_header_268435459_customize_popup_standardfields_alignment']")
		private static WebElement department_Settings_StandardFieldsBtn_Alignmentdropdown;

		@FindBy(xpath = "//input[@id='id_header_268435459_customize_popup_standardfields_width']")
		private static WebElement department_Settings_StandardFieldsBtn_WidthTxt;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement department_Settings_StandardFieldsBtn_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement department_Settings_StandardFieldsBtn_CancelBtn;

		@FindBy(xpath = "//div[@id='id_header_268435459_search_container']//input[2]")
		private static WebElement department_Settings_DeleteColumnBtn;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement department_Settings_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement department_Settings_CancelBtn;

		@FindBy(xpath = "//*[@id='id_header_6']")
		private static WebElement maturityDateTxt;

		@FindBy(xpath = "//*[@id='id_header_6_input_image']/span']")
		private static WebElement maturityDateTxt_CalenderBtn;

		@FindBy(xpath = "//input[@id='id_header_3']")
		private static WebElement purchaseAccountTxt;

		@FindBy(xpath = "//*[@id='id_header_3_input_image']/span")
		private static WebElement purchaseAccount_ExpansionBtn;

		@FindBy(xpath = "//*[@id='id_header_3_input_settings']/span")
		private static WebElement purchaseAccount_SettingsBtn;

		@FindBy(xpath = "//div[@id='id_header_3_customize_popup_container']")
		private static WebElement purchaseAccount_Settings_Container;

		@FindBy(xpath = "//div[@id='id_header_3_customize_popup_footer']//input[1]")
		private static WebElement purchaseAccount_Settings_StandardFieldsBtn;

		@FindBy(xpath = "//select[@id='id_header_3_customize_popup_standardfields_list']")
		private static WebElement purchaseAccount_Settings_StandardFieldsBtn_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_header_3_customize_popup_standardfields_header']")
		private static WebElement purchaseAccount_Settings_StandardFieldsBtn_HeaderTxt;

		@FindBy(xpath = "//select[@id='id_header_3_customize_popup_standardfields_alignment']")
		private static WebElement purchaseAccount_Settings_StandardFieldsBtn_Alignmentdropdown;

		@FindBy(xpath = "//input[@id='id_header_3_customize_popup_standardfields_width']")
		private static WebElement purchaseAccount_Settings_StandardFieldsBtn_WidthTxt;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement purchaseAccount_Settings_StandardFieldsBtn_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement purchaseAccount_Settings_StandardFieldsBtn_CancelBtn;

		@FindBy(xpath = "//div[@id='id_header_3_search_container']//input[2]")
		private static WebElement purchaseAccount_Settings_DeleteColumnBtn;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement purchaseAccount_Settings_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement purchaseAccount_Settings_CancelBtn;

		@FindBy(xpath = "//input[@id='id_header_4']")
		private static WebElement vendorAccountTxt;

		@FindBy(xpath = "//*[@id='id_header_4_input_image']/span")
		private static WebElement vendorAccount_ExpansionBtn;

		@FindBy(xpath = "//*[@id='id_header_4_input_settings']/span")
		private static WebElement vendorAccount_SettingsBtn;

		@FindBy(xpath = "//div[@id='id_header_4_customize_popup_container']")
		private static WebElement vendorAccount_Settings_Container;

		@FindBy(xpath = "//div[@id='id_header_4_customize_popup_footer']//input[1]")
		private static WebElement vendorAccount_Settings_StandardFieldsBtn;

		@FindBy(xpath = "//select[@id='id_header_4_customize_popup_standardfields_list']")
		private static WebElement vendorAccount_Settings_StandardFieldsBtn_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_header_4_customize_popup_standardfields_header']")
		private static WebElement vendorAccount_Settings_StandardFieldsBtn_HeaderTxt;

		@FindBy(xpath = "//select[@id='id_header_4_customize_popup_standardfields_alignment']")
		private static WebElement vendorAccount_Settings_StandardFieldsBtn_Alignmentdropdown;

		@FindBy(xpath = "//input[@id='id_header_4_customize_popup_standardfields_width']")
		private static WebElement vendorAccount_Settings_StandardFieldsBtn_WidthTxt;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement vendorAccount_Settings_StandardFieldsBtn_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement vendorAccount_Settings_StandardFieldsBtn_CancelBtn;

		@FindBy(xpath = "//div[@id='id_header_4_search_container']//input[2]")
		private static WebElement vendorAccount_Settings_DeleteColumnBtn;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement vendorAccount_Settings_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement vendorAccount_Settings_CancelBtn;

		@FindBy(xpath = "//input[@id='id_header_3']")
		private static WebElement salesAccountTxt;

		@FindBy(xpath = "//*[@id='id_header_3_input_image']/span")
		private static WebElement salesAccountExpansionBtn;

		@FindBy(xpath = "//*[@id='id_header_3_input_settings']/span")
		private static WebElement salesAccount_SettingsBtn;

		@FindBy(xpath = "//div[@id='id_header_3_customize_popup_container']")
		private static WebElement salesAccount_Settings_Container;

		@FindBy(xpath = "//div[@id='id_header_3_customize_popup_footer']//input[1]")
		private static WebElement salesAccount_Settings_StandardFieldsBtn;

		@FindBy(xpath = "//select[@id='id_header_3_customize_popup_standardfields_list']")
		private static WebElement salesAccount_Settings_StandardFieldsBtn_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_header_3_customize_popup_standardfields_header']")
		private static WebElement salesAccount_Settings_StandardFieldsBtn_HeaderTxt;

		@FindBy(xpath = "//select[@id='id_header_3_customize_popup_standardfields_alignment']")
		private static WebElement salesAccount_Settings_StandardFieldsBtn_Alignmentdropdown;

		@FindBy(xpath = "//input[@id='id_header_3_customize_popup_standardfields_width']")
		private static WebElement salesAccount_Settings_StandardFieldsBtn_WidthTxt;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement salesAccount_Settings_StandardFieldsBtn_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement salesAccount_Settings_StandardFieldsBtn_CancelBtn;

		@FindBy(xpath = "//div[@id='id_header_3_search_container']//input[2]")
		private static WebElement salesAccount_Settings_DeleteColumnBtn;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement salesAccount_Settings_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement salesAccount_Settings_CancelBtn;

		@FindBy(xpath = "//input[@id='id_header_4']")
		private static WebElement customerAccountTxt;

		@FindBy(xpath = "//*[@id='id_header_4_input_image']/span")
		private static WebElement customerAccount_ExpansionBtn;

		@FindBy(xpath = "//*[@id='id_header_4_input_settings']/span")
		private static WebElement customerAccount_SettingsBtn;

		@FindBy(xpath = "//div[@id='id_header_4_customize_popup_container']")
		private static WebElement customerAccount_Settings_Container;

		@FindBy(xpath = "//div[@id='id_header_4_customize_popup_footer']//input[1]")
		private static WebElement customerAccount_Settings_StandardFieldsBtn;

		@FindBy(xpath = "//select[@id='id_header_4_customize_popup_standardfields_list']")
		private static WebElement customerAccount_Settings_StandardFieldsBtn_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_header_4_customize_popup_standardfields_header']")
		private static WebElement customerAccount_Settings_StandardFieldsBtn_HeaderTxt;

		@FindBy(xpath = "//select[@id='id_header_4_customize_popup_standardfields_alignment']")
		private static WebElement customerAccount_Settings_StandardFieldsBtn_Alignmentdropdown;

		@FindBy(xpath = "//input[@id='id_header_4_customize_popup_standardfields_width']")
		private static WebElement customerAccount_Settings_StandardFieldsBtn_WidthTxt;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement customerAccount_Settings_StandardFieldsBtn_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement customerAccount_Settings_StandardFieldsBtn_CancelBtn;

		@FindBy(xpath = "//div[@id='id_header_4_search_container']//input[2]")
		private static WebElement customerAccount_Settings_DeleteColumnBtn;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement customerAccount_Settings_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement customerAccount_Settings_CancelBtn;

		@FindBy(xpath = "//input[@id='id_header_12']")
		private static WebElement accountTxt;

		@FindBy(xpath = "//*[@id='id_header_12_input_image']/span")
		private static WebElement account_ExpansionBtn;

		@FindBy(xpath = "//*[@id='id_header_4_input_settings']/span")
		private static WebElement account_SettingsBtn;

		@FindBy(xpath = "//div[@id='id_header_4_customize_popup_container']")
		private static WebElement account_Settings_Container;

		@FindBy(xpath = "//div[@id='id_header_4_customize_popup_footer']//input[1]")
		private static WebElement account_Settings_StandardFieldsBtn;

		@FindBy(xpath = "//select[@id='id_header_4_customize_popup_standardfields_list']")
		private static WebElement account_Settings_StandardFieldsBtn_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_header_4_customize_popup_standardfields_header']")
		private static WebElement account_Settings_StandardFieldsBtn_HeaderTxt;

		@FindBy(xpath = "//select[@id='id_header_4_customize_popup_standardfields_alignment']")
		private static WebElement account_Settings_StandardFieldsBtn_Alignmentdropdown;

		@FindBy(xpath = "//input[@id='id_header_4_customize_popup_standardfields_width']")
		private static WebElement account_Settings_StandardFieldsBtn_WidthTxt;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement account_Settings_StandardFieldsBtn_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement account_Settings_StandardFieldsBtn_CancelBtn;

		@FindBy(xpath = "//div[@id='id_header_4_search_container']//input[2]")
		private static WebElement account_Settings_DeleteColumnBtn;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement account_Settings_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement account_Settings_CancelBtn;

		@FindBy(xpath = "//input[@id='id_header_31']")
		private static WebElement issuesReceiptsTxt;

		@FindBy(xpath = "//input[@id='id_header_268435460']")
		private static WebElement wareHouseTxt;

		@FindBy(xpath = "//*[@id='id_header_268435460_input_image']/span")
		private static WebElement wareHouse_ExpansionBtn;

		@FindBy(xpath = "//*[@id='id_header_268435460_input_settings']/span")
		private static WebElement wareHouse_SettingsBtn;

		@FindBy(xpath = "//div[@id='id_header_268435460_customize_popup_container']")
		private static WebElement wareHouse_Settings_Container;

		@FindBy(xpath = "//div[@id='id_header_268435460_customize_popup_footer']//input[1]")
		private static WebElement wareHouse_Settings_StandardFieldsBtn;

		@FindBy(xpath = "//select[@id='id_header_268435460_customize_popup_standardfields_list']")
		private static WebElement wareHouse_Settings_StandardFieldsBtn_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_header_268435460_customize_popup_standardfields_header']")
		private static WebElement wareHouse_Settings_StandardFieldsBtn_HeaderTxt;

		@FindBy(xpath = "//select[@id='id_header_268435460_customize_popup_standardfields_alignment']")
		private static WebElement wareHouse_Settings_StandardFieldsBtn_Alignmentdropdown;

		@FindBy(xpath = "//input[@id='id_header_268435460_customize_popup_standardfields_width']")
		private static WebElement wareHouse_Settings_StandardFieldsBtn_WidthTxt;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement wareHouse_Settings_StandardFieldsBtn_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement wareHouse_Settings_StandardFieldsBtn_CancelBtn;

		@FindBy(xpath = "//div[@id='id_header_268435460_search_container']//input[2]")
		private static WebElement wareHouse_Settings_DeleteColumnBtn;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement wareHouse_Settings_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement wareHouse_Settings_CancelBtn;

		@FindBy(xpath = "//span[contains(text(),'sName')]")
		private static WebElement customize_sName;

		@FindBy(xpath = "//span[@class='vcenter'][contains(text(),'Code')]")
		private static WebElement customize_sCode;

		@FindBy(xpath = "//span[contains(text(),'Alias')]")
		private static WebElement customize_Alias;

		@FindBy(xpath = "//span[@class='vcenter'][contains(text(),'Group')]")
		private static WebElement customize_bGroup;

		@FindBy(xpath = "//span[contains(@class,'vcenter')][contains(text(),'Balance')]")
		private static WebElement customize_Balance;

		@FindBy(xpath = "//input[@id='id_header_67108920']")
		private static WebElement appropriateBasedOndropdown;

		@FindBy(xpath = "//input[@id='id_header_67108921']")
		private static WebElement additionalValueTxt;

		@FindBy(xpath = "//input[@id='id_header_21']")
		private static WebElement raiseReceiptsChkBox;

		@FindBy(xpath = "//input[@id='id_header_67108865']")
		private static WebElement receipts_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108868']")
		private static WebElement payments_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108871']")
		private static WebElement pettyCash_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108874']")
		private static WebElement postDatedReceipts_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108865']")
		private static WebElement postDatedPayments_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108880']")
		private static WebElement purchaseVouchers_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108881']")
		private static WebElement purchaseReturns_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108884']")
		private static WebElement salesInvoice_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108886']")
		private static WebElement salesOrder_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108928']")
		private static WebElement salesInvoiceNewNarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108885']")
		private static WebElement salesReturns_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108913']")
		private static WebElement cashSales_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108914']")
		private static WebElement hirePurchaseSales_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108890']")
		private static WebElement journalEntries_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108916']")
		private static WebElement forexJV_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108910']")
		private static WebElement interDepartmentalJV_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108888']")
		private static WebElement nonStandardJournalEntries_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108892']")
		private static WebElement debitNotes_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108911']")
		private static WebElement debitNotesLineWise_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108893']")
		private static WebElement creditNotes_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108912']")
		private static WebElement creditNotesLineWisePayroll_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108894']")
		private static WebElement openingBalance_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108891']")
		private static WebElement fixedAssetDepreciationVoucher_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108909']")
		private static WebElement requestForQuote_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108903']")
		private static WebElement purchasesQuotations_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108882']")
		private static WebElement purchasesOrders_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108882']")
		private static WebElement materialReceiptsNotes_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108904']")
		private static WebElement salesQuotationss_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108887']")
		private static WebElement deliveryNotes_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108917']")
		private static WebElement posSales_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108918']")
		private static WebElement productionOrder_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108886']")
		private static WebElement salesOrders_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108895']")
		private static WebElement shortagesinStock_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108896']")
		private static WebElement excessesinStocks_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108897']")
		private static WebElement stockTransfers_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108898']")
		private static WebElement openingStocks_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108923']")
		private static WebElement openingStocksNEW_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108908']")
		private static WebElement materialRequisition_NarrationTxt;

		@FindBy(xpath = "//input[@id='id_header_67108866']")
		private static WebElement receipts_ChequeNoTxt;

		@FindBy(xpath = "//input[@id='id_header_67108869']")
		private static WebElement payments_ChequeNoTxt;

		@FindBy(xpath = "//input[@id='id_header_67108872']")
		private static WebElement pettyCash_ChequeNoTxt;

		@FindBy(xpath = "//input[@id='id_header_67108875']")
		private static WebElement postDatedReceipts_ChequeNoTxt;

		@FindBy(xpath = "//input[@id='id_header_67108878']")
		private static WebElement postDatedPayments_ChequeNoTxt;

		// Voucher Entry Page Body Fields

		@FindBy(xpath = "//input[@id='id_body_536870916']")
		private static WebElement pvWareHouseTxt;

		@FindBy(xpath = "//td[@id='id_body_536870916_input_image']//span[@class='icon-down-arrow optioncontrol_arrow_margin']")
		private static WebElement pvWareHouse_ExpansionBtn;

		@FindBy(xpath = "//td[@id='id_body_536870916_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement pvWareHouse_SettingsBtn;

		@FindBy(xpath = "//div[@id='id_body_536870916_customize_popup_container']")
		private static WebElement pvWareHouse_Settings_Container;

		@FindBy(xpath = "//div[@id='id_body_536870916_customize_popup_footer']//input[1]")
		private static WebElement pvWareHouse_Settings_StandardFieldsBtn;

		@FindBy(xpath = "//select[@id='id_body_536870916_customize_popup_standardfields_list']")
		private static WebElement pvWareHouse_Settings_StandardFieldsBtn_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_body_536870916_customize_popup_standardfields_header']")
		private static WebElement pvWareHouse_Settings_StandardFieldsBtn_HeaderTxt;

		@FindBy(xpath = "//select[@id='id_body_536870916_customize_popup_standardfields_alignment']")
		private static WebElement pvWareHouse_Settings_StandardFieldsBtn_Alignmentdropdown;

		@FindBy(xpath = "//input[@id='id_body_536870916_customize_popup_standardfields_width']")
		private static WebElement pvWareHouse_Settings_StandardFieldsBtn_WidthTxt;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement pvWareHouse_Settings_StandardFieldsBtn_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement pvWareHouse_Settings_StandardFieldsBtn_CancelBtn;

		@FindBy(xpath = "//div[@id='id_body_536870916_search_container']//input[2]")
		private static WebElement pvWareHouse_Settings_DeleteColumnBtn;

		@FindBy(xpath = "//div[@class='panel']//input[3]")
		private static WebElement pvWareHouse_Settings_OkBtn;

		@FindBy(xpath = "//div[@class='panel']//input[4]")
		private static WebElement pvWareHouse_Settings_CancelBtn;

		// Grid Elements
		@FindBy(xpath = "//th[@id='id_transaction_entry_detail_table_control_heading_1']")
		private static WebElement grid_Header_Account;

		@FindBy(xpath = "//th[@id='id_transaction_entry_detail_table_control_heading_1']")
		private static WebElement grid_Header_Item;

		@FindBy(xpath = "//th[@id='id_transaction_entry_detail_table_control_heading_1']")
		private static WebElement grid_Header_DebitAC;

		@FindBy(xpath = "//th[@id='id_transaction_entry_detail_table_control_heading_2']")
		private static WebElement grid_Header_CreditAC;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[1]")
		private static WebElement firstRowIndex;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[1]")
		private static WebElement secondRowIndex;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[1]")
		private static WebElement thirdRowIndex;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[1]")
		private static WebElement fourthRowIndex;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[1]")
		private static WebElement fifthRowIndex;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[1]")
		private static WebElement sixthRowIndex;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[1]")
		private static WebElement seventhRowIndex;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[1]")
		private static WebElement eigthRowIndex;

		@FindBy(xpath = "//span[@class='icon-delete icon-font8']")
		private static WebElement deleteRowBtn;

		@FindBy(xpath = "//span[@class='icon-insertrow icon-font8']")
		private static WebElement insertRowBtn;

		@FindBy(xpath = "//span[@class='icon-selectall icon-font8']")
		private static WebElement selectRowBtn;

		@FindBy(xpath = "//span[@class='icon-unselectall1 icon-font8']")
		private static WebElement unSelectRowBtn;

		@FindBy(xpath = "//span[@class='icon-clone icon-font8']")
		private static WebElement duplicateRowBtn;

		                  
		
		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[2]")
		private static WebElement select1stRow_1stColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[3]")
		private static WebElement select1stRow_2ndColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[4]")
		private static WebElement select1stRow_3rdColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[5]")
		private static WebElement select1stRow_4thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[6]")
		private static WebElement select1stRow_5thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[7]")
		private static WebElement select1stRow_6thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[8]")
		private static WebElement select1stRow_7thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[9]")
		private static WebElement select1stRow_8thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[10]")
		private static WebElement select1stRow_9thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[11]")
		private static WebElement select1stRow_10thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[12]")
		private static WebElement select1stRow_11thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[13]")
		private static WebElement select1stRow_12thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[14]")
		private static WebElement select1stRow_13thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[15]")
		private static WebElement select1stRow_14thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[16]")
		private static WebElement select1stRow_15thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[17]")
		private static WebElement select1stRow_16thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[2]")
		private static WebElement select2ndRow_1stColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[3]")
		private static WebElement select2ndRow_2ndColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[4]")
		private static WebElement select2ndRow_3rdColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[5]")
		private static WebElement select2ndRow_4thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[6]")
		private static WebElement select2ndRow_5thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[7]")
		private static WebElement select2ndRow_6thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[8]")
		private static WebElement select2ndRow_7thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[9]")
		private static WebElement select2ndRow_8thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[10]")
		private static WebElement select2ndRow_9thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[12]")
		private static WebElement select2ndRow_11thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[13]")
		private static WebElement select2ndRow_12thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[11]")
		private static WebElement select2ndRow_10thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[15]")
		private static WebElement select2ndRow_14thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[16]")
		private static WebElement select2ndRow_15thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[17]")
		private static WebElement select2ndRow_16thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[2]")
		private static WebElement select3rdRow_1stColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[3]")
		private static WebElement select3rdRow_2ndColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[4]")
		private static WebElement select3rdRow_3rdColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[5]")
		private static WebElement select3rdRow_4thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[6]")
		private static WebElement select3rdRow_5thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[7]")
		private static WebElement select3rdRow_6thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[8]")
		private static WebElement select3rdRow_7thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[9]")
		private static WebElement select3rdRow_8thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[10]")
		private static WebElement select3rdRow_9thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[11]")
		private static WebElement select3rdRow_10thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[12]")
		private static WebElement select3rdRow_11thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[13]")
		private static WebElement select3rdRow_12thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[14]")
		private static WebElement select3rdRow_13thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[15]")
		private static WebElement select3rdRow_14thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[16]")
		private static WebElement select3rdRow_15thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[17]")
		private static WebElement select3rdRow_16thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[2]")
		private static WebElement select4thRow_1stColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[3]")
		private static WebElement select4thRow_2ndColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[4]")
		private static WebElement select4thRow_3rdColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[5]")
		private static WebElement select4thRow_4thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[6]")
		private static WebElement select4thRow_5thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[7]")
		private static WebElement select4thRow_6thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[8]")
		private static WebElement select4thRow_7thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[9]")
		private static WebElement select4thRow_8thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[10]")
		private static WebElement select4thRow_9thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[11]")
		private static WebElement select4thRow_10thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[12]")
		private static WebElement select4thRow_11thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[13]")
		private static WebElement select4thRow_12thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[14]")
		private static WebElement select4thRow_13thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[15]")
		private static WebElement select4thRow_14thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[16]")
		private static WebElement select4thRow_15thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[17]")
		private static WebElement select4thRow_16thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[2]")
		private static WebElement select5thRow_1stColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[3]")
		private static WebElement select5thRow_2ndColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[4]")
		private static WebElement select5thRow_3rdColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[5]")
		private static WebElement select5thRow_4thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[6]")
		private static WebElement select5thRow_5thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[7]")
		private static WebElement select5thRow_6thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[8]")
		private static WebElement select5thRow_7thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[9]")
		private static WebElement select5thRow_8thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[10]")
		private static WebElement select5thRow_9thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[11]")
		private static WebElement select5thRow_10thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[12]")
		private static WebElement select5thRow_11thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[13]")
		private static WebElement select5thRow_12thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[14]")
		private static WebElement select5thRow_13thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[15]")
		private static WebElement select5thRow_14thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[16]")
		private static WebElement select5thRow_15thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[17]")
		private static WebElement select5thRow_16thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[2]")
		private static WebElement select6thRow_1stColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[3]")
		private static WebElement select6thRow_2ndColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[4]")
		private static WebElement select6thRow_3rdColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[5]")
		private static WebElement select6thRow_4thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[6]")
		private static WebElement select6thRow_5thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[7]")
		private static WebElement select6thRow_6thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[8]")
		private static WebElement select6thRow_7thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[9]")
		private static WebElement select6thRow_8thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[10]")
		private static WebElement select6thRow_9thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[11]")
		private static WebElement select6thRow_10thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[12]")
		private static WebElement select6thRow_11thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[13]")
		private static WebElement select6thRow_12thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[14]")
		private static WebElement select6thRow_13thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[15]")
		private static WebElement select6thRow_14thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[16]")
		private static WebElement select6thRow_15thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[17]")
		private static WebElement select6thRow_16thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[2]")
		private static WebElement select7thRow_1stColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[3]")
		private static WebElement select7thRow_2ndColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[4]")
		private static WebElement select7thRow_3rdColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[5]")
		private static WebElement select7thRow_4thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[6]")
		private static WebElement select7thRow_5thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[7]")
		private static WebElement select7thRow_6thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[8]")
		private static WebElement select7thRow_7thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[9]")
		private static WebElement select7thRow_8thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[10]")
		private static WebElement select7thRow_9thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[2]")
		private static WebElement select8thRow_1stColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[3]")
		private static WebElement select8thRow_2ndColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[4]")
		private static WebElement select8thRow_3rdColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[5]")
		private static WebElement select8thRow_4thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[6]")
		private static WebElement select8thRow_5thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[7]")
		private static WebElement select8thRow_6thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[8]")
		private static WebElement select8thRow_7thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[9]")
		private static WebElement select8thRow_8thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[10]")
		private static WebElement select8thRow_9thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[2]")
		private static WebElement select9thRow_1stColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[3]")
		private static WebElement select9thRow_2ndColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[4]")
		private static WebElement select9thRow_3rdColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[5]")
		private static WebElement select9thRow_4thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[6]")
		private static WebElement select9thRow_5thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[7]")
		private static WebElement select9thRow_6thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[8]")
		private static WebElement select9thRow_7thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[9]")
		private static WebElement select9thRow_8thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[10]")
		private static WebElement select9thRow_9thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[2]")
		private static WebElement select10thRow_1stColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[3]")
		private static WebElement select10thRow_2ndColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[4]")
		private static WebElement select10thRow_3rdColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[5]")
		private static WebElement select10thRow_4thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[6]")
		private static WebElement select10thRow_5thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[7]")
		private static WebElement select10thRow_6thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[8]")
		private static WebElement select10thRow_7thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[9]")
		private static WebElement select10thRow_8thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[10]")
		private static WebElement select10thRow_9thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[2]")
		private static WebElement select11thRow_1stColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[3]")
		private static WebElement select11thRow_2ndColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[4]")
		private static WebElement select11thRow_3rdColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[5]")
		private static WebElement select11thRow_4thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[6]")
		private static WebElement select11thRow_5thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[7]")
		private static WebElement select11thRow_6thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[8]")
		private static WebElement select11thRow_7thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[9]")
		private static WebElement select11thRow_8thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[10]")
		private static WebElement select11thRow_9thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[2]")
		private static WebElement select12thRow_1stColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[3]")
		private static WebElement select12thRow_2ndColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[4]")
		private static WebElement select12thRow_3rdColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[5]")
		private static WebElement select12thRow_4thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[6]")
		private static WebElement select12thRow_5thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[7]")
		private static WebElement select12thRow_6thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[8]")
		private static WebElement select12thRow_7thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[9]")
		private static WebElement select12thRow_8thColumn;

		@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[10]")
		private static WebElement select12thRow_9thColumn;

		@FindBy(xpath = "//input[@id='id_body_12']")
		private static WebElement enter_AccountTxt;

		@FindBy(xpath = "//*[@id='id_body_12_input_image']/span")
		private static WebElement enter_Account_ExpansionBtn;

		@FindBy(xpath = "//*[@id='id_body_12_input_settings']/span")
		private static WebElement enter_Account_SettingsBtn;

		@FindBy(xpath = "//div[@id='id_header_12_customize_popup_container']")
		private static WebElement enter_Account_Settings_Container;

		@FindBy(xpath = "//div[@id='id_header_12_customize_popup_footer']//input[1]")
		private static WebElement enter_Account_Settings_StandardFieldsBtn;

		@FindBy(xpath = "//select[@id='id_header_12_customize_popup_standardfields_list']")
		private static WebElement enter_Account_Settings_StandardFieldsBtn_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_header_12_customize_popup_standardfields_header']")
		private static WebElement enter_Account_Settings_StandardFieldsBtn_HeaderTxt;

		@FindBy(xpath = "//select[@id='id_header_12_customize_popup_standardfields_alignment']")
		private static WebElement enter_Account_Settings_StandardFieldsBtn_Alignmentdropdown;

		@FindBy(xpath = "//input[@id='id_header_12_customize_popup_standardfields_width']")
		private static WebElement enter_Account_Settings_StandardFieldsBtn_WidthTxt;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[3]")
		private static WebElement enter_Account_Settings_StandardFieldsBtn_OkBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[4]")
		private static WebElement enter_Account_Settings_StandardFieldsBtn_CancelBtn;

		@FindBy(xpath = "//div[@id='id_header_12_search_container']//input[2]")
		private static WebElement enter_Account_Settings_DeleteColumnBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[3]")
		private static WebElement enter_Account_Settings_OkBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[4]")
		private static WebElement enter_Account_Settings_CancelBtn;

		@FindBy(xpath = "//input[@id='id_body_16']")
		private static WebElement enter_Amount;

		@FindBy(xpath = "//textarea[@id='id_body_16777219']")
		private static WebElement enter_Receipts_Remarks;

		@FindBy(xpath = "//textarea[@id='id_body_16777222']")
		private static WebElement enter_Payments_Remarks;

		@FindBy(xpath = "//textarea[@id='id_body_16777225']")
		private static WebElement enter_PettyCash_Remarks;

		@FindBy(xpath = "//textarea[@id='id_body_16777228']")
		private static WebElement enter_PostDatedReceipts_Remarks;

		@FindBy(xpath = "//textarea[@id='id_body_16777231']")
		private static WebElement enter_PostDatedPaymentss_Remarks;

		@FindBy(xpath = "//input[@id='id_body_23']")
		private static WebElement enter_ItemTxt;
		
		
		@FindBy(xpath = "//textarea[@id='id_body_16777360']")
		private static WebElement enterDescrptionTxt;
		
		@FindBy(xpath = "//*[@id='id_body_23_input_image']/span")
		private static WebElement item_ExpansionBtn;

		@FindBy(xpath = "//*[@id='id_body_23_input_settings']/span")
		private static WebElement item_SettingsBtn;

		@FindBy(xpath = "//div[@id='id_body_23_customize_popup_container']")
		private static WebElement item_Settings_Container;

		@FindBy(xpath = "//div[@id='id_body_23_customize_popup_footer']//input[1]")
		private static WebElement item_Settings_StandardFieldsBtn;

		@FindBy(xpath = "//select[@id='id_body_23_customize_popup_standardfields_list']")
		private static WebElement item_Settings_StandardFieldsBtn_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_body_23_customize_popup_standardfields_header']")
		private static WebElement item_Settings_StandardFieldsBtn_HeaderTxt;

		@FindBy(xpath = "//select[@id='id_body_23_customize_popup_standardfields_alignment']")
		private static WebElement item_Settings_StandardFieldsBtn_Alignmentdropdown;

		@FindBy(xpath = "//input[@id='id_body_23_customize_popup_standardfields_width']")
		private static WebElement item_Settings_StandardFieldsBtn_WidthTxt;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[3]")
		private static WebElement item_Settings_StandardFieldsBtn_OkBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[4]")
		private static WebElement item_Settings_StandardFieldsBtn_CancelBtn;

		@FindBy(xpath = "//div[@id='id_body_23_search_container']//input[2]")
		private static WebElement item_Settings_DeleteColumnBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[3]")
		private static WebElement item_Settings_OkBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[4]")
		private static WebElement item_Settings_CancelBtn;

		@FindBy(xpath = "//input[@id='id_body_24']")
		private static WebElement enter_UnitTxt;

		@FindBy(xpath = "//*[@id='id_body_24_input_image']/span")
		private static WebElement unit_ExpansionBtn;

		@FindBy(xpath = "//*[@id='id_body_24_input_settings']/span")
		private static WebElement unit_SettingsBtn;

		@FindBy(xpath = "//div[@id='id_body_24_customize_popup_container']")
		private static WebElement unit_Settings_Container;

		@FindBy(xpath = "//div[@id='id_body_24_customize_popup_footer']//input[1]")
		private static WebElement unit_Settings_StandardFieldsBtn;

		@FindBy(xpath = "//select[@id='id_body_24_customize_popup_standardfields_list']")
		private static WebElement unit_Settings_StandardFieldsBtn_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_body_24_customize_popup_standardfields_header']")
		private static WebElement unit_Settings_StandardFieldsBtn_HeaderTxt;

		@FindBy(xpath = "//select[@id='id_body_24_customize_popup_standardfields_alignment']")
		private static WebElement unit_Settings_StandardFieldsBtn_Alignmentdropdown;

		@FindBy(xpath = "//input[@id='id_body_24_customize_popup_standardfields_width']")
		private static WebElement unit_Settings_StandardFieldsBtn_WidthTxt;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[3]")
		private static WebElement unit_Settings_StandardFieldsBtn_OkBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[4]")
		private static WebElement unit_Settings_StandardFieldsBtn_CancelBtn;

		@FindBy(xpath = "//div[@id='id_body_24_search_container']//input[2]")
		private static WebElement unit_Settings_DeleteColumnBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[3]")
		private static WebElement unit_Settings_OkBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[4]")
		private static WebElement unit_Settings_CancelBtn;

		@FindBy(xpath = "//input[@id='id_body_26']")
		private static WebElement enter_Quantity;

		@FindBy(xpath = "//input[@id='id_body_27']")
		private static WebElement enter_Rate;

		@FindBy(xpath = "//input[@id='id_body_28']")
		private static WebElement enter_Gross;

		@FindBy(xpath = "//input[@id='id_body_13']")
		private static WebElement enter_Batch;

		@FindBy(xpath = "//input[@id='id_body_12']")
		private static WebElement enter_DebitACTxt;

		@FindBy(xpath = "//*[@id='id_body_12_input_image']/span")
		private static WebElement enter_DebitAC_ExpansionBtn;

		@FindBy(xpath = "//*[@id='id_body_12_input_settings']/span")
		private static WebElement enter_DebitAC_SettingsBtn;

		@FindBy(xpath = "//div[@id='id_header_12_customize_popup_container']")
		private static WebElement enter_DebitAC_Settings_Container;

		@FindBy(xpath = "//div[@id='id_header_12_customize_popup_footer']//input[1]")
		private static WebElement enter_DebitAC_Settings_StandardFieldsBtn;

		@FindBy(xpath = "//select[@id='id_header_12_customize_popup_standardfields_list']")
		private static WebElement enter_DebitAC_Settings_StandardFieldsBtn_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_header_12_customize_popup_standardfields_header']")
		private static WebElement enter_DebitAC_Settings_StandardFieldsBtn_HeaderTxt;

		@FindBy(xpath = "//select[@id='id_header_12_customize_popup_standardfields_alignment']")
		private static WebElement enter_DebitAC_Settings_StandardFieldsBtn_Alignmentdropdown;

		@FindBy(xpath = "//input[@id='id_header_12_customize_popup_standardfields_width']")
		private static WebElement enter_DebitAC_Settings_StandardFieldsBtn_WidthTxt;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[3]")
		private static WebElement enter_DebitAC_Settings_StandardFieldsBtn_OkBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[4]")
		private static WebElement enter_DebitAC_Settings_StandardFieldsBtn_CancelBtn;

		@FindBy(xpath = "//div[@id='id_header_12_search_container']//input[2]")
		private static WebElement enter_DebitAC_Settings_DeleteColumnBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[3]")
		private static WebElement enter_DebitAC_Settings_OkBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[4]")
		private static WebElement enter_DebitAC_Settings_CancelBtn;

		@FindBy(xpath = "//input[@id='id_body_39']")
		private static WebElement enter_CreditACTxt;

		@FindBy(xpath = "//*[@id='id_body_39_input_image']/span")
		private static WebElement enter_CreditAC_ExpansionBtn;

		@FindBy(xpath = "//*[@id='id_body_39_input_settings']/span")
		private static WebElement enter_CreditAC_SettingsBtn;

		@FindBy(xpath = "//div[@id='id_header_39_customize_popup_container']")
		private static WebElement enter_CreditAC_Settings_Container;

		@FindBy(xpath = "//div[@id='id_header_39_customize_popup_footer']//input[1]")
		private static WebElement enter_CreditAC_Settings_StandardFieldsBtn;

		@FindBy(xpath = "//select[@id='id_header_39_customize_popup_standardfields_list']")
		private static WebElement enter_CreditAC_Settings_StandardFieldsBtn_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_header_39_customize_popup_standardfields_header']")
		private static WebElement enter_CreditAC_Settings_StandardFieldsBtn_HeaderTxt;

		@FindBy(xpath = "//select[@id='id_header_39_customize_popup_standardfields_alignment']")
		private static WebElement enter_CreditAC_Settings_StandardFieldsBtn_Alignmentdropdown;

		@FindBy(xpath = "//input[@id='id_header_39_customize_popup_standardfields_width']")
		private static WebElement enter_CreditAC_Settings_StandardFieldsBtn_WidthTxt;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[3]")
		private static WebElement enter_CreditAC_Settings_StandardFieldsBtn_OkBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[4]")
		private static WebElement enter_CreditAC_Settings_StandardFieldsBtn_CancelBtn;

		@FindBy(xpath = "//div[@id='id_header_39_search_container']//input[2]")
		private static WebElement enter_CreditAC_Settings_DeleteColumnBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[3]")
		private static WebElement enter_CreditAC_Settings_OkBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[4]")
		private static WebElement enter_CreditAC_Settings_CancelBtn;

		@FindBy(xpath = "//input[@id='id_body_18']")
		private static WebElement enter_DebitTxt;

		@FindBy(xpath = "//input[@id='id_body_19']")
		private static WebElement enter_CreditTxt;

		@FindBy(xpath = "//textarea[@id='id_body_16777241']")
		private static WebElement enter_NonStandardJournalEntries_RemarksTxt;

		@FindBy(xpath = "//input[@id='id_body_536870916']")
		private static WebElement enter_Warehouse2Txt;

		@FindBy(xpath = "//*[@id='id_body_87_input_image']/span")
		private static WebElement enter_Warehouse2_ExpansionBtn;

		@FindBy(xpath = "//*[@id='id_body_87_input_settings']/span")
		private static WebElement enter_Warehouse2_SettingsBtn;

		@FindBy(xpath = "//div[@id='id_header_87_customize_popup_container']")
		private static WebElement enter_WareHouse2_Settings_Container;

		@FindBy(xpath = "//div[@id='id_header_87_customize_popup_footer']//input[1]")
		private static WebElement enter_WareHouse2_Settings_StandardFieldsBtn;

		@FindBy(xpath = "//select[@id='id_header_87_customize_popup_standardfields_list']")
		private static WebElement enter_WareHouse2_Settings_StandardFieldsBtn_Fielddropdown;

		@FindBy(xpath = "//input[@id='id_header_87_customize_popup_standardfields_header']")
		private static WebElement enter_WareHouse2_Settings_StandardFieldsBtn_HeaderTxt;

		@FindBy(xpath = "//select[@id='id_header_87_customize_popup_standardfields_alignment']")
		private static WebElement enter_WareHouse2_Settings_StandardFieldsBtn_Alignmentdropdown;

		@FindBy(xpath = "//input[@id='id_header_87_customize_popup_standardfields_width']")
		private static WebElement enter_WareHouse2_Settings_StandardFieldsBtn_WidthTxt;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[3]")
		private static WebElement enter_WareHouse2_Settings_StandardFieldsBtn_OkBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[4]")
		private static WebElement enter_WareHouse2_Settings_StandardFieldsBtn_CancelBtn;

		@FindBy(xpath = "//div[@id='id_header_87_search_container']//input[2]")
		private static WebElement enter_WareHouse2_Settings_DeleteColumnBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[3]")
		private static WebElement enter_WareHouse2_Settings_OkBtn;

		@FindBy(xpath = "//div[contains(@class,'panel')]//input[4]")
		private static WebElement enter_WareHouse2_Settings_CancelBtn;

		@FindBy(xpath = "//*[@id='ISOCurrencyCode_input_image']/span")
		private static WebElement isoCurrencyCodeComboboxButton;
		
		@FindBy(xpath = "//input[@id='sFileName1']")
		private static WebElement extraFields_DocumentTab_FileName;

		@FindBy(xpath = "//input[@id='biDocument1']")
		private static WebElement extraFields_DocumentTab_Document;

		@FindBy(xpath = "//tr[1]//td[5]//div[1]//div[1]//div[1]")
		private static WebElement extraFields_DocumentTab_UploadDocument;

		@FindBy(xpath = "//a[@id='SaveImg']")
		private static WebElement extraFields_DocumentTab_SaveDocument;

		@FindBy(xpath = "//i[@class='fa fa-eraser RemoveImg']")
		private static WebElement extraFields_DocumentTab_DeleteDocument;

		@FindBy(xpath = "//input[@id='iFileSize1']")
		private static WebElement extraFields_DocumentTab_FileSize;

		// Right Panel InfoSide Bar
		@FindBy(xpath = "//*[@id='id_transactionentry_infopanel_container']/div[1]/div[2]/span")
		private static WebElement infoSideBarCustomizeBtn;

		@FindBy(xpath = "//*[@id='id_transactionentry_infopanel_container']/div[1]/div[2]/span")
		private static WebElement customizeCloseBtn;

		@FindBy(xpath = "//div[@id='id_transactionentry_infopanel_customize_container']//li[@id='Dash_Save']")
		private static WebElement customizeSaveBtn;

		@FindBy(xpath = "//div[@id='id_transactionentry_infopanel_customize_container']//li[@id='Dash_Close']")
		private static WebElement customizeCancelBtn;

		@FindBy(xpath = "//input[@id='searchBoxTrans']")
		private static WebElement customizeSearchTxt;

		@FindBy(xpath = "//div[@id='Dashboard_Graph_panelID_Trans']//i[@class='icon-search searchicon']")
		private static WebElement customizeSearchBtn;

		@FindBy(xpath = "//div[@id='Dashboard_Graph_panelID_Trans']//i[@id='g']")
		private static WebElement graphBtn;

		@FindBy(xpath = "//div[@id='Dashboard_Graph_panelID_Trans']//li[@id='Newgraph']")
		private static WebElement newGraphOption;

		@FindBy(xpath = "//div[@id='Dashboard_Graph_panelID_Trans']//i[@id='r']")
		private static WebElement reportBtn;

		@FindBy(xpath = "//div[@id='Dashboard_Graph_panelID_Trans']//li[@id='NewReport']")
		private static WebElement newReportOption;

		@FindBy(xpath = "//div[@id='Dashboard_Graph_panelID_Trans']//i[@id='i']")
		private static WebElement infoPanelBtn;

		@FindBy(xpath = "//div[@id='Dashboard_Graph_panelID_Trans']//li[@id='NewInfopanel']")
		private static WebElement newInfoPanelOption;

		@FindBy(xpath = "//div[@id='Dashboard_Graph_panelID_Trans']//i[@id='w']")
		private static WebElement workFlowBtn;

		@FindBy(xpath = "//div[@id='Dashboard_Graph_panelID_Trans']//li[@id='NewWorkflow']")
		private static WebElement newWorkFlowOption;

		@FindBy(xpath = "//a[@id='DocumentInfoOption']")
		private static WebElement otherDashlets;

		@FindBy(xpath = "//*[@id='id_Dashlet0']")
		private static WebElement documentInfo;

		@FindBy(xpath = "//span[contains(@class,'icon-left-and-right-panel-icon icon-font6 no_padding_left_right')]")
		private static WebElement infoSideBarMinimizeExpandBtn;

		@FindBy(xpath = "//span[@class='icon-left-and-right-panel-icon icon-font6 no_padding_left_right']")
		private static WebElement infoSideBarExpandBtn;

		// Footer Section
		@FindBy(xpath = "//label[@id='id_transactionentry_footer_panel_summary_value_net']")
		private static WebElement netLabel;

		@FindBy(xpath = "//*[@id='id_transactionentry_footer_panel_summary_value_net']/span[2]")
		private static WebElement netAmount;

		// Login Page
		@FindBy(xpath = "//*[@id='txtUsername']")
		private static WebElement username;

		@FindBy(id = "txtPassword")
		private static WebElement password;

		@FindBy(id = "btnSignin")
		private static WebElement signIn;

		@FindBy(id = "ddlCompany")
		private static WebElement companyDropDownList;

		@FindBy(xpath = "//*[@id='mainHeader_MainLayout']/nav/div/ul/li[6]/a")
		private static WebElement userNameDisplay;

		@FindBy(xpath = "//*[@id='companyLogo']")
		private static WebElement companyLogo;

		@FindBy(xpath = "//*[@id='ulCompanyDetails_HomePage']/li[1]")
		private static WebElement companyName;

		@FindBy(xpath = "//*[@id='userprofile']/li/span[2]")
		private static WebElement logoutOption;

		@FindBy(xpath = "//input[@id='donotshow']")
		private static WebElement doNotShowCheckbox;

		@FindBy(xpath = "//span[@class='pull-right']")
		private static WebElement closeBtnInDemoPopupScreen;

		    
		// Server Error Elements
		@FindBy(xpath = "//div[@id='id_focus_msgbox_main']")
		private static WebElement serverErrorPopup;

		@FindBy(xpath = "//div[@id='id_focus_msgbox_main']/div[2]/button")
		private static WebElement serverErrorClose;


		 public static boolean checkErrorMessageIsDisplayingOrNotDisplaying() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
		   {     
			   
			 System.out.println("************************   checkErrorMessageIsDisplayingOrNotDisplaying Method Executes.............    ******************************");
			   
			 try 
			 { 
				boolean actValidationMessageIsEmptyDisplay      = validationConfirmationMessage.getText().isEmpty();
		        boolean expValidationMessageIsEmptyDisplay      = true;   
				
		        System.out.println("Error Message Is Displaying Or Not Displaying Value Actual : " + actValidationMessageIsEmptyDisplay + " Value Expected : " + expValidationMessageIsEmptyDisplay);
		        
				if (actValidationMessageIsEmptyDisplay == expValidationMessageIsEmptyDisplay)
				{
					//System.out.println("No Error Message Is Displaying On Opening Page,Click On Add, Edit, Delete Etc...... In Method : " + validationConfirmationMessage.getText());
					return true;
				} 
				else 
				{
					//System.err.println("Error Message Is Displaying On Opening Page, Click, Add, Edit, Delete Etc..... In Method : " + validationConfirmationMessage.getText());
					newErrorMessageCloseBtn.click();
					return false;
				}
				
			 } 
			 catch (NoSuchElementException nse)
			 {
				System.err.println("CATCH BLOCK : NO SUCH ELEMENT MESSAGE IS DISPLAY  :   "+ nse.getMessage());
				return false;
			 }
		  }



		
		
		
		// Customize Master In Accounts Master Page
		@FindBy(xpath = "//a[@href='#HeaderDetails_0']")
		private static WebElement accountsCustomizeGeneralHeaderDetailsTab;

		@FindBy(xpath = "//a[@href='#HeaderDetails_1']")
		private static WebElement accountsCustomizeSettingsHeaderDetailsTab;

		@FindBy(xpath = "//a[@href='#BodyDetails_1']")
		private static WebElement accountsCustomizeSettingsBodyDetailsTab;

		@FindBy(xpath = "//a[@href='#HeaderDetails_2']")
		private static WebElement accountsCustomizeDetailsHeaderDetailsTab;

		@FindBy(xpath = "//a[@href='#BodyDetails_2']")
		private static WebElement accountsCustomizeDetailsBodyDetailsTab;

		@FindBy(xpath = "//a[@href='#HeaderDetails_3']")
		private static WebElement accountsCustomizePrintLayoutHeaderDetailsTab;

		@FindBy(xpath = "//a[@href='#BodyDetails_3']")
		private static WebElement accountsCustomizePrintLayoutBodyDetailsTab;

		@FindBy(xpath = "//a[@href='#HeaderDetails_4']")
		private static WebElement accountsCustomizeTreeTabHeaderDetailsTab;

		@FindBy(xpath = "//a[@href='#BodyDetails_4']")
		private static WebElement accountsCustomizeTreeTabBodyDetailsTab;

		@FindBy(xpath = "//a[@href='#HeaderDetails_5']")
		private static WebElement accountCustomizeDocumetTabHeaderDetailsTab;

		@FindBy(xpath = "//a[@href='#BodyDetails_5']")
		private static WebElement accountCustomizeDocumetTabBodyDetailsTab;

		@FindBy(xpath = "//a[@href='#HeaderDetails_6']")
		private static WebElement accountCustomizeCreateTabHeaderDetailsTab;

		@FindBy(xpath = "//a[@href='#BodyDetails_6']")
		private static WebElement accountCustomizeCreateTabBodyDetailsTab;

		@FindBy(xpath = "//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[9]/td[1]/i[1]")
		private static WebElement accountSettingsHeaderEditNinethRowBtn;
		
		@FindBy(xpath = "//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[10]/td[1]/i[1]")
		private static WebElement accountSettingsHeaderEditTenthRowBtn;
		
		@FindBy(xpath="//*[@id='navigation_menu']/li/a/div/div")
		private static List<WebElement> menusList;
		
		@FindBy(xpath="//*[@id='navigation_menu']/li[1]/ul/li/a/span")
		private static List<WebElement> homeMenuList;
		
		@FindBy(xpath="//*[@id='navigation_menu']/li[1]/ul/li[3]/ul/li/a/span")
		private static List<WebElement> mastersMenuList;
		
		@FindBy(xpath="//*[@id='ulCommonlyUsedRibbon']/li/span")
		private static List<WebElement> mastersRibbonControlList;
		
		@FindBy(xpath ="//*[@id='ulRibbonControl']/li/span")
		private static List<WebElement> ribbonControlExpandList;
		
		@FindBy(xpath = "//span[@id='updateButton']")
		private static WebElement updateBtn;
		
		// Home Menu
		@FindBy(xpath = "//*[@id='1']/div/span")
		public static WebElement homeMenu;

		//Financial Menu
		@FindBy (xpath="//span[@class='icon-financial icon-font1']")
	    private static WebElement financialMenu;
		
		//Inventory Menu
		/*@FindBy (xpath="//span[@class='icon-inventory icon-font1']")
	    private static WebElement inventoryMenu;
		*/
		//Fixed Assets Menu
		@FindBy (xpath="//a[@id='1127']//span[contains(text(),'Fixed Assets')]")
	    private static WebElement fixedAssetsMenu;
		
		//Production Menu
		@FindBy (xpath="//span[@class='icon-production icon-font1']")
	    private static WebElement productionMenu;
		
		//Point Of Sale Menu
		@FindBy (xpath="//div[contains(text(),'Point of Sale')]")
	    private static WebElement pointOfSaleMenu;
		
		//Quality Of Control Menu
		@FindBy (xpath="//span[@class='icon-quotationanalysis icon-font1']")
	    private static WebElement qualityControlMenu;
		
		//Settings Menu
		@FindBy (xpath="//span[@class='icon-settings icon-font1']")
		private static WebElement SettingsMenu;
		
		  //Company Menu
				@FindBy (xpath="//*[@id='2']/span")
			    private static WebElement companyMenu;

			    //Security Menu	
				@FindBy (xpath="//*[@id='6']/span")
			    private static WebElement securityMenu;

		        //Masters Menu	
				@FindBy (xpath="//*[@id='1000']/span")
			    private static WebElement mastersMenu;

		        //DataManagement menu
				@FindBy (xpath="//*[@id='25']/span")
			    private static WebElement dataMangementMenu;

		        //Utilities Menu		
				@FindBy (xpath="//*[@id='20']/span")
			    private static WebElement utilities;
			
				@FindBy (xpath="//*[@id='114']/span")
			    private static WebElement repostVoucher;
				
				
				
				//In Masters Menu
				
				// Accounts
				@FindBy(xpath = "//a[@id='1104']//span[contains(text(),'Account')]")
				public static WebElement accounts;

		        @FindBy (xpath="//*[@id='220']/span")
				private static WebElement currency;

		        @FindBy (xpath="//*[@id='221']/span")
				private static WebElement ItemMenu;

		        
		        @FindBy (xpath="//a[@id='1105']//span[contains(text(),'Item')]")
				private static WebElement Item_ItemMenu;
		        
		      
		        
		        @FindBy (xpath="//*[@id='224']/span")
				private static WebElement mrp;

		        @FindBy (xpath="//a[@id='225']//span[contains(text(),'QC')]")
				private static WebElement qc;

		        @FindBy (xpath="//*[@id='251']/span")
				private static WebElement maintainance;

		        @FindBy (xpath="//*[@id='8001']/span")
				private static WebElement pointOfSale;

		        @FindBy (xpath="//*[@id='1106']/span")
				private static WebElement departments;

				@FindBy (xpath="//*[@id='1107']/span")
				private static WebElement warehouse;

				@FindBy (xpath="//*[@id='1108']/span")
				private static WebElement costcenter;

				@FindBy (xpath="//*[@id='1109']/span")
				private static WebElement location;

				@FindBy (xpath="//*[@id='1110']/span")
				private static WebElement region;

				@FindBy (xpath="//*[@id='1111']/span")
				private static WebElement country;

				@FindBy (xpath="//*[@id='1112']/span")
				private static WebElement state;

				@FindBy (xpath="//*[@id='1113']/span")
				private static WebElement city;

				@FindBy (xpath="//*[@id='1115']/span")
				private static WebElement bins;

				@FindBy (xpath="//*[@id='1126']/span")
				private static WebElement insurance;

				@FindBy (xpath="//*[@id='1127']/span")
				private static WebElement fixedassets;
				
				@FindBy (xpath="//*[@id='1161']/span")
				private static WebElement taxCode;
				
		        @FindBy (xpath="//*[@id='1103']/span")
				private static WebElement payroll;
			
		        @FindBy (xpath="//*[@id='8002']/span")
		  		private static WebElement schemes;
				
		        @FindBy(xpath="//ul[@id='preferenceUL']//a[@id='37']")
				private static WebElement  PronghornBtn;
				
				

				@FindBy(xpath="//*[@id='preferenceUL']/li/a")
				private static List<WebElement>  preferenceList;
				

				
				@FindBy(xpath="//input[@id='chkARTag']")
				private static WebElement arTagChkBox; 
				
				@FindBy(xpath="//input[@id='optARTag']")
				private static WebElement arTagTxt;
				
				@FindBy(xpath="//input[@id='chkARSalesAC']")
				private static WebElement salesAccountChkBox;
				
				@FindBy(xpath="//input[@id='chkARProduct']")
				private static WebElement  arItemChkBox;
				
				@FindBy(xpath="//input[@id='chkEnableCreditLimitCheck']")
				private static WebElement enableCreditLimitCheckChkBox;
				
				@FindBy(xpath="//input[@id='chkIncludePendingSalesOrderValueInCreditLimitCheck']")
				private static WebElement includePendingSalesOrderChkBox;
				
				@FindBy(xpath="//input[@id='chkCheckcreditlimitinorders']")
				private static WebElement checkCreditLimitInOrdersChkBox; 
				
				@FindBy(xpath="//input[@id='chkDefineCreditLimitByDepartment']")
				private static WebElement defineCreditLimitByDepartmentChkBox;
				
				@FindBy(xpath="//input[@id='chkAllowCreditLimitAuthorizationMappingToCustomer Group']")
				private static WebElement allowCreditLimitAuthorizationmappingChkBox;
				
				@FindBy(xpath="//input[@id='dvExceed_0']")
				private static WebElement  whenCreditLimitexceeded_WarnAndAllowRadio;
				
				@FindBy(xpath="//input[@id='dvExceed_1']")
				private static WebElement whenCreditLimitexceeded_StopRadio;
				
				@FindBy(xpath="//input[@id='dvExceed_2']")
				private static WebElement  RequestCreditLimitIncreaseRadio;

				@FindBy(xpath="//input[@id='chkPickCreditDayFromSalesAC']")
				private static WebElement pickCreditDayFromChkBox; 
				
				@FindBy(xpath="//input[@id='dvPickCreditDayFromSalesAC_0']")
				private static WebElement customerVendorACRadio;
				
				@FindBy(xpath="//input[@id='dvPickCreditDayFromSalesAC_1']")
				private static WebElement paymentsTermsRadio;
				
				@FindBy(xpath="//input[@id='dvPickCreditDayFromSalesAC_2']")
				private static WebElement  salesACRadio;
				
				@FindBy(xpath="//input[@id='chkAPTag']")
				private static WebElement apTagChkBox; 
				
				@FindBy(xpath="//input[@id='optAPTag']")
				private static WebElement apTagTxt;
				
				@FindBy(xpath="//input[@id='chkAPPurchaseAC']")
				private static WebElement purchaseAccountChkBox;
				
				@FindBy(xpath="//input[@id='chkAPProduct']")
				private static WebElement  apItemChkBox;
				
				@FindBy(xpath="//input[@id='chkDependsonmaintainbillwiseforAR_AP']")
				private static WebElement dependsOnMaintainBillWiseForARAPChkBox;
				
				@FindBy(xpath="//input[@id='chkMaintainARTransactionCurrencies']")
				private static WebElement maintainARTransactionCurrenciesChkBox;
				
				@FindBy(xpath="//input[@id='chkMaintainAPTransactionCurrencies']")
				private static WebElement maintainAPTransactionCurrenciesChkBox; 
				
				@FindBy(xpath="//input[@id='chkInputNarrationInReferences']")
				private static WebElement inputNarrationInReferenceChkBox;
				
				@FindBy(xpath="//input[@id='chkEnableOverDueCheck']")
				private static WebElement enableOverDueCheckChkBox;
				
				@FindBy(xpath="//input[@id='chkCreditDaysInsteadOfDueDateForLimitCheck']")
				private static WebElement  useCreditDaysInsteadOfDueDateChkBox;
				
				@FindBy(xpath="//input[@id='dvOverdueExceed_0']")
				private static WebElement whenCreditDaysLimitexceeded_WarnAndAllowRadio;

				@FindBy(xpath="//input[@id='dvOverdueExceed_1']")
				private static WebElement whenCreditDaysLimitexceeded_StopRadio;

				
				@FindBy(xpath="//input[@id='chkAddFreeItems']")
				private static WebElement addfreeItemsonaNewLineChkBox; 
				
				
				
				@FindBy(xpath="//input[@id='chkCalDueDtLRDt']")
				private static WebElement calculatedueDateFromLRDateChkBox;
				
				@FindBy(xpath="//input[@id='chkCreateCustomerProfilefields']")
				private static WebElement createCustomerProfileFieldsChkBox;
				
				@FindBy(xpath="//input[@id='chkDialogbasedentryinvouchers']")
				private static WebElement  dialogbasedentryinvouchersChkBox;
				
				@FindBy(xpath="//input[@id='chkDontRefreshDescACProdDoc']")
				private static WebElement dontRefreshDescACProdDocChkBox; 
				
				@FindBy(xpath="//input[@id='chkIncUnCommitTransRpts']")
				private static WebElement includeUnCommitedTransactionChkBox;
				
				@FindBy(xpath="//input[@id='chkMntnLnksStkTransfer']")
				private static WebElement maintainLinksforOnlyOneSideChkBox;
				
				@FindBy(xpath="//input[@id='chkShowStatus']")
				private static WebElement  showStatusMessageInPopUpChkox;
				
				@FindBy(xpath="//input[@id='chkPrefixLocationCode']")
				private static WebElement prefixLocationCodeWhileImportingChkBox; 
				
				@FindBy(xpath="//input[@id='chkShowTransactionInFifo']")
				private static WebElement showTransactionDateInFifo;
				
				@FindBy(xpath="//input[@id='chkEnableLocalCurrency']")
				private static WebElement enableLocalCurrencyChkBox;
				
				@FindBy(xpath="//input[@id='optLocalCurrency']")
				private static WebElement  localCurrencyTxt;
				
				@FindBy(xpath="//input[@id='chkDontShowOpeningBalLedger']")
				private static WebElement dontShowOpeningBalLedgerChkBox; 
				
				@FindBy(xpath="//input[@id='chkDontStoreDatesEntries']")
				private static WebElement dontStoreDatesEntriesChkBox;
				
				@FindBy(xpath="//input[@id='chkShowexchangeratedifferenceinledger']")
				private static WebElement showexchangeratedifferenceinledgerChkBox;
				
				@FindBy(xpath="//input[@id='chkEnableHijriDate']")
				private static WebElement  enableHijriDateChkBox;
				
				@FindBy(xpath="//input[@id='chkEnableprofitabilitycheckbyproduct']")
				private static WebElement enableprofitabilitycheckbyproductChkBox; 
				
				@FindBy(xpath="//input[@id='chkDonotLoadDocInExlusivemode']")
				private static WebElement donotLoadDocInExlusivemodeChkBox;
				
				@FindBy(xpath="//input[@id='chkCreateMasterinTransactionEntry']")
				private static WebElement createMasterinTransactionEntryChkBox;
				
				@FindBy(xpath="//input[@id='chkNextVoucherNoAftSaveVoucher']")
				private static WebElement  stayOnSameVoucherNumberAfterDeleteingChkBox;
				

				@FindBy(xpath="//input[@id='chkDonotModifyDoc']")
				private static WebElement doNotModifiedDocumentIfTriggerIsRaised;
				
				@FindBy(xpath="//input[@id='chkOpenSearchIfMasterNotFound']")
				private static WebElement openSearchIfMasterNotFoundChkBox; 
				
				@FindBy(xpath="//input[@id='chkloadDateBasesonLastSavedVoucher']")
				private static WebElement loadDateBasesonLastSavedVoucherChkBox;
				
				@FindBy(xpath="//input[@id='idfiltercurrencyBasedondepartment']")
				private static WebElement filtercurrencyBasedondepartmentChkBox;
				
				@FindBy(xpath="//input[@id='optCalendar']")
				private static WebElement  defaultCalanderTxt;
				 
				
				@FindBy(xpath="//select[@id='cmbImageformat']")
				private static WebElement imageFormatDropdown;
				
				@FindBy(xpath="//select[@id='cmbNumericSep']")
				private static WebElement numericSeperatorDropdown1;
				
				@FindBy(xpath="//select[@id='cmbFormula']")
				private static WebElement  numericSeperatorDropdown2;
				
				@FindBy(xpath="//input[@id='rdoDocTypeOrClass_0']")
				private static WebElement docTypeRadio; 
				
				@FindBy(xpath="//input[@id='rdoDocTypeOrClass_1']")
				private static WebElement docClassRadio;
				
				@FindBy(xpath="//td[@id='gPrefMisExcludeVoucher_col_1-0']")
				private static WebElement gridFirstRowIndexBtn;
				
				@FindBy(xpath="//td[@id='gPrefMisExcludeVoucher_col_2-0']")
				private static WebElement  gridSecondRowIndexBtn;
				
				@FindBy(xpath="//td[@id='gPrefMisExcludeVoucher_col_1-1']")
				private static WebElement gridVoucherFirstRow; 
				
				@FindBy(xpath="//td[@id='gPrefMisExcludeVoucher_col_2-1']")
				private static WebElement gridVoucherSecondRow;
				
				@FindBy(xpath="//input[@id='rdoSendMailClass_0']")
				private static WebElement sendAsAttachmentRadio;
				
				@FindBy(xpath="//input[@id='rdoSendMailClass_1']")
				private static WebElement  sendASBodyRadio;
				
				@FindBy(xpath="//input[@id='optMisExcludeVoucher']")
				private static WebElement  VoucherTxt;
				
		
		

				
					@FindBy (xpath="//*[@id='26']/span")
					private static WebElement backup;
						
					@FindBy (xpath="//*[@id='27']/span")
					private static WebElement restore;
					
					@FindBy (xpath="//*[@id='111']/span")
					private static WebElement reindex;
					
					@FindBy (xpath="//*[@id='115']/span")
					private static WebElement eraseAll;
					
					@FindBy (xpath="//*[@id='216']/span")
					private static WebElement lockAndUnlock;
					
					
				@FindBy (xpath="//button[@class='btn Fbutton']")
				private static WebElement folderpathExpandBtn;
				
				@FindBy (xpath="//div[@id='RestoreControls']//div[1]//div[1]")
				private static WebElement restoreCompanyBtn;
				
				@FindBy (xpath="//button[@class='Fbutton'][contains(text(),'Yes')]")
				private static WebElement overRideYesBtn;
				
				@FindBy (xpath="//button[@class='Fbutton'][contains(text(),'No')]")
				private static WebElement overRideNoBtn;
				
				@FindBy (xpath="//button[@class='Fbutton'][contains(text(),'Cancel')]")
				private static WebElement overRideCancelBtn;


				@FindBy(xpath="//*[@id='divBtnGroup1']/div/a")
				private static WebElement  deperciationTab;
				
				@FindBy(xpath="//*[@id='divBtnGroup2']/div/a")
				private static WebElement  accountingLinksTab;
				
				
				@FindBy(xpath="//*[@id='divBtnGroup3']/div/a")
				private static WebElement  insuranceDetailsTab;
				
				@FindBy(xpath="//*[@id='divBtnGroup4']/div/a")
				private static WebElement  maintanceTab;
				
				@FindBy(xpath="//input[@id='iDepreciationInYears']")
				private static WebElement  depaerciationYearsTxt;
				
				@FindBy(xpath="//input[@id='iDateOfCommencement']")
				private static WebElement  dateOfCommencementTxt;
				
				@FindBy(xpath="//select[@id='iDepreciationMethod']")
				private static WebElement  depreciationMethodDrpdwn;
				
				@FindBy(xpath="//select[@id='iUnitOfUsage']")
				private static WebElement  unitsOfUsage;
				

				
		public static boolean checkLoginAndRestoreOptionsCompanyAndLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException 
		{
			
			LoginPage lp=new LoginPage(getDriver()); 
			
		    String unamelt="su";
		  
		    String pawslt="su";
		    
		    lp.enterUserName(unamelt);
		    
		    lp.enterPassword(pawslt);
		  
		     
		     lp.clickOnSignInBtn();
		    
		     Thread.sleep(8000);
		     
		    String actUserInfo1=userNameDisplay.getText();
				
			System.out.println("User Info  : "+actUserInfo1);
			
			System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());
			
			
			String expUserInfo1           ="SU";
			String expLoginCompanyName1   ="Automation Options ";
		     
		   
			System.out.println("UserInfo1             : ."+actUserInfo1            +". Value Expected : "+expUserInfo1);
			
			
			if(actUserInfo1.contains(expUserInfo1))
			{
		
				return true;
			}
			else
			{
				return false;
			}
		}


				

	public boolean checkFixedAssetsTabUnderPreference() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		
		System.out.println("********************* checkPreferencesOption Method Executes.............   *****************************");
			
            
		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();
			
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigureTransactionBtn));
		ConfigureTransactionBtn.click();
		
		
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FixedAssetsBtn));
 		
        
        JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("arguments[0].scrollIntoView();", FixedAssetsBtn);
		
        
 		boolean FixedAssets  =FixedAssetsBtn.isDisplayed();
 		
 		String actFixedAssetsBtn=Boolean.toString(FixedAssets);
 		String expFixedAssetsBtn  = "true";
 		
 		
 		System.out.println("*********************************checkFixedAssetsTabUnderPreference**************************");
 		
 		System.out.println("MiselleanousBtn  : "+actFixedAssetsBtn+" Value Expected : "+expFixedAssetsBtn);
		
		
 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FixedAssetsBtn));
 		FixedAssetsBtn.click();
 		
 		if (actFixedAssetsBtn.equalsIgnoreCase(expFixedAssetsBtn))
 		{
 			System.out.println(" Test Pass: Displayed Fixed Assets Menu under Setting");
			return true;
		}
 		else
 		{
 			System.out.println(" Test Fail: Displayed Fixed Assets Menu under Setting");
 			return false;
		}
	 }
 		
	
	public boolean checkSavingGroupAsBasedOnUsage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		
		System.out.println("*******************checkSavingGroupAsBasedOnUsage   Method Executes***************");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		
		Thread.sleep(2000);
        JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("arguments[0].scrollIntoView();", fixedAssetsMenu);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssetsMenu));
		fixedAssetsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAddGroupBtn));
		masterAddGroupBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		
		nameTxt.sendKeys("FixedAssetsBasedOnUsage");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.sendKeys("FixedAssetsBasedOnUsage");
		
		Thread.sleep(2000);
		codeTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deperciationTab));
		deperciationTab.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(depreciationMethodDrpdwn));
		depreciationMethodDrpdwn.click();
		
		Select s = new Select(depreciationMethodDrpdwn);
		
		s.selectByVisibleText("Based on Usage");
		depreciationMethodDrpdwn.sendKeys(Keys.TAB);
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accCreationPageSaveBtn));
		accCreationPageSaveBtn.click();
		
		String expMessage="Saved Successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			System.out.println("Test Pass : Conditions As Expected ");
			return true;
		}
		else
		{
			System.out.println("Test Fail : Conditions As Expected ");
			return false;
		}
	}
	
	
	
	public boolean checkSavingGroupAsStraightLine() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		
		System.out.println("*******************checkSavingGroupAsBasedOnUsage   Method Executes***************");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.sendKeys("FixedAssetsStraightLine");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.sendKeys("FixedAssetsStraightLine");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deperciationTab));
		deperciationTab.click();
		
		Thread.sleep(2000);
		
		Select s = new Select(depreciationMethodDrpdwn);
		s.selectByVisibleText("Straigh tLine");
		depreciationMethodDrpdwn.sendKeys(Keys.TAB);
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accCreationPageSaveBtn));
		accCreationPageSaveBtn.click();
		
		String expMessage="Saved Successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			System.out.println("Test Pass : Conditions As Expected ");
			return true;
		}
		else
		{
			System.out.println("Test Fail : Conditions As Expected ");
			return false;
		}
	}
	
	
	
	

	public boolean checkDeletionGroupInFixedAssets() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		
		System.out.println("*******************checkDeletionGroupInFixedAssets   Method Executes***************");
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.sendKeys("GroupForDeletion");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.sendKeys("GroupForDeletion");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accCreationPageSaveBtn));
		accCreationPageSaveBtn.click();
		
		String expMessage="Saved Successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MasterCloseBtn));
		MasterCloseBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterRow1Chkbox));
		masterRow1Chkbox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
		deleteBtn.click();
		
 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelete));
 		clickOnOkInDelete.click();
		
 		System.out.println("*******************Click on Ok ***************8");
 		
        String expDeletionMessage="Record Deleted Successfully.";
		
		String actDeletionMessage=checkValidationMessage(expDeletionMessage);
		
		System.out.println("actDeletionMessage   :"+actDeletionMessage  +" Value Expected  "+expDeletionMessage);
		
		if(actMessage.equalsIgnoreCase(expMessage) &&actDeletionMessage.equalsIgnoreCase(expDeletionMessage) )
		{
			System.out.println("Test Pass : Conditions As Expected ");
			return true;
		}
		else
		{
			System.out.println("Test Fail : Conditions As Expected ");
			return false;
		}
		
	}
	
	
	
	@FindBy(xpath="//*[@id='LandingGridBody']/tr[1]/td[8]")
	public static WebElement masterRow1Chkbox;
	
	
	@FindBy(xpath="//*[@id='GenerateNewMasterModel']/div/div[1]/div[2]/div/ul/li[2]/button[5]")
	public static WebElement MasterCloseBtn;
	
	@FindBy(xpath="//i[@class='icon-font6 icon-add-group']")
	public static WebElement masterAddGroupBtn; 
	 
	@FindBy(xpath="//i[@class='icon-font6 icon-edit']")
	public static WebElement masterEditBtn; 

	
	public boolean checkConditionsinFixedAssetsScreen() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		System.out.println("********************* checkConditionsinFixedAssetsScreen Method Executes.............   *****************************");
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigureTransactionBtn));
		ConfigureTransactionBtn.click();
		
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FixedAssetsBtn));
		FixedAssetsBtn.click();
		
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FixedAssets_FixedAssetsLocationDrrpdwn));
		FixedAssets_FixedAssetsLocationDrrpdwn.click();
			
			
		Select oSelect = new Select(FixedAssets_FixedAssetsLocationDrrpdwn);
		
		
		List<WebElement> elementCount = oSelect.getOptions();

		
		int actLocationList = elementCount.size();
		
		String actLocationCount=Integer.toString(actLocationList);
		String expAccountTypeSize = "69";
		
		System.err.println("actLocationList  : "+actLocationList);
		
		
		if (FixedAssets_CreationOfFixedAssetInMandatoryFromPurchaseChkbox.isSelected()==false)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FixedAssets_CreationOfFixedAssetInMandatoryFromPurchaseChkbox));
			FixedAssets_CreationOfFixedAssetInMandatoryFromPurchaseChkbox.click();
			
		}
		

		if (FixedAssets_AlwaysDefineDepreciationRateAsPercentageChkbox.isSelected()==true)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FixedAssets_AlwaysDefineDepreciationRateAsPercentageChkbox));
			FixedAssets_AlwaysDefineDepreciationRateAsPercentageChkbox.click();
			
		}
		
		
		if (FixedAssets_IncludeAssetsInPurchaseAccountSelectionChkbox.isSelected()==false)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FixedAssets_IncludeAssetsInPurchaseAccountSelectionChkbox));
			FixedAssets_IncludeAssetsInPurchaseAccountSelectionChkbox.click();
			
		}
		
		if (FixedAssetsPostDepreciationDefaultRadioBtn.isSelected()==false)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FixedAssetsPostDepreciationDefaultRadioBtn));
			FixedAssetsPostDepreciationDefaultRadioBtn.click();
			
		}
		
		
		if (PostentryduringTransferChkbox.isSelected()==false)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PostentryduringTransferChkbox));
			PostentryduringTransferChkbox.click();
			
		}
		
		if (PostentryduringDisposalChkbox.isSelected()==false)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PostentryduringDisposalChkbox));
			PostentryduringDisposalChkbox.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PostentryduringTransferTxt));
		PostentryduringTransferTxt.click();
		PostentryduringTransferTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		PostentryduringTransferTxt.sendKeys("Fixed Asset Depreciation Voucher");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PostentryduringDisposalTxt));
		PostentryduringDisposalTxt.click();
		PostentryduringDisposalTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		PostentryduringDisposalTxt.sendKeys("Fixed Asset Depreciation Voucher");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PostDepreciationToTxt));
		PostDepreciationToTxt.click();
		PostDepreciationToTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		PostDepreciationToTxt.sendKeys("Fixed Asset Depreciation Voucher");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UpdateBtn));
		UpdateBtn.click();
		
		getWaitForAlert();
		
		String actAlertMgs  =getAlert().getText();
		String expAlertMgs  ="Do you want to save the changes?";
		
        Thread.sleep(2000);
        
        getAlert().accept();
		
		String expMessage="Data saved Successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("AlertMessage     : "+actAlertMgs    +" Value Expected : "+expAlertMgs);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
		CloseBtn.click();
		
		if(actAlertMgs.equalsIgnoreCase(expAlertMgs) && 
				actMessage.equalsIgnoreCase(expMessage) && actLocationCount.equalsIgnoreCase(expAccountTypeSize))
		{
			System.out.println("Test Pass : Conditions As Expected ");
			return true;
		}
		else
		{
			System.out.println("Test Fail : Conditions As Expected ");
			return false;
		}
	 }
 		
	@FindBy(xpath = "//input[@id='optidPostDepreciationTo']")
	private static WebElement PostDepreciationToTxt;
	
	@FindBy(xpath="//*[@id='optidOptionPostentryduringTransfer']")
	public static WebElement PostentryduringTransferTxt; 

	@FindBy(xpath="//input[@id='optidOptionPostentryduringDisposal']")
	public static WebElement PostentryduringDisposalTxt; 

	
	public boolean checkSavedOptionsinFixedAssetsScreen() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		
		System.out.println("********************* checkSavedOptionsinFixedAssetsScreen Method Executes.............   *****************************");
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigureTransactionBtn));
		ConfigureTransactionBtn.click();
		
		Thread.sleep(1000);
		
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FixedAssetsBtn));
 		
        
        JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("arguments[0].scrollIntoView();", FixedAssetsBtn);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FixedAssetsBtn));
		FixedAssetsBtn.click();
			
		
		boolean CreationOfFixedAssetInMandatoryFromPurchaseChkbox=FixedAssets_CreationOfFixedAssetInMandatoryFromPurchaseChkbox.isSelected();
		
		String actCreationOfFixedAssetInMandatoryFromPurchaseChkbox=Boolean.toString(CreationOfFixedAssetInMandatoryFromPurchaseChkbox);
		String expCreationOfFixedAssetInMandatoryFromPurchaseChkbox="true";

		boolean IncludeAssetsInPurchaseAccountSelectionChkbox=FixedAssets_IncludeAssetsInPurchaseAccountSelectionChkbox.isSelected();
		
		String actIncludeAssetsInPurchaseAccountSelectionChkbox=Boolean.toString(IncludeAssetsInPurchaseAccountSelectionChkbox);
		String expIncludeAssetsInPurchaseAccountSelectionChkbox="true";
		
		

		boolean fixedAssetsPostDepreciationDefaultRadioBtn=FixedAssetsPostDepreciationDefaultRadioBtn.isSelected();
		
		String actFixedAssetsPostDepreciationDefaultRadioBtn=Boolean.toString(fixedAssetsPostDepreciationDefaultRadioBtn);
		String expFixedAssetsPostDepreciationDefaultRadioBtn="true";
		
		
        String actpostEntryTransfer=PostentryduringTransferTxt.getAttribute("value");
		String exppostEntryTransfer="Fixed Asset Depreciation Voucher";
        
        String actPostentryduringDisposalTxt=PostentryduringDisposalTxt.getAttribute("value");
		String expPostentryduringDisposalTxt="Fixed Asset Depreciation Voucher";
        
        
        String actPostDepreciationToTxt=PostDepreciationToTxt.getAttribute("value");
		String expPostDepreciationToTxt="Fixed Asset Depreciation Voucher";
        
				
        System.out.println(" CreationOfFixedAsset     : "+actCreationOfFixedAssetInMandatoryFromPurchaseChkbox    +" Value Expected : "+expCreationOfFixedAssetInMandatoryFromPurchaseChkbox);
        System.out.println("IncludeAssetsInPurchase   : "+actIncludeAssetsInPurchaseAccountSelectionChkbox        +" Value Expected : "+expIncludeAssetsInPurchaseAccountSelectionChkbox);
        System.out.println("DefaultRadioBtn           : "+actFixedAssetsPostDepreciationDefaultRadioBtn           +" Value Expected : "+expFixedAssetsPostDepreciationDefaultRadioBtn);
	    
        System.out.println("actpostEntryTransfer      :"+actpostEntryTransfer                +" Value Expected  : "+exppostEntryTransfer);
        System.out.println("actPostentryduringDisTxt  :"+actPostentryduringDisposalTxt       +" Value Expected  : "+expPostentryduringDisposalTxt);
        System.out.println("actPostDepreciationToTxt      :"+actPostDepreciationToTxt        +" Value Expected  : "+expPostDepreciationToTxt);
				
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
		CloseBtn.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
		userNameDisplay.click();
			  
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();
		
		
		if(actCreationOfFixedAssetInMandatoryFromPurchaseChkbox.equalsIgnoreCase(expCreationOfFixedAssetInMandatoryFromPurchaseChkbox) && 
				actIncludeAssetsInPurchaseAccountSelectionChkbox.equalsIgnoreCase(expIncludeAssetsInPurchaseAccountSelectionChkbox)
				&& actFixedAssetsPostDepreciationDefaultRadioBtn.equalsIgnoreCase(expFixedAssetsPostDepreciationDefaultRadioBtn) && 
				actPostDepreciationToTxt.equalsIgnoreCase(expPostDepreciationToTxt) && actPostentryduringDisposalTxt.equalsIgnoreCase(expPostentryduringDisposalTxt) && 
				actpostEntryTransfer.equalsIgnoreCase(exppostEntryTransfer))
		{
			System.out.println("Test Pass : Conditions As Expected ");
			return true;
		}
		else
		{
			System.out.println("Test Fail : Conditions As Expected ");
			return false;
		}
	 }
 		
 		
	
    
@FindBy(xpath="//*[@id='61']/span")
private static WebElement  financialsTransactionMenu; 	


@FindBy(xpath="//*[@id='2007']/span")
private static WebElement  financialsTransactionsPurchaseMenu; 

@FindBy(xpath="//*[@id='2053']/span")
private static WebElement  purchaseVouchersVat;
	
@FindBy(xpath="//*[@id='2058']")
private static WebElement  pvvatViewMenu;

@FindBy(xpath="//input[@id='id_body_16777313']")
private static WebElement  pvvanterPurVouVATaxcode;

@FindBy(xpath="//*[@id='id_body_12_table_data_body']/tr")
private static List<WebElement>  pvvatPurchaseAccList;


	public boolean checkPurchaseAccountListAfterEnableFixedAssetsOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		
		System.out.println("********************* checkPurchaseAccountListAfterEnableFixedAssetsOption Method Executes.............   *****************************");
			
		
		System.out.println("Login in Company After Options Enabled ");
		
        LoginPage lp=new LoginPage(getDriver()); 
		
	    String unamelt="su";
	  
	    String pawslt="su";
	      
	    lp.enterUserName(unamelt);
	    
	    lp.enterPassword(pawslt);
	    
        lp.clickOnSignInBtn();
		
		
		Thread.sleep(5000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
		financialMenu.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		financialsTransactionsPurchaseMenu.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
		purchaseVouchersVat.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		
		checkUserFriendlyMessage();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Warehouse2Txt));
		enter_Warehouse2Txt.click();
		enter_Warehouse2Txt.sendKeys("HYDERABAD");
		Thread.sleep(2000);
		enter_Warehouse2Txt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys("BR COGS ITEM");
		Thread.sleep(3000);
		enter_ItemTxt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvvanterPurVouVATaxcode));
		pvvanterPurVouVATaxcode.sendKeys(Keys.TAB);
		
		enter_AccountTxt.click();
		enter_AccountTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME,Keys.SPACE);
		
		
		ArrayList<String>pvvatPurchaseAccListArray=new ArrayList<String>();
		
		int pvvatPurchaseAccListCount=pvvatPurchaseAccList.size();
		
		System.out.println(" pvvatPurchaseAccListCount  :  "+pvvatPurchaseAccListCount);
		
		for (int i = 0; i < pvvatPurchaseAccListCount; i++)
		{
			String data=pvvatPurchaseAccList.get(i).getText();
			pvvatPurchaseAccListArray.add(data);
		}
		
		String actpvvatPurchaseAccList=pvvatPurchaseAccListArray.toString();
		String exppvvatPurchaseAccList="[BR COGS ACC INV, COGS POSTING ACC, Cost of goods sold - Computers, Cost of goods sold - Electronics, Cost of goods sold - HA, EXCESS COGS POSTING ACC, FIFO COGS ACC INV, Fixed Deposit Investment, Inventory  Trade, Inventory FG, Inventory RM, Journal Entries Control A/C, Opening Balances Control A/C, Profit/Loss A/C, Purchase, SHORTAGE COGS POSTING ACC, STD RATE COGS ACC INV, VAT INPUT, WA COGS ACC INV]";
		
		System.err.println("  Actual    pvvatPurchaseAccList    : "+actpvvatPurchaseAccList);
		System.err.println("  Expected  pvvatPurchaseAccList    : "+exppvvatPurchaseAccList);
		
		enter_AccountTxt.sendKeys(Keys.TAB);
		
		if (actpvvatPurchaseAccList.equalsIgnoreCase(exppvvatPurchaseAccList) )
		{
			return true;
		} 
		else
		{
			return false;
		}
		
	}
	
    @FindBy(xpath="//div[contains(text(),'Purchase Account')]")
	private static WebElement voucherBodyPurchaseAccountName;	
		  
    @FindBy(xpath="//*[@id='btnMasterSaveClick'][2]")
	public static WebElement vatCreationPageSaveBtn;
	  
    @FindBy(xpath="//*[@id='btnMasterSaveClick'][1]")
   	public static WebElement accCreationPageSaveBtn;
    
    @FindBy(xpath="//div[contains(text(),'Item')]")
	private static WebElement voucherBodyItemName;
	
	
	public boolean checkCreationOfPurchaseAccountAndItemFromPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		System.out.println("********************* chekCreationOfPurchaseAccountFromPurchaseVoucherNew Method Executes.............   *****************************");
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherBodyPurchaseAccountName));
		voucherBodyPurchaseAccountName.click();
		
		
		getAction().doubleClick(voucherBodyPurchaseAccountName).build().perform();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.sendKeys("Fixed Assets Purchase");
		
		nameTxt.sendKeys(Keys.TAB);
		
		codeTxt.sendKeys("Fixed Assets Purchase");
		codeTxt.sendKeys(Keys.TAB);
		
		Select accountTypeDropdownSelect = new Select(accountTypeDropdown);
		
		accountTypeDropdownSelect.selectByVisibleText("Fixed Assets");
		
		accountTypeDropdown.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatCreationPageSaveBtn));
		vatCreationPageSaveBtn.click();
     	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherBodyItemName));
		voucherBodyItemName.click();
		
        getAction().doubleClick(voucherBodyItemName).build().perform();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.sendKeys("FAItem");
		nameTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
        codeTxt.sendKeys("FAItem");
		codeTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemNewUnitsBtn));
		ItemNewUnitsBtn.click();
		itemUnitsDefaultBaseUnitDropdown.click();
		itemUnitsDefaultBaseUnitDropdown.sendKeys("Pcs");
		itemUnitsDefaultBaseUnitDropdown.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatCreationPageSaveBtn));
	
		vatCreationPageSaveBtn.click();
    
		for (int i = 0; i < 2; i++) 
		{
			if (getIsAlertPresent()) 
			{
				System.err.println(i+" "+":::After Save:::");
				System.err.println(getAlert().getText());
				getAlert().accept();
			}
		}
		
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
		if (getIsAlertPresent()) 
		{
			getAlert().accept();
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemMenu));
		ItemMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Item_ItemMenu));
		Item_ItemMenu.click();
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cmbSearchTxt));
		cmbSearchTxt.click();
		
		cmbSearchTxt.sendKeys("FAItem");
		
		Thread.sleep(2000);
		
		cmbSearchTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(propertiesBtn));
		propertiesBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dontMaintainStockbyBatchCheckbox));
		if (dontMaintainStockbyBatchCheckbox.isSelected()==false)
		{
			dontMaintainStockbyBatchCheckbox.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dontinputBinCheckbox));
		if (dontinputBinCheckbox.isSelected()==false)
		{
			dontinputBinCheckbox.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dontMaintainStocksByRMACheckbox));
		if (dontMaintainStocksByRMACheckbox.isSelected()==false)
		{
			dontMaintainStocksByRMACheckbox.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(quantitiesRadioBtn));
		if (quantitiesRadioBtn.isSelected()==false)
		{
			quantitiesRadioBtn.click();
		}
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemProperties_okBtn));
		itemProperties_okBtn.click();
		
		
		if (dontMaintainStockbyBatchCheckbox.isSelected()==true && dontinputBinCheckbox.isSelected()==true
				&& quantitiesRadioBtn.isSelected()==true)
		{
			System.out.println(" Test Pass: Purchase Account and Item Saved Successfully");
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
			userNameDisplay.click();
				  
			System.out.println("*********Logout Successfully********************************8");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			logoutOption.click();
			return true;
		} 
		else
		{
			System.out.println(" Test Fail: Purchase Account and Item Saved Successfully");
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
			userNameDisplay.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			logoutOption.click();
			return false;
		}
	}
	
	@FindBy(xpath="//i[@class='icon-font7 icon-ok']")
	public static WebElement itemProperties_okBtn;
	
	@FindBy(xpath="//input[@id='rbnDefault']")
	public static WebElement  defaultRadioBtn;

	@FindBy(xpath="//label[@id='lblNone']")
	public static WebElement reservationNone;
	
	@FindBy(xpath="//input[@id='rbnQualities']")
	public static WebElement  quantitiesRadioBtn;
	
	
	@FindBy(xpath="//input[@id='chkDontMaintainStockbyBatch']")
	public static WebElement dontMaintainStockbyBatchCheckbox ;
	
	@FindBy(xpath="//input[@id='chkDontinputBin']")
	public static WebElement  dontinputBinCheckbox;
	
	@FindBy(xpath="//input[@id='chkDontMaintainStocksByRMA']")
	public static WebElement dontMaintainStocksByRMACheckbox ;
	

@FindBy(xpath="//tbody[@id='id_header_4_table_data_body']/tr/td")
private static List<WebElement> vendorAccountListCount; 

@FindBy(xpath="//tbody[@id='id_header_268435459_table_data_body']/tr/td")
private static List<WebElement> departmentListCount;
	
	
@FindBy(xpath="//*[@id='2008']/span")
private static WebElement  purchaseVoucher;


@FindBy(xpath="//input[@id='txtUsername']")
private static WebElement  loginUserNameTxt;

@FindBy(xpath="//input[@id='txtPassword']")
private static WebElement  loginPasswordTxt;

@FindBy(xpath="//button[@id='btnSignin']")
private static WebElement  signInBtn;


@FindBy(xpath="//tbody[@id='id_header_3_table_data_body']/tr/td")
private static List<WebElement> purchaseAccountListCount; 
		

@FindBy(xpath="//tbody[@id='id_header_268435460_table_data_body']/tr/td")
private static List<WebElement> wareHouseListCount;
	
	public boolean checkSavingVoucherinPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		
		System.out.println("********************* checkSavingVoucherinPurchaseVoucherVAt Method Executes.............   *****************************");
			

	    LoginPage lp=new LoginPage(getDriver()); 
		
	    String unamelt="su";
	  
	    String pawslt="su";
	      
	    lp.enterUserName(unamelt);
	    
	    lp.enterPassword(pawslt);
	    
	    lp.clickOnSignInBtn();
		
		Thread.sleep(5000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
		financialMenu.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		financialsTransactionsPurchaseMenu.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucher));
		purchaseVoucher.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		
	    checkUserFriendlyMessage();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
		purchaseAccountTxt.click();
		
		purchaseAccountTxt.sendKeys(Keys.SPACE);
		
		ArrayList<String>  actPVPurchaseAccount= new ArrayList<String>(); 
		int count=purchaseAccountListCount.size();
		
		System.err.println(count);
		
		for(int i=0 ; i < count ;i++)
		{
			String data=purchaseAccountListCount.get(i).getText();
			
			if(data.equalsIgnoreCase("Fixed Assets Purchase"))
			{
				purchaseAccountListCount.get(i).click();
				
				break;
			}
			
		}
		
		purchaseAccountTxt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();
		vendorAccountTxt.sendKeys(Keys.SPACE);
		
		int vendorcount=vendorAccountListCount.size();
		
		System.err.println(vendorcount);
		
		for(int i=0 ; i < vendorcount ;i++)
		{
			String data=vendorAccountListCount.get(i).getText();
			
			if(data.equalsIgnoreCase("Vendor A"))
			{
				vendorAccountListCount.get(i).click();
				
				break;
			}
		}
		
		vendorAccountTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		
		departmentTxt.sendKeys(Keys.SPACE);
			
		int departmentcount=departmentListCount.size();
		
		System.err.println(departmentcount);
			
		for(int i=0 ; i < departmentcount ;i++)
		{
			String data=departmentListCount.get(i).getText();
			
			if(data.equalsIgnoreCase("India"))
			{
				departmentListCount.get(i).click();
				
				break;
			}
		}
			
		departmentTxt.sendKeys(Keys.TAB);	
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(wareHouseTxt));
		wareHouseTxt.click();
		
		wareHouseTxt.sendKeys(Keys.SPACE);
		
		int wareHouseListCountCount = wareHouseListCount.size();
		
		ArrayList<String> wareHouseListCountArray = new ArrayList<String>();
		
		for(int i=0;i<wareHouseListCountCount;i++)
		{
			String data=wareHouseListCount.get(i).getText();
			if (data.equalsIgnoreCase("HYDERABAD")) 
			{
				wareHouseListCount.get(i).click();
			}
		}
		
		wareHouseTxt.sendKeys(Keys.TAB);

		
        Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys("FAItem");
		
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		select1stRow_4thColumn.click();
		
		enter_Quantity.sendKeys("1");
		
		enter_Quantity.sendKeys(Keys.TAB);
		
		enter_Rate.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME,Keys.BACK_SPACE);
		
		enter_Rate.sendKeys("20000");
		Thread.sleep(2000);
		
		enter_Rate.sendKeys(Keys.TAB);
		
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(VoucherSaveBtn));
		VoucherSaveBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		boolean actFixedAssets=vatCreationPageSaveBtn.isDisplayed();
		
		System.out.println("actFixedAssets  : "+actFixedAssets);
		
		
		if (actFixedAssets==true )
		{
			System.out.println(" Test Pass: Fixed Assets Pop up Displayed ");
			return true;
		} 
		else
		{
			System.out.println(" Test Fail: Fixed Assets Pop up Displayed ");
			return false;
		}
	}
	
	@FindBy(xpath="//input[@id='txtNewReference']")
	private static WebElement  billRefNewReferenceTxt;
	
	@FindBy(xpath="//span[@id='btnPick']")
	private static WebElement  billRefPickIcon;		
	
	@FindBy(xpath="//span[@id='btnOk']")
	private static WebElement  billRefOkBtn;	
	
	@FindBy(xpath="//*[@id='id_transactionentry_save']/div[2]")
	private static WebElement  VoucherSaveBtn;
	
	@FindBy(xpath="//*[@id='GenerateNewMasterModel']/div/div[2]/div[2]/div[2]/div/ul/li[2]/button[5]")
	private static WebElement vatCreationPageCloseBtn;
	
	
	@FindBy(xpath="//input[@id='oc_Group']")
	private static WebElement  groupTXt;
				
				
		

	public boolean checkFixedAssetsPopUpScreenGenaralTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		System.out.println("********************* checkFixedAssetsPopUpScreen Method Executes.............   *****************************");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(groupTXt));
		
		groupTXt.click();
		groupTXt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME,Keys.SPACE);
		
		groupTXt.sendKeys("FixedAssetsBasedOnUsage");
		
		Thread.sleep(2000);
		groupTXt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		
		String actName=nameTxt.getAttribute("value");
		String actCode=codeTxt.getAttribute("value");
		
		Select s = new Select(assertTypeDrpdwn);
		s.selectByVisibleText("Computers");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentdrpdwn));
		departmentdrpdwn.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
		departmentdrpdwn.sendKeys("India");
		
		Thread.sleep(2000);
		
		departmentdrpdwn.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentdrpdwn));
		String actAssertTypeDrpdwn = s.getFirstSelectedOption().getText();
		String actDepTypeDrpdwn    = departmentdrpdwn.getAttribute("value");
		String actvendorDrpdwn     = vendorDrpdwn.getAttribute("value");
		String actitemDrpdwn       = itemDrpdwn.getAttribute("value");
		String actPurchaseDate     = purchaseDateTxt.getAttribute("value");
		String actPrice            = priceTxt.getAttribute("value");
		String actQuantity         = quantityTxt.getAttribute("value");
		
		
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
	    Date date=new Date();
	    System.out.println(df.format(date));
		
	    int year=date.getYear();
	    
		
		String expName             = "FAItem";
		String expCode             = "FAItem";
		String expAssertTypeDrpdwn = "Computers";
		String expDepTypeDrpdwn    = "India";
		String expvendorDrpdwn     = "Vendor A";
		String expitemDrpdwn       = "FAItem";
		String expPurchaseDate     = df.format(date);
		String expPrice            = "20000";
		String expQuantity         = "1";
		
		System.out.println("actName               : "+actName             +" Value Expected : "+expName);
		System.out.println("actCode               : "+actCode             +" Value Expected : "+expCode);
		System.out.println("actAssertTypeDrpdwn   : "+actAssertTypeDrpdwn +" Value Expected : "+expAssertTypeDrpdwn);
		System.out.println("actDepTypeDrpdwn      : "+actDepTypeDrpdwn    +" Value Expected : "+expDepTypeDrpdwn);
		System.out.println("actvendorDrpdwn       : "+actvendorDrpdwn     +" Value Expected : "+expvendorDrpdwn);
		System.out.println("actitemDrpdwn         : "+actitemDrpdwn       +" Value Expected : "+expitemDrpdwn);
		System.out.println("actPurchaseDate       : "+actPurchaseDate     +" Value Expected : "+expPurchaseDate);
		System.out.println("actPrice              : "+actPrice             +" Value Expected : "+expPrice);
		System.out.println("actQuantity           : "+actQuantity          +" Value Expected : "+expQuantity);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseDateTxt));
		purchaseDateTxt.click();
		purchaseDateTxt.sendKeys(Keys.HOME);
		
		purchaseDateTxt.sendKeys("0101"+year);
		
		System.err.println("0101"+year);
		
		
		purchaseDateTxt.sendKeys(Keys.TAB);
		
		if (actName.equalsIgnoreCase(expName) && actCode.equalsIgnoreCase(expCode) && actAssertTypeDrpdwn.equalsIgnoreCase(expAssertTypeDrpdwn) && 
				actDepTypeDrpdwn.equalsIgnoreCase(expDepTypeDrpdwn) && actvendorDrpdwn.equalsIgnoreCase(expvendorDrpdwn) && 
				actitemDrpdwn.equalsIgnoreCase(expitemDrpdwn) && actPurchaseDate.equalsIgnoreCase(expPurchaseDate) && 
				actPrice.equalsIgnoreCase(expPrice) && actQuantity.equalsIgnoreCase(expQuantity)) 
		{
			
			System.out.println(" Test Pass :  Values as Expected in Genaral Tab");
			return true;
		} 
		else 
		{

			System.out.println(" Test Fail :  Values as Expected in Genaral Tab");
			return false;
		}
	}
	

				

	@FindBy(xpath="//input[@id='MaximumusageLimit']")
	private static WebElement  maxUsageLimit;
	
	
	@FindBy(xpath="//*[@id='newMasterDiv2']/div/div[2]/div[1]/label/a")
	private static WebElement  assetACCode;
	
	@FindBy(xpath="//input[@id='iAssetACCode']")
	private static WebElement  assetACCodeTxt;
	
	
	
	
	@FindBy(xpath="//a[contains(text(),'Accumulated Depreciation')]")
	private static WebElement  accumulatedDepreciation;
	
	
	@FindBy(xpath="//input[@id='iAccumulatedDepreciation']")
	private static WebElement  accumulatedDepreciationTxt;
	
	
	
	@FindBy(xpath="//a[contains(text(),'Depreciation P')]")
	private static WebElement  DepreciationPLCode;
	
	@FindBy(xpath="//input[@id='iDepreciationPLCode']")
	private static WebElement  DepreciationPLCodeTXt;
	
	
	
	
	@FindBy(xpath="//a[contains(text(),'Sales Account')]")
	private static WebElement  salesAccount;
	
	@FindBy(xpath="//input[@id='iSalesAccount']")
	private static WebElement  ALsalesAccountTxt;
	
	
	
	@FindBy(xpath="//a[contains(text(),'Disposal Gain')]")
	private static WebElement  disposalGain;
	
	@FindBy(xpath="//input[@id='iDisposalGain']")
	private static WebElement  disposalGainTxt;
	
	
	
	
	@FindBy(xpath="//a[contains(text(),'Disposal Loss')]")
	private static WebElement  disposalLoss;
	
	@FindBy(xpath="//input[@id='iDisposalLoss']")
	private static WebElement  disposalLossTxt;
	
	
	
	
	public boolean checkFixedAssetsPopUpScreenDepreciationTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		
		System.out.println("********************* checkFixedAssetsPopUpScreenDepreciationTab Method Executes.............   *****************************");
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deperciationTab));
		deperciationTab.click();
		
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(depreciationMethodDrpdwn));
		Select s = new Select(depreciationMethodDrpdwn);
		
        s.selectByValue("0");
		
		
		depreciationMethodDrpdwn.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(depaerciationYearsTxt));
		depaerciationYearsTxt.click();
		
		depaerciationYearsTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
		depaerciationYearsTxt.sendKeys("0");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dateOfCommencementTxt));
		dateOfCommencementTxt.click();
		dateOfCommencementTxt.sendKeys(Keys.HOME);
		
		dateOfCommencementTxt.sendKeys("0101");
		dateOfCommencementTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(maxUsageLimit));
		maxUsageLimit.click();
		
		maxUsageLimit.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
		maxUsageLimit.sendKeys("50");
		
		maxUsageLimit.sendKeys(Keys.TAB);
		
		System.out.println(" ********Account Links TAB ****************");
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountingLinksTab));
		accountingLinksTab.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assetACCodeTxt));
		assetACCodeTxt.click();
		
		assetACCodeTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
		
		assetACCodeTxt.sendKeys("Fixed Assets Purchase");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accumulatedDepreciation));
		accumulatedDepreciation.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
        nameTxt.click();
        
        nameTxt.sendKeys("ALAccumulatedDeprection");
        
        Thread.sleep(2000);
        
        nameTxt.sendKeys(Keys.TAB);
        
        codeTxt.sendKeys("ALAccumulatedDeprection");
		
        
        Select accountTypeDropdownSelect = new Select(accountTypeDropdown);
		
		accountTypeDropdownSelect.selectByVisibleText("Expenses");
		
		
		accountTypeDropdown.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatCreationPageSaveBtn));
		vatCreationPageSaveBtn.click();
     	
		System.out.println("*************************accumulatedDepreciation Account Saved Successfully***********");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepreciationPLCode));
		DepreciationPLCode.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
        nameTxt.click();
        
        nameTxt.sendKeys("ALDepreciationPLCode");
        
        Thread.sleep(2000);
        
        nameTxt.sendKeys(Keys.TAB);
        
        codeTxt.sendKeys("ALDepreciationPLCode");
		
        
        Select accountTypeDropdownSelect1 = new Select(accountTypeDropdown);
		
		accountTypeDropdownSelect1.selectByVisibleText("Assets");
		
		
		accountTypeDropdown.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatCreationPageSaveBtn));
		vatCreationPageSaveBtn.click();
		
		System.out.println("*************************Depreciation PL Code Account Saved Successfully***********");
		
        Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesAccount));
		salesAccount.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
        nameTxt.click();
        
        nameTxt.sendKeys("ALsalesAccount");
        
        Thread.sleep(2000);
        
        nameTxt.sendKeys(Keys.TAB);
        
        codeTxt.sendKeys("ALsalesAccount");
		
        
        Select accountTypeDropdownSelect2 = new Select(accountTypeDropdown);
		
		accountTypeDropdownSelect2.selectByVisibleText("Sales");
		
		
		accountTypeDropdown.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatCreationPageSaveBtn));
		vatCreationPageSaveBtn.click();
		
		System.out.println("*************************Sales Account Account Saved Successfully***********");
		
        Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(disposalGain));
		disposalGain.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
        nameTxt.click();
        
        nameTxt.sendKeys("ALDisposalGain");
        
        Thread.sleep(2000);
        
        nameTxt.sendKeys(Keys.TAB);
        
        codeTxt.sendKeys("ALDisposalGain");
		
        
        Select accountTypeDropdownSelect3 = new Select(accountTypeDropdown);
		
		accountTypeDropdownSelect3.selectByVisibleText("Expenses");
		
		accountTypeDropdown.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatCreationPageSaveBtn));
		vatCreationPageSaveBtn.click();
		
		System.out.println("*************************Disposal Gain Account Saved Successfully***********");
		
		
		Thread.sleep(2000);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(disposalLoss));
		disposalLoss.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
        nameTxt.click();
        
        nameTxt.sendKeys("ALDisposalLoss");
        
        Thread.sleep(2000);
        
        nameTxt.sendKeys(Keys.TAB);
        
        codeTxt.sendKeys("ALDisposalLoss");
		
        
        Select accountTypeDropdownSelect4 = new Select(accountTypeDropdown);
		accountTypeDropdownSelect4.selectByVisibleText("Expenses");
		
		accountTypeDropdown.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatCreationPageSaveBtn));
		vatCreationPageSaveBtn.click();
			
		System.out.println("*************************Disposal Loss Account Saved Successfully***********");
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accumulatedDepreciationTxt));
		accumulatedDepreciationTxt.click();
		
		accumulatedDepreciationTxt.sendKeys("ALAccumulatedDeprection");
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepreciationPLCodeTXt));
		DepreciationPLCodeTXt.click();
		
		DepreciationPLCodeTXt.sendKeys("ALDepreciationPLCode");
		
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ALsalesAccountTxt));
		ALsalesAccountTxt.click();
		
		ALsalesAccountTxt.sendKeys("ALsalesAccount");
		
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(disposalGainTxt));
		disposalGainTxt.click();
		
		disposalGainTxt.sendKeys("ALDisposalGain");
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(disposalLossTxt));
		disposalLossTxt.click();
		
		disposalLossTxt.sendKeys("ALDisposalLoss");
		
        Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatCreationPageSaveBtn));
		vatCreationPageSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String docno = documentNumberTxt.getAttribute("value");
		
		boolean savingMessage=checkVoucherSavingMessage(docno);
		   
		String actSavingMessage=Boolean.toString(savingMessage);
		String expSavingMessage="true";
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
		
		System.out.println("actSavingMessage   : "+actSavingMessage  +" Value Expected : "+expSavingMessage);
		
		if (actSavingMessage.equalsIgnoreCase(expSavingMessage)  )
		{
			System.out.println(" Test Pass: Voucher Saved SuccessFully");
			return true;
			
		} 
		else
		{
			System.out.println(" Test Fail: Voucher Saved SuccessFully");
			return false;
		}
	}
	
	  @FindBy(xpath="//*[@id='2023']/span")
	  private static WebElement  finTransJournalsMenu;
	  
	  @FindBy(xpath="//a[@id='2074']")
	  private static WebElement  FADepreciationTransferVoucherMenu;
	  
	
	  
			
	  @FindBy(xpath="//a[@id='2046']//span[contains(text(),'Fixed Asset Depreciation Voucher')]")
	  private static WebElement  fixedAssetDepreciationVoucherMenu;
		
	  @FindBy(xpath="//div[@id='myNavbar']//span[@class='icon-settings icon-font6']")
	  private static WebElement settingBtn;
	
	  
	  @FindBy(xpath="//td[@id='doc_TagsTable_col_2-1']")
	  private static WebElement  masters2;
		
	  @FindBy(xpath="//td[@id='doc_TagsTable_col_2-2']")
	  private static WebElement  position2;
	  
	  @FindBy(xpath="//select[@id='doc_TagsTableMasterDropDown']")
	  private static WebElement  masterDropDown;
		
	  @FindBy(xpath="//select[@id='doc_TagsPositionDropDown']")
	  private static WebElement  positionDropdown;
	  
	
	  
	public boolean checkAddingFixedAssetInHeaderOFFixedAssetDepreciationVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
		financialMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finTransJournalsMenu));
		finTransJournalsMenu.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssetDepreciationVoucherMenu));
		fixedAssetDepreciationVoucherMenu.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingBtn));
		settingBtn.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masters2)); 
		masters2.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDropDown));
		masterDropDown.click();
		
		Select s1=new Select(masterDropDown);
		s1.selectByVisibleText("Fixed Asset");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(position2));
		position2.click();
		
		positionDropdown.click();
		
		Select s2=new Select(positionDropdown);
		s2.selectByVisibleText("Header");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		
	    String expMessage="Data saved successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingCloseIcon));
		settingCloseIcon.click();
		
		if ( actMessage.equalsIgnoreCase(expMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	@FindBy(xpath="//div[contains(text(),'Fixed Asset')]")
	private static WebElement  fixedAssets_Menu;
	  
	
	@FindBy(xpath="//a[@id='138']//span[contains(text(),'Transactions')]")
	private static WebElement  fixedAssetsTransMenu;
	  
	
	@FindBy(xpath="//a[@id='50002']//span[contains(text(),'Asset Usage')]")
	private static WebElement  transAssetsUsageMenu;
	  
	
	@FindBy(xpath="//*[@id='AssetUsage']/ul/li[3]")
	private static WebElement  assertUsageNewBtn;
	
	
	@FindBy(xpath="//input[@id='txtDocNo']")
	private static WebElement  addAssultValueDOCTxt;
	
	
	
	@FindBy(xpath="//input[@id='txtdocno']")
	private static WebElement  assertUsageDocNoTxt;
	  
	@FindBy(xpath="//input[@id='AssetId']")
	private static WebElement  assertUsageAssertDrpdwn;
	  
	@FindBy(xpath="//input[@id='opcTag0']")
	private static WebElement  assertUsageDepDrpdwn;
	  
	@FindBy(xpath="//input[@id='id_date_english']")
	private static WebElement  assertUsageDateTxt;
	  
	@FindBy(xpath="//input[@id='txtUnitsOfUsage']")
	private static WebElement  assertUsageUnitsOFUnits;
	  
	@FindBy(xpath="//input[@id='txtMaximumUsage']")
	private static WebElement  assertUsageMaxUsage;
	  
	@FindBy(xpath="//input[@id='txtNewUnitHrs']")
	private static WebElement  assertUsageNewUnitsTxt;
	  
	@FindBy(xpath="//input[@id='txtLastReadingHrs']")
	private static WebElement  assertUsageLastReadingTxt;
	  
	@FindBy(xpath="//input[@id='txtNewReadingHrs']")
	private static WebElement  assertUsageNewReadingTxt;
	  
	@FindBy(xpath="//span[@class='icon-new icon-font7']")
	private static WebElement  assertUsageNew_NewBtn;
	  
	@FindBy(xpath="//div[@id='btnSaveHeader']//span[contains(@class,'icon-font7')]")
	private static WebElement  assertUsageSaveBtn;
	  
	@FindBy(xpath="//div[contains(text(),'Delete')]")
	private static WebElement  assertUsageDeleteBtn;
	  
	@FindBy(xpath="//div[@name='btnClose']")
	private static WebElement  assertUsageCloseBtn;
	  
	@FindBy(xpath="//div[@class='toolbar_button']//span[@class='icon-close icon-font7']")
	private static WebElement  disposalOFAssertCLoseBtn;
	  
	
	
	
	public boolean checkSavingVoucherInAssertUsage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssets_Menu));
		fixedAssets_Menu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssetsTransMenu));
		fixedAssetsTransMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transAssetsUsageMenu));
		transAssetsUsageMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageNewBtn));
		assertUsageNewBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageAssertDrpdwn));
		assertUsageAssertDrpdwn.click();
		assertUsageAssertDrpdwn.sendKeys("FAItem");
		
		Thread.sleep(2000);
		assertUsageAssertDrpdwn.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageDateTxt));
		assertUsageDateTxt.click();
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
	    Date date=new Date();
        		
	    String docdate = df.format(date);  
		Calendar c = Calendar.getInstance();
		c.add(Calendar.DAY_OF_MONTH, -20);  
		
		String FilterDate = df.format(c.getTime());  
		
		System.out.println("FilterDate  : "+FilterDate);
		
		assertUsageDateTxt.sendKeys(FilterDate);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageNewUnitsTxt));
		assertUsageNewUnitsTxt.click();
		assertUsageNewUnitsTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
		assertUsageNewUnitsTxt.sendKeys("10");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageSaveBtn));
		assertUsageSaveBtn.click();
		
		
        String expMessage="Record Saved Succesfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	public boolean checkEditingInSavedVoucherInAssertUsage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageDocNoTxt));
		assertUsageDocNoTxt.click();
		
		assertUsageDocNoTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
		assertUsageDocNoTxt.sendKeys("1");
		Thread.sleep(2000);
		assertUsageDocNoTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageNewUnitsTxt));
		assertUsageNewUnitsTxt.click();
		assertUsageNewUnitsTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		assertUsageNewUnitsTxt.sendKeys("5");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageSaveBtn));
		assertUsageSaveBtn.click();
		
        String expMessage="Record Saved Succesfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	
	
	@FindBy(xpath="//button[contains(text(),'YES')]")
	private static WebElement  yesBtn;
	
	
	public boolean checkDeletingVoucherInAssertUsage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		
		System.out.println("***********************************checkDeletingVoucherInAssertUsage  Excuted Starts***************");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageDocNoTxt));
		assertUsageDocNoTxt.click();
		
		assertUsageDocNoTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
		assertUsageDocNoTxt.sendKeys("1");
		Thread.sleep(2000);
		assertUsageDocNoTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageDeleteBtn));
		assertUsageDeleteBtn.click();
		
		Thread.sleep(1000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesBtn));
		yesBtn.click();
		
		
        String expMessage="Record Deleted Successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	
	public boolean checkSavingVoucherInAssertUsageAfterDeletion() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageAssertDrpdwn));
		assertUsageAssertDrpdwn.click();
		assertUsageAssertDrpdwn.sendKeys("FAItem");
		Thread.sleep(2000);
		assertUsageAssertDrpdwn.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageDateTxt));
		assertUsageDateTxt.click();
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
	    Date date=new Date();
        		
	    String docdate = df.format(date);  
		Calendar c = Calendar.getInstance();
		c.add(Calendar.DAY_OF_MONTH, -20);  
		
		String FilterDate = df.format(c.getTime());  
		
		System.out.println("FilterDate  : "+FilterDate);
		
		assertUsageDateTxt.sendKeys(FilterDate);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageNewUnitsTxt));
		assertUsageNewUnitsTxt.click();
		assertUsageNewUnitsTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
		assertUsageNewUnitsTxt.sendKeys("10");
		Thread.sleep(2000);
		assertUsageNewUnitsTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageSaveBtn));
		assertUsageSaveBtn.click();
		
		
        String expMessage="Record Saved Succesfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	
	
	public boolean checkSavingVoucherInAssertUsageWithSameAssertUsage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageAssertDrpdwn));
		assertUsageAssertDrpdwn.click();
		assertUsageAssertDrpdwn.sendKeys("FAItem");
		
		Thread.sleep(2000);
		assertUsageAssertDrpdwn.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageDateTxt));
		assertUsageDateTxt.click();
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
	    Date date=new Date();
        		
	    String docdate = df.format(date);  
		Calendar c = Calendar.getInstance();
		c.add(Calendar.DAY_OF_MONTH, -20);  
		
		String FilterDate = df.format(c.getTime());  
		
		System.out.println("FilterDate  : "+FilterDate);
		
		assertUsageDateTxt.sendKeys(FilterDate);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageNewUnitsTxt));
		assertUsageNewUnitsTxt.click();
		assertUsageNewUnitsTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		assertUsageNewUnitsTxt.sendKeys("5");
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageSaveBtn));
		assertUsageSaveBtn.click();
		
        String expMessage="Record Saved Succesfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	@FindBy(xpath="//*[@id='id_footer']/button[1]")
	private static WebElement informationYesBtn;
	

	
	public boolean checkSavedVoucherInAssertUsage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		
		System.out.println("***************************************checkSavedVoucherInAssertUsage  Excutes ****************" );
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageDocNoTxt));
		assertUsageDocNoTxt.click();
		
		assertUsageDocNoTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
		assertUsageDocNoTxt.sendKeys("2");
		Thread.sleep(2000);
		assertUsageDocNoTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageDocNoTxt));
		String actDocNo                           =assertUsageDocNoTxt.getAttribute("value");
		String actAssertUsage                     =assertUsageAssertDrpdwn.getAttribute("value");
		String actDepartment                      =assertUsageDepDrpdwn.getAttribute("value");
		String actDate							  =assertUsageDateTxt.getAttribute("value");
		String actUnitsOfUssage					  =assertUsageUnitsOFUnits.getAttribute("value");
		String actMaxUsage						  =assertUsageMaxUsage.getAttribute("value");
		String actnewUnitHrs					  =assertUsageNewUnitsTxt.getAttribute("value");
		String actassertUsageLastReadingTxt		  =assertUsageLastReadingTxt.getAttribute("Value");
		String actassertUsageNewReadingTxt 		  =assertUsageNewReadingTxt.getAttribute("value");
		
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
	    Date date=new Date();
        		
	    String docdate = df.format(date);  
		Calendar c = Calendar.getInstance();
		c.add(Calendar.DAY_OF_MONTH, -20);  
		
		String FilterDate = df.format(c.getTime());  
		
		System.out.println("FilterDate  : "+FilterDate);
		
		
		String expDocNo							="2";
		String expAssertUsage					="FAItem";
		String expDepartment					="INDIA";
		
		String expDate							=FilterDate;
		String expUnitsOfUssage					="Hour";
		String expMaxUsage						="50.00";
		String expnewUnitHrs					="5.00";
		String expassertUsageLastReadTxt		="10.00";
		String expassertUsageNewReadTxt			="15.00";
		
		
		System.out.println("actDocNo                  :"+actDocNo                 +" Value Expected   : "+expDocNo);
		System.out.println("actAssertUsage            :"+actAssertUsage           +" Value Expected   : "+expAssertUsage);
		System.out.println("actDepartment             :"+actDepartment            +" Value Expected   : "+expDepartment);
		System.out.println("actDate                   :"+actDate                  +" Value Expected   : "+expDate);
		System.out.println("actUnitsOfUssage          :"+actUnitsOfUssage         +" Value Expected   : "+expUnitsOfUssage);
		System.out.println("actMaxUsage               :"+actMaxUsage              +" Value Expected   : "+expMaxUsage);
		System.out.println("actnewUnitHrs             :"+actnewUnitHrs            +" Value Expected   : "+expnewUnitHrs);
		System.out.println("actassertUsageLastReadTxt :"+actassertUsageLastReadingTxt +" Value Expected   : "+expassertUsageLastReadTxt);
		System.out.println("actassertUsageNewReadTxt  :"+actassertUsageNewReadingTxt  +" Value Expected   : "+expassertUsageNewReadTxt);
		
		if (actDocNo.equalsIgnoreCase(expDocNo) && 
				actAssertUsage.equalsIgnoreCase(actAssertUsage) && 
				actDate.equalsIgnoreCase(expDate) && 
				actUnitsOfUssage.equalsIgnoreCase(expUnitsOfUssage) && 
				actMaxUsage.equalsIgnoreCase(expMaxUsage) && 
				actnewUnitHrs.equalsIgnoreCase(expnewUnitHrs) && 
				actassertUsageLastReadingTxt.equalsIgnoreCase(expassertUsageLastReadTxt) && 
				actassertUsageNewReadingTxt.equalsIgnoreCase(expassertUsageNewReadTxt) ) 
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	
	
	@FindBy(xpath="//a[@id='50009']//span[contains(text(),'Post Depreciation Entries')]")
	private static WebElement  postDepreciationEntriesMenu;
	
	@FindBy(xpath="//input[@id='chkPostIndividualVouchers']")
	private static WebElement  postIndividualVoucherChkbox;
	
	@FindBy(xpath="//input[@id='AssetId']")
	private static WebElement  listingPostingToTheGroupdrpdwn;
	
	@FindBy(xpath="//span[@class='icon-ok icon-font7']")
	private static WebElement  postDepreciationOkBtn;
	
	@FindBy(xpath="//span[@class='icon-close icon-font7']")
	private static WebElement  postDepreciationCancelBtn;
	
	@FindBy(xpath="//span[@class='icon-close icon-font7']")
	private static WebElement  faAuthCloseBtn;
	
	
	
	
	public boolean checkPostingInPostDepreciationEntries() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssets_Menu));
		fixedAssets_Menu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssetsTransMenu));
		fixedAssetsTransMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postDepreciationEntriesMenu));
		postDepreciationEntriesMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postIndividualVoucherChkbox));
		
		if (postIndividualVoucherChkbox.isSelected()==true)
		{
			postIndividualVoucherChkbox.click();
			
		}
		
		Thread.sleep(3000);
		
		listingPostingToTheGroupdrpdwn.sendKeys("FixedAssetsBasedOnUsage");
		
		Thread.sleep(3000);
		
		listingPostingToTheGroupdrpdwn.sendKeys(Keys.TAB);
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postDepreciationOkBtn));
		postDepreciationOkBtn.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(informationYesBtn));
		informationYesBtn.click();
		
        String expMessage="Voucher Posted Sucessfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postDepreciationCancelBtn));
		postDepreciationCancelBtn.click();
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	

	public boolean checkPostedVoucherInFixedAssetDepreciationVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
		 financialMenu.click();
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		 financialsTransactionMenu.click();
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finTransJournalsMenu));
		 finTransJournalsMenu.click();
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssetDepreciationVoucherMenu));
		 fixedAssetDepreciationVoucherMenu.click();
		
		 Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		 editBtn.click();
		
		 boolean loading=checkLoadingMessage();
		
		
		 Thread.sleep(2000);

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		
		 String actDocno      = documentNumberTxt.getAttribute("value");
		 String actDate       = dateTxt.getAttribute("value");
		 String actDepartment = departmentTxt.getAttribute("value");
		 String actFixedAsset = fixedAssetTxt.getAttribute("value");
		 
		 Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		 String actAccRow1    = select1stRow_1stColumn.getText();
		 String actDebitRow1  = select1stRow_2ndColumn.getText();
		 
		 Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		 String actAccRow2    = select2ndRow_1stColumn.getText();
		 String actCreditRow2 = select2ndRow_3rdColumn.getText();
		
		 
		 DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
	     Date date=new Date();
        		
	     String docdate = df.format(date); 
	 
	     System.out.println("Current Date  : "+docdate);
		 
		 String expDocno      = "Dep:1";
		 String expDate       = docdate;
		 String expDepartment = "INDIA";
		 String expFixedAsset = "FAItem";
		 String expAccRow1    = "ALDepreciationPLCode";
		 String expDebitRow1  = "6,000.00";
		 String expAccRow2    = "ALAccumulatedDeprection";
		 String expCreditRow2 = "6,000.00";
		 
		 
		 System.out.println("actDocno       :"+actDocno       +" Value Expected  :" +expDocno);
		 System.out.println("actDate        :"+actDate        +" Value Expected  :" +expDate);
		 System.out.println("actDepartment  :"+actDepartment  +" Value Expected  :" +expDepartment);
		 System.out.println("actFixedAsset  :"+actFixedAsset  +" Value Expected  :" +expFixedAsset);
		 System.out.println("actAccRow1     :"+actAccRow1     +" Value Expected  :" +expAccRow1);

		 System.out.println("actDebitRow1   :"+actDebitRow1   +" Value Expected  :" +expDebitRow1);
		 System.out.println("actAccRow2     :"+actAccRow2     +" Value Expected  :" +expAccRow2);
		 System.out.println("actCreditRow2  :"+actCreditRow2  +" Value Expected  :" +expCreditRow2);
		 
		
			if (actDocno.equalsIgnoreCase(expDocno) && 
					actDate.equalsIgnoreCase(expDate) && 
					actFixedAsset.equalsIgnoreCase(expFixedAsset) && 
					actAccRow1.equalsIgnoreCase(expAccRow1) && 
					actDebitRow1.equalsIgnoreCase(expDebitRow1) && 
					actAccRow2.equalsIgnoreCase(expAccRow2) && 
					actCreditRow2.equalsIgnoreCase(expCreditRow2) ) 
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
	@FindBy(xpath="//span[@id='btnNew']")
	public static WebElement masterNewBtn;
	
	
	public boolean checkAddingAccountUnderGroupInFixedAssets() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
	
		System.out.println("********************* checkAddingAccountUnderGroupInFixedAssets Method Executes.............   *****************************");
			
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		 homeMenu.click();
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		 mastersMenu.click();
		
		 
		 Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssetsMenu));
		 fixedAssetsMenu.click();
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterRow1Chkbox));
		 masterRow1Chkbox.click();
		
		 getAction().doubleClick(masterRow1Chkbox).build().perform();
		
		 Thread.sleep(2000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		 masterNewBtn.click();
		 
		 System.out.println("**********General Tab *****************************");
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.sendKeys("Bike");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 codeTxt.sendKeys("Bike");
		 
		 Select s = new Select(assertTypeDrpdwn);
	     s.selectByVisibleText("Vehicle");
		 
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentdrpdwn));
	     departmentdrpdwn.click();
	     departmentdrpdwn.sendKeys("AMERICA");
	     Thread.sleep(2000);
	     
	     departmentdrpdwn.sendKeys(Keys.TAB);
	     
		 
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorDrpdwn));
	     vendorDrpdwn.click();
	     vendorDrpdwn.sendKeys("Test Master");
	     Thread.sleep(2000);
	     
	     vendorDrpdwn.sendKeys(Keys.TAB);
	     
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseDateTxt));
	     purchaseDateTxt.click();
	     
	     purchaseDateTxt.sendKeys(Keys.HOME);
	     
	     
	     
	     
	     purchaseDateTxt.sendKeys("0101");
	     
	     purchaseDateTxt.sendKeys(Keys.TAB);
	     
	     priceTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	     priceTxt.sendKeys("100000.00");
	     Thread.sleep(2000);
	     
	     quantityTxt.click();
	     
	     quantityTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	     
	     quantityTxt.sendKeys("1");
	     
	     Thread.sleep(2000);
	     
	     quantityTxt.sendKeys(Keys.TAB);
	     
	     System.out.println("********** Entered deperciationTab Tab *****************************");
	     
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deperciationTab));
		 deperciationTab.click();
			
			
		 Thread.sleep(3000);
			
		 getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(depreciationMethodDrpdwn));
		 
		 Select s1 = new Select(depreciationMethodDrpdwn);
		
		 s1.selectByVisibleText("Straigh tLine");
		
		 Thread.sleep(2000);
		 
		 /*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitsOfUsage));
		 Select s2 = new Select(unitsOfUsage);
		 s1.selectByVisibleText(excelReader.getCellData(xlSheetName, 144, 6));*/
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(depaerciationYearsTxt));
	     depaerciationYearsTxt.click();
		
		 depaerciationYearsTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
		 depaerciationYearsTxt.sendKeys("5");
		
		 Thread.sleep(1000);
		
		
		 depaerciationYearsTxt.sendKeys(Keys.TAB);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dateOfCommencementTxt));
		 dateOfCommencementTxt.click();
		 dateOfCommencementTxt.sendKeys(Keys.HOME);
		 dateOfCommencementTxt.sendKeys("0101");
		 dateOfCommencementTxt.sendKeys(Keys.TAB);
		
	    System.out.println("********************* Accounting Links Tab  Executes.............   *****************************");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountingLinksTab));
		accountingLinksTab.click();
		
		
		System.out.println("*******************Account Creation Starts From Here *********************************");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assetACCode));
		assetACCode.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
        nameTxt.click();
       
        nameTxt.sendKeys("FAAccountType");
       
        Thread.sleep(2000);
       
        nameTxt.sendKeys(Keys.TAB);
       
        codeTxt.sendKeys("FAAccountType");
		
       
        Select accountTypeDropdownSelect01 = new Select(accountTypeDropdown);
		
        accountTypeDropdownSelect01.selectByVisibleText("Fixed Assets");
		
		
		accountTypeDropdown.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatCreationPageSaveBtn));
		vatCreationPageSaveBtn.click();
    	
		System.out.println("*************************assetACCode Account Saved Successfully***********");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accumulatedDepreciation));
		accumulatedDepreciation.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
        nameTxt.click();
       
        nameTxt.sendKeys("FA Accumalted Exp AC");
       
        Thread.sleep(2000);
       
        nameTxt.sendKeys(Keys.TAB);
       
        codeTxt.sendKeys("FA Accumalted Exp AC");
		
       
        Select accountTypeDropdownSelect11 = new Select(accountTypeDropdown);
		
        accountTypeDropdownSelect11.selectByVisibleText("Expenses");
		
		
		accountTypeDropdown.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatCreationPageSaveBtn));
		vatCreationPageSaveBtn.click();
    	
		System.out.println("*************************accumulatedDepreciation Account Saved Successfully***********");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepreciationPLCode));
		DepreciationPLCode.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
        nameTxt.click();
       
        nameTxt.sendKeys("FA Dep PL AC");
       
        Thread.sleep(2000);
       
        nameTxt.sendKeys(Keys.TAB);
       
        codeTxt.sendKeys("FA Dep PL AC");
		
       
        Select accountTypeDropdownSelect1 = new Select(accountTypeDropdown);
		
		accountTypeDropdownSelect1.selectByVisibleText("Assets");
		
		
		accountTypeDropdown.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatCreationPageSaveBtn));
		vatCreationPageSaveBtn.click();
		
		System.out.println("*************************Depreciation PL Code Account Saved Successfully***********");
		
        Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesAccount));
		salesAccount.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
        nameTxt.click();
       
        nameTxt.sendKeys("FA SALES AC");
       
        Thread.sleep(2000);
       
        nameTxt.sendKeys(Keys.TAB);
       
        codeTxt.sendKeys("FA SALES AC");
		
       
        Select accountTypeDropdownSelect2 = new Select(accountTypeDropdown);
		
		accountTypeDropdownSelect2.selectByVisibleText("Sales");
		
		
		accountTypeDropdown.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatCreationPageSaveBtn));
		vatCreationPageSaveBtn.click();
		
		System.out.println("*************************Sales Account Account Saved Successfully***********");
		
        Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(disposalGain));
		disposalGain.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
        nameTxt.click();
        nameTxt.sendKeys("FA DISPOSAL GAIN AC");
       
        Thread.sleep(2000);
       
        nameTxt.sendKeys(Keys.TAB);
        codeTxt.sendKeys("FA DISPOSAL GAIN AC");
		
        Select accountTypeDropdownSelect3 = new Select(accountTypeDropdown);
		accountTypeDropdownSelect3.selectByVisibleText("Expenses");
		
		accountTypeDropdown.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatCreationPageSaveBtn));
		vatCreationPageSaveBtn.click();
		
		System.out.println("*************************Disposal Gain Account Saved Successfully***********");
		
		
		Thread.sleep(2000);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(disposalLoss));
		disposalLoss.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
        nameTxt.click();
       
        nameTxt.sendKeys("FA DISPOSAL LOSS AC");
       
        Thread.sleep(2000);
       
        nameTxt.sendKeys(Keys.TAB);
       
        codeTxt.sendKeys("FA DISPOSAL LOSS AC");
		
       
        Select accountTypeDropdownSelect4 = new Select(accountTypeDropdown);
		accountTypeDropdownSelect4.selectByVisibleText("Expenses");
		
		accountTypeDropdown.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vatCreationPageSaveBtn));
		vatCreationPageSaveBtn.click();
			
		System.out.println("*************************Disposal Loss Account Saved Successfully***********");
		
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assetACCodeTxt));
		assetACCodeTxt.click();
		
		assetACCodeTxt.sendKeys("FAAccountType");
		
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accumulatedDepreciationTxt));
		accumulatedDepreciationTxt.click();
		
		accumulatedDepreciationTxt.sendKeys("FA Accumalted Exp AC");
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepreciationPLCodeTXt));
		DepreciationPLCodeTXt.click();
		
		DepreciationPLCodeTXt.sendKeys("FA Dep PL AC");
		
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ALsalesAccountTxt));
		ALsalesAccountTxt.click();
		
		ALsalesAccountTxt.sendKeys("FA SALES AC");
		
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(disposalGainTxt));
		disposalGainTxt.click();
		
		disposalGainTxt.sendKeys("FA DISPOSAL GAIN AC");
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(disposalLossTxt));
		disposalLossTxt.click();
		
		disposalLossTxt.sendKeys("FA DISPOSAL LOSS AC");
		
       Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
        String expSavingMessage="Saved Successfully";
		
		String actSavingMessage=checkValidationMessage(expSavingMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();
		
		System.out.println("actSavingMessage   : "+actSavingMessage  +" Value Expected : "+expSavingMessage);
		
		if (actSavingMessage.equalsIgnoreCase(expSavingMessage)  )
		{
			System.out.println(" Test Pass: Account Saved SuccessFully");
			return true;
			
		} 
		else
		{
			System.err.println(" Test Fail: Account Saved SuccessFully");
			return false;
		}
		
	}
	


	public boolean checkPostingWithStraightLineInPostDepreciationEntries() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssets_Menu));
		fixedAssets_Menu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssetsTransMenu));
		fixedAssetsTransMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postDepreciationEntriesMenu));
		postDepreciationEntriesMenu.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postIndividualVoucherChkbox));
		
		if (postIndividualVoucherChkbox.isSelected()==true)
		{
			postIndividualVoucherChkbox.click();
			
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assertUsageDateTxt));
		assertUsageDateTxt.click();
		
		assertUsageDateTxt.sendKeys(Keys.HOME);
		
		
		
		assertUsageDateTxt.sendKeys("0201");
		
		Thread.sleep(3000);
		
		listingPostingToTheGroupdrpdwn.sendKeys("FixedAssetsStraightLine");
		
		Thread.sleep(3000);
		
		listingPostingToTheGroupdrpdwn.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postDepreciationOkBtn));
		postDepreciationOkBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(informationYesBtn));
		
		if (informationYesBtn.isDisplayed()==true)
		{
			informationYesBtn.click();
			
		}
		
        String expMessage="Voucher Posted Sucessfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(postDepreciationCancelBtn));
		postDepreciationCancelBtn.click();
		
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	
	@FindBy(xpath="//*[@id='tblBodyTransRender']/tr/td[4]")
    private static List<WebElement>  voucherBodyGridVouRowCount;

	
	@FindBy(xpath="//*[@id='tblBodyTransRender']/tr/td[2]/input")
    private static List<WebElement>  voucherGridIndexChkboxlist;
	
	public boolean checkPostedVoucherInFixedAssetDepreciationVoucherWithStraightLine() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
		 financialMenu.click();
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		 financialsTransactionMenu.click();
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finTransJournalsMenu));
		 finTransJournalsMenu.click();
		
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssetDepreciationVoucherMenu));
		 fixedAssetDepreciationVoucherMenu.click();
		 
		 Thread.sleep(2000);
		
		 int VoucherBodyGridVouRowCount=voucherBodyGridVouRowCount.size();
		 
		 
		 
			for (int i = 0; i < VoucherBodyGridVouRowCount; i++) 
			{
				String docNo=voucherBodyGridVouRowCount.get(i).getText();
				
				
				if (docNo.equalsIgnoreCase("Dep:2")) 
				{
					voucherGridIndexChkboxlist.get(i).click();
					
					break;
				}
			}
		
		 Thread.sleep(2000);
			
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		 editBtn.click();
		 
		 
		 boolean loading=checkLoadingMessage();
		
		

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		 String actDocno      =documentNumberTxt.getAttribute("value");
		 String actDate       =dateTxt.getAttribute("value");
		 String actDepartment =departmentTxt.getAttribute("value");
		 String actFixedAsset =fixedAssetTxt.getAttribute("value");
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		 String actAccRow1    =select1stRow_1stColumn.getText();
		 String actDebitRow1  =select1stRow_2ndColumn.getText();
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		 String actAccRow2    =select2ndRow_1stColumn.getText();
		 String actCreditRow2 =select2ndRow_3rdColumn.getText();
		
		 
		 DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
	     Date date=new Date();
        		
	     String docdate = df.format(date); 
	 
	     System.out.println("Current Date  : "+docdate);
	     
		 
		 String expDocno      ="Dep:2";
		 String expFixedAsset ="Bike";
		 String expAccRow1    ="FA Dep PL AC";
		 String expDebitRow1  ="54.64";
		 String expAccRow2    ="FA Accumalted Exp AC";
		 String expCreditRow2 ="54.64";
		 
		 
		 System.out.println("actDocno       :"+actDocno       +" Value Expected     :" +expDocno);
		 System.out.println("actFixedAsset  :"+actFixedAsset  +" Value Expected     :" +expFixedAsset);
		 System.out.println("actAccRow1     :"+actAccRow1     +" Value Expected     :" +expAccRow1);
		 System.out.println("actDebitRow1   :"+actDebitRow1   +" Value Expected     :" +expDebitRow1);
		 System.out.println("actAccRow2     :"+actAccRow2     +" Value Expected     :" +expAccRow2);
		 System.out.println("actCreditRow2  :"+actCreditRow2  +" Value Expected     :" +expCreditRow2);
		 
		
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		 new_CloseBtn.click();
		 
			if (actDocno.equalsIgnoreCase(expDocno) && 
					actFixedAsset.equalsIgnoreCase(expFixedAsset) && 
					actAccRow1.equalsIgnoreCase(expAccRow1) && 
					actDebitRow1.equalsIgnoreCase(expDebitRow1) && 
					actAccRow2.equalsIgnoreCase(expAccRow2) && 
					actCreditRow2.equalsIgnoreCase(expCreditRow2) ) 
			{
				return true;
			}
			else
			{
				return false;
			}
		}
	
	


	//Reports Starts From Here 

	  @FindBy(xpath = "//a[@id='50010']//span[contains(text(),'Reports')]")
	  private static WebElement assetReportsMenu;
		
	  @FindBy(xpath = "//a[@id='614']")
	  private static WebElement transferOfAssetReportMenu;
		
	  @FindBy(xpath = "//a[@id='615']")
	  private static WebElement assetUsageReportMenu;
		
	  @FindBy(xpath = "//a[@id='616']")
	  private static WebElement addAssetsValueReportMenu;
		
	  @FindBy(xpath = "//a[@id='618']")
	  private static WebElement componentAddedReportMenu;
		
	  @FindBy(xpath = "//a[@id='617']")
	  private static WebElement reduceAssetValueReportMenu;
		
	  @FindBy(xpath = "//a[@id='619']")
	  private static WebElement componentReduceReportMenu;
		
	  @FindBy(xpath = "//a[@id='620']")
	  private static WebElement disposalOfAssetReportMenu;
		
	  @FindBy(xpath = "//a[@id='612']")
	  private static WebElement fixedAssetReportMenu;
		
	  @FindBy(xpath = "//a[@id='613']")
	  private static WebElement depreciationScheduleReportMenu;
		
	  

		@FindBy(xpath="//div[@id='spnhiererchy']//select[@id='iTreeId']")
		private static WebElement sl_Dropdown1;
		
		@FindBy(xpath="//div[@id='spnhiererchy']//select[@id='iTreeViewId']")
		private static WebElement sl_Dropdown2;
		
		@FindBy(xpath="//label[@id='oncheckaccorderan']")
		private static WebElement sl_oncheckBtn;
		
		@FindBy(xpath="//input[@id='txtsrch-term']")
		private static WebElement sl_SearhItemTxt;
		
		@FindBy(xpath="//i[contains(@class,'icon-search theme_color-inverse')]")
		private static WebElement sl_SearchBtn;
		
		@FindBy(xpath="//input[@id='cmbUserTypeMaster']")
		private static WebElement sl_MasterTypeTxt;
		
		@FindBy(xpath="//span[contains(@class,'icon-font6 icon-external-module theme_icon-color')]")
		private static WebElement sl_AutoAdjustColumnsBtn;
		
		@FindBy(xpath="//a[contains(@class,'clsMasterName')][contains(text(),'ITEMS GROUP')]")
		private static WebElement sl_ItemGroupsBtn;
		
		@FindBy(xpath="//input[@id='liSelectAllMasters']")
		private static WebElement sl_HeaderSelectChkBox;
		
		@FindBy(xpath="//i[contains(@class,'icon-font6 icon-schedule')]")
		private static WebElement sl_ScheduleBtn;
		
		@FindBy(xpath="//i[contains(@class,'icon-font6 icon-custamize')]")
		private static WebElement sl_CustomizeBtn;
		
		@FindBy(xpath="//i[contains(@class,'icon-font6 icon-filter')]")
		private static WebElement sl_FilterBtn;
		
		@FindBy(xpath="//i[contains(@class,'icon-font6 icon-ok')]")
		private static WebElement sl_OkBtn;
		
		@FindBy(xpath="//i[contains(@class,'icon-font6 icon-close')]")
		private static WebElement sl_CloseBtn;
		
		@FindBy(xpath="//select[@id='DateOptions_']")
		private static WebElement sl_DateOptionDropdown;
		
		@FindBy(xpath="//input[@id='MasterSingle__1']")
		private static WebElement reportAssetDrpDwn;
		
		
		@FindBy(xpath="//input[@id='id_starting_date_']")
		private static WebElement sl_StartDateTxt;
		
		@FindBy(xpath="//td[@id='id_starting_date__input_image']//span[contains(@class,'icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow')]")
		private static WebElement sl_StartDateCalanderBtn;
		
		@FindBy(xpath="//input[@id='id_ending_date_']")
		private static WebElement sl_EndDateTxt;
		
		@FindBy(xpath="//td[@id='id_ending_date__input_image']//span[contains(@class,'icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow')]")
		private static WebElement sl_EndDateCalanderBtn;
		
		@FindBy(xpath="//input[@id='RITCheckbox__3']")
		private static WebElement sl_IncludeServiceTypeItemChkBox;
		
		@FindBy(xpath="//select[@id='RITLayout_']")
		private static WebElement sl_LayoutDropdown;
		
		@FindBy(xpath="//select[@id='RITOutput_']")
		private static WebElement sl_OutputDropdown;
		
		@FindBy(xpath="//input[@id='selectAllMasters_']")
		private static WebElement sl_SelectAllItemsChkBox;
		
		@FindBy(xpath="//div[contains(@class,'first')]//input[contains(@class,'btn btn-xs Fbutton')]")
		private static WebElement sl_FirstBtn;
		
		@FindBy(xpath="//div[contains(@class,'prev')]//input[contains(@class,'btn btn-xs Fbutton')]")
		private static WebElement sl_PreviousBtn;
		
		@FindBy(xpath="//input[@id='btn1']")
		private static WebElement sl_PageBtn;
		
		@FindBy(xpath="//div[contains(@class,'next')]//input[contains(@class,'btn btn-xs Fbutton')]")
		private static WebElement sl_NextBtn;
		
		@FindBy(xpath="//div[contains(@class,'End')]//input[contains(@class,'btn btn-xs Fbutton')]")
		private static WebElement sl_EndBtn;
		
		@FindBy(xpath="//tr[1]//td[8]//div[1]//label[1]")
		private static WebElement sl_1stRowChkBox;
		
		@FindBy(xpath="//tr[2]//td[8]//div[1]//label[1]")
		private static WebElement sl_2ndRowChkBox;
		
		

		@FindBy(xpath="//i[contains(@class,'theme_button_color dropdown-toggle')]")
		private static WebElement report_sortingBtn;
				
		@FindBy(xpath="//i[contains(@class,'icon-font6 icon-refresh')]")
		private static WebElement report_RefreshBtn;
				
		@FindBy(xpath="//i[contains(@class,'icon-font6 icon-print')]")
		private static WebElement report_PrintBtn;
				
		@FindBy(xpath="//i[contains(@class,'icon-font6 icon-export')]")
		private static WebElement report_ExportBtn;
				
		@FindBy(xpath="//i[contains(@class,'icon-font6 icon-graph ImagesinArabic')]")
		private static WebElement report_GraphBtn;
				
		@FindBy(xpath="//i[contains(@class,'icon-font6 icon-backtrack')]")
		private static WebElement report_BackTrackBtn;
				
		@FindBy(xpath="//span[@id='reportCustomize_']//i[contains(@class,'icon-font6 icon-custamize')]")
		private static WebElement report_ReportCustomizeBtn;
				
		@FindBy(xpath="//span[@id='filterIcon_']//i[contains(@class,'icon-font6 icon-filter')]")
		private static WebElement report_FilterBtn;
		
		@FindBy(xpath="//span[@id='idFilterCustomizeIcon']")
		private static WebElement  report_FilterCustomizeBtn;
		
		@FindBy(xpath="//span[@id='a']")
		private static WebElement  report_Filter_FilterBtn;
		
		@FindBy(xpath="//*[@id='551_0_AdvanceFilter']/table/tbody/tr/td[1]/select")
		private static WebElement  reportsAdvanced_WhereDropDown;
		
		@FindBy(xpath="//tbody//input[@placeholder='Select Field']")
		private static WebElement  reportsAdvanced_SelectFielddropdown;
		
		@FindBy(xpath="//table/tbody/tr/td[3]/select")
		private static WebElement  reportsAdvanced_SelectOperatorDropDown;
		
		@FindBy(xpath="//table/tbody/tr/td[4]/select")
		private static WebElement  reportsAdvanced_ComparewithDropDown;
		
		@FindBy(xpath="//table/tbody/tr/td[5]/input")
		private static WebElement  reportsAdvanced_ValueTxt;
		
		@FindBy(xpath="//input[@id='advancefilter_master_551_0']")
		private static WebElement  reportsAdvanced_EnterValue;
		
		@FindBy(xpath="//table/tbody/tr/td[7]/span")
		private static WebElement  reportsAdvanced_AddIconBtn;
		
		@FindBy(xpath="//table/tbody/tr/td[6]/span")
		private static WebElement  reportsAdvanced_DeleteIcon1Btn;
		
		@FindBy(xpath="//table/tbody/tr[2]/td[6]/span")
		private static WebElement  reportsAdvanced__DeleteIcon2Btn;
		
		@FindBy(xpath="//div[@id='filterTree_551_0_AdvanceFilter']//a[@id='536870916']")
		private static WebElement  reports_AdvancedFilter_SelectField_WH_ExpandBtn;
		
		@FindBy(xpath="//div[@id='filterTree_551_0_AdvanceFilter']//a[@id='5058']")
		private static WebElement  reports_AdvancedFilter_SelectField_WH_NameBtn;
		
		@FindBy(xpath="//input[@id='chkSaveFilter']")
		private static WebElement  report_FilterSaveFilterPermanentlyChkBox;
		
		@FindBy(xpath="//span[@id='filterRefresh']")
		private static WebElement  report_Filter_RefreshBtn;
		
		@FindBy(xpath="//i[contains(@class,'icon icon-ok')]")
		private static WebElement  report_FilterOkBtn;
		
		@FindBy(xpath="//i[contains(@class,'icon icon-close')]")
		private static WebElement  report_FilterCancelBtn;
				
		@FindBy(xpath="//div[@id='dvfilter_']//span[contains(@class,'icon-down-arrow icon-font8 dropdown-toggle theme_button_color')]")
		private static WebElement report_FilterDownBtn;
				
		@FindBy(xpath="//i[contains(@class,'icon-font6 icon-analyze')]")
		private static WebElement report_AnalyzeBtn;
				
		@FindBy(xpath="//i[contains(@class,'icon-font6 icon-cross-reference')]")
		private static WebElement report_CrossReferenceBtn;
				
		@FindBy(xpath="//i[contains(@class,'icon-font6 icon-options')]")
		private static WebElement report_OptionsBtn;
				
		@FindBy(xpath="//div[@id='dvoptions_']//span[contains(@class,'icon-down-arrow icon-font8 dropdown-toggle theme_button_color')]")
		private static WebElement report_OptionsDownBtn;
				
		@FindBy(xpath="//span[@id='reportClose_']//i[contains(@class,'icon-font6 icon-close')]")
		private static WebElement report_CloseBtn;
				
		@FindBy(xpath="//button[@id='frstPage_']")
		private static WebElement report_FirstBtn;
				
		@FindBy(xpath="//button[@id='prvPage_']")
		private static WebElement report_PreviousBtn;
				
		@FindBy(xpath="//button[@id='pgnumber_']")
		private static WebElement report_PageNoBtn;
				
		@FindBy(xpath="//button[@id='nxtPage_']")
		private static WebElement report_NextBtn;
				
		@FindBy(xpath="//button[@id='lstPage_']")
		private static WebElement report_LastBtn;
		
		@FindBy(xpath="//input[@id='txtSearchReport']")
		private static WebElement report_SearchTxt;
		
		@FindBy(xpath="//*[@id='tblBodyReportRender']/tr/td")
		private static List<WebElement> reportList;
		
	  
		@FindBy(xpath="//*[@id='tblBodyReportRender']/tr[1]/td")
		private static List<WebElement> report1stRowList;
		
		@FindBy(xpath="//*[@id='tblBodyReportRender']/tr[2]/td")
		private static List<WebElement> report2ndRowList;
		
		@FindBy(xpath="//*[@id='tblBodyReportRender']/tr[3]/td")
		private static List<WebElement> report3rdRowList;
		
		@FindBy(xpath="//*[@id='tblBodyReportRender']/tr[4]/td")
		private static List<WebElement> report4thRowList;
		
		@FindBy(xpath="//*[@id='tblBodyReportRender']/tr[5]/td")
		private static List<WebElement> report5thRowList;
		
		@FindBy(xpath="//*[@id='tblBodyReportRender']/tr[6]/td")
		private static List<WebElement> report6thRowList;
		
		

public boolean checkAssetUsageReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	
	System.out.println("****************************checkAssetUsageReport Method Executes ****");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssets_Menu));
	fixedAssets_Menu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssetsTransMenu));
	fixedAssetsTransMenu.click();
	
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assetReportsMenu));
	assetReportsMenu.click();
	
	
	JavascriptExecutor js = (JavascriptExecutor) getDriver();
	js.executeScript("arguments[0].scrollIntoView();", depreciationScheduleReportMenu);
		
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assetUsageReportMenu));
	assetUsageReportMenu.click();
	
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
	sl_DateOptionDropdown.click();
	Select s=new Select(sl_DateOptionDropdown);
	
	s.selectByValue("1");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reportAssetDrpDwn));
	reportAssetDrpDwn.click();
	reportAssetDrpDwn.sendKeys("bike");
	
	Thread.sleep(2000);
	reportAssetDrpDwn.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
	sl_OkBtn.click();
	
	
    boolean novalidationConfirmationMessage =validationConfirmationMessage.getText().isEmpty();
	
	String actvalidationConfirmationMessage = Boolean.toString(novalidationConfirmationMessage);
	String expvalidationConfirmationMessage = "true";
	
	
	int row1Count = report1stRowList.size();
	ArrayList<String> row1Array = new ArrayList<String>();
	
    System.out.println("***********row1Count*********"+row1Count);
   
	for(int i=0;i<row1Count;i++)
	{
		String data = report1stRowList.get(i).getText();
		
		if(i==3)
		{
			data="Date Field";
		}
		row1Array.add(data);
	}
	
	String actRow1List = row1Array.toString();
	String expRow1List = "[1, FAItem, FAItem, Date Field, , 10.00, 10.00]";
	
	
	int row2Count = report2ndRowList.size();
	ArrayList<String> row2Array = new ArrayList<String>();
	
    System.out.println("***********row2Count*********"+row2Count);
   
	for(int i=0;i<row2Count;i++)
	{
		String data = report2ndRowList.get(i).getText();
		
		if(i==3)
		{
			data="Date Field";
		}
		row2Array.add(data);
	}
	
	String actRow2List = row2Array.toString();
	String expRow2List = "[2, FAItem, FAItem, Date Field, , 10.00, 10.00]";
	
	
	int row3Count = report3rdRowList.size();
	ArrayList<String> row3Array = new ArrayList<String>();
	
    System.out.println("***********row3Count*********"+row3Count);
   
	for(int i=0;i<row3Count;i++)
	{
		String data = report3rdRowList.get(i).getText();
		
		if(i==3)
		{
			data="Date Field";
		}
		row3Array.add(data);
	}
	
	String actRow3List = row3Array.toString();
	String expRow3List = "[3, FAItem, FAItem, Date Field, 10.00, 5.00, 15.00]";
	
	
	int row4Count = report4thRowList.size();
	ArrayList<String> row4Array = new ArrayList<String>();
	
    System.out.println("***********row4Count*********"+row4Count);
   
	for(int i=0;i<row4Count;i++)
	{
		String data = report4thRowList.get(i).getText();
		
		if(i==3)
		{
			data="Date Field";
		}
		row4Array.add(data);
	}
	
	String actRow4List = row4Array.toString();
	String expRow4List = "[4, Grand Total, , Date Field, 10.00, 25.00, 35.00]";
		
	System.out.println("actRow1List  : "+actRow1List);
	System.out.println("expRow1List  : "+expRow1List);
	
	
	System.out.println("actRow2List  : "+actRow2List);
	System.out.println("expRow2List  : "+expRow2List);
	
	System.out.println("actRow3List  : "
	+actRow3List);
	System.out.println("expRow3List  : "+expRow3List);
	
	System.out.println("actRow4List  : "+actRow4List);
	System.out.println("expRow4List  : "+expRow4List);

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
	report_CloseBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
	sl_CloseBtn.click();
	
	
			
	if(actRow1List.equalsIgnoreCase(expRow1List) && 
			actRow2List.equalsIgnoreCase(expRow2List) && 
			actRow3List.equalsIgnoreCase(expRow3List) && 
			actRow4List.equalsIgnoreCase(expRow4List) )
	{
		System.out.println("Test Pass : Report Screen Displayed  ");
		return true;
	}
	else
	{
		System.out.println("Test Fail : Report Screen Displayed  ");
		return false;
	}
}

	




public boolean checkFixedAssetReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	System.out.println("****************************checkFixedAssetReportOptions Method Executes ****");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssets_Menu));
	fixedAssets_Menu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssetsTransMenu));
	fixedAssetsTransMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assetReportsMenu));
	assetReportsMenu.click();
	
	JavascriptExecutor js = (JavascriptExecutor) getDriver();
	js.executeScript("arguments[0].scrollIntoView();", depreciationScheduleReportMenu);
		
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fixedAssetReportMenu));
	fixedAssetReportMenu.click();
	
	boolean novalidationConfirmationMessage =validationConfirmationMessage.getText().isEmpty();
	
	String actvalidationConfirmationMessage = Boolean.toString(novalidationConfirmationMessage);
	String expvalidationConfirmationMessage = "true";
	
	System.out.println("validationConfirmationMessage : "+actvalidationConfirmationMessage+" Value Expected : "+expvalidationConfirmationMessage);
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
	sl_DateOptionDropdown.click();
	Select s=new Select(sl_DateOptionDropdown);
	
	s.selectByValue("1");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
	sl_SelectAllItemsChkBox.click();
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
	sl_OkBtn.click();
	
	
    boolean novalidationConfirmationMessage1 =validationConfirmationMessage.getText().isEmpty();
	
	String actvalidationConfirmationMessage1 = Boolean.toString(novalidationConfirmationMessage1);
	String expvalidationConfirmationMessage1 ="true";

	
	System.out.println("validationConfirmationMessage1 : "+actvalidationConfirmationMessage1+" Value Expected : "+expvalidationConfirmationMessage1);

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CrossReferenceBtn));
	
    int reportVocGridListCount = reportVocGridList.size();
	
	
	String actreportVocGridListCount=Integer.toString(reportVocGridListCount);
	String expreportVocGridListCount="5";
	
	System.err.println("reportVocGridListCount   : "+actreportVocGridListCount);
	
	

	int row1Count = report1stRowList.size();
	ArrayList<String> row1Array = new ArrayList<String>();
	
    System.out.println("***********row1Count*********"+row1Count);
   
	for(int i=2;i<row1Count;i++)
	{
		String data = report1stRowList.get(i).getText();
		
		if(i==9)
		{
			data="Date Field";
		}
		row1Array.add(data);
	}
	
	String actRow1List = row1Array.toString();
	String expRow1List = "[FixedAssetsStraightLine, FixedAssetsStraightLine, , , , , , Date Field, , , , , Straight Line, , , ]";
	
	
	int row2Count = report2ndRowList.size();
	ArrayList<String> row2Array = new ArrayList<String>();
	
    System.out.println("***********row2Count*********"+row2Count);
   
	for(int i=2;i<row2Count;i++)
	{
		String data = report2ndRowList.get(i).getText();
		
		if(i==9)
		{
			data="Date Field";
		}
		if(i==16)
		{
			data="Date Field";
		}
		row2Array.add(data);
	}
	
	String actRow2List = row2Array.toString();
	String expRow2List = "[Bike, Bike, Vehicle, AMERICA, , , , Date Field, 1.00, 5.00, 1,00,000.00, , Straight Line, FA Dep PL AC, Date Field, ]";
	
	
	int row3Count = report3rdRowList.size();
	ArrayList<String> row3Array = new ArrayList<String>();
	
    System.out.println("***********row3Count*********"+row3Count);
   
	for(int i=2;i<row3Count;i++)
	{
		String data = report3rdRowList.get(i).getText();
		
		if(i==9)
		{
			data="Date Field";
		}
		
		row3Array.add(data);
	}
	
	String actRow3List = row3Array.toString();
	String expRow3List = "[FixedAssetsBasedOnUsage, FixedAssetsBasedOnUsage, , , , , , Date Field, , , , , Based on Usage, , , ]";
	
	
	int row4Count = report4thRowList.size();
	ArrayList<String> row4Array = new ArrayList<String>();
	
    System.out.println("***********row4Count*********"+row4Count);
   
	for(int i=2;i<row4Count;i++)
	{
		String data = report4thRowList.get(i).getText();
		
		if(i==9)
		{
			data="Date Field";
		}
		if(i==16)
		{
			data="Date Field";
		}
		row4Array.add(data);
	}
	
	String actRow4List = row4Array.toString();
	String expRow4List = "[FAItem, FAItem, Building, INDIA, , , Vendor A, Date Field, 1.00, , 20,000.00, , Based on Usage, ALDepreciationPLCode, Date Field, ]";
		
	int row5Count = report5thRowList.size();
	ArrayList<String> row5Array = new ArrayList<String>();
	
    System.out.println("***********row5Count*********"+row5Count);
   
	for(int i=0;i<row5Count;i++)
	{
		String data = report5thRowList.get(i).getText();
		
		row5Array.add(data);
	}
	
	String actRow5List = row5Array.toString();
	String expRow5List = "[5, Grand Total, , , , , , , , , 2.00, 5.00, 1,20,000.00, , , , , ]";
		
	System.out.println("actRow1List  : "+actRow1List);
	System.out.println("expRow1List  : "+expRow1List);
	
	
	System.out.println("actRow2List  : "+actRow2List);
	System.out.println("expRow2List  : "+expRow2List);
	
	System.out.println("actRow3List  : "+actRow3List);
	System.out.println("expRow3List  : "+expRow3List);
	
	System.out.println("actRow4List  : "+actRow4List);
	System.out.println("expRow4List  : "+expRow4List);

	System.out.println("actRow5List  : "+actRow5List);
	System.out.println("expRow5List  : "+expRow5List);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
	report_CloseBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
	sl_CloseBtn.click();
	
	
			
	if(actRow1List.equalsIgnoreCase(expRow1List) && 
			actRow2List.equalsIgnoreCase(expRow2List) && 
			actRow3List.equalsIgnoreCase(expRow3List) && 
			actRow4List.equalsIgnoreCase(expRow4List)  && 
			actRow5List.equalsIgnoreCase(expRow5List))
	{
		System.out.println("Test Pass : Report Screen Displayed  ");
		return true;
	}
	else
	{
		System.out.println("Test Fail : Report Screen Displayed  ");
		return false;
	}
}





@FindBy(xpath="//*[@id='tblBodyReportRender']/tr/td[3]")
private static List<WebElement> reportVocGridList;





	
	
 public  SanityFixedAssetsPage(WebDriver driver)
{
	
	 PageFactory.initElements(driver, this);
	
}
	
}
