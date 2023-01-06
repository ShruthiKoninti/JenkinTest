package com.focus.repository;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;

public class ElementRepository 
{
	// Elements 

	@FindBy(xpath="//*[@id='txtUsername']")
	public static WebElement username;

	@FindBy(id="txtPassword")
	public static WebElement password;

	@FindBy(id="ddlCompany")
	public static WebElement companyDropDownList;

	@FindBy(id="btnSignin")
	public static WebElement signIn;

	@FindBy(xpath="//i[@title='Create Company']")
	public static WebElement companyCreateBtn;		

	@FindBy(xpath="//i[@title='Keyboard ']")
	public static WebElement keyboardBtn;

	@FindBy(xpath="//i[@title='Refresh']")
	public static WebElement refreshBtn;

	@FindBy(xpath="//input[@id='chkRememberMe']/following-sibling::span")
	public static WebElement rememberMeChk;

	@FindBy(xpath="//input[@id='chkRememberPwd']/following-sibling::span")
	public static WebElement rememberPwdChk;

	@FindBy(id="frgtPwd")
	public static WebElement frgtPwdLnk;

	@FindBy(xpath="//span[contains(@class,'icon-ok icon-font6')]")
	public static WebElement okButtonInCreateCompany;

	@FindBy(xpath="//span[contains(@class,'icon-close icon-font6')]")
	public static WebElement cancelButtonInCreateCompany;

	@FindBy(xpath="//div[@id='Layout_Popup']//input[2]")
	public static WebElement noIn;

	@FindBy(linkText="Create Company")
	public static WebElement createCompanyScreen;

	@FindBy(xpath="//*[@id='errmsgDiv']")
	public static WebElement mandatoryMsgs;

	//Fields of Company su Home Page
	@FindBy(xpath="//*[@id='companyLogo']/parent::a/following-sibling::ul/li[1]")
	public static WebElement companyName;

	@FindBy(xpath="//*[@id='mainHeader_MainLayout']/a/img")
	public static WebElement focusLogo;

	@FindBy(xpath="//*[@id='id_focus8_wrapper_default']/aside/section")
	public static WebElement menuBar;

	@FindBy(xpath="//*[@id='dashName']")
	public static WebElement dashboardName;

	@FindBy(xpath="//*[@id='dashIcons']")
	public static WebElement dashboardIcons;

	@FindBy(xpath="(//*[contains(text(),'SU')])[1]")
	public static WebElement userNameDisplay;
	
	@FindBy(xpath="(//*[contains(text(),'DemoUser')])[1]")
	public static WebElement userNameDemoDisplay;
	
	@FindBy(xpath="(//*[contains(text(),'Test')])[1]")
	public static WebElement userNameTestDisplay;
	
	@FindBy(xpath="//*[contains(text(),'UserAllOptionsST')]")
	public static WebElement userNameUserDisplay;
	
	@FindBy(xpath="//*[contains(text(),'UserAccountRestrictions')]")
	public static WebElement userNameUserRestrictionDisplay;

	/*@FindBy(xpath="//div[@id='id_mainlayoutmenu']/ul[2]/li[6]")
	public static WebElement userNameDisplayLogo;*/

	@FindBy(xpath="//*[@id='mainHeader_MainLayout']//li[7]/a")
		public static WebElement userNameDisplay_User;

	//div[@id='id_mainlayoutmenu']/ul[2]/li[7]
	@FindBy(xpath="//*[@id='companyLogo']")
	public static WebElement companyLogo;

	@FindBy(xpath="//div[@id='id_mainlayoutmenu']/ul[2]/li[6]//ul/li[1]")
	public static WebElement changePassword;

	@FindBy(xpath="//ul[@id='id_user_profile_icons']/li[3]/div/div/select")
	public static WebElement languageDropdownInLogout;

	@FindBy(xpath="//*[contains(text(),'Logout')]")
	public static WebElement logoutOption;
	
	
	
			@FindBy(xpath="//*[@id='id_user_profile_icons']/li[1]")
			public static WebElement changePasswordOption;
			
			
			@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[1]/a/i")
			public static WebElement changePasswordSaveBtn;

			/*@FindBy(xpath="//*[@id='myNavbar']/ul/li/span[2]/i")
					public static WebElement changePasswordCloseBtn;*/


			@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[2]/a/i")
			public static WebElement changePasswordCloseBtn;

			@FindBy(xpath="//*[@id='oldPassword']")
			public static WebElement changePasswordOldPasswordTxt;

			@FindBy(xpath="//*[@id='newPassword']")
			public static WebElement changePasswordNewPasswordTxt;

			@FindBy(xpath="//*[@id='confirmPassword']")
			public static WebElement changePasswordConfirmTxt;


			

	@FindBy(xpath="//div[@id='DashboardRenderHeader']/ul/li[1]/a")
	public static WebElement dashboard;


	@FindBy(xpath="//div[@id='DashboardRenderHeader']/ul/li[2]/a")
	public static WebElement dashBoardFirst;


	@FindBy(xpath="//ul[@id='ulCompanyDetails_HomePage']")
	public static WebElement companydetails;

	//Dash board Page
	@FindBy(xpath="//*[@id='Dashboard_ddlList']//a")
	public static WebElement labelDashboard ;

	@FindBy(xpath="//*[@id='Select_dash']")
	public static WebElement selectDashboard ;

	@FindBy(xpath="//*[@id='Dashboard_AddDash']")
	public static WebElement newAddDashBoard;

	@FindBy(xpath="//*[@id='Dashboard_Dash_Config']")
	public static WebElement dashboardCustomizationSettings;

	@FindBy(xpath="//div[@id='login_div']/div[1]//ul/li[3]//i")
	public static WebElement  loginPageRestoreBtn;

	@FindBy(xpath="//div[@id='login_div']/div[1]//ul/li[5]//i")
	public static WebElement  loginPageSecurityProtectedCompanyBtn;

	@FindBy(xpath="//input[@id='donotshow']")
	public static WebElement doNotShowCheckbox;

	@FindBy(xpath="//span[@class='pull-right']")
	public static WebElement closeBtnInDemoPopupScreen;

	//Financial Menu
	@FindBy (xpath="//span[@class='icon-financial icon-font1']")
	public static WebElement financialMenu;

	//Fixed Assets Menu
	@FindBy (xpath="//span[@class='icon-assets icon-font1']")
	public static WebElement fixedAssetsMenu;

	//Production Menu
	@FindBy (xpath="//span[@class='icon-production icon-font1']")
	public static WebElement productionMenu;

	//Point Of Sale Menu
	@FindBy (xpath="//div[contains(text(),'Point of Sale')]")
	public static WebElement pointOfSaleMenu;

	//Quality Of Control Menu
	@FindBy (xpath="//span[@class='icon-quotationanalysis icon-font1']")
	public static WebElement qualityControlMenu;

	//Settings Menu
	@FindBy (xpath="//a[@id='16']")
	public static WebElement SettingsMenu;

	@FindBy(xpath="//ul[@id='navigation_menu']/li[8]/ul/li/a/span")
	public static List<WebElement> SettingsSubMenusList;

	//Inventory Menu and Sub Menus	
	
	@FindBy(xpath="//a[@id='135']")
	public static WebElement  inventoryMenu; 

	@FindBy(xpath = "//div[@id='id_left_mainmenu']/ul/li/a")
	public static List<WebElement> menusList;

	@FindBy(xpath = "//*[@id='mainHeader_MainLayout']/nav/div/ul/li")
	public static List<WebElement> navbarList;

	@FindBy (xpath="//a[@id='dropdown03']")
	public static WebElement aboutQuickAcessMenu;

	@FindBy (xpath="//input[@id='id_ai_search_input']")
	public static WebElement aiSearchTextBox;

	@FindBy (xpath="(//div[@id='mainHeader_MainLayout']//div/input)[3]")
	public static WebElement searchTxtBox;

	@FindBy (xpath="//div[@id='id_mainlayoutmenu']/ul[2]/li[1]/a")
	public static WebElement productTour;

	@FindBy (xpath="//div[@id='id_mainlayoutmenu']/ul[2]/li[3]/a")
	public static WebElement alertsOptions;

	@FindBy (xpath="//div[@id='id_mainlayoutmenu']/ul[2]/li[2]/a")
	public static WebElement recentMenus;




	// AccountsPage

	@FindBy (xpath="//*[@id='1']/span")
	public static WebElement homeMenu;

	//Masters Menu	
	@FindBy (xpath="//*[@id='1000']")
	public static WebElement mastersMenu;

	//Accounts 
	@FindBy (xpath="//*[@id='1104']")
	public static WebElement accounts;


	//Master Main Header Fields		
	@FindBy(xpath="//a[@id='btnNew']/i")
	public static WebElement masterNewBtn;

	@FindBy(xpath="//a[@id='btnAddGroup']/i")
	public static WebElement masterAddGroupBtn; 

	@FindBy(xpath="//a[@id='btnMasterEdit']/i")
	public static WebElement masterEditBtn; 

	@FindBy(xpath="//a[@id='btnClone']/i")
	public static WebElement masterCloneBtn; 

	@FindBy(xpath="//a[@id='btnProperties']/i")
	public static WebElement masterPropertiesBtn; 

	@FindBy(xpath="//a[@id='btnDelete']/i")
	public static WebElement masterDeleteBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-group']")
	public static WebElement masterGroupBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-delete-all']")
	public static WebElement masterDeleteAllBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-closed-account']")
	public static WebElement masterCloseAccountOrItemOrUnitsOrDepartmentOrWareHouseOrStateOrCity; 

	@FindBy(xpath="//i[@class='icon-font6 icon-open-close-account']")
	public static WebElement masteropenCloseAccountOrItemOrDepartmentOrWareHouseOrState; 

	@FindBy(xpath="//i[@class='icon-font6 icon-import']")
	public static WebElement masterAdvanceMasterImportORExportBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-info']")
	public static WebElement masterAuthorInfoBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-up-arrow']")
	public static WebElement masterMoveUpBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-down-arrow']")
	public static WebElement masterMoveDownBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-sort']")
	public static WebElement masterSortBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-mass-update']")
	public static WebElement masterMassUpdateBtn; 

	@FindBy(xpath="//*[@id='btnXMLImport']")
	public static WebElement masterImportFromXmlBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-xmlexport']")
	public static WebElement masterExportFormatToXmlBtn; 

	@FindBy(xpath="//a[@id='btnCustomizeMaster']")
	public static WebElement masterCustamizemasterBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-options']")
	public static WebElement masterCustamizeViewBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-tree']")
	public static WebElement mastercustamizeTreeBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-stock-ledger']")
	public static WebElement masterLedgerBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-financial-1']")
	public static WebElement masterManageCreditBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-department-appropriation']")
	public static WebElement masterDepartmentAppropriationBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-budget']")
	public static WebElement masterBudgetBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-backtrack']")
	public static WebElement masterBackTrackBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-transfer']")
	public static WebElement masterTranferBtn; 

	@FindBy(xpath="//*[@id='btnNewSets']")
	public static WebElement masterNewSetBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-modifier']")
	public static WebElement masterModifierBtn; 

	@FindBy(xpath="//*[@id='btnSetType']")
	public static WebElement masterSetTypeBtn; 

	@FindBy(xpath="//i[@class='icon-convert icon-font6']")
	public static WebElement masterUnitConversionBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-alternate-product-1']")
	public static WebElement masterAlternateItemBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-related-product']")
	public static WebElement masterRelatedItemBtn; 

	@FindBy(xpath="//i[@class='icon-font6 icon-stock-ledger']")
	public static WebElement masterStockLedgerBtn; 



	@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[11]")
	public static List<WebElement> masterAccountsList;
	
	@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[10]")
	public static List<WebElement> masterWarehouseList;

	//New General Elements
	//Name
	@FindBy(xpath="//input[@id='sName']")
	public static WebElement nameTxt;

	//AccountType	
	@FindBy(xpath="//select[@id='iAccountType']")
	public static WebElement accountTypeDropdown;


	//CreditLimit
	@FindBy(xpath="//input[@id='fCreditLimit']")
	public static WebElement creditLimitTxt;

	//CreditDays
	@FindBy(xpath="//input[@id='iCreditDays']")
	public static WebElement creditdaysTxt;


	//ChequeDiscountLimit
	@FindBy(xpath="//input[@id='fChequeDiscountLimit']")
	public static WebElement chequeDiscountLimitTxt;


	//Rate Of Interest
	@FindBy(xpath="//input[@id='fRateofinterest']")
	public static WebElement rateofinterestTxt;


	//BankAccount
	@FindBy(xpath="//input[@id='iBankAc']")
	public static WebElement bankAccountTxt;

	@FindBy(xpath="//*[@id='iBankAc_input_image']/span")
	public static WebElement bankAccount_ExpansionBtn;

	@FindBy(xpath="//*[@id='iBankAc_input_settings']/span")
	public static WebElement bankAccount_SettingBtn;



	@FindBy(xpath="(//*[@id='btnMasterSaveClick'])[1]")
	public static WebElement saveBtn;

	@FindBy(xpath="//*[@id='btnCustomizeClose']/i")
	//@FindBy(xpath="(//i[@class='icon-close-1 hiconright2'])[1]")
	public static WebElement closeBtn;

	@FindBy(xpath="//*[@id='btnCustomizeClose']/i")
	public static WebElement Settings_closeBtn;

	@FindBy(xpath="//*[@id='btnDelete']/a/i")
	public static WebElement deleteBtn;

	@FindBy(xpath="//div[@id='idGlobalError']")
	public static WebElement validationConfirmationMessage;



	@FindBy(xpath="//*[@id='GenerateNewMasterModel']/div/div[1]/div[2]/div/ul/li[2]/button[5]")
	public static WebElement masterCloseBtn;

	@FindBy(xpath="//*[@id='LandingGridBody']/tr[1]/td[11]")
	public static WebElement accountCreation;

	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/ol[1]/li[4]/span[1]")
	public static WebElement clickOnVendorGroup;

	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/form[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/input[1]")
	public static WebElement getVendorGroupName;

	@FindBy(xpath="//button[@class='btn-link theme_icon-color btnGenerateMasterclose closeMasterPartial clsMasterSaveClose']//i[@class='icon-font6 icon-close']")
	public static WebElement closeAccountCreationScreen;

	@FindBy(xpath="//input[@id='chkRetainSelection']")
	public static WebElement accountMasterRetain;

	@FindBy(xpath="//input[@id='liSelectAllMasters']")
	public static WebElement accountMasterSelect;

	@FindBy(xpath="(//label[@id='lblnchecked']/input)[1]")
	public static WebElement accountFirstCheckBox;

	@FindBy(xpath="(//label[@id='lblnchecked']/input)[2]")
	public static WebElement accountSecondCheckBox;

	@FindBy(xpath="//button[@class='btn-link theme_icon-color btnGenerateMasterclose closeMasterPartial clsMasterSaveClose']")
	public static WebElement closeAccMasterCreationScreen;


	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/ul[1]/li[2]/button[5]/i[1]")
	public static WebElement closeMaster;


	@FindBy(xpath="//div[@id='innerDivDeleteMasterModal']/div[2]/div[1]//label")
	public static WebElement getMsgOnDelete;

	@FindBy(xpath="//button[@id='btnOkForDelete']")
	public static WebElement clickOnOkInDelete;

	@FindBy(xpath="//button[@id='btnCancelForDelete']")
	public static WebElement clickOnCancelInDelete;

	@FindBy(xpath = "//*[@id='idGlobalError']/div/div[2]")
	public static WebElement errorMessage;

	@FindBy(xpath = "(//*[@id='idGlobalError']/div/div[2])[2]")
	public static WebElement ExceptionerrorMessage;

	@FindBy(xpath = "(//*[@id='idGlobalError']/div/div[2])[1]")
	public static WebElement ExceptionerrorMessage1;
	
	@FindBy(xpath = "//*[@id='idGlobalError']/div/div[1]/button")
	public static WebElement errorMessageCloseBtn;
	
	@FindBy(xpath = "//*[@id='idGlobalError']/div[2]/div[1]/button")
	public static WebElement ExceptionerrorMessageCloseBtn;
	

	@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[8]//input")
	public static List<WebElement> masterGridBodyChkbox;

	@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[11]")
	public static List<WebElement> masterGridBodyName;

	@FindBy (xpath="//span[@id='spnHeaderText']")
	public static WebElement accountsTitle;


	// Item Page

	@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[12]")
	public static List<WebElement> masterItemGridBodyName;

	@FindBy(xpath="//*[@id='iTreeId']")
	public static WebElement masterTreeIdDropdown; 

	@FindBy(xpath="//*[@id='spnhiererchy']/div/div[1]/div/div/button/span")
	public static WebElement masterTreeIDToggleBtn; 

	@FindBy(xpath="//*[@id='spnhiererchy']/div/div[1]/div/div/ul/li[1]/a")
	public static WebElement masterTreeIDToggle_CreateTreeBtn; 


	@FindBy(xpath="//*[@id='spnhiererchy']/div/div[1]/div/div/ul/li[2]")
	public static WebElement masterTreeIDToggle_EditBtn; 

	@FindBy(xpath="//*[@id='spnhiererchy']/div/div[1]/div/div/ul/li[3]")
	public static WebElement masterTreeIDToggle_DeleteBtn; 

	@FindBy(xpath="//*[@id='spnhiererchy']/div/div[1]/div/div/ul/li[4]")
	public static WebElement masterTreeIDToggle_setDefaultBtn;

	@FindBy(xpath="//*[@id='221']")
	public static WebElement  homeMasterItemMenu;

	@FindBy(xpath="//*[@id='1105']")
	public static WebElement  homeMasterItem_ItemMenu;



	// Units Test

	@FindBy (xpath="//a[@id='1114']")
	public static WebElement unitsMenu;

	@FindBy(xpath="//input[@id='iNoOfDecimals']")
	public static WebElement noOfDecimalsTxt;

	@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[10]")
	public static List<WebElement> masterItemUnitsList;


	// Currency Test

	@FindBy(xpath="//a[@id='220']")
	public static WebElement currencyMenu;

	@FindBy(xpath="//span[contains(text(),'Currency Master')]")
	public static WebElement currencyMasterMenu;

	@FindBy(xpath="//a[@id='71']")
	public static WebElement exchangeRateMenu;

	@FindBy(xpath="//span[contains(text(),'Exchange Rate History')]")
	public static WebElement exchangeRateHistoryMenu;

	@FindBy(xpath="//*[@id='btnSave']")
	public static WebElement saveIcon;

	@FindBy(xpath="//input[@id='CoinsName']")
	public static WebElement coinsNameTxt;

	@FindBy(xpath="//li[@id='btnClose']")
	public static WebElement currencyMasterCancelBtn;

	@FindBy(xpath="//*[@id='ISOCurrencyCode']")
	public static WebElement ISOCurrencyCodeTxt;




	// Exchange Rate and Exchange Rate History

	@FindBy(xpath="//*[@id='navbarSupportedContent2']//li[4]")
	public static WebElement exchangeRateDefinitionSaveIcon;

	@FindBy(xpath="//*[@id='navbarSupportedContent2']//li[2]")
	public static WebElement exchnageRateDeleteOption;

	@FindBy(xpath="//input[@id='id_option']")
	public static WebElement baseCurrencyInExchangeRateHistory;

	@FindBy(xpath="//input[@id='ShowData']")
	public static WebElement loadIcon;

	@FindBy(xpath="//*[@id='id_option_table_body']/tr/td[2]")
	public static List<WebElement> baseCurrencyListCount;


	@FindBy(xpath="//input[@id='id_option']")
	public static WebElement baseCurrencyTxt;

	@FindBy(xpath="//input[@id='EffectiveDate']")
	public static WebElement withEffectiveDateTxt;


	@FindBy(xpath="//input[@id='ExchangeTable_DefinedAs']")
	public static WebElement gridDefineAsTxt;

	@FindBy(xpath="//input[@id='ExchangeTable_Rate']")
	public static WebElement gridRateTxt;

	@FindBy(xpath="//input[@id='ExchangeTable_CurrencyName']")
	public static WebElement gridcurrencyName1Txt;

	@FindBy(xpath="//li[@id='closeExchangeRate']")
	public static WebElement exchangeRateDefinitionPageCloseIcon;

	@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li/a/i")
	public static WebElement exchangeRateHistoryCloseIcon;

	@FindBy(xpath="//tbody[@id='HistoryTableBody']/tr[2]/td[3]")
	public static WebElement valueUSD;

	@FindBy(xpath="//tbody[@id='HistoryTableBody']/tr[2]/td[4]")
	public static WebElement valueAED;

	@FindBy(xpath="//td[@id='ExchangeTable_col_1-2']")
	public static WebElement selCurrencyOne;

	@FindBy(xpath="//td[@id='ExchangeTable_col_2-2']")
	public static WebElement selCurrecnyTwo;

	@FindBy(xpath="//td[@id='ExchangeTable_col_1-3']")
	public static WebElement defCurrecnyOne;

	@FindBy(xpath="//td[@id='ExchangeTable_col_2-3']")
	public static WebElement defCurrecnyTwo;



	// Departments Test

	@FindBy (xpath="//a[@id='1106']")
	public static WebElement departmentMenu;

	@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[10]")
	public static List<WebElement> masterDepartmentGridBodyName;


	// Bins Test

	@FindBy (xpath="//a[@id='1115']")
	public static WebElement binsMenu;

	@FindBy(xpath="//*[@id='navigation_menu']/li[1]/ul/li[4]/ul/li")
	public static List<WebElement> masterSubMenusList;

	@FindBy(xpath="//input[@id='iColNo']")
	public static WebElement binColNoTxt;

	@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[8]/div[1]/label/input")
	public static List<WebElement> masterItemUnitsChkboxList;


	// Unit Conversion Test

	@FindBy(xpath="//a[@id='72']/span")
	public static WebElement  homeMasterItem_UnitConversionMenu;

	@FindBy(xpath="//input[@id='optionBaseUnit']")
	public static WebElement unitConversionBaseUnitTxt;

	@FindBy(xpath="//input[@id='optionProduct']")
	public static WebElement unitConversionItemTxt;

	@FindBy(xpath="//td[@id='myTagsTable_col_1-2']")
	public static WebElement unitConversionTableRow1Col1;

	@FindBy(xpath="//td[@id='myTagsTable_col_1-3']")
	public static WebElement unitConversionTableRow1Col2;

	@FindBy(xpath="//td[@id='myTagsTable_col_1-4']")
	public static WebElement unitConversionTableRow1Col3;

	@FindBy(xpath="//td[@id='myTagsTable_col_1-5']")
	public static WebElement unitConversionTableRow1Col4;

	@FindBy(xpath="//td[@id='myTagsTable_col_1-6']")
	public static WebElement unitConversionTableRow1Col5;

	@FindBy(xpath="//input[@id='myTagsTable_UnitName']")
	public static WebElement unitConversion_EnterUnitNameTxt;

	@FindBy(xpath="//input[@id='myTagsTable_XFactor']")
	public static WebElement unitConversion_EnterXFactorTxt;

	@FindBy(xpath="//*[@id='ucCancel']/i")
	public static WebElement unitConversionCancelBtn;

	//@FindBy(xpath="//*[@id='btnDeleteUnitConversion']/span")
	@FindBy(xpath="//*[@id='btnDelete']")
	public static WebElement unitConversionDeleteBtn;


	// Configure Masters Test
	
	
	@FindBy(xpath="//a[@id='DefinationTab']")
	public static WebElement definationTab;
	

	@FindBy(xpath="//*[@id='16']/span")
	public static WebElement settingsmenuBtn;
	
	@FindBy(xpath="//*[@id='19']/span[text() = 'Configure Masters']")
	public static WebElement settingsConfigureMasters;

	@FindBy(xpath="//*[@id='navigation_menu']/li[8]/ul/li")
	public static List<WebElement> settingsSubMenusList;

	@FindBy(xpath="//div[@id='myNavbar']/ul/li/a")
	public static List<WebElement> configMastersRibbonControl;

	@FindBy(xpath="//*[@id='MasterTabs']/li[2]/a")
	public static WebElement customizationTab;

	@FindBy(xpath="//select[@id='ddlModule']")
	public static WebElement moduleDropdown;

	@FindBy(xpath="//input[@id='txtMasterName']")
	public static WebElement nameComboBox;

	@FindBy(xpath="//tbody[@id='txtMasterName_table_body']/tr/td[2]")
	public static List<WebElement> nameComboBoxList;

	@FindBy(xpath="//input[@id='txtMasterCaption']")
	public static WebElement captionTxt;

	@FindBy(xpath="//input[@id='rdoSearchByName_master']")
	public static WebElement searchByNameRadioBtn;

	@FindBy(xpath="//input[@id='rdoSearchByCode_master']")
	public static WebElement searchByCodeRadioBtn;

	@FindBy(xpath="//input[@id='chkAllowRivisions']")
	public static WebElement allowReservationChkBox;

	@FindBy(xpath="//*[@id='chkIsDefault']")
	public static WebElement defaultChkBox;
	
	@FindBy(xpath="//*[@id='67']")
	public static WebElement masterAuthorization;

	@FindBy(xpath="//input[@id='chkGroupMandatory']")
	public static WebElement groupMandatoryChkBox;

	@FindBy(xpath="//input[@id='chkPickCode']")
	public static WebElement pickCodeChkBox;

	@FindBy(xpath="//*[@id='divDependency_Master']/table/thead/tr/td")
	public static List<WebElement> showDependencyHeader;

	@FindBy(xpath="//a[contains(text(),'Dependent')]")
	public static WebElement dependentTab;

	@FindBy(xpath="//a[contains(text(),'Used')]")
	public static WebElement usedTab;

	@FindBy(xpath="//div[@class='msTreeMenuContainer']//div//span[@class='icon-left-and-right-panel-icon icon-font3 clsMenus pull-right']")
	public static WebElement menuMinimizeBtn;

	@FindBy(xpath="//span[@id='altMenu_Icon']")
	public static WebElement menuExpansionBtn;

	@FindBy(xpath="//div[@id='masterTreeMenu']//ul//li//a//span")
	public static List<WebElement> menuList;

	@FindBy(xpath="//div[@id='masterTreeMenu']/ul//li/a/i")
	public static List<WebElement> menuPlus;
	
	@FindBy(xpath="//*[@id='DefinationTab']")
	public static WebElement cmDefinitionTab;

	@FindBy(xpath="//*[@id='btnSave_ConfigureMasters']/a/i")
	public static WebElement cmSaveBtn;

	@FindBy(xpath="//*[@id='btnClose_ConfigureMasters']/div[1]/span")
	public static WebElement cmCloseBtn;

	@FindBy(xpath="(//a[contains(@class,'menu-selected theme_color-inverse')])[2]/span[contains(text(),'Masters')]")
	public static WebElement cmMastermenu;

	//@FindBy(xpath="//*[@id='btnDelete_ConfigureForms']/a/i")
	@FindBy(xpath="//*[@id='btnDelete_ConfigureMasters']/a/i")
	public static WebElement cmDeleteBtn;

	@FindBy(xpath="//label[@id='lblMessage_Modal']")
	public static WebElement cmGetMgsonDelete;

	@FindBy(xpath="//input[@id='btnDeleteMaster_Ok']")
	public static WebElement cmclickOnOkInDelete;

	@FindBy(xpath="//*[@id='1182']/span")
	public static WebElement ConfigMasterInMastersList;

	@FindBy(xpath="//*[@id='sName']")
	public static WebElement NameTxt; 

	@FindBy(xpath="//*[@id='sCode']")
	public static WebElement codeTxt; 

	@FindBy(xpath="//*[@id='accordionExample']/div[2]/input[1]")
	public static WebElement addGroupBtnInconfigureMaster;

	@FindBy(xpath="//*[@id='accordionExample']/div[2]/input[2]")
	public static WebElement deleteGroupBtnInconfigureMaster;

	@FindBy(xpath="//*[@id='txtNewGroupName']")
	public static WebElement groupNameTxt;

	@FindBy(xpath="//*[@id='btnAddGroup_Ok']")
	public static WebElement okBtnInAddgroupPopup;

	@FindBy(xpath="//li[@id='mMenu1']/a/i")
	public static WebElement homeMenuExpandBtn;

	@FindBy(xpath="//li[@id='mMenu1000']/a/i")
	public static WebElement masterMenuExpandBtn;

	@FindBy(xpath="//li[@id='mMenu1000']/a/span")
	public static WebElement masterMenuBtn;

	@FindBy(xpath="//span[text()='GroupMaster']")
	public static WebElement GroupMasterMenuBtn;

	@FindBy(xpath="(//span[text()='GroupMaster'])[1]")
	public static WebElement GroupMasterBtn;

	@FindBy(xpath="//span[text()='GroupMasterToDelete']")
	public static WebElement GroupMasterToDeleteMenuBtn;



	// Seller Price book Test

	@FindBy(xpath="//*[@id='80']/span")
	public static WebElement  homeMasterItem_buyerPriceBookMenu;

	@FindBy(xpath="//*[@id='76']/span")
	public static WebElement  homeMasterItem_sellerPriceBookMenu;

	@FindBy(xpath="//*[@id='ctrlOptionProPriceBookH']")
	public static WebElement buyerPriceBook_PriceBook;


	//First Row

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[2]")
	public static WebElement itemFirstCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[4]")
	public static WebElement wareHouseFirstCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[5]")
	public static WebElement startingDateFirstCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[6]")
	public static WebElement endingDateFirstCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[7]")
	public static WebElement vendorFirstCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[8]")
	public static WebElement currencyFirstCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[9]")
	public static WebElement minQtyFirstCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[10]")
	public static WebElement maxQtyFirstCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[11]")
	public static WebElement unitFirstCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[12]")
	public static WebElement rateFirstCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[13]")
	public static WebElement PoFirstCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[14]")
	public static WebElement PtFirstCellSelect; 	


	//Second Row
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[2]")
	public static WebElement itemSecondCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[4]")
	public static WebElement wareHouseSecondCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[5]")
	public static WebElement satrtingDateSecondCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[6]")
	public static WebElement endingDateSecondCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[7]")
	public static WebElement vendorSecondCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[8]")
	public static WebElement currencySecondCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[9]")
	public static WebElement minQtySecondCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[10]")
	public static WebElement maxQtySecondCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[11]")
	public static WebElement unitSecondCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[12]")
	public static WebElement rateSecondCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[13]")
	public static WebElement PoSecondCellSelect; 	

	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[14]")
	public static WebElement PtSecondCellSelect; 


	//Third Row
	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-1']")
	public static WebElement itemThirdCellSelect; 	

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-3']")
	public static WebElement wareHouseThirdCellSelect; 	

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-4']")
	public static WebElement satrtingDateThirdCellSelect; 	

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-5']")
	public static WebElement endingDateThirdCellSelect; 	

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-6']")
	public static WebElement vendorThirdCellSelect; 	

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-7']")
	public static WebElement currencyThirdCellSelect; 	

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-8']")
	public static WebElement minQtyThirdCellSelect; 	

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-9']")
	public static WebElement maxQtyThirdCellSelect; 	

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-10']")
	public static WebElement unitThirdCellSelect; 	

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-11']")
	public static WebElement rateThirdCellSelect; 	

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-12']")
	public static WebElement PoThirdCellSelect; 	

	@FindBy(xpath="//td[@id='PriceBookWebGrid_col_3-13']")
	public static WebElement PtThirdCellSelect; 	


	//To get Value From Cell

	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_2']")
	public static WebElement itemValue; 	

	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_4']")
	public static WebElement wareHouseValue; 	

	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_StartingDate']")
	public static WebElement startingDateValue; 	

	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_EndingDate']")
	public static WebElement endingDateValue; 	

	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_1']")
	public static WebElement vendorValue; 	

	@FindBy(xpath="//input[@id='Currency_option']")
	public static WebElement currencyOptionValue; 	

	@FindBy(xpath="//input[@id='PriceBookWebGrid_minqty']")
	public static WebElement minQtyValue; 	

	@FindBy(xpath="//input[@id='PriceBookWebGrid_maxqty']")
	public static WebElement maxQtyValue; 	

	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_11']")
	public static WebElement unitValue; 	

	@FindBy(xpath="//input[@id='PriceBookWebGrid_val0']")
	public static WebElement rateValue; 	

	@FindBy(xpath="//input[@id='PriceBookWebGrid_val1']")
	public static WebElement PoValue; 	

	@FindBy(xpath="//input[@id='PriceBookWebGrid_val2']")
	public static WebElement PtValue; 	

	@FindBy(xpath="//*[@id='chkDateRange']")
	public static WebElement buyerPriceBook_CustamizeDateRangeCheckBox;

	@FindBy(xpath="//*[@id='chkCustomer']")
	public static WebElement buyerPriceBook_CustomizeVendorCheckBox;
	
	@FindBy(xpath="//*[@id='chkCustomer']/following-sibling::span")
	public static WebElement buyerPriceBook_CustomizeVendorCheckBoxSelected;

	@FindBy(xpath="//*[@id='chkDepartment']")
	public static WebElement buyerPriceBook_WarehouseCheckBox;
	
	@FindBy(xpath="//*[@id='chkDepartment']/following-sibling::span")
	public static WebElement buyerPriceBook_WarehouseCheckBoxSelected;

	@FindBy(xpath="//*[@id='chkQtyRange']")
	public static WebElement buyerPriceBook_QualityRangeCheckBox;
	
	@FindBy(xpath="//*[@id='chkQtyRange']/following-sibling::span")
	public static WebElement buyerPriceBook_QualityRangeCheckBoxSelected;

	@FindBy(xpath="//*[@id='chkCurrency']")
	public static WebElement buyerPriceBook_CurrencyCheckBox;

	@FindBy(xpath="//*[@id='chkUnit']")
	public static WebElement buyerPriceBook_UnitCheckBox;

	@FindBy(xpath="//*[@id='txtAbbreviation']")
	public static WebElement buyerPriceBook_AbbreviationBtn;

	@FindBy(xpath="//*[@id='PriceBookContainer']/div[2]/nav/div/div[1]/a/div/span[2]")
	public static WebElement buyerPriceBookLabel;

	@FindBy(xpath="//*[@id='btnClear']")
	public static WebElement buyerPriceBookClearBtn;

	@FindBy(xpath="//*[@id='btnPaste']")
	public static WebElement buyerPriceBookPasteBtn;

	@FindBy(xpath="//*[@id='btnCopy']")
	public static WebElement buyerPriceBookCopyBtn;

	@FindBy(xpath="//*[@id='btnSelectAll']")
	public static WebElement buyerPriceBookSelectAllBtn;

	@FindBy(xpath="//*[@id='btnClose']")
	public static WebElement buyerPriceBookCloseBtn;

	@FindBy(xpath="//*[@id='btnSave']")
	public static WebElement buyerPriceBookSaveBtn;

	@FindBy(xpath="//*[@id='ctrlOptionProPriceBookH_table_body']/tr/td[2]")
	public static List<WebElement> pricebookComboBoxList;

	@FindBy(xpath="//*[@id='PriceBookWebGrid_option_2_table_body']/tr/td[2]")
	public static List<WebElement> itemComboBoxList;

	@FindBy(xpath="//*[@id='PriceBookWebGrid_option_4_table_body']/tr/td[2]")
	public static List<WebElement> warehouseComboBoxList;

	@FindBy(xpath="//*[@id='PriceBookWebGrid_option_1_table_body']/tr/td[2]")
	public static List<WebElement> vendorComboBoxList;

	@FindBy(xpath="//*[@id='Currency_option_table_body']/tr/td[2]")
	public static List<WebElement> currencyComboBoxList;

	@FindBy(xpath="//*[@id='PriceBookWebGrid_option_11_table_body']/tr/td[2]")
	public static List<WebElement> unitComboBoxList;

	@FindBy(xpath="//input[@id='btnLoad']")
	public static WebElement priceBookFilterAndLoadBtn;



	@FindBy (xpath="//input[@id='btnExporttoExcel']")
	public static WebElement priceBookExportToExcel;

	@FindBy(xpath="//input[@id='btnPBImportFromExcel']")
	public static WebElement buyerPriceBook_ImportFromExcel;

	@FindBy(xpath="//*[@id='btnImportData']")
	public static WebElement importButton;

	@FindBy(xpath="//button[@id='btnCloseImport']")
	public static WebElement closeButtonInImport;

	//@FindBy(xpath="//span[@class='btn btn-sm theme_background-color fileUpload theme_color']")
	@FindBy(xpath="//*[@id='formLoadDataExcel']/div/span/span")
	public static WebElement browseButton;

	@FindBy(xpath="//input[@id='rdoFilterName']")
	public static WebElement nameRadioButton;

	@FindBy(xpath="//input[@id='rdoFilterCode']")
	public static WebElement codeRadioButton;

	@FindBy(xpath="//input[@id='rdoFilterBoth']")
	public static WebElement bothRadioButton;


	@FindBy(xpath="//input[@id='chkOverride']")
	public static WebElement overrideCheckbox;


	@FindBy(xpath="//span[contains(text(),'Select Sheet')]")
	public static WebElement sheetTitle;

	@FindBy(xpath="//li[@class='clsSheetName']")
	public static WebElement sheetName;

	@FindBy(xpath="//input[@id='btnSheetOk']")
	public static WebElement sheetOkButton;

	@FindBy(xpath="//button[@id='btnSheetClose']")
	public static WebElement sheetCloseButton;


	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[1]//td[3]")
	public static WebElement sheetItemHeading;

	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList0']")
	public static WebElement sheetItem;

	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[2]//td[3]")
	public static WebElement sheetAccountHeading;

	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList1']")
	public static WebElement sheetAccount;

	@FindBy(xpath="(//div[@id='myModalImportFromExcel']//tr[3]//td[3])[2]")
	public static WebElement sheetTagHeading;

	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList2']")
	public static WebElement sheetTag;

	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[4]//td[3]")
	public static WebElement sheetCurrencyHeading;

	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList3']")
	public static WebElement sheetCurrency;

	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[5]//td[3]")
	public static WebElement sheetStartDateHeading;

	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList4']")
	public static WebElement sheetStartDate;

	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[6]//td[3]")
	public static WebElement sheetEndDateHeading;

	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList5']")
	public static WebElement sheetEndDate;

	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[7]//td[3]")
	public static WebElement sheetMinQtyHeading;

	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList6']")
	public static WebElement sheetMinQty;

	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[8]//td[3]")
	public static WebElement sheetMaxQtyHeading;

	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList7']")
	public static WebElement sheetMaxQty;

	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[9]//td[3]")
	public static WebElement sheetUnitHeading;

	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList8']")
	public static WebElement sheetUnitQty;

	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[11]//td[3]")
	public static WebElement sheetRateHeading;

	//@FindBy(xpath="//select[@id='ddlImportFromExcelItemList9']")
	@FindBy(xpath="//select[@id='ImportGrid_ExternalFields']")
	public static WebElement sheetRate;

	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[12]//td[3]")
	public static WebElement sheetPOHeading;

	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList10']")
	public static WebElement sheetPO;

	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[13]//td[3]")
	public static WebElement sheetPTHeading;

	@FindBy(xpath="//div[@id='myModalImportFromExcel']//tr[14]//td[3]")
	public static WebElement sheetHeading;

	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList11']")
	public static WebElement sheetPT;

	@FindBy(xpath="//select[@id='ddlImportFromExcelItemList12']")
	public static WebElement sheetP;

	@FindBy(xpath="//*[@id='toggleDropDown']/i")
	public static WebElement buyerPriceBook_CustamizeSettingBtn;

	@FindBy (xpath="//ul[@id='ddlMenu']/li/label/input")
	public static List<WebElement> buyerPriceBook_CustomizeCheckBoxList;
	
	@FindBy (xpath="//ul[@id='ddlMenu']/li/label/span")
	public static List<WebElement> buyerPriceBook_CustomizeCheckBoxListSelected;


	@FindBy (xpath="//tbody[@id='PriceBookWebGrid_body']/tr[1]/td")
	public static List<WebElement> buyerPriceBook_TableRow1List;

	@FindBy (xpath="//tbody[@id='PriceBookWebGrid_body']/tr[2]/td")
	public static List<WebElement> buyerPriceBook_TableRow2List;


	@FindBy(xpath="//select[@id='ImportGrid_ExternalFields']")
	public static WebElement buyerPriceBook_MappingDrpdwn;




	// Security Test

	//Home Menu Security Menu
	@FindBy(xpath="//*[@id='6']")
	public static WebElement securityMenu; 

	//Home Menu Password Policy Menu
	@FindBy(xpath="//span[contains(text(),'Password policy')]")
	public static WebElement  homeSecurityPasswordPolicyMenu;  

	@FindBy(xpath="//a[@id='11']/span")
	public static WebElement createUserMenu;

	@FindBy (xpath="//li//span[text()='Password Policy']")
	public static WebElement passwordPolicyTitle;

	@FindBy (xpath="//*[@id='PolicyListDiv']")
	public static WebElement passwordPolicyList;

	//@FindBy (xpath="//*[@id='addPasswordPolicybtn']/i")
	@FindBy(xpath="//*[@id='navbarSupported3']/ul/li[1]")
	public static WebElement passwordPolicyAddBtn;

	//Password Policy DeleteBtn
	//@FindBy (xpath="//*[@id='controlBtns']/span[2]/i")
	@FindBy (xpath="//*[@id='navbarSupported3']/ul/li[2]")
	public static WebElement passwordPolicyDeletebtn;

	//Password Policy CancelBtn
	@FindBy (xpath="//*[@id='controlBtns']/span[3]/i")
	public static WebElement passwordPolicyCancelBtn;

	//Password Policy SaveBtn
	@FindBy (xpath="//*[@id='btnSave']//i")
	public static WebElement SaveBtn;


	//Password Policy CloseBtn
	@FindBy (xpath="//*[@id='CancelPasswordPolicy']/i")
	public static WebElement passwordPolicyCloseBtn;

	//Password Policy policyName
	@FindBy (xpath="//*[@id='policyName']")
	public static WebElement policyName;

	//Password Policy passwordLength
	@FindBy (xpath="//*[@id='MinPassLength']")
	public static WebElement passwordLength;

	//Password Policy passwordComplexity
	@FindBy (xpath="//*[@id='complexity']")
	public static WebElement passwordComplexity;

	//Password Policy doNotAllowPreviousPassword
	@FindBy (xpath="//*[@id='Donotallprevious']")
	public static WebElement doNotAllowPrevious;

	//Password Policy passwordExpiryDays
	@FindBy (xpath="//*[@id='passwordexpirydays']")
	public static WebElement passwordExpiryDays;

	//Password Policy passwordOTPExpiryInMins
	@FindBy (xpath="//*[@id='OTPexpiryinminute']")
	public static WebElement passwordExpiryInMins;

	//Password Policy passwordInvalidAttempts
	@FindBy (xpath="//*[@id='NoOfInvalidAttemps']")
	public static WebElement passwordInvalidAttempts;

	//Password Policy passwordLockOut
	@FindBy (xpath="//*[@id='Lockoutperiod']")
	public static WebElement passwordLockOut;

	//Password Policy passwordLockOutUnits
	@FindBy (xpath="//*[@id='Units']")
	public static WebElement passwordLockOutUnits;

	//Password Policy sendMailCheck
	@FindBy (xpath="//*[@id='Sendemailchk']")
	public static WebElement sendMailCheck;

	//Password Policy sendMailOnSuccess
	@FindBy (xpath="//*[@id='sendemailonloginsuccesschk']")
	public static WebElement sendMailOnSuccess;


	//Password Policy sendMailOnFailure
	@FindBy (xpath="//*[@id='Sendemailonloginfailurechk']")
	public static WebElement sendMailOnFailure;

	//Password Policy sendMailOnFailure
	@FindBy (xpath="//*[@id='PolicyDiv']/div[14]/div/label")
	public static WebElement changePasswordAfterFirstLogin;

	//Password Policy sendMailOnFailure
	@FindBy (xpath="//*[@id='PolicyDiv']/div[15]/div/label")
	public static WebElement cannotChangePassword;

	//Password Policy sendMailOnFailure
	@FindBy (xpath="//*[@id='PolicyDiv']/div[16]/div/label")
	public static WebElement otpBasedLogin;

	//Password Policy sendMailOnFailure
	@FindBy (xpath="//*[@id='sendEmailId']")
	public static WebElement sendEmailId;

	//Password Policy sendMailOnFailure
	@FindBy (xpath="//*[@id='sendemailonloginsuccessId']")
	public static WebElement sendEmailOnLoginSuccess;

	//Password Policy sendMailOnFailure
	@FindBy (xpath="//*[@id='SendemailonloginfailureId']")
	public static WebElement sendEmailOnLoginFailure;

	@FindBy(xpath="//*[@id='PolicyListDiv']/li")
	public static List<WebElement> passwordList;

	@FindBy (xpath="//a[@id='9']")
	public static WebElement createProfile;

	@FindBy(xpath="//input[@id='profileName']")
	public static WebElement  createProfileProfileNameCombo; 

	@FindBy(xpath="//ul[@id='Profilemenu']/li/a")
	public static List<WebElement> profileMenuPanelList;

	@FindBy(xpath="//div[@id='controlBtns']/div/span[1]/i")
	public static WebElement  createProfileSelectAllIcon; 

	@FindBy(xpath="//div[@id='controlBtns']/div/span[2]/i")
	public static WebElement  createProfileUnSelectAllIcon; 

	@FindBy(xpath="//span[contains(text(),'Save')]/i")
	public static WebElement  createProfileSaveIcon; 

	@FindBy(xpath="//*[@id='CancelCreateProfile']/i")
	public static WebElement  createProfileCloseIcon; 

	@FindBy(xpath="//a[@id='btnLoadForm']/i")
	public static WebElement createProfileLoadFromIcon;

	@FindBy(xpath="//ul[@id='ProfileloadFromDiv']/li/span")
	public static List<WebElement> createProfileLoadFromPopupProfilesList;

	@FindBy(xpath="//*[@id='LoadProfile']/div[2]/div/div[3]/div/input[1]")
	public static WebElement loadPopupOkBtn;

	@FindBy(xpath="//*[@id='LoadProfile']/div[2]/div/div[3]/div/input[2]")
	public static WebElement loadPopupCancelBtn;

	@FindBy(xpath="//*[@id='Profilemenu']/li[2]/i")
	public static WebElement  createProfileFinancalsExpandBtn; 

	@FindBy(xpath="//ul[@id='Profilemenu']/li[2]/ul/li[1]/i")
	public static WebElement financialsTransactionsExpandBtn;

	@FindBy(xpath="//*[@id='Profilemenu']/li[2]/ul/li[1]/ul/li/a")
	public static List<WebElement> financialsTransactionsList;

	@FindBy(xpath="//li[@class='treeview']//a[@id='2001']//span[contains(text(),'Cash and Bank')]")
	public static WebElement financialsTransactionsCashAndBank;

	@FindBy(xpath="//li[@class='treeview']//a[@id='2014']//span[contains(text(),'Sales')]")
	public static WebElement financialsTransactionsSales;

	@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Auto Postings')]")
	public static WebElement financialsTransactionsAutoPostings;



	@FindBy(xpath="//*[@id='Profilemenu']/li[2]/ul/li[1]/ul/li[3]/a/span")
	public static WebElement financialsTransactionsSalesExpandBtn;

	@FindBy(xpath="//*[@id='Profilemenu']/li[2]/ul/li[1]/ul/li[3]/ul/li/a/span")
	public static List<WebElement> financialsTransactionsSalesOptionsList;

	@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Sales Invoices')]")
	public static WebElement financialsTransactionsSalesInvoiceOption;

	@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Sales Returns')]")
	public static WebElement financialsTransactionsSalesReturnsOption;

	@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Auto Sales Invoice')]")
	public static WebElement financialsTransactionsAutoSalesInvoiceOption;

	@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Sales Invoice N')]")
	public static WebElement financialsTransactionsSalesInvoiceNOption;

	@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Cash Sales')]")
	public static WebElement financialsTransactionsCashSalesOption;

	@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Hire-Purchase Sales')]")
	public static WebElement financialsTransactionsHirePurchaseSalesOption;

	@FindBy(xpath="//ul[@id='Profilemenu']/li[2]/ul/li[1]/ul/li[2]/i")
	public static WebElement financialsTransactionsPurchasesExpandBtn;

	@FindBy(xpath="//ul[@id='Profilemenu']/li[2]/ul/li[1]/ul/li[2]/ul/li[1]/a")
	public static WebElement financialsTransactionsPurchasesVoucherOption;



	@FindBy(xpath="//ul[@id='Profilemenu']/li[2]/ul/li[2]/a/span")
	public static WebElement financialsCreditManagementExpandBtn;

	@FindBy(xpath="//ul[@id='Profilemenu']/li[2]/ul/li[3]/a/span")
	public static WebElement finincalsReportsExpandBtn;

	@FindBy(xpath="//ul[@id='Profilemenu']/li[2]/ul/li[4]/a/span")
	public static WebElement finalAccountsExpandBtn;

	@FindBy(xpath="//ul[@id='Profilemenu']/li[2]/ul/li[5]/a/span")
	public static WebElement financialsReceivableandPayableAnalysisExpandBtn;

	@FindBy(xpath="//ul[@id='Profilemenu']/li[2]/ul/li[6]/a/span")
	public static WebElement financialsBudgetExpandBtn;

	@FindBy(xpath="//ul[@id='Profilemenu']/li[2]/ul/li[7]/a/span")
	public static WebElement financialsOnlinePaymentsExpandBtn;

	@FindBy(xpath="//li[1]//div[1]//label[1]//input[1]")
	public static WebElement faAccessChkbox;

	@FindBy(xpath="//li[2]//div[1]//label[1]//input[1]")
	public static WebElement faAddchkbox;

	@FindBy(xpath="//li[3]//div[1]//label[1]//input[1]")
	public static WebElement faEditChkBox;

	@FindBy(xpath="//li[4]//div[1]//label[1]//input[1]")
	public static WebElement faDeleteChkbox;

	@FindBy(xpath="//li[5]//div[1]//label[1]//input[1]")
	public static WebElement faSaveChkbox;

	@FindBy(xpath="//li[2]//div[1]//label[1]//input[1]")
	public static WebElement faPOstDepreciationChkbox;

	@FindBy(xpath="//li[2]//div[1]//label[1]//input[1]")
	public static WebElement faMultipleAssetUsageDeleteChkbox;

	@FindBy(xpath="//li[3]//div[1]//label[1]//input[1]")
	public static WebElement faMultipleusageAssetSaveChkbox;

	@FindBy(xpath="//li[2]//div[1]//label[1]//input[1]")
	public static WebElement searchChkbox;

	@FindBy(xpath="//li[3]//div[1]//label[1]//input[1]")
	public static WebElement printChkbox;

	@FindBy(xpath="//li[4]//div[1]//label[1]//input[1]")
	public static WebElement viewChkbox;

	@FindBy(xpath="//li[5]//div[1]//label[1]//input[1]")
	public static WebElement exportChkbox;

	@FindBy(xpath="//li[6]//div[1]//label[1]//input[1]")
	public static WebElement emailReportChkbox;

	@FindBy(xpath="//li[7]//div[1]//label[1]//input[1]")
	public static WebElement editPrintlayoutChkbox;

	@FindBy(xpath="//li[8]//div[1]//label[1]//input[1]")
	public static WebElement viewCustomizeLayoutChkbox;

	@FindBy(xpath="//li[9]//div[1]//label[1]//input[1]")
	public static WebElement addCustomizeLAyoutChkbox;

	@FindBy(xpath="//li[10]//div[1]//label[1]//input[1]")
	public static WebElement BackTrackChkbox;

	@FindBy(xpath="//li[11]//div[1]//label[1]//input[1]")
	public static WebElement emailChkbox;

	@FindBy(xpath="//li[12]//div[1]//label[1]//input[1]")
	public static WebElement viewGraphChkbox;

	@FindBy(xpath="//li[13]//div[1]//label[1]//input[1]")
	public static WebElement filterChkbox;

	@FindBy(xpath="//li[14]//div[1]//label[1]//input[1]")
	public static WebElement analyzeChkbox;

	@FindBy(xpath="//li[15]//div[1]//label[1]//input[1]")
	public static WebElement crossReferenceChkbox;

	@FindBy(xpath="//li[16]//div[1]//label[1]//input[1]")
	public static WebElement AddAnalyzeReportChkbox;

	@FindBy(xpath="//*[@id='restrictionsDiv']/li/label/span")
	public static List<WebElement> restrictionTabChkBoxList;
	
	@FindBy(xpath="//*[@id='restrictionsDiv']/li/label/input")
	public static List<WebElement> restrictionTabChkBoxIsSelectedList;

	@FindBy(xpath="//*[@id='restrictionsDiv']/li/label")
	public static List<WebElement> restrictionTabChkBoxNameList;

	@FindBy(xpath="//tbody[@id='profileName_table_body']/tr/td[2]")
	public static List<WebElement> createProfileProfileNameList;

	@FindBy(xpath="//*[@id='deleteBtn']//i")
	public static WebElement  createProfileDeleteIcon; 

	@FindBy(xpath="//span[contains(text(),'Create Role')]")
	public static WebElement  homeSecurityCreateRoleMenu;

	@FindBy(xpath="//input[@id='roleName']")
	public static WebElement createRoleRoleNameCombo; 

	@FindBy(xpath="//tbody[@id='roleName_table_body']/tr/td[2]")
	public static List<WebElement> createRoleRoleNameComboList;

	@FindBy(xpath="//select[@id='passwordPolicy']")
	public static WebElement  createRolePasswordPolicyDropdown; 

	@FindBy(xpath="//*[@id='btnSave']/i")
	public static WebElement  createRoleSaveicon; 

	@FindBy(xpath="//div[@id='myNavbar']//span[2]")
	public static WebElement  createRoleDeleteIcon; 

	@FindBy(xpath="//*[@id='CancelCreateUser']/i")
	public static WebElement  createRoleCloseIcon; 

	@FindBy(xpath="//span[2]//span[1]")
	public static WebElement  createRoleMOveFiledsFronLeftSideToRightSide; 


	@FindBy(xpath="//*[contains(text(),'All Profile')]")
	public static WebElement avaliableAllProfile ;

	@FindBy(xpath="//*[contains(text(),'All Profile ST')]")
	public static WebElement avaliableAllProfileSt ;

	@FindBy(xpath="//*[contains(text(),'Purchase Profile')]")
	public static WebElement avaliableProfilePurchaseProfile ;

	@FindBy(xpath="//*[contains(text(),'SalesOrderProfile')]")
	public static WebElement avaliableProfileSalesOrderProfile ;

	@FindBy(xpath="//*[contains(text(),'SaleAccessProfile')]")
	public static WebElement avaliableProfileSaleAccessProfile ;

	@FindBy(xpath="//*[contains(text(),'SaleAddProfile')]")
	public static WebElement avaliableProfileSaleAddProfile ;

	@FindBy(xpath="//*[contains(text(),'SaleEditProfile')]")
	public static WebElement avaliableProfileSaleEditProfile ;

	@FindBy(xpath="//*[contains(text(),'SaleDeleteProfile')]")
	public static WebElement avaliableProfileSaleDeleteProfile ;

	@FindBy(xpath="//*[contains(text(),'SalePrintProfile')]")
	public static WebElement avaliableProfileSalePrintProfile ;

	@FindBy(xpath="//*[contains(text(),'SaleRePrintProfile')]")
	public static WebElement avaliableProfileSaleRePrintProfile ;

	@FindBy(xpath="//*[contains(text(),'SalesPrintUnAuthorizeProfile')]")
	public static WebElement avaliableProfileSalesPrintUnAuthorizeProfile ;

	@FindBy(xpath="//*[contains(text(),'SalesEditMiscProfile')]")
	public static WebElement avaliableProfileSalesEditMiscProfile ;

	@FindBy(xpath="//*[contains(text(),'SalesAlwaysSuspendingProfile')]")
	public static WebElement avaliableProfileSalesAlwaysSuspendingProfile ;

	@FindBy(xpath="//*[contains(text(),'SalesChangePrintLayoutProfile')]")
	public static WebElement avaliableProfileSalesChangePrintLayoutProfile ;

	@FindBy(xpath="//*[contains(text(),'SalesAuthorizeProfile')]")
	public static WebElement avaliableProfileSalesAuthorizeProfile ;

	@FindBy(xpath="//*[contains(text(),'SalesRejectProfile')]")
	public static WebElement avaliableProfileSalesRejectProfile ;

	@FindBy(xpath="//*[contains(text(),'SalesHomePageProfile')]")
	public static WebElement avaliableProfileSalesHomePageProfile ;

	@FindBy(xpath="//*[contains(text(),'SalesStopProfile')]")
	public static WebElement avaliableProfileSalesStopProfile ;

	@FindBy(xpath="//*[contains(text(),'SalesModifySettingsProfile')]")
	public static WebElement avaliableProfileSalesModifySettingsProfile ;

	@FindBy(xpath="//*[contains(text(),'SalesViewDocumentsEnterByOther')]")
	public static WebElement avaliableProfileSalesViewDocumentsEnterByOther ;

	@FindBy(xpath="//*[contains(text(),'SalesCloseLinkProfile')]")
	public static WebElement avaliableProfileSalesCloseLinkProfile ;

	@FindBy(xpath="//*[contains(text(),'SalesSaveRevisionProfile')]")
	public static WebElement avaliableProfileSalesSaveRevisionProfile ;

	@FindBy(xpath="//*[contains(text(),'SalesSuspendProfile')]")
	public static WebElement avaliableProfileSalesSuspendProfile ;

	@FindBy(xpath="//*[contains(text(),'EditExportedSyncDocProfile')]")
	public static WebElement avaliableProfileEditExportedSyncDocProfile ;

	@FindBy(xpath="//*[contains(text(),'EditImportedSyncDocProfile')]")
	public static WebElement avaliableProfileEditImportedSyncDocProfile ;

	@FindBy(xpath="//*[contains(text(),'SaleHideProfile')]")
	public static WebElement avaliableProfileSaleHideProfile ;

	@FindBy(xpath="//*[contains(text(),'FAReportsAllOptions')]")
	public static WebElement avaliableProfileFAReportsAllOptions ;

	@FindBy(xpath="//*[contains(text(),'LedgerProfile')]")
	public static WebElement avaliableProfileLedgerProfile ;

	@FindBy(xpath="//*[contains(text(),'SLAcessProfile')]")
	public static WebElement avaliableProfileSLAcessProfile ;

	@FindBy(xpath="//*[contains(text(),'SLSearchProfile')]")
	public static WebElement avaliableProfileSLSearchProfile ;

	@FindBy(xpath="//*[contains(text(),'SLPrintProfile')]")
	public static WebElement avaliableProfileSLPrintProfile ;

	@FindBy(xpath="//*[contains(text(),'SLViewProfile')]")
	public static WebElement avaliableProfileSLViewProfile ;

	@FindBy(xpath="//*[contains(text(),'SLExportProfile')]")
	public static WebElement avaliableProfileSLExportProfile ;

	@FindBy(xpath="//*[contains(text(),'SLEmailProfile')]")
	public static WebElement avaliableProfileSLEmailProfile ;

	@FindBy(xpath="//*[contains(text(),'SLViewGraphProfile')]")
	public static WebElement avaliableProfileSLViewGraphProfile ;

	@FindBy(xpath="//*[contains(text(),'SLFilterProfile')]")
	public static WebElement avaliableProfileSLFilterProfile ;

	@FindBy(xpath="//*[contains(text(),'SLAnalyzeProfile')]")
	public static WebElement avaliableProfileSLAnalyzeProfile ;

	@FindBy(xpath="//*[contains(text(),'SLCrossReferenceProfile')]")
	public static WebElement avaliableProfileSLCrossReferenceProfile ;

	@FindBy(xpath="//*[contains(text(),'SLAddAnalyzeProfile')]")
	public static WebElement avaliableProfileSLAddAnalyzeProfile ;

	@FindBy(xpath="//*[contains(text(),'AllMasterProfile')]")
	public static WebElement avaliableProfileAllMasterProfile ;

	@FindBy(xpath="//*[contains(text(),'AllProductProfile')]")
	public static WebElement avaliableProfileAllProductProfile ;

	@FindBy(xpath="//*[contains(text(),'ProfileAddProduct')]")
	public static WebElement avaliableProfileProfileAddProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileEditProduct')]")
	public static WebElement avaliableProfileProfileEditProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileCopyProduct')]")
	public static WebElement avaliableProfileProfileCopyProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileSearchProduct')]")
	public static WebElement avaliableProfileProfileSearchProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileDeleteProduct')]")
	public static WebElement avaliableProfileProfileDeleteProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfilePrintProduct')]")
	public static WebElement avaliableProfileProfilePrintProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileExportProduct')]")
	public static WebElement avaliableProfileProfileExportProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileCustomizeViewProduct')]")
	public static WebElement avaliableProfileProfileCustomizeViewProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileSortProduct')]")
	public static WebElement avaliableProfileProfileSortProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileMassUpdateProduct')]")
	public static WebElement avaliableProfileProfileMassUpdateProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileAuthorizeProduct')]")
	public static WebElement avaliableProfileProfileAuthorizeProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileTransferProduct')]")
	public static WebElement avaliableProfileProfileTransferProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfilePropertiesProduct')]")
	public static WebElement avaliableProfileProfilePropertiesProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileLedgerProduct')]")
	public static WebElement avaliableProfileProfileLedgerProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileBackTrackProduct')]")
	public static WebElement avaliableProfileProfileBackTrackProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileBudgetsProduct')]")
	public static WebElement avaliableProfileProfileBudgetsProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileHelpProduct')]")
	public static WebElement avaliableProfileProfileHelpProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileCreateTreeProduct')]")
	public static WebElement avaliableProfileProfileCreateTreeProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileDeleteTreeProduct')]")
	public static WebElement avaliableProfileProfileDeleteTreeProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileCreateViewProduct')]")
	public static WebElement avaliableProfileProfileCreateViewProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileEditViewProduct')]")
	public static WebElement avaliableProfileProfileEditViewProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileDeleteViewProduct')]")
	public static WebElement avaliableProfileProfileDeleteViewProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileSelectViewProduct')]")
	public static WebElement avaliableProfileProfileSelectViewProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileSelectTreeProduct')]")
	public static WebElement avaliableProfileProfileSelectTreeProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileCloneProduct')]")
	public static WebElement avaliableProfileProfileCloneProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileAddGroupProduct')]")
	public static WebElement avaliableProfileProfileAddGroupProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileGroupMasterProduct')]")
	public static WebElement avaliableProfileProfileGroupMasterProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileSetTypeProduct')]")
	public static WebElement avaliableProfileProfileSetTypeProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileDeleteAllProduct')]")
	public static WebElement avaliableProfileProfileDeleteAllProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileMoveUpProduct')]")
	public static WebElement avaliableProfileProfileMoveUpProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileMoveDownProduct')]")
	public static WebElement avaliableProfileProfileMoveDownProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileCreditManagementProduct')]")
	public static WebElement avaliableProfileProfileCreditManagementProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileDepAppropirationProduct')]")
	public static WebElement avaliableProfileProfileDepAppropirationProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileCustomizeTreeProduct')]")
	public static WebElement avaliableProfileProfileCustomizeTreeProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileCustomizeMasterProduct')]")
	public static WebElement avaliableProfileProfileCustomizeMasterProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileCanChangeGroupProduct')]")
	public static WebElement avaliableProfileProfileCanChangeGroupProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileCanNavigateProduct')]")
	public static WebElement avaliableProfileProfileCanNavigateProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileAddInfoPanelProduct')]")
	public static WebElement avaliableProfileProfileAddInfoPanelProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileEditInfoPanelProduct')]")
	public static WebElement avaliableProfileProfileEditInfoPanelProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileViewInfoPanelProduct')]")
	public static WebElement avaliableProfileProfileViewInfoPanelProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileRejectProduct')]")
	public static WebElement avaliableProfileProfileRejectProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileImportProduct')]")
	public static WebElement avaliableProfileProfileImportProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileShowHomePageProduct')]")
	public static WebElement avaliableProfileProfileShowHomePageProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileDeleteInfoPanelProduct')]")
	public static WebElement avaliableProfileProfileDeleteInfoPanelProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileStopProduct')]")
	public static WebElement avaliableProfileProfileStopProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileSaveRevisionProduct')]")
	public static WebElement avaliableProfileProfileSaveRevisionProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileSetDeafultTreeProduct')]")
	public static WebElement avaliableProfileProfileSetDeafultTreeProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileModifierProduct')]")
	public static WebElement avaliableProfileProfileModifierProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileEditTreeProduct')]")
	public static WebElement avaliableProfileProfileEditTreeProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileGeneralProduct')]")
	public static WebElement avaliableProfileProfileGeneralProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileUnitsProduct')]")
	public static WebElement avaliableProfileProfileUnitsProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileSettingsProduct')]")
	public static WebElement avaliableProfileProfileSettingsProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileClassificationProduct')]")
	public static WebElement avaliableProfileProfileClassificationProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileOtherDetailsProduct')]")
	public static WebElement avaliableProfileProfileOtherDetailsProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileReplenishmentProduct')]")
	public static WebElement avaliableProfileProfileReplenishmentProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileReorderProduct')]")
	public static WebElement avaliableProfileProfileReorderProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileOutletProduct')]")
	public static WebElement avaliableProfileProfileOutletProduct ;

	@FindBy(xpath="//*[contains(text(),'ProfileAllAcountsOption')]")
	public static WebElement avaliableProfileProfileAllAcountsOption ;

	@FindBy(xpath="//*[contains(text(),'ProfileAllDepartmentOption')]")
	public static WebElement avaliableProfileProfileAllDepartmentOption ;

	@FindBy(xpath="//*[contains(text(),'ProfileLoadFrom')]")
	public static WebElement avaliableProfileProfileLoadFrom ;

	@FindBy(xpath="//*[@id='profileAdditionbtns']/div[2]/span/span")
	public static WebElement  createRoleMOveFiledsFromLeftSideToRightSide; 

	@FindBy(xpath="//*[@id='profileAdditionbtns']/div[3]/span/span")
	public static WebElement createRoleMOveFiledsFromRigheSideToLeftSide;

	@FindBy(xpath="//ul[@id='assignedProfiles']/li")
	public static List<WebElement> assignedProfileList;

	@FindBy(xpath="//a[contains(text(),'Assigned Profiles')]")
	public static WebElement  assignedProfilesTab;

	@FindBy(xpath="//a[contains(text(),'Additions')]")
	public static WebElement  additionTab;

	@FindBy(xpath="//a[contains(text(),'Exclusions')]")
	public static WebElement  exclusionsTab;

	/*	@FindBy(xpath="//a[contains(text(),'Restriction for entry')]")
public static WebElement  restrictionforEntryTab;*/

	@FindBy(xpath="//a[contains(text(),'Restriction for trees')]")
	public static WebElement  restrictionforTreesTab;

	@FindBy(xpath="//a[contains(text(),'Transaction Rights')]")
	public static WebElement  transactionRightsTab;

	@FindBy(xpath="//a[contains(text(),'AI Rights')]")
	public static WebElement  AIRightsTab;

	@FindBy(xpath="//div[@id='profileAdditionbtns']/span[1]")
	public static WebElement  CreateRoleAssignAllToRight;

	@FindBy(xpath="//div[@id='profileAdditionbtns']/span[2]")
	public static WebElement  CreateRoleAssignSelectedToRight;

	@FindBy(xpath="//div[@id='profileAdditionbtns']/span[3]")
	public static WebElement  CreateRoleAssignSelectedToLeft;

	@FindBy(xpath="//div[@id='profileAdditionbtns']/span[4]")
	public static WebElement  CreateRoleAssignAllToLeft;



	@FindBy(xpath="//ul[@id='availableProfiles']")
	public static WebElement  availableProfilesArea;

	@FindBy(xpath="//ul[@id='assignedProfiles']")
	public static WebElement  assignedProfilesArea;


	//Addtions Tab
	@FindBy(xpath="//ul[@id='Addmenu']/li/a/span")
	public static List<WebElement> createRoleAddTabMenuList;

	@FindBy(xpath="//ul[@id='Addmenu']/li[1]/a/span")
	public static WebElement createRoleAddTabHomeMenu;

	@FindBy(xpath="//ul[@id='Addmenu']/li[2]/a")
	public static WebElement createRoleAddTabFinancialsMenu;
	
	@FindBy(xpath="//ul[@id='Addmenu']/li[2]/i")
	public static WebElement createRoleAddTabFinancialsMenuExpand;

	@FindBy(xpath="//ul[@id='Addmenu']/li[3]/a/span")
	public static WebElement createRoleAddTabInventoryMenu;

	@FindBy(xpath="//ul[@id='Addmenu']/li[4]/a/span")
	public static WebElement createRoleAddTabFixedAssetsMenu;

	@FindBy(xpath="//ul[@id='Addmenu']/li[5]/a/span")
	public static WebElement createRoleAddTabProductionMenu;

	@FindBy(xpath="//ul[@id='Addmenu']/li[6]/a/span")
	public static WebElement createRoleAddTabPointOfSaleMenu;

	@FindBy(xpath="//ul[@id='Addmenu']/li[7]/a/span")
	public static WebElement createRoleAddTabQualityControlMenu;

	@FindBy(xpath="//ul[@id='Addmenu']/li[8]/a/span")
	public static WebElement createRoleAddTabSettingsMenu;

	//@FindBy(xpath="//div[@id='Additions']//div[@id='controlBtns']/div/i[1]")
	@FindBy(xpath="(//*[@id='controlBtns']/i[1])[1]")
	public static WebElement createRoleAddTabSelectAllChkBox;

	@FindBy(xpath="//div[@id='Additions']//div[@id='controlBtns']/div/i[2]")
	public static WebElement createRoleAddTabUnSelectAllChkBox;


	// Exclusions Tab
	@FindBy(xpath="//ul[@id='Exclmenu']/li/a/span")
	public static List<WebElement> createRoleExclTabMenuList;

	@FindBy(xpath="//ul[@id='Exclmenu']/li[1]/a/span")
	public static WebElement createRoleExclTabHomeMenu;

	@FindBy(xpath="//ul[@id='Exclmenu']/li[2]/a/span")
	public static WebElement createRoleExclTabFinancialsMenu;

	@FindBy(xpath="//ul[@id='Exclmenu']/li[3]/a")
	public static WebElement createRoleExclTabInventoryMenu;

	@FindBy(xpath="//ul[@id='Exclmenu']/li[4]/a/span")
	public static WebElement createRoleExclTabFixedAssetsMenu;

	@FindBy(xpath="//ul[@id='Exclmenu']/li[5]/a/span")
	public static WebElement createRoleExclTabProductionMenu;

	@FindBy(xpath="//ul[@id='Exclmenu']/li[6]/a/span")
	public static WebElement createRoleExclTabPointOfSaleMenu;

	@FindBy(xpath="//ul[@id='Exclmenu']/li[7]/a/span")
	public static WebElement createRoleExclTabQualityControlMenu;

	@FindBy(xpath="//ul[@id='Exclmenu']/li[8]/a/span")
	public static WebElement createRoleExclTabSettingsMenu;

	@FindBy(xpath="//div[@id='Exclusions']//div[@id='controlBtns']/div/i[1]")
	public static WebElement createRoleExclTabSelectAllChkBox;

	//@FindBy(xpath="//div[@id='Exclusions']//div[@id='controlBtns']/div/i[2]")
	@FindBy(xpath="(//*[@id='controlBtns']/i[2])[2]")
	public static WebElement createRoleExclTabUnSelectAllChkBox;



	// Restriction for Entry Tab
	@FindBy(xpath="//*[@id='RestrictionEntryMasters']//li")
	public static List<WebElement> createRoleRestrictionForEntryTabMastersList;

	@FindBy(xpath="//input[@id='chkExclusion']/following-sibling::span")
	public static WebElement createRoleRestrictionForEntryTabExclChkBox;

	@FindBy(xpath="//table[@id='restrictionEntryTable']")
	public static WebElement createRoleRestrictionForEntryTabTable;

	@FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[1]/td[2]")
	public static WebElement restrictionForEntryTabTableSelectRow1Col1;

	@FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[1]/td[3]//span")
	public static WebElement restrictionForEntryTabRow1EntryChkBox;

	@FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[1]/td[4]//span")
	public static WebElement restrictionForEntryTabRow1ReportChkBox;

	@FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[1]/td[5]//span")
	public static WebElement restrictionForEntryTabRow1ViewChkBox;


	@FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[2]/td[2]")
	public static WebElement restrictionForEntryTabTableSelectRow2Col1;

	@FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[2]/td[3]//span")
	public static WebElement restrictionForEntryTabRow2EntryChkBox;

	@FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[2]/td[4]//span")
	public static WebElement restrictionForEntryTabRow2ReportChkBox;

	@FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[2]/td[5]//span")
	public static WebElement restrictionForEntryTabRow2ViewChkBox;


	@FindBy(xpath="//input[@id='MasteroptionControl']")
	public static WebElement restrictionForEntryTabTableEnterMasterTxt;

	@FindBy(xpath="//tbody[@id='MasteroptionControl_table_body']/tr/td[2]")
	public static List<WebElement> restrictionForEntryTabTableMasterComboList;


	//Restriction for trees Tab
	@FindBy(xpath="//div[@id='RestrictionTreeMasters']/ul/li")
	public static List<WebElement> createRoleRestrictionForTreesTabMastersList;

	@FindBy(xpath="//*[@id='RestrictionTreeMastersTree']/li/label/span")
	public static WebElement restrictionForTreesTabDefaultChkBox;

	// AI Rights
	@FindBy(xpath="//input[@id='EnableVUI']/following-sibling::span")
	public static WebElement AIRightsEnableVUIChkBox;

	@FindBy(xpath="//input[@id='EnableAI']/following-sibling::span")
	public static WebElement AIRightsEnableAIChkBox;

	@FindBy(xpath="//input[@id='Donotapplythetagrestrictiontotheresults']/following-sibling::span")
	public static WebElement doNotApplyTheTagResTrictionToTheResultsChkBox;

	@FindBy(xpath="//input[@id='Noofvoicecommandspermitted']")
	public static WebElement NoOfVoiceCommandsPermittedTxt;

	@FindBy(xpath="//input[@id='Nooftextcommandspermitted']")
	public static WebElement NoOfTextCommandsPermittedTxt;

	@FindBy(xpath="//input[@id='Sales']/following-sibling::span")
	public static WebElement AIRightsSalesChkBox;

	@FindBy(xpath="//input[@id='Salesorder']/following-sibling::span")
	public static WebElement AIRightsSalesOrderChkBox;

	@FindBy(xpath="//input[@id='Stock']/following-sibling::span")
	public static WebElement AIRightsStockChkBox;

	@FindBy(xpath="//input[@id='Purchases']/following-sibling::span")
	public static WebElement AIRightsPurchasesChkBox;

	@FindBy(xpath="//input[@id='Purchaesorders']/following-sibling::span")
	public static WebElement AIRightsPurchasesOrdersChkBox;

	@FindBy(xpath="//input[@id='accountsreceivables']")
	public static WebElement AIRightsAccountsReceivablesChkBox;

	@FindBy(xpath="//input[@id='accountsPayables']")
	public static WebElement AIRightsAccountsPayablesChkBox;

	@FindBy(xpath="//input[@id='Finance']")
	public static WebElement AIRightsFinanceChkBox;


	@FindBy(xpath="//*[@id='Addmenu']/li[2]/ul/li/a")
	public static List<WebElement> additionTabFinancialsMenuList;

	@FindBy(xpath="//*[@id='Addmenu']/li[2]/ul/li[1]/a/span")
	public static WebElement additionTabFinancialsTransaction;

	@FindBy(xpath="//*[@id='Addmenu']/li[2]/ul/li[1]/ul/li/a")
	public static List<WebElement> additionTabFinancialsTransactionMenusList;

	@FindBy(xpath="//*[@id='Addmenu']/li[2]/ul/li[1]/ul/li[2]/ul/li/a")
	public static List<WebElement> additionTabFinancialsTransactionsPurchasesMenusList;

	@FindBy(xpath="//*[@id='AddrestrictionsDiv']/li[1]/label/input")
	public static WebElement CRAddAlwaysSuspendOnSavingChkBoxIsSelected;
	
	@FindBy(xpath="//*[@id='AddrestrictionsDiv']/li[1]/label/input/following-sibling::span")
	public static WebElement CRAddAlwaysSuspendOnSavingChkBox;

	@FindBy(xpath="//*[@id='AddrestrictionsDiv']/li[2]/label/input/following-sibling::span")
	public static WebElement CRAddHideMenuChkBox;
	
	@FindBy(xpath="//*[@id='AddrestrictionsDiv']/li[2]/label/input")
	public static WebElement CRAddHideMenuChkBoxIsSelected;

	@FindBy(xpath="//div[@class='Fcheckbox']/label/input[@name='Deletedocumentscreatedbyothers']")
	public static WebElement CRAddDeleteDocumentsCreatedByOthersChkBox;

	@FindBy(xpath="//div[@class='Fcheckbox']/label/input[@name='Deleteauthorizeddocuments']")
	public static WebElement CRAddDeleteAuthorizedDocumentsChkBox;

	@FindBy(xpath="//div[@class='Fcheckbox']/label/input[@name='EditSuspendedDocuments']")
	public static WebElement CRAddEditSuspendedDocumentsChkBox;

	@FindBy(xpath="//*[@id='Exclmenu']/li[3]/ul/li/a")
	public static List<WebElement> exclusionsTabInventoryMenusList;

	@FindBy(xpath="//*[@id='Exclmenu']/li[3]/ul/li[1]/ul/li/a")
	public static List<WebElement> exclusionsTabInventoryTransactionsMenusList;

	@FindBy(xpath="//div[@class='Fcheckbox']/label/input[@name='Save']")
	public static WebElement CreateRoleSaveChkBox;

	@FindBy(xpath="//div[@class='Fcheckbox']/label/input[@name='Access']")
	public static WebElement CreateRoleAccessChkBox;

	@FindBy(xpath="//div[@class='Fcheckbox']/label/input[@name='Edit']")
	public static WebElement CreateRoleEditChkBox;


	@FindBy(xpath="//a[contains(text(),'Restriction for entry')]")
	public static WebElement createRoleRestrictionForEntryTab;

	@FindBy(xpath="//div[@id='RestrictionEntryMasters']//li[@id='1']")
	public static WebElement resOFEntryAccount;

	@FindBy(xpath="//div[@id='RestrictionEntryMasters']//li[@id='2']")
	public static WebElement resOFEntryItem;

	@FindBy(xpath="//div[@id='RestrictionEntryMasters']//li[@id='3']")
	public static WebElement resOFEntryDepartment;

	@FindBy(xpath="//div[@id='RestrictionEntryMasters']//li[@id='4']")
	public static WebElement resOFEntryWarehouse;

	@FindBy(xpath="//input[@id='chkExclusion']")
	public static WebElement resOFEntryExclusionchkBox;

	@FindBy(xpath="//div[contains(text(),'Masters')]")
	public static WebElement  resOFEntryGridMasterColoumn;

	@FindBy(xpath="//table[@id='restrictionEntryTable']")
	public static WebElement  resOFEntryTable;

	//@FindBy(xpath="//*[@id='TransactionRights']/div/div/div[1]/div[1]/div[2]/label/span")
	@FindBy(xpath="//*[@id='TransactionRights']/div/div[1]/div[1]/div[2]/label/span")
	public static WebElement allowBillWiseOnAccountChkBox;

	@FindBy(xpath="//select[@id='BudgetLimitWarning']")
	public static WebElement budgetLimitWarningDropdown;

	@FindBy(xpath="//select[@id='CreditLimitWarning']")
	public static WebElement creditLimitWarningDropdown;

	@FindBy(xpath="//select[@id='NegativeCashCheck']")
	public static WebElement negativeCashCheckDropdown;

	@FindBy(xpath="//select[@id='NegativeStockCheck']")
	public static WebElement negativeStockCheckDropdown;

	@FindBy(xpath="//*[@id='editOptionDiv']/div[2]/label/span")
	public static WebElement cantaddfutureTransChkbox;

	@FindBy(xpath="//input[@id='CantAdTransThtAremrethan']")
	public static WebElement cannotAddTransThatAreMoreThanTxt;

	@FindBy(xpath="//input[@id='CantEditPrevTransaremorethn']")
	public static WebElement cannotEditPreviousMonthEntriesAfterTxt;

	@FindBy(xpath="//input[@id='CantEditTranstharemorethn']")
	public static WebElement CantEditTransthatAreremorethanTxt;

	@FindBy(xpath="//input[@id='CantAddPrevMnthEntriesaft']")
	public static WebElement CantAddPrevMnthEntriesaftTxt;

	@FindBy(xpath="//input[@id='CantPrintAfterValue']")
	public static WebElement CantPrintAfterValueTxt;

	@FindBy(xpath="//select[@id='CantPrintAfterOption']")
	public static WebElement cantPrintAfterValueDropdown;

	@FindBy(xpath="//input[@id='CantRe-PrintAfterValue']")
	public static WebElement CantRePrintAfterValueTxt;

	@FindBy(xpath="//select[@id='CantRe-PrintAfterOption']")
	public static WebElement CantRePrintAfterDropdown;

	@FindBy(xpath="//*[@id='editOptionDiv']/div[9]/label/span")
	public static WebElement doNotAllowMasterCustomization;


	@FindBy(xpath="//div[@id='Additions']//label[@class='Flabel font-4'][contains(text(),'Menu')]")
	public static WebElement menuLabel;

	@FindBy(xpath="//ul[@id='Addmenu']/li[1]/i")
	public static WebElement  addHomeExpandBtn; 

	@FindBy(xpath="//ul[@id='Addmenu']/li[2]/i")
	public static WebElement  addFininicalExpandBtn; 

	@FindBy(xpath="//ul[@id='Addmenu']/li[3]/i")
	public static WebElement  addInventoryExpandBtn; 

	@FindBy(xpath="//ul[@id='Addmenu']/li[4]/i")
	public static WebElement  addFixedAssestsExpandBtn; 

	@FindBy(xpath="//ul[@id='Addmenu']/li[5]/i")
	public static WebElement  addProductionExpandBtn; 

	@FindBy(xpath="//ul[@id='Addmenu']/li[6]/i")
	public static WebElement  addPointOfSaleExpandBtn; 

	@FindBy(xpath="//ul[@id='Addmenu']/li[7]/i")
	public static WebElement  addQualityExpandBtn; 

	@FindBy(xpath="//ul[@id='Addmenu']/li[8]/i")
	public static WebElement  addSettingExpandBtn; 

	@FindBy(xpath="//ul[@id='Addmenu']/li[2]/ul/li[1]/i")
	public static WebElement  addFinTransationExpandBtn;

	@FindBy(xpath="//ul[@id='Addmenu']/li[2]/ul/li[1]/ul/li[2]/i")
	public static WebElement addFinTransationpurchaseExpandBtn;

	@FindBy(xpath="(//*[@id='2008'])[2]")
	public static WebElement addFinTransPurPurchasesVouchers;

	@FindBy(xpath="//li[@class='active']//span[contains(text(),'Purchases Returns')]")
	public static WebElement addFinTransPurPurchasesReturns;

	@FindBy(xpath="//li[@class='active']//span[contains(text(),'Purchase Vouchers N')]")
	public static WebElement addFinTransPurPurchaseVouchersN;

	@FindBy(xpath="//label[contains(text(),'Always suspend on saving')]//input[@id='2008']")
	public static WebElement purAlwaysSuspendChkBox;

	@FindBy(xpath="//label[contains(text(),'Hide Menu')]//input[@id='2008']")
	public static WebElement purHideMenuChkbox;

	@FindBy(xpath="//li[1]//div[1]//label[1]//input[1]")
	public static WebElement finPurAccessChkbox;

	@FindBy(xpath="//li[2]//div[1]//label[1]//input[1]")
	public static WebElement finPurAddChkbox;

	@FindBy(xpath="//li[3]//div[1]//label[1]//input[1]")
	public static WebElement finPurEditChkbox;

	@FindBy(xpath="//label[contains(text(),'Delete')]//input[@id='2008']")
	public static WebElement finPurDeleteChkbox;

	@FindBy(xpath="//li[5]//div[1]//label[1]//input[1]")
	public static WebElement finPurPrintChkbox;

	@FindBy(xpath="//li[6]//div[1]//label[1]//input[1]")
	public static WebElement finPurRePrintChkbox;

	@FindBy(xpath="//label[contains(text(),'Print Barcode')]//input[@id='2008']")
	public static WebElement finPurPrintBarcodeChkbox;

	@FindBy(xpath="//label[contains(text(),'Print Un-Authorized documents')]//input[@id='2008']")
	public static WebElement finPurPrintUnAuthorizedDocumentsChkbox;

	@FindBy(xpath="//label[contains(text(),'Export')]//input[@id='2008']")
	public static WebElement finPurExportChkbox;

	@FindBy(xpath="//label[contains(text(),'Edit documents entered by other')]//input[@id='2008']")
	public static WebElement finPurEditDocumentsEnteredbyOtherChkbox;

	@FindBy(xpath="//label[contains(text(),'Edit documents that are checked')]//input[@id='2008']")
	public static WebElement finPurEditDocumentsThatAreCheckedChkbox;

	@FindBy(xpath="//label[contains(text(),'Edit Reconciled documents')]//input[@id='2008']")
	public static WebElement finPurEditReconciledDocumentsChkbox;

	@FindBy(xpath="//label[contains(text(),'Edit documents authorization by higher ups')]//input[@id='2008']")
	public static WebElement finPurEditDocumentsAuthorizationByHigherUpsChkbox;

	@FindBy(xpath="//label[contains(text(),'Enter documents that exceed limit')]//input[@id='2008']")
	public static WebElement finPurEnterDocumentsThatExceedLimitChkbox;

	@FindBy(xpath="//label[contains(text(),'Enter documents that make Cash or Bank Balance Neg')]//input[@id='2008']")
	public static WebElement finPurChkboxEnterDocumentsThatMakeCashorBankBalanceNeg;

	@FindBy(xpath="//label[contains(text(),'Edit documents that are Re-Printed')]//input[@id='2008']")
	public static WebElement finPurEditDocumentsThatAreRePrintedChkbox;

	@FindBy(xpath="//label[contains(text(),'Access through API')]//input[@id='2008']")
	public static WebElement finPurAccessThroughAPIChkbox;

	@FindBy(xpath="//label[contains(text(),'Always suspend on saving')]//input[@id='2008']")
	public static WebElement finPurAlwaysSuspendOnSavingChkbox;

	@FindBy(xpath="//label[contains(text(),'Change print layout')]//input[@id='2008']")
	public static WebElement finPurChangePrintLayoutChkbox;

	@FindBy(xpath="//li[20]//div[1]//label[1]//input[1]")
	public static WebElement finPurAuthorizeChkbox;

	@FindBy(xpath="//label[contains(text(),'Reject')]//input[@id='2008']")
	public static WebElement finPurRejectChkbox;

	@FindBy(xpath="//label[contains(text(),'Show home page')]//input[@id='2008']")
	public static WebElement finPurShowHomepageChkbox;

	@FindBy(xpath="//label[contains(text(),'Stop')]//input[@id='2008']")
	public static WebElement finPurStopChkbox;

	@FindBy(xpath="//label[contains(text(),'Modify Settings')]//input[@id='2008']")
	public static WebElement finPurModifySettingsChkbox;

	@FindBy(xpath="//label[contains(text(),'View document entered by others')]//input[@id='2008']")
	public static WebElement finPurViewDocumentEnteredByOthersChkbox;

	@FindBy(xpath="//label[contains(text(),'Close Links')]//input[@id='2008']")
	public static WebElement finPurCloseLinksChkbox;

	@FindBy(xpath="//label[contains(text(),'Save Revision')]//input[@id='2008']")
	public static WebElement finPurSaveRevisionChkbox;

	@FindBy(xpath="//label[contains(text(),'Show Posting Details')]//input[@id='2008']")
	public static WebElement finPurShowPostingDetailsChkbox;

	@FindBy(xpath="//label[contains(text(),'Suspend')]//input[@id='2008']")
	public static WebElement finPurSuspendChkbox;

	@FindBy(xpath="//label[contains(text(),'Edit exported (Sync) documents')]//input[@id='2008']")
	public static WebElement finPurEditExportedChkbox;

	@FindBy(xpath="//label[contains(text(),'Edit imported (Sync) documents')]//input[@id='2008']")
	public static WebElement finPurEditImportedChkbox;

	@FindBy(xpath="//label[contains(text(),'Hide Menu')]//input[@id='2008']")
	public static WebElement finPurHideMenuChkbox;

	@FindBy(xpath="//label[contains(text(),'Amend')]//input[@id='2008']")
	public static WebElement finPurAmendChkbox;

	@FindBy(xpath="//label[contains(text(),'Customize link')]//input[@id='2008']")
	public static WebElement finPurCustomizeLinkChkbox;

	@FindBy(xpath="//label[contains(text(),'Filter home page')]//input[@id='2008']")
	public static WebElement finPurFilterHomePageChkbox;

	@FindBy(xpath="//label[contains(text(),'Add to stock')]//input[@id='2008']")
	public static WebElement finPurAddToStockChkbox;

	@FindBy(xpath="//label[contains(text(),'Reverse Entry')]//input[@id='2008']")
	public static WebElement finPurReverseEntryChkbox;

	@FindBy(xpath="//label[contains(text(),'Customize info panel')]//input[@id='2008']")
	public static WebElement finPurCustomizeInfoPanelChkbox;

	@FindBy(xpath="//label[contains(text(),'Raise Cheque Return')]//input[@id='2008']")
	public static WebElement finPurRaiseChequeReturnChkbox;

	@FindBy(xpath="//li[@class='active']//span[contains(text(),'Request for Quote')]")
	public static WebElement addRequestForQuoteOption;

	@FindBy(xpath="//li[@class='active']//span[contains(text(),'Purchases Quotations')]")
	public static WebElement addPurchaseQuotationOption;

	@FindBy(xpath="//li[@class='active']//a[@id='2012']//span[contains(text(),'Purchases Orders')]")
	public static WebElement addPurchaseorderOption;

	@FindBy(xpath="//li[@class='active']//span[contains(text(),'Material Receipt Notes')]")
	public static WebElement addMaterialRecepitNotes;

	@FindBy(xpath="//*[@id='btndelete']/i")
	public static WebElement CreateRoleDeleteButtonImage; 



	@FindBy(xpath="//*[@id='createUserBtn']/i")
	public static WebElement createUserIcon;

	@FindBy(xpath="//*[@id='createGroupBtn']/i")
	public static WebElement createUserCreateGroupIcon;

	@FindBy(xpath="//*[@id='navbarSupported3']/ul/li[3]/a/i")
	public static WebElement createUserCancelIcon;

	@FindBy(xpath="//input[@id='loginName']")
	public static WebElement createUserLoginNameTxt;

	@FindBy(xpath="//select[@id='ERPRoles']")
	public static WebElement createUserERPRoleDropDown;

	@FindBy(xpath="//select[@id='passwordPolicy']")
	public static WebElement createUserPasswordPolicyDropdown;

	@FindBy(xpath="//input[@id='password']")
	public static WebElement createUserPasswordTxt;

	@FindBy(xpath="//input[@id='confirmpassword']")
	public static WebElement createUserConfirmPasswordTxt;

	@FindBy(xpath="//input[@id='Username']")
	public static WebElement createUserNameTxt;

	@FindBy(xpath="//input[@id='lginAbbr']")
	public static WebElement createUserLoginAbbrivationTxt;

	@FindBy(xpath="//select[@id='SecurityQues']")
	public static WebElement createUserSecurityQuestionDropDown;

	@FindBy(xpath="//input[@id='SecurityAns']")
	public static WebElement createUserSecurityAnswerTxt;

	@FindBy(xpath="//select[@id='language']")
	public static WebElement createUserLanguageTxt;

	@FindBy(xpath="//select[@id='altLanguage']")
	public static WebElement createUserAlternateLanguageTxt;

	@FindBy(xpath="//i[@class='icon-unlock-user icon-font6']")
	public static WebElement createUserUnlockUserIcon;

	@FindBy(xpath="//i[@class='icon-move-user icon-font6']")
	public static WebElement createUserMoveUserIcon;

	@FindBy(xpath="//i[@class='icon-reset icon-font6']")
	public static WebElement createUserResetIcon;

	@FindBy(xpath="//*[@id='btnLoadFrom']/i")
	public static WebElement createUserLoadFromIcon;	

	@FindBy(xpath="//span[@id='btnSave']//i[@class='icon-save icon-font6']")
	public static WebElement createUserSaveIcon;

	@FindBy(xpath="//*[@id='CancelCreateUser']/i")
	public static WebElement createUserCloseIcon;

	@FindBy(xpath="//input[@id='emailId']")
	public static WebElement createUserEmailTxt;

	@FindBy(xpath="//input[@id='Emailpassword']")
	public static WebElement createUserEmailPasswordTxt;

	@FindBy(xpath="//input[@id='phone']")
	public static WebElement createUserPhoneTxt;

	@FindBy(xpath="//input[@id='mobile']")
	public static WebElement createUserMobileTxt;

	@FindBy(xpath="//input[@id='IsEmailAuthPermission']")
	public static WebElement createUserSetPermissionForEmailAuthorizationChkBox;

	@FindBy(xpath="//select[@id='DominSecurityMping']")
	public static WebElement createUserDomainUserMappingDrpDwn;

	@FindBy(xpath="//input[@id='domainUserName']")
	public static WebElement createUserDomainUserMappingTxt;

	@FindBy(xpath="//input[@id='ipAddr1']")
	public static WebElement createUserUserValuesTxt1;

	@FindBy(xpath="//input[@id='ipAddr2']")
	public static WebElement createUserUserValuesTxt2;

	@FindBy(xpath="//input[@id='ipAddr3']")
	public static WebElement createUserUserValuesTxt3;

	@FindBy(xpath="//input[@id='ipAddr4']")
	public static WebElement createUserUserValuesTxt4;

	@FindBy(xpath="//input[@id='ipAddr5']")
	public static WebElement createUserUserValuesTxt5;

	@FindBy(xpath="//select[@id='typeOfUser']")
	public static WebElement createUserUserTypeDrpDwn;

	@FindBy(xpath="//input[@id='cmbUserTypeMaster']")
	public static WebElement createUserUserTypeCombo;

	@FindBy(xpath="//select[@id='crmRoles']")
	public static WebElement createUserCRMRolesDrpDwn;

	@FindBy(xpath="//input[@id='UserTypeMobile']")
	public static WebElement createUserMobileChkBox;

	@FindBy(xpath="//input[@id='UserTypeOffline']")
	public static WebElement createUserOfflineChkBox;

	@FindBy(xpath="//input[@id='UserTypeWeb']")
	public static WebElement createUserWebChkBox;

	@FindBy(xpath="//select[@id='ERPRoles']/option")
	public static List<WebElement> createUserERPRolesDropdownOptions;

	@FindBy(xpath="//select[@id='passwordPolicy']//option[6]")
	public static WebElement getNumericValue;

	@FindBy(xpath="//select[@id='language']")
	public static WebElement getLanguage;

	@FindBy(xpath="//select[@id='altLanguage']")
	public static WebElement getAltLanguage;

	@FindBy(xpath="//select[@id='language']/option")
	public static List<WebElement> createUserLanguageDropdownList;

	@FindBy(xpath="//select[@id='altLanguage']/option")
	public static List<WebElement> createUserAltLanguageDropdownList;

	@FindBy(xpath="//*[contains(text(),'Additional Info')]")
	public static WebElement createUserAdditionalInfoTab;

	@FindBy(xpath="//div[@id='userAndGroups']/ul/li/ul/li/div/span/span[1]")
	public static List<WebElement> createUserSavedUsersList;

	@FindBy(xpath="//i[@class='icon-delete icon-font6']")
	public static WebElement createUserDeleteIcon;

	@FindBy(xpath="//*[contains(text(),'Restriction for entry')]")
	public static WebElement createUserRestrictionforentryTab;

	@FindBy(xpath="//div[@id='id_login_div']//div[contains(text(),'Invalid Password')]")
	public static WebElement invalidPasswordMsgs;

	@FindBy(xpath="//*[@id='60']")
	public static WebElement financialsMenu;

	@FindBy(xpath="//*[@id='61']")
	public static WebElement  financialsTransactionMenu;

	@FindBy(xpath="//a[@id='2007']//span[contains(text(),'Purchases')]")
	public static WebElement  purchasesExpandBtn;

	@FindBy(xpath="//span[contains(text(),'Purchases Vouchers')]")
	public static WebElement  purchaseVouchersBtn;

	@FindBy(xpath="//*[@id='id_transaction_homescreen_new']/i")
	public static WebElement newBtn;

	//@FindBy(xpath="//*[@id='id_header_4_input_image']/span")
	@FindBy(xpath="//*[@id='id_header_4_input_container']/div[2]/table/tbody/tr/td[1]/i")
	public static WebElement vendorAccountComboExpand;

	//@FindBy(xpath="//*[@id='id_header_4_input_image']/span")
	@FindBy(xpath="//*[@id='id_header_4_input_container']/div[2]/table/tbody/tr/td[1]/i")
	public static WebElement customerAccountComboExpand;

	@FindBy(xpath="//*[@id='id_header_4_table_body']/tr/td[2]")
	public static List<WebElement> vendorAccountList;

	@FindBy(xpath="//*[@id='id_header_4_table_body']/tr/td[2]")
	public static List<WebElement> customerAccountList;

	@FindBy(xpath="//*[@id='2014']/span")
	public static WebElement salesMenu;

	@FindBy(xpath="//*[@id='2055']/span")
	public static WebElement salesInvoicesVatBtn;

	@FindBy(xpath="//*[@id='81']/span")
	public static WebElement  financialsReportsMenu; 

	@FindBy(xpath="//*[@id='500']/span")
	public static WebElement  ledger;

	@FindBy (xpath="//*[@id='txtsrch-term']")
	public static WebElement accountSearchTextArea;

	@FindBy (xpath="//tbody[@id='LandingGridBody']//td[@class='tdName padding0 ']")
	public static List<WebElement> accountsListInLedger;

	@FindBy(xpath="//select[@id='timezone']")
	public static WebElement userCreationTimeZoneDrpdwn;



	// Document Customization



	@FindBy (xpath="//table[@id='editScreen_footerTable']/tbody/tr[2]/td/span[1]")
	public static WebElement editScreenFooterRow2Edit;

	@FindBy (xpath="//table[@id='editScreen_footerTable']/tbody/tr[2]/td/span[2]")
	public static WebElement editScreenFooterRow2Delete;

	@FindBy (xpath="//table[@id='editScreen_footerTable']/tbody/tr[2]/td[3]")
	public static WebElement editScreenFooterRow2Caption;

	@FindBy (xpath="//table[@id='editScreen_footerTable']/tbody/tr[2]/td[8]")
	public static WebElement editScreenFooterRow2Formula;

	@FindBy (xpath="//table[@id='editScreen_footerTable']/tbody/tr")
	public static List <WebElement> editScreenFooterRowsCount;


	@FindBy (xpath="//table[@id='editLayout_headerTable']/tbody/tr")
	public static List <WebElement> editLayoutHeaderRowsCount;


	@FindBy (xpath="//table[@id='editScreen_bodyTable']/tbody/tr[2]/td/span[1]")
	public static WebElement editScreenBodyRow2Edit;
	
	@FindBy (xpath="//table[@id='editScreen_bodyTable']/tbody/tr[3]/td/span[1]")
	public static WebElement editScreenBodyRow3Edit;

	@FindBy (xpath="//table[@id='editScreen_bodyTable']/tbody/tr[3]/td/span[2]")
	public static WebElement editScreenBodyRow3Delete;

	@FindBy (xpath="//table[@id='editScreen_bodyTable']/tbody/tr[3]/td[3]")
	public static WebElement editScreenBodyRow3Caption;

	@FindBy (xpath="//table[@id='editScreen_bodyTable']/tbody/tr[3]/td[8]")
	public static WebElement editScreenBodyRow3Formula;


	@FindBy (xpath="//table[@id='editScreen_bodyTable']/tbody/tr")
	public static List <WebElement> editScreenBodyRowsCount;

	@FindBy(xpath="//input[@id='views_existingViews']")
	public static WebElement  ViewExistingViewTxt;

	@FindBy(xpath="//*[@id='views_User_1']/following-sibling::span")
	public static WebElement  viewSuChkbox;
	
	@FindBy(xpath="//*[@id='views_User_1']")
	public static WebElement  viewSuChkboxIsSelected;

	//@FindBy(xpath="//ul[@id='id_documentcustomization_views_list_users']/li[2]/div/label/input/following-sibling::span")
	
//	@FindBy(xpath="//*[@id='id_documentcustomization_views_list_users']/li[2]/div/label")
//	@FindBy(xpath="(//*[@id='views_User_3'])[1]")
	@FindBy(xpath="//*[@id='views_User_3']/following-sibling::span")
	public static WebElement  viewsUserAllOptionsSTChkbox;
	
	
//	@FindBy(xpath="//*[@id='id_documentcustomization_views_list_users']/li[2]/div/label")
//	@FindBy(xpath="//*[@id='id_documentcustomization_views_list_users']/li[2]/div/label/span")
	@FindBy(xpath="//*[@id='views_User_3']")
	public static WebElement  viewsUserAllOptionsSTChkboxSelected;
	
	@FindBy(xpath="//*[@id='views_User_4']/following-sibling::span")
	public static WebElement  viewsUserRestrictChkbox;
	
	@FindBy(xpath="//*[@id='views_User_4']")
	public static WebElement  viewsUserRestrictChkboxSelected;
	
	
	//@FindBy(xpath="//ul[@id='id_documentcustomization_views_list_users']/li[3]/div/label/input")
	@FindBy(xpath="(//*[@id='views_User_3'])[1]")
	public static WebElement  viewsUserAllOptionsSTChkboxIsSelected;

	@FindBy(xpath="//*[@id='views_tabContent']/div/div/div[1]/a[9]/i")
	public static WebElement  viewsDeleteViewIcon;

	@FindBy(xpath="//*[@id='spnSaveView']/i")
	public static WebElement  viewSaveView;

	//@FindBy(xpath="//*[@id='panelsStayOpen-collapseOne']/div[2]/div/label[1]")
	@FindBy(xpath="//*[@id='views_chkauthorizeOnly']")
	public static WebElement  viewsAuthorizeOnlyChkboxIsSelected;
	
	@FindBy(xpath="//*[@id='views_chkauthorizeOnly']/following-sibling::span")
	//@FindBy(xpath="(//*[@id='panelsStayOpen-collapseOne']/div[2])[3]//div//label[1]//span")
	public static WebElement  viewsAuthorizeOnlyChkbox;



	@FindBy(xpath="//input[@id='triggers_existTriggers']")
	public static WebElement  editLayoutTriggerName;

	@FindBy(xpath="//td[@id='triggers_existTriggers_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
	public static WebElement  triggersExistTriggersSettingBtn;

	@FindBy(xpath="//*[@id='trigger_Controlbuttons']/a[1]/i")
	public static WebElement  editLayoutTriggerDeleteBtn;

	@FindBy(xpath="//*[@id='trigger_Controlbuttons']/a[2]/i")
	public static WebElement  editLayoutTriggersSaveBtn;

	@FindBy(xpath="//*[@id='17_1_AdvanceFilter_']/table/tbody/tr/td[1]/select")////*[@id="17_1_AdvanceFilter_"]/table/tbody/tr/td[1]/select
	public static WebElement  triggersConjuctionDrpdwn;

	@FindBy(xpath="//*[@id='17_1_AdvanceFilter_']/table/tbody/tr/td[2]/input")
	public static WebElement  triggersSelectFieldTxt;

	@FindBy(xpath="//*[@id='17_1_AdvanceFilter_']/table/tbody/tr/td[3]/select")
	public static WebElement  triggersSelectOperdrpdwn;

	@FindBy(xpath="//*[@id='17_1_AdvanceFilter_']/table/tbody/tr/td[4]/select")
	public static WebElement  triggersCompareWithdrpdwn;

//	@FindBy(xpath="//tbody//input[@placeholder='Value']")
	@FindBy(xpath="//*[@id='advancefilter_master_17_1_']")
	public static WebElement  triggersValueTxt;


	@FindBy(xpath="//*[@id='advancefilter_master_17_1_']")
	public static WebElement  triggersValueToEnterTxt;
	//*[@id='advancefilter_master_17_1_']



	@FindBy(xpath="//tbody//span[@class='icon-close icon-font6']")
	public static WebElement  triggersTrigCreteriaCloseBtn;

	@FindBy(xpath="//tbody//span[@class='icon-pluse icon-font6']")
	public static WebElement  triggersTrigCreteriaAddBtn;

	@FindBy(xpath="//tr[2]//td[6]//span[1]")
	public static WebElement  triggersTrigCretriaRow2CloseBtn;

	@FindBy(xpath="//input[@id='triggers_chkraiseNewDoc']")
	public static WebElement editLayoutTriggerRaiseDoc;  	

//	@FindBy(xpath="//input[@id='triggers_chkraiseNewDoc']")
	@FindBy(xpath="//*[@id='triggers_chkraiseNewDoc']/following-sibling::span")
	public static WebElement  triggersRiseANewDocChkbox;

	@FindBy(xpath="//*[@id='triggers_chksuspend']/following-sibling::span")
	public static WebElement  triggersSuspendedChkboxs;
	
	@FindBy(xpath="//*[@id='triggers_chksuspend']")
	public static WebElement  triggersSuspendedChkboxIsSelected;

	@FindBy(xpath="//*[@id='triggers_chksuspend']/following-sibling::span")
	public static WebElement  triggersSuspended;

	@FindBy(xpath="//input[@id='rdbOnSave']")
	public static WebElement  triggersOnSaveRadioBtn;

	@FindBy(xpath="//input[@id='rdbOnAuthorization']")
	public static WebElement  triggersOnAuthorizatinRadioBtn;

	@FindBy(xpath="//input[@id='triggers_chksendAlert']")
	public static WebElement  triggersSendAlertChkbox;

	@FindBy(xpath="//select[@id='triggers_sendAlert']")
	public static WebElement  triggersSendAlertDrpdwn;

	@FindBy(xpath="//*[@id='triggers_FieldsDiv']/div[2]/div[3]/div[1]/button")
	public static WebElement  triggersSendAlertExpansionBtn;

	@FindBy(xpath="//input[@id='id_trigger_alertdateformula']")
	public static WebElement  triggersAlertDataFormula;

	@FindBy(xpath="//input[@id='triggers_chksendEmail']")
	public static WebElement  triggersSendEmailChkbox;

	@FindBy(xpath="//input[@id='Triggers_LayoutID']")
	public static WebElement  triggersSendEmailDrpdwn;

	@FindBy(xpath="//div[@class='col-xs-12 col-sm-6']//div[2]//button[1]")
	public static WebElement  triggersSendEmailExpansionBtn;

	@FindBy(xpath="//input[@id='triggers_chksendSMS']")
	public static WebElement  triggersSendSMSChkbox;

	@FindBy(xpath="//select[@id='triggers_sendSMS']")
	public static WebElement  triggersSendSMSDrpdwn;

	@FindBy(xpath="//div[@class='col-xs-12 col-sm-6']//div[@class='col-xs-12']//div[3]//button[1]")
	public static WebElement  triggersSendSMSExpansionBtn;

	@FindBy(xpath="//input[@id='triggers_smsconfirm']")
	public static WebElement  triggersConformBeforeSendingChkbox;

	@FindBy(xpath="//input[@id='triggers_smsonedit']")
	public static WebElement  triggersSendSMSonEdit;

	@FindBy(xpath="//select[@id='triggers_PhoneNo']")
	public static WebElement  triggersPhoneNoDrpdwn;

	@FindBy(xpath="//input[@id='triggers_chksendTypeCustomeVendor']")
	public static WebElement  triggersCustmerORVendorChkbox;

	@FindBy(xpath="//input[@id='triggers_chksendTypeLogin']")
	public static WebElement  triggersLoginChkbox;

	@FindBy(xpath="//input[@id='triggers_chksendTypeExtraField']")
	public static WebElement  triggersExtraFieldChkbox;

	@FindBy(xpath="//input[@id='triggers_Role']")
	public static WebElement  triggersRoleChkbox;

	@FindBy(xpath="//input[@id='triggers_Auth']")
	public static WebElement  triggersAuthorisationChkbox;

	@FindBy(xpath="//select[@id='triggers_Auth_DD']")
	public static WebElement  triggersDefaultDrpdwn;


	@FindBy(xpath="//*[@id='docNum_Table_head']/tr/th/div")
	public static List<WebElement>  docNumberingGridHeader;	

	@FindBy(xpath="//*[@id='docNum_Table_body']/tr/td")
	public static List<WebElement>  docNumberingGridBody;	

	@FindBy(xpath="//td[@id='docNum_Table_col_1-0']")
	public static WebElement  docNUmneringFirstRowIndex;	

	@FindBy(xpath="//td[@id='docNum_Table_col_1-1']")
	public static WebElement  typeRow1;	

	@FindBy(xpath="//td[@id='docNum_Table_col_1-2']")
	public static WebElement  fieldValueRow1;	

	@FindBy(xpath="//td[@id='docNum_Table_col_1-3']")
	public static WebElement  startingCharacterRow1;	

	@FindBy(xpath="//td[@id='docNum_Table_col_1-4']")
	public static WebElement  noofCharactersRow1;	

	@FindBy(xpath="//td[@id='docNum_Table_col_2-0']")
	public static WebElement  docNUmneringSecondRowIndex;	

	@FindBy(xpath="//td[@id='docNum_Table_col_2-1']")
	public static WebElement  typeRow2;	

	@FindBy(xpath="//td[@id='docNum_Table_col_2-2']")
	public static WebElement  fieldValueRow2;	

	@FindBy(xpath="//td[@id='docNum_Table_col_2-3']")
	public static WebElement  startingCharacterRow2;	

	@FindBy(xpath="//td[@id='docNum_Table_col_2-4']")
	public static WebElement  noofCharactersRow2;	

	@FindBy(xpath="//td[@id='docNum_Table_col_3-0']")
	public static WebElement  docNUmneringThkirdRowIndex;	

	@FindBy(xpath="//td[@id='docNum_Table_col_3-1']")
	public static WebElement  typeRow3;	

	@FindBy(xpath="//td[@id='docNum_Table_col_3-2']")
	public static WebElement  fieldValueRow3;	

	@FindBy(xpath="//td[@id='docNum_Table_col_3-3']")
	public static WebElement  startingCharacterRow3;	

	@FindBy(xpath="//td[@id='docNum_Table_col_3-4']")
	public static WebElement  noofCharactersRow3;	

	@FindBy(xpath="//td[@id='docNum_Table_col_4-0']")
	public static WebElement  docNUmneringFourthRowIndex;	

	@FindBy(xpath="//td[@id='docNum_Table_col_4-1']")
	public static WebElement  typeRow4;	

	@FindBy(xpath="//td[@id='docNum_Table_col_4-2']")
	public static WebElement  fieldValueRow4;	

	@FindBy(xpath="//td[@id='docNum_Table_col_4-3']")
	public static WebElement  startingCharacterRow4;	

	@FindBy(xpath="//td[@id='docNum_Table_col_4-4']")
	public static WebElement  noofCharactersRow4;	

	@FindBy(xpath="//td[@id='docNum_Table_col_5-0']")
	public static WebElement  docNUmneringFifthRowIndex;	

	@FindBy(xpath="//td[@id='docNum_Table_col_5-1']")
	public static WebElement  typeRow5;	

	@FindBy(xpath="//td[@id='docNum_Table_col_5-2']")
	public static WebElement  fieldValueRow5;	

	@FindBy(xpath="//td[@id='docNum_Table_col_5-3']")
	public static WebElement  startingCharacterRow5;	

	@FindBy(xpath="//td[@id='docNum_Table_col_5-4']")
	public static WebElement  noofCharactersRow5;	

	@FindBy(xpath="//td[@id='docNum_Table_col_6-0']")
	public static WebElement  docNUmneringSixthRowIndex;	

	@FindBy(xpath="//td[@id='docNum_Table_col_6-1']")
	public static WebElement  typeRow6;	

	@FindBy(xpath="//td[@id='docNum_Table_col_6-2']")
	public static WebElement  fieldValueRow6;	

	@FindBy(xpath="//td[@id='docNum_Table_col_6-3']")
	public static WebElement  startingCharacterRow6;	

	@FindBy(xpath="//td[@id='docNum_Table_col_6-4']")
	public static WebElement  noofCharactersRow6;	

	@FindBy(xpath="//td[@id='docNum_Table_col_7-0']")
	public static WebElement  docNUmneringSeventhRowIndex;	

	@FindBy(xpath="//td[@id='docNum_Table_col_7-1']")
	public static WebElement  typeRow7;	

	@FindBy(xpath="//td[@id='docNum_Table_col_7-2']")
	public static WebElement  fieldValueRow7;	

	@FindBy(xpath="//td[@id='docNum_Table_col_7-3']")
	public static WebElement  startingCharacterRow7;	

	@FindBy(xpath="//td[@id='docNum_Table_col_7-4']")
	public static WebElement  noofCharactersRow7;	



	@FindBy(xpath="//select[@id='docNum_TypeDropDown']")
	public static WebElement  typeDropDown;	

	@FindBy(xpath="//input[@id='docNum_FieldORValueDropDown']")
	public static WebElement  fieldValueTxt;	

	@FindBy(xpath="//select[@id='docNum_FieldORValueDropDown']")
	public static WebElement  fieldValueDropDown;	

	@FindBy(xpath="//input[@id='docNum_startCharacterTxtbox']")
	public static WebElement  startingCharacterTxt;	

	@FindBy(xpath="//input[@id='docNum_noofCharacterTxtbox']")
	public static WebElement  noOfCharacterTxt;

	//@FindBy(xpath="//label[@class='Flabel col-xs-12 col-sm-5 col-md-3']")
	
	@FindBy(xpath="//*[@id='docNum_documentNumberAs']")
	public static WebElement  docNumberAppearAsLabel;

	@FindBy(xpath="//label[@id='docNum_documentNumberAs']")
	public static WebElement  docNumberAppearAsValue;

	@FindBy (xpath="//*[@id='docNum_Table_col_1-0']")
	public static WebElement docNumberingRow1Col1;

	@FindBy (xpath="//*[@id='docNum_Table_col_2-0']")
	public static WebElement docNumberingRow2Col1;
	
	
	
	@FindBy (xpath="//*[@id='DocumentNumbering']//div[2]/input[1]")
	public static WebElement docNumberingDeleteRowBtn;
	

//	@FindBy (xpath="//div[contains(text(),'Delete Row')]")
	@FindBy(xpath="//*[@id='DocumentNumbering']/div/div/div/div[2]/input[1]")
	public static WebElement docNumberingRow1DeleteBtn;


	@FindBy(xpath = "//*[@id='lnkRule_div_DocCustRules']/input")
	public static WebElement rulesTabAddRuleBtn;

	@FindBy(xpath = "//input[@id='txtMasterRuleName_DocCustRules']")
	public static WebElement rulesTabRuleNameTxt;

	@FindBy(xpath = "//td[@id='txtMasterRuleName_DocCustRules_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
	public static WebElement rulesTabRuleNameSettingBtn;

	@FindBy(xpath = "//*[@id='Conditions-Context_Menu_DocCustRules']/li[2]/a")
	public static WebElement rulesTabElseTab;


	@FindBy(xpath = "//*[@id='Conditions-Context_Menu_DocCustRules']/li[1]/a")
	public static WebElement rulesTabIFTab;

	@FindBy(xpath = "//input[@id='txtField_RC_1_1_DocCustRules']")
	public static WebElement rulesTabIFSelectTabTxt;

	@FindBy(xpath = "//select[@id='ddlMasterFields_RULES_DocCustRules']")
	public static WebElement rulesTabIFSelectTabDrpdwn;


	@FindBy(xpath = "//select[@id='ddlExtraFields_1_1_DocCustRules']")
	public static WebElement rulesTabExtraFieldDrpdwn;


	@FindBy(xpath = "//select[@id='ddlOperators_RC_1_1_DocCustRules']")
	public static WebElement rulesTabOperatorsDrpdwn;

	@FindBy(xpath = "//select[@id='ddlCompareWith_RC_1_1_DocCustRules']")
	public static WebElement rulesTabCompareWithDrpdwn;

	@FindBy(xpath = "//input[@id='txtCompareType_RC_1_1_DocCustRules']")
	public static WebElement rulesTabValueTxt;


	@FindBy(xpath = "//input[@id='txtCompareType_RC_1_1_DocCustRules']")
	public static WebElement rulesTabConditionTxttoEnter;

	@FindBy(xpath = "//input[@id='optControl_RulesCondition_DocCustRules']	")
	public static WebElement rulesTabConditionTxt;





	@FindBy(xpath = "//input[@id='optControl_RulesCondition_DocCustRules']")
	public static WebElement rulesTabConditioDrpdwn;




	@FindBy(xpath = "//select[@id='ddlConjunction_RC_1_1_DocCustRules']")
	public static WebElement rulesTabConjunctionDrpdwn;

	//@FindBy(xpath = "//input[@id='chkNoRuleCondition_DocCustRules']")
	
	@FindBy(xpath="//*[@id='chkNoRuleCondition_DocCustRules']/following-sibling::span")
	public static WebElement rulesTabNoConditionChkbbox;

	@FindBy(xpath = "//i[@id='iconRuleSettings_DocCustRules']")
	public static WebElement rulesTabSettingsAddIcon;

	@FindBy(xpath = "//span[@class='icon-close icon-font7']")
	public static WebElement rulesTabIfConditionCloseBtn;

	@FindBy(xpath = "//*[@id='Rule_IF_DocCustRules']/div[2]/div[1]/ul/li[1]/a/span")
	public static WebElement rulesTab_IFFormattingTab;

	//@FindBy(xpath = "//*[@id='Rule_IF_DocCustRules']/div[2]/div[1]/ul/li[2]/a/span")
	@FindBy(xpath="//*[@id='tab_Rule_IF_Messaging_DocCustRules']/a")
	public static WebElement rulesTab_IFMesssageTab;

	@FindBy(xpath = "//*[@id='Rule_IF_DocCustRules']/div[2]/div[1]/ul/li[3]/a/span")
	public static WebElement rulesTab_IFAlertsTab;

	@FindBy(xpath = "//*[@id='Rule_Else_DocCustRules']/div/div[1]/ul/li[1]/a/span")
	public static WebElement rulesTab_ElseFormattingTab;

	@FindBy(xpath = "//*[@id='Rule_Else_DocCustRules']/div/div[1]/ul/li[2]/a/span")
	public static WebElement rulesTab_ElseMessageTab;

	@FindBy(xpath = "//*[@id='Rule_Else_DocCustRules']/div/div[1]/ul/li[3]/a/span")
	public static WebElement rulesTab_ElseAlertsTab;

	@FindBy(xpath = "//*[@id='btnAddRule']")
	public static WebElement rulesTab_SaveRuleBtn;

	@FindBy(xpath = "//button[contains(text(),'Delete Rule')]")
	public static WebElement rulesTab_DeleteRuleBtn;

	@FindBy (xpath="//*[@id='DeleteRule_Modal_DocCustRules']/div[2]/div/div[3]/div/div/input[1]")
	public static WebElement deleteRuleYesBtn;


	@FindBy(xpath="//*[@id='configTransMenu']/ul/li[2]/span[2]")
	public static WebElement  docCustCloseBtn;

	@FindBy(xpath="//li/span")
	public static WebElement searchMenuTextClick;

	@FindBy(xpath="//div[@id='docCustomization']")
	public static WebElement  documentCustomization45Btn;

	@FindBy(xpath="//*[@id='tbPreference']")
	public static WebElement  configureTransactionPreferencesMinus;

	//@FindBy(xpath="//*[@id='configureTransactions']/div[2]/div[1]/div/div/ul/li[2]/a/i")
	@FindBy(xpath="//*[@id='configureTransactions']/div/div[1]/div/div/ul/li[2]/a/i")
	public static WebElement  configureTransactionDocumentCustomiztionMinus;
	
	@FindBy(xpath="//a[@id='769'][text()='Purchase Voucher VAT']")
	public static WebElement documentCustomizationPurchaseVoucherVAT;
	

	@FindBy(xpath="//span[@class='icon-general icon-font6 theme_button_color']")
	public static WebElement  createVoucherBtn;

	@FindBy(xpath="//span[@id='spanDeleteVoucher']")
	public static WebElement  deletevoucherBtn;

	@FindBy(xpath="//a[@id='4608']")
	public static WebElement  receiptsBtn;

	@FindBy(xpath="//a[@id='4864']")
	public static WebElement  paymentsBtn;

	@FindBy(xpath="//a[@id='5120']")
	public static WebElement  pettyCashBtn;

	@FindBy(xpath="//a[@id='5888']")
	public static WebElement  postDatedreceiptsBtn;

	@FindBy(xpath="//a[@id='7168']")
	public static WebElement  postDatedpaymentsBtn;

//	@FindBy(xpath="//*[@id='docCustomizationUList']/li[6]/a")
	@FindBy(xpath="//*[@id='768']")
	public static WebElement  docCustPurchaseVouchersBtn;

	@FindBy(xpath="//a[@id='1023']")
	public static WebElement  stockAdjustmentBtn;

	@FindBy(xpath="//a[@id='6400']")
	public static WebElement  purchaseReturnsBtn;

	@FindBy(xpath="//a[@id='2560']")
	public static WebElement  purcahseOrdersBtn;

	@FindBy(xpath="//a[@id='2561']")
	public static WebElement  jobOrdersBtn;

	@FindBy(xpath="//a[@id='1280']")
	public static WebElement  materialReceiptsNotesBtn;

	@FindBy(xpath="//a[@id='3328']//i[contains(@class,'icon-font7')]")
	public static WebElement  salesInvoiceBtn;

	@FindBy(xpath="//a[@id='3329']")
	public static WebElement  cashSalesBtn;

	@FindBy(xpath="//a[@id='3330']")
	public static WebElement  hirePurcahseSalesBtn;

	@FindBy(xpath="//a[@id='3331']")
	public static WebElement  posSalesBtn;

	@FindBy(xpath="//a[@id='1792']")
	public static WebElement  salesReturnsBtn;

	@FindBy(xpath="//a[@id='5632']")
	public static WebElement  salesOrdersBtn;

	@FindBy(xpath="//a[@id='5633']")
	public static WebElement  productionOrderBtn;

	@FindBy(xpath="//a[@id='2019']")
	public static WebElement  deliveryNotesBtn;

	@FindBy(xpath="//a[@id='3584']")
	public static WebElement  nonStandardjournalEntriesBtn;

	@FindBy(xpath="//a[@id='3585']")
	public static WebElement  interDepartmentalJVBtn;

	@FindBy(xpath="//a[@id='3586']")
	public static WebElement  debitNotesLineWiseBtn;

	@FindBy(xpath="//a[@id='3587']")
	public static WebElement  creditNotesLineWiseBtn;

	@FindBy(xpath="//a[@id='3588']")
	public static WebElement  fixedAssetDepreciationVoucherBtn;

	@FindBy(xpath="//a[@id='8704']")
	public static WebElement  journalEntriesBtn;
	
	

	@FindBy(xpath="//a[@id='3840']")
	public static WebElement  debitNotesBtn;

	@FindBy(xpath="//a[@id='4096']")
	public static WebElement  creditNotesBtn;

	@FindBy(xpath="//a[@id='4097']")
	public static WebElement  payRollPostingsBtn;

	@FindBy(xpath="//a[@id='256']")
	public static WebElement  openingBalancesBtn;

	@FindBy(xpath="//a[@id='5376']")
	public static WebElement  shortageInStocksBtn;

	@FindBy(xpath="//div[@class='col-xs-12']//a[@id='2048']")
	public static WebElement  excessInStocksBtn;

	@FindBy(xpath="//a[@id='3072']")
	public static WebElement  stockTransfersBtn;

	@FindBy(xpath="//div[@class='col-xs-12']//a[@id='512']")
	public static WebElement  openingStocksBtn;

	@FindBy(xpath="//a[@id='6656']")
	public static WebElement  issuesToProductionBtn;

	@FindBy(xpath="//a[@id='1024']")
	public static WebElement  receiptsFromProductionBtn;

	@FindBy(xpath="//a[@id='1056']")
	public static WebElement  returnsFromShopfloorBtn;

	@FindBy(xpath="//a[@id='6912']")
	public static WebElement  jobWorkIssuesBtn;

	@FindBy(xpath="//a[@id='1536']")
	public static WebElement  jobWorkReeceiptsBtn;

	@FindBy(xpath="//a[@id='2304']")
	public static WebElement  purchaseQuotationsBtn;

	@FindBy(xpath="//a[@id='7424']")
	public static WebElement  salesQuotationsBtn;

	@FindBy(xpath="//a[@id='7680']")
	public static WebElement  jobOrderBtn;

	@FindBy(xpath="//a[@id='2816']")
	public static WebElement  productionProcessBtn;

	@FindBy(xpath="//a[@id='7936']")
	public static WebElement  materialRequisitionBtn;

	@FindBy(xpath="//a[@id='8192']")
	public static WebElement  requestForQuoteBtn;

	@FindBy(xpath="//a[@id='8448']")
	public static WebElement  forecxJVBtn;

	@FindBy(xpath="//a[@id='updateButton']/i")
	public static WebElement  updateBtn;

	@FindBy(xpath="//*[@id='navigationTabs']/li[1]")
	public static WebElement  documentsTab;

	@FindBy(xpath="//*[@id='panelsStayOpen-headingOne']/button")
	public static WebElement docCustDocumentsTabGeneralBtn;
	
	@FindBy(xpath="//*[@id='doc_title']")
	public static WebElement  titleTxt;

	@FindBy(xpath="//td[@id='doc_TagsTable_col_1-1']")
	public static WebElement  masters1;

	@FindBy(xpath="//td[@id='doc_TagsTable_col_1-2']")
	public static WebElement  position1;

	@FindBy(xpath="//td[@id='doc_TagsTable_col_1-3']")
	public static WebElement  showDependency1;

	@FindBy(xpath="//td[@id='doc_TagsTable_col_1-4']")
	public static WebElement  group1;

	@FindBy(xpath="//td[@id='doc_TagsTable_col_1-5']")
	public static WebElement  filter1;

	@FindBy(xpath="//td[@id='doc_TagsTable_col_1-6']")
	public static WebElement  mandatory1;


	@FindBy(xpath="//td[@id='doc_TagsTable_col_2-1']")
	public static WebElement  masters2;

	@FindBy(xpath="//td[@id='doc_TagsTable_col_2-2']")
	public static WebElement  position2;

	@FindBy(xpath="//td[@id='doc_TagsTable_col_2-3']")
	public static WebElement  showDependency2;

	@FindBy(xpath="//td[@id='doc_TagsTable_col_2-4']")
	public static WebElement  group2;

	@FindBy(xpath="//td[@id='doc_TagsTable_col_2-5']")
	public static WebElement  filter2;

	@FindBy(xpath="//td[@id='doc_TagsTable_col_2-6']")
	public static WebElement  mandatory2;

	@FindBy(xpath="//select[@id='doc_TagsTableMasterDropDown']")
	public static WebElement  masterDropDown;

	@FindBy(xpath="//select[@id='doc_TagsPositionDropDown']")
	public static WebElement  positionDropdown;


	// Miscellaneous Tab

	@FindBy(xpath="//*[@id='navigationtab2']")
	public static WebElement  miscellaneousTab;
	
	//@FindBy(xpath="//*[@id='Miscellaneous']//div/h2[1]/button")
	@FindBy(xpath="(//*[@id='panelsStayOpen-headingOne']/button)[2]")
	public static WebElement miscellaneousTabVoucherAccountSettingsBtn;

	@FindBy(xpath="//input[@id='misc_Abbrivation']")
	public static WebElement  abbreviationTxt;


	@FindBy(xpath="//*[@id='panelsStayOpen-headingTwo']/button")
	public static WebElement miscellaneusTabPostingDetailsMenu;
	
	@FindBy(xpath="//*[@id='panelsStayOpen-headingten']/button")
	public static WebElement miscellaneusTabARAPBtn;
	
	@FindBy(xpath="//select[@id='misc_arapDueDate']")
	public static WebElement dueDateDropdown;

	@FindBy(xpath="//*[@id='Miscellaneous']//div/h2[3]/button")
	public static WebElement miscellaneousTabCurrencyMenu;

	@FindBy(xpath="//*[@id='panelsStayOpen-collapseTwo']/label[1]/span")
	public static WebElement updateFAChkBox;
	
	
	
	@FindBy(xpath="//*[@id='panelsStayOpen-collapseTwo']/label[1]/input")
	public static WebElement updateFAChkBoxIsSelected;
	
	//@FindBy(xpath="//*[@id='panelsStayOpen-collapseTwo']/label[4]/span")
	@FindBy(xpath="//*[@id='id_misc_updatestock_section']/span")
	public static WebElement updateStockChkBox;
	
	//@FindBy(xpath="//*[@id='panelsStayOpen-collapseTwo']/label[4]/input")
	@FindBy(xpath="//*[@id='id_misc_updatestock_section']/input")
	public static WebElement updateStockChkBoxIsSelected;

	@FindBy(xpath="//select[@id='misc_CurrencyAddCurrencyIn']")
	public static WebElement addCurrencyInDropdown;

	@FindBy(xpath="//*[@id='panelsStayOpen-collapseThree']/label[2]/span")
	public static WebElement inputExchangeRateChkBox;
	
	@FindBy(xpath="//*[@id='panelsStayOpen-collapseThree']/label[2]/input")
	public static WebElement inputExchangeRateChkBoxIsSelected;

	@FindBy(xpath="//ul[@id='navigationTabs']/li[4]/a")
	public static WebElement  editScreenTab; 

	@FindBy(xpath="//*[@id='editScreen_tabContent']//nav/div/div/ul//li[1]")
	public static WebElement editScreenaddBtn;

	@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_Caption']")
	public static WebElement  editScreenCaptionTxt;

	@FindBy(xpath="//select[@id='editScreen_FieldsCustomization_behaviour']")
	public static WebElement editScreenAddFieldsBehaviourDropdown;

	//@FindBy(xpath="//*[@id='editScreen_addFieldsPopupDiv']/div/div[1]/div/span[1]/i")
	@FindBy(xpath="//*[@id='dvCustomizeEditScreenField']/div[2]/div/div[3]/input[1]")
	public static WebElement editScreenApplyBtn;  
	
	@FindBy(xpath="//*[@id='editScreen_bodyTable']/tbody/tr/td[3]")
	public static List<WebElement> editScreenbodyCaption;
	
	@FindBy(xpath="//*[@id='editScreen_bodyTable']/tbody/tr/td/span[1]")
	public static List<WebElement> editScreenbodyEditBtn;
	
	@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_Formula_textbox']")
	public static WebElement  editScreenAddBehaviourformulaTxt;

	@FindBy(xpath="//*[@id='editScreen_tabContent']//ul/li[2]/a/i")
	public static WebElement editScreenLoadFieldsBtn;

	@FindBy(xpath="//select[@id='editScreen_voucherDropDown']")
	public static WebElement editScreenLoadFieldsDropdown;  

	@FindBy(xpath="//div[@id='dvVWCommonDetails']/div[2]/div/div[3]/input[2]")
	public static WebElement editScreenokBtn;  



	@FindBy(xpath="//*[@id='BodyTabMenu']")
	public static WebElement  editScreenBodyTab;

	@FindBy(xpath="//*[@id='FooterTabMenu']")
	public static WebElement  editScreenFooterTab;  

	@FindBy(xpath="//ul[@id='navigationTabs']/li[3]/a")
	public static WebElement editLayoutTab;

	@FindBy(xpath="//*[@id='editLayout_tabContent']//nav/div/div/ul//li[1]")
	public static WebElement editLayoutAddFieldsBtn;

	@FindBy(xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_Caption']")
	public static WebElement  editLayoutCaptionTxt;

	@FindBy(xpath="//*[@id='dvCustomizeEditLayoutField']/div[2]/div/div[3]/input[1]")
	public static WebElement editLayoutApplyBtn;  
	

	@FindBy(xpath="//td[@id='editLayout_headerrow_2 _4']")
	public static WebElement  editLayoutHeader2ndRowCaption;

	@FindBy(xpath="//td[@id='editLayout_bodyrow_1 _4']")
	public static WebElement  editLayoutbody1stRowCaption;

	@FindBy(xpath="//*[@id='editLayout_tabContent']//ul/li[2]/a/i")
	public static WebElement editLayoutLoadFieldsBtn;  

	@FindBy(xpath="//*[@id='editLayout_voucherDropDown']")
	public static WebElement editLayoutLoadFieldsDropdown;  

	@FindBy(xpath="(//*[@id='dvVWCommonDetails']/div[2])[1]/div/div[3]/input[1]")
	public static WebElement editLayoutselectAllBtn;  

	@FindBy(xpath="(//*[@id='dvVWCommonDetails']/div[2])[1]/div/div[3]/input[2]")
	public static WebElement editLayoutokBtn;     
	
	@FindBy(xpath="//*[@id='navigationtab5']")
	public static WebElement externalFunctionsBtn;
	
	@FindBy(xpath="//*[@id='navigationtab15']")
	public static WebElement schemesBtn;
	
	@FindBy(xpath="//*[@id='navigationtab6']")
	public static WebElement  exportFieldsBtn;
	
	@FindBy(xpath="//*[@id='navigationtab10']")
	public static WebElement  hirePurchaseBtn;
	
	@FindBy(xpath="//*[@id='navigationtab12']")
	public static WebElement  reportsBtn;


	@FindBy(xpath="//*[@id='HeaderTabMenu']")
	public static WebElement  editLayoutHeaderTab;  

	@FindBy(xpath="//*[@id='editLayout_headerrow_2 _1']/span[1]")
	public static WebElement  editLayoutHeader2ndRowEditBtn; 

	@FindBy(xpath="//*[@id='editLayout_headerrow_2 _1']/span[2]")
	public static WebElement  editLayoutHeader2ndRowDeleteBtn; 

	@FindBy(xpath="//*[@id='BodyTabMenu']")
	public static WebElement  editLayoutBodyTab; 

	@FindBy(xpath="//*[@id='editLayout_bodyrow_1 _1']/span[1]")
	public static WebElement  editLayoutBody1stRowEditBtn; 

	@FindBy(xpath="//*[@id='editLayout_bodyrow_1 _1']/span[2]")
	public static WebElement  editLayoutBody1stRowDeleteBtn; 

	//@FindBy(xpath="//input[@id='chkNewRecord_Rule_DocCustRules']")
	@FindBy(xpath="//*[@id='chkNewRecord_Rule_DocCustRules']/following-sibling::span")
	public static WebElement rulesNewRecordChekbox ;

	//@FindBy(xpath="//input[@id='chkEdit_Rule_DocCustRules']")
	@FindBy(xpath="//*[@id='chkEdit_Rule_DocCustRules']/following-sibling::span")
	public static WebElement rulesEditChekbox ;

	//@FindBy(xpath="//input[@id='chkLoad_Rule_DocCustRules']")
	@FindBy(xpath="//*[@id='chkLoad_Rule_DocCustRules']/following-sibling::span")
	public static WebElement rulesLoadChekbox ;

	//@FindBy(xpath="//input[@id='chkOnEnter_Rule_DocCustRules']")
	@FindBy(xpath="//*[@id='chkOnEnter_Rule_DocCustRules']/following-sibling::span")
	public static WebElement rulesOnEnterChekbox ;

//	@FindBy(xpath="//*[@id='Rule_IF_DocCustRules']/div[2]/div[1]/ul/li[2]/a/span")
	
	@FindBy(xpath="//*[@id='tab_Rule_IF_Messaging_DocCustRules']/a")
	public static WebElement rulesIFMessageTab ;

	@FindBy(xpath="//input[@id='txtGeneralMsg_RuleIfMsg_DocCustRules']")
	//@FindBy(xpath="//*[@id="txtGeneralMsg_RuleIfMsg_DocCustRules"]")
	public static WebElement rulesIFMessagesGenralMessageTxt;

	@FindBy(xpath="//ul[@id='navigationTabs']/li[8]/a")
	public static WebElement  inventoryoptionsTab;
	
	@FindBy(xpath="//*[@id='InventoryOptions']//div/h2[1]/button")
	public static WebElement  inventoryoptionsTabInventorySettingBtn;

	@FindBy(xpath="//input[@id='inventoryOpt_chkApplyRateSchemes']")
	public static WebElement applyrateSchemesChkBoxIsSelected;
	
	/*@FindBy(xpath="//input[@id='inventoryOpt_chkApplyRateSchemes']/following-sibling::span")
	public static WebElement applyrateSchemesChkBox;
	*/
	
	@FindBy(xpath="//*[@id='panelsStayOpen-collapseOne']/div[3]//label[1]/span")
	public static WebElement applyrateSchemesChkBox;
	
	
	

	@FindBy(xpath="//ul[@id='navigationTabs']/li[14]/a")
	public static WebElement  viewsBtn;
	
	@FindBy(xpath="//*[@id='views_tabContent']//div/h2[1]/button")
	public static WebElement viewTabViewSettingsBtn;
	
	@FindBy(xpath="//*[@id='panelsStayOpen-headingTwo']/button")
	public static WebElement viewTabLoginBtn;

	@FindBy(xpath="//*[@id='navigationTabs']/li[13]/a")
	public static WebElement  triggersBtn;
	
	@FindBy(xpath="(//*[@id='accordionPanelsStayOpenExample']//div/h2[1]/button)[1]")
	public static WebElement  triggerTabTriggersBtn;
	
	@FindBy(xpath="(//*[@id='accordionPanelsStayOpenExample']//div/h2[1]/button)[2]")
	public static WebElement  triggerTabActionBtn;
	
	@FindBy(xpath="(//*[@id='panelsStayOpen-headingTwo']/button)[3]")
	public static WebElement triggerTabAlertBtn;
	
	@FindBy(xpath="//ul[@id='navigationTabs']/li[2]/a")
	public static WebElement  documentNumberingBtn;

	@FindBy(xpath="//ul[@id='navigationTabs']/li[11]/a")
	public static WebElement  rulesBtn;
	
	@FindBy(xpath="//input[@id='misc_PostingSelectAccInLine']")
	public static WebElement selectAccountInEveryLineChkBoxIsSelected;
	
	@FindBy(xpath="//input[@id='misc_PostingSelectAccInLine']/following-sibling::span")
	public static WebElement selectAccountInEveryLineChkBox;

	@FindBy(xpath="//*[@id='id_misc_inputlocalexchangerate_section']/label/span")
	public static WebElement inputLocalExchangeRateChkBox;
	
	@FindBy(xpath="//*[@id='id_misc_inputlocalexchangerate_section']/label/input")
	public static WebElement inputLocalExchangeRateChkBoxIsSelected;

	@FindBy (xpath="//div[@id='dvSearchBx']/div[2]/span[1]")
	public static WebElement docCustCreateVoucherBtn;

	@FindBy(xpath="//*[@id='769']")
	public static WebElement docCustCreatedVoucher;
	
	@FindBy(xpath="//*[@id='editScreen_copyFieldsOptFields']/div/label")
	public static List<WebElement>  editScreenLoadFieldTxtList;

	@FindBy(xpath="//*[@id='editScreen_copyFieldsOptFields']/div/label/span")
	public static List<WebElement>  editScreenLoadFieldchkboxList;

	@FindBy(xpath="//*[@id='105']")
	public static WebElement  settingsWizardMenu;

	@FindBy(xpath="//a[@id='1'][text()='VAT']")
	public static WebElement settingsWizardVATBtn;

	@FindBy(xpath="//div[@id='dvRightSettingWizard']/div/div[2]/div[1]/div/div/label")
	public static List<WebElement> settingsWizardVATVouchersList;

	//@FindBy(xpath="//div[@id='dvRightSettingWizard']/div[2]/div[1]/div/div/label/input")
	@FindBy(xpath="//div[@id='dvRightSettingWizard']//div[2]//div//label")
	public static List<WebElement> settingsWizardVATVoucherChkboxList;
	
	@FindBy(xpath="//div[@id='dvRightSettingWizard']//div[2]//div//span")
	public static List<WebElement> settingsWizardVATVoucherChkboxListSelected;

	@FindBy(xpath="//div[@id='dvSelectedVouchers']/div/label")
	public static List<WebElement> settingsWizardVATSelectedVouchersList;

	@FindBy(xpath="//label[contains(text(),'Gross')]/span")
	public static WebElement settingsWizardVATGrossChkbox;

	@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[1]")
	public static WebElement settingsWizardUpdateBtn;

	@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[2]")
	public static WebElement settingsWizardVAtCloseBtn;
	
	@FindBy(xpath="//*[@id='cmbBehaviour']")
	public static WebElement behaviourSelect;


	// Transaction Authorization

	@FindBy(xpath="//*[@id='68']/span")
	public static WebElement  transactionAuthorizationBtn;

	@FindBy(xpath="//*[@id='id_menu_tree_Purchases Returns']/a/span")
	public static WebElement purchasesReturns; 

	@FindBy(xpath="//*[@id='id_menu_tree_Purchases Returns (1)']/a/span")
	public static WebElement purchasesReturns1; 

	@FindBy(xpath="//*[@id='txtDoc']")
	public static WebElement documentNameDropdown; 

	@FindBy(xpath="//*[@id='spnAddLevel']/i")
	public static WebElement addLevelBtn; 

	@FindBy(xpath="//*[@id='txtDescription']")
	public static WebElement descriptionTxt; 

	//@FindBy(xpath="//*[@id='chkConditionNotRequired']")
	@FindBy(xpath="//*[@id='ConditionPanel']/div/div[1]")
	public static WebElement conditionNotRequiredChkBox; 

	@FindBy(xpath="//*[@id='id_report_popup_body']/div[2]/div[1]/ul[1]/li[2]/a")
	public static WebElement userSelectionTab;

	@FindBy(xpath="//*[@id='id_report_popup_body']/div[2]/div[1]/ul[1]/li[3]/a")
	public static WebElement alertsTab;

	@FindBy(xpath="//u[contains(text(),'Email Header Template')]")
	public static WebElement emailHeaderTemplateLink;

	@FindBy(xpath="//input[@id='smsTemplateOptCtrl']")
	public static WebElement templateTxt;

	@FindBy(xpath="//tbody[@id='smsTemplateOptCtrl_table_body']/tr/td[2]")
	public static List<WebElement> custTemplatesList;

	@FindBy(xpath="//span[@id='btnCloseSMSModal']")
	public static WebElement custTemplateCloseBtn;

	@FindBy(xpath="//*[@id='ddlUserSelection']")
	public static WebElement userselectionDropdown;

	@FindBy(xpath="//*[@id='ddlCriteria']")
	public static WebElement criteriaDropdown;

	@FindBy(xpath="//*[@id='id_tblUserSelection_User']")
	public static WebElement userDropdown;

	@FindBy(xpath="//td[@id='id_tblUserSelection_col_1-1']")
	public static WebElement userrow1;

	@FindBy(xpath="//td[@id='id_tblUserSelection_col_2-1']")
	public static WebElement userrow2;

	@FindBy(xpath="//*[@id='id_report_popup_ok']")
	public static WebElement definitionOkBtn;					

	@FindBy(xpath="//span[@id='id_report_popup_cancel']")
	public static WebElement definitionCloseBtn;	

	@FindBy(xpath="//*[@id='ulAuthBtns']/li[5]/span/i")
	public static WebElement tranAuthCancelBtn; 

	@FindBy(xpath="//*[@id='btnSaveAuthorization']")
	public static WebElement tranAuthSaveBtn; 

	@FindBy(xpath="//*[@id='spndeleteProfile']")
	public static WebElement tranAuthDeleteBtn; 

	@FindBy(xpath="//*[@id='ctrlAuthorizationH']")
	public static WebElement transAuthNameTxt;
	
	@FindBy(xpath="(//*[@id='divLevel'])[2]/div[1]/i")
	public static WebElement Level1Icon;

	@FindBy(xpath="(//*[@id='divLevel'])[2]/div[2]/ul/li[5]")
	public static WebElement Level1editBtn;

	@FindBy(xpath="//*[@id='68_0_AdvanceFilter_']/table/tbody/tr[1]/td[1]/select")
	public static WebElement Row1ConjunctionDropdown; 

	@FindBy(xpath="//*[@id='68_0_AdvanceFilter_']/table/tbody/tr[1]/td[2]/input")
	public static WebElement Row1SelectFieldTxt; 

	@FindBy(xpath="//*[@id='68_0_AdvanceFilter_']/table/tbody/tr[1]/td[3]/select")
	public static WebElement Row1SelectOperatorDropdown; 

	@FindBy(xpath="//*[@id='68_0_AdvanceFilter_']/table/tbody/tr[1]/td[4]/select")
	public static WebElement Row1CompareWithDropdown; 

	@FindBy(xpath="//*[@id='68_0_AdvanceFilter_']/table/tbody/tr[1]/td[5]/input")
	public static WebElement Row1ValueTxt; 

	@FindBy(xpath="//input[@id='advancefilter_master_68_0_']")
	public static WebElement Row1VendorValueTxt; 

	@FindBy(xpath="//*[@id='advancefilter_master_68_0__table_body']/tr/td[2]")
	public static List<WebElement> Row1ValueList; 

	@FindBy(xpath="//input[@id='advancefilter_master_68_0_']")
	public static WebElement InputValueTxt; 

	@FindBy(xpath="//div[@id='filterTree_68_0_AdvanceFilter_']//a[@id='4']")
	public static WebElement selectFieldVendorAC;

	@FindBy(xpath="//*[@id='filterTree_68_0_AdvanceFilter_']/ul/li[4]/ul/li[1]/a")
	public static WebElement selectFieldVendorACName;

	@FindBy(xpath="//div[@id='filterTree_68_0_AdvanceFilter_']//a[@id='23']")
	public static WebElement selectFieldItem;

	@FindBy(xpath="//div[@id='filterTree_68_0_AdvanceFilter_']//a[@id='5021']")
	public static WebElement selectFieldItemName;

	@FindBy(xpath="//div[@id='filterTree_68_0_AdvanceFilter_']//a[@id='26']")
	public static WebElement selectFieldQuantity;


	// Master Authorization

	@FindBy(xpath="//a[@id='id_showAllUnAuth']/span")
	public static WebElement masterShowAllUnauthorisedRecords;

	@FindBy(xpath="//label[@class='msgInfo']")
	public static WebElement clickOnEmptyGroup;

	@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[8]//input")
	public static List<WebElement> masterAccountsChkboxList;

	@FindBy (xpath="//a[@id='1107']")
	public static WebElement warehouseMenu;

	@FindBy(xpath="//*[@id='67_0_AdvanceFilter_']/table/tbody/tr/td[1]/select")
	public static WebElement conjunctionDropdown;

	@FindBy(xpath="//tbody//input[@placeholder='Select Field']")
	public static WebElement selectFieldTxt; 

	@FindBy(xpath="(//a[@id='5058'])[1]")
	public static WebElement sName;

	@FindBy(xpath="//*[@id='67_0_AdvanceFilter_']/table/tbody/tr/td[3]/select")
	public static WebElement selectOperatorDropdown; 

	@FindBy(xpath="//*[@id='67_0_AdvanceFilter_']/table/tbody/tr/td[4]/select")
	public static WebElement compareWithDropdown; 

	@FindBy(xpath="//*[@id='67_0_AdvanceFilter_']/table/tbody/tr/td[5]/input")
	public static WebElement valueTxt;

	@FindBy(xpath="//*[@id='67']/span")
	public static WebElement  masterAuthorizationBtn;

	@FindBy(xpath="//input[@id='txtDoc']")
	public static WebElement  masterDropdown;

	@FindBy(xpath="//li[contains(text(),'Warehouse')]")
	public static WebElement  masterAuthWarehouse;

	@FindBy(xpath="//ul[@id='MasterMenu']/li[4]")
	public static WebElement  masterAuthWarehouse1;

	@FindBy (xpath="//th[@id='id_transaction_entry_detail_table_control_heading_1']/div[1]")
	public static WebElement headerWarehouse;

	@FindBy(xpath="//div[@id='GenerateNewMasterModel']//button[2]")
	public static WebElement account_saveBtn;






	// Design Workflow

	@FindBy(xpath="//i[@id='icon_5']")
	public static WebElement purchaseVouchersExpandBtn;

	@FindBy(xpath="//ul[@id='Submenu_5']/li[3]/span")
	public static WebElement purchasesVouchersVATDragAndDrop; 

	@FindBy(xpath="//div[contains(text(),'NDT57')]")
	public static WebElement draggedPurchasesVouchersVAT; 

	@FindBy(xpath="//div[contains(text(),'Material Requisition')]")
	public static WebElement draggedMaterialRequisition ; 

	@FindBy(xpath="//div[contains(text(),'PurOrd')]")
	public static WebElement draggedPurchasesOrders; 

	@FindBy(xpath="//*[@id='WorkflowCheckboxDiv']/div/div/label")
	public static List<WebElement> definitionTabChkBoxNameList;

	@FindBy(xpath="//*[@id='LeftSideChkbox']/div/label/input/following-sibling::span")
	public static List<WebElement> definitionTabChkBoxList;

	@FindBy(xpath="//a[@id='74']//span[contains(text(),'Design Workflow')]")
	public static WebElement designWorkflowbtn; 

	@FindBy(xpath="//div[@class='navText']//span[contains(text(),'Design Workflow')]")
	public static WebElement designWorkflowLabel; 

	@FindBy(xpath="//a[@id='btnSave']/i")
	public static WebElement dwSaveBtn; 

	@FindBy(xpath="//i[@class='icon icon-clear icon-font6']")
	public static WebElement clearBtn; 

	@FindBy(xpath="//div[@id='workflowButtons']/ul/li[3]/a/i")
	public static WebElement dwDeleteBtn; 

	@FindBy(xpath="//*[@id='btnCancel']/i")
	public static WebElement dwCloseBtn;

	@FindBy(xpath="//input[@id='optWorkflow']")
	public static WebElement workflowNameTxt; 



	@FindBy(xpath="//*[@id='optWorkflow_input_settings']/span")
	public static WebElement workFlowNameSettingsBtn; 

	@FindBy(xpath="//span[contains(text(),'Request For Quote')]")
	public static WebElement requestForQuoteDragAndDrop;

	@FindBy(xpath="//span[@class='DragLi ui-draggable ui-draggable-handle'][contains(text(),'Purchases Quotations')]")
	public static WebElement purchasesQuotationsDragAndDrop; 

	@FindBy(xpath="//div[@class='flowchart-operator-title ui-draggable-handle']")  //As it Is First One Selected Drag On FlowChart Is Same Id Property For Every Element
	public static WebElement FirstDraggedElement; 

	@FindBy(xpath="(//div[@class='flowchart-operator-title ui-draggable-handle'])[2]")  //As it Is First One Selected Drag On FlowChart Is Same Id Property For Every Element
	public static WebElement SecondDraggedElement; 




	@FindBy(xpath="//div[@id='idNode_0']")
	public static WebElement draggedNode1;

	@FindBy(xpath="//div[@id='idNode_1']")
	public static WebElement draggedNode2; 

	@FindBy(xpath="//div[contains(text(),'ReqQuo')]")
	public static WebElement draggedPurchasesQuotations; 

	@FindBy(xpath="/html[1]/body[1]/ul[1]/li[1]")
	public static WebElement draggedlinkBtn; 

	@FindBy(xpath="/html[1]/body[1]/ul[1]/li[2]")
	public static WebElement draggeddeleteBtn; 


	@FindBy(xpath="//*[@class='flowchart-link'][@id='line0']")
	//@FindBy(xpath="//*[@id='line1']")
	public static WebElement link0TO1; 

	@FindBy(xpath="//*[@class='flowchart-link'][@id='line1']")
	public static WebElement link1TO2; 

	@FindBy(xpath="//*[@class='flowchart-link'][@id='line2']")
	public static WebElement link2TO3; 

	@FindBy(xpath="//*[@class='flowchart-link'][@id='line3']")
	public static WebElement link3TO4; 

	@FindBy(xpath="//*[@class='flowchart-link'][@id='line4']")
	public static WebElement link4TO5; 

	@FindBy(xpath="//*[@class='flowchart-link'][@id='line5']")
	public static WebElement link5TO6; 

	@FindBy(xpath="//*[@class='flowchart-link'][@id='line6']")
	public static WebElement link6TO7; 

	@FindBy(xpath="//*[@class='flowchart-link'][@id='line7']")
	public static WebElement link7TO8; 

	@FindBy(xpath="//*[@class='flowchart-link'][@id='line8']")
	public static WebElement link8TO9; 




	@FindBy(xpath="//*[@id='idNode_0']/div[2]/div[1]/div/div/div[2]")
	public static WebElement linkNode0Left; 

	@FindBy(xpath="//*[@id='idNode_1']/div[2]/div[1]/div/div/div[2]")
	public static WebElement linkNode1Left; 

	@FindBy(xpath="//*[@id='idNode_2']/div[2]/div[1]/div/div/div[2]")
	public static WebElement linkNode2Left;

	@FindBy(xpath="//*[@id='idNode_3']/div[2]/div[1]/div/div/div[2]")
	public static WebElement linkNode3Left; 

	@FindBy(xpath="//*[@id='idNode_4']/div[2]/div[1]/div/div/div[2]")
	public static WebElement linkNode4Left; 

	@FindBy(xpath="//*[@id='idNode_5']/div[2]/div[1]/div/div/div[2]")
	public static WebElement linkNode5Left;

	@FindBy(xpath="//*[@id='idNode_6']/div[2]/div[1]/div/div/div[2]")
	public static WebElement linkNode6Left; 

	@FindBy(xpath="//*[@id='idNode_7']/div[2]/div[1]/div/div/div[2]")
	public static WebElement linkNode7Left; 

	@FindBy(xpath="//*[@id='idNode_8']/div[2]/div[1]/div/div/div[2]")
	public static WebElement linkNode8Left;





	@FindBy(xpath="//*[@id='idNode_0']/div[2]/div[2]/div/div/div[2]")
	public static WebElement linkNode0Right; 

	@FindBy(xpath="//*[@id='idNode_1']/div[2]/div[2]/div/div/div[2]")
	public static WebElement linkNode1Right; 

	@FindBy(xpath="//*[@id='idNode_2']/div[2]/div[2]/div/div/div[2]")
	public static WebElement linkNode2Right; 

	@FindBy(xpath="//*[@id='idNode_3']/div[2]/div[2]/div/div/div[2]")
	public static WebElement linkNode3Right; 

	@FindBy(xpath="//*[@id='idNode_4']/div[2]/div[2]/div/div/div[2]")
	public static WebElement linkNode4Right; 

	@FindBy(xpath="//*[@id='idNode_5']/div[2]/div[2]/div/div/div[2]")
	public static WebElement linkNode5Right; 

	@FindBy(xpath="//*[@id='idNode_6']/div[2]/div[2]/div/div/div[2]")
	public static WebElement linkNode6Right; 

	@FindBy(xpath="//*[@id='idNode_7']/div[2]/div[2]/div/div/div[2]")
	public static WebElement linkNode7Right; 

	@FindBy(xpath="//*[@id='idNode_8']/div[2]/div[2]/div/div/div[2]")
	public static WebElement linkNode8Right; 




	@FindBy(xpath="//div[@id='settingIcon']//span[@class='iconbar']")
	public static WebElement rightSideExpandBtn; 

	@FindBy(xpath="//a[contains(text(),'Definition')]")
	public static WebElement DefinitionTab; 

	@FindBy(xpath="//select[@id='DDLLinkValue']")
	public static WebElement DefinitionLinkValueDropdown; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[1]//li[1]//input[1]")
	public static WebElement dateChkBox; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']/div[2]/label/span")
	public static WebElement customerAccountChkbox; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[4]//li[1]//input[1]")
	public static WebElement salesOrderandInvoicewareHouseChkbox; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[3]//li[1]//input[1]")
	public static WebElement autosalesOrderandInvoicewareHouseChkbox; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[3]//li[1]//input[1]")
	public static WebElement narrationChkbox; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[4]//li[1]//input[1]")
	public static WebElement salesOrderandInvoicenarrationChkbox;

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[4]//li[1]//input[1]")
	public static WebElement itemchkbox; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']/div[4]/label/span")
	public static WebElement materialRequiitionAndPquotationsitemchkbox; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[5]//li[1]//input[1]")
	public static WebElement salesOrderandInvoiceitemchkbox;

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[5]//li[1]//input[1]")
	public static WebElement purchaseOrderandPvoucheritemchkbox;

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[6]//li[1]//input[1]")
	public static WebElement purchaseVoucherAndPReturnsitemchkbox;

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[7]//li[1]//input[1]")
	public static WebElement purchaseVoucherNAndMRNitemchkbox;

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[6]//li[1]//input[1]")
	public static WebElement mrnAndPurchaseReturnsitemchkbox;

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[10]//li[1]//input[1]")
	public static WebElement salesInvoicenAndDeliveryNotesitemchkbox;

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[11]//li[1]//input[1]")
	public static WebElement salesInvoiceAndSalesRetutrnsitemchkbox;

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[6]//li[1]//input[1]")
	public static WebElement stockTransferAndSalesInvoiceNitemchkbox;

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[5]//li[1]//input[1]")
	public static WebElement unitsChkbox; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[6]//li[1]//input[1]")
	public static WebElement salesOrderandInvoiceunitsChkbox;

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[6]//li[1]//input[1]")
	public static WebElement quantityChkbox; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[7]//li[1]//input[1]")
	public static WebElement salesOrderandInvoicequantityChkbox; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[7]//li[1]//input[1]")
	public static WebElement voucherQtyhkbox; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[8]//li[1]//input[1]")
	public static WebElement salesOrderandInvoicevoucherQtyhkbox;

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[8]//li[1]//input[1]")
	public static WebElement reserveChkbox; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[9]//li[1]//input[1]")
	public static WebElement salesOrderandInvoicereserveChkbox;

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[9]//li[1]//input[1]")
	public static WebElement rateChkbox; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[10]//li[1]//input[1]")
	public static WebElement salesOrderandInvoicerateChkbox; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[10]//li[1]//input[1]")
	public static WebElement grossChkbox; 

	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[11]//li[1]//input[1]")
	public static WebElement salesOrderandInvoicegrossChkbox; 

	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[1]//li[1]//input[1]")
	public static WebElement voucherGrossChkbox; 

	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[2]//li[1]//input[1]")
	public static WebElement netChkbox; 

	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[3]//li[1]//input[1]")
	public static WebElement voucherNetChkbox; 

	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[4]//li[1]//input[1]")
	public static WebElement userChkbox; 

	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[5]//li[1]//input[1]")
	public static WebElement roleChkbox; 

	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[6]//li[1]//input[1]")
	public static WebElement user1Chkbox; 

	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[7]//li[1]//input[1]")
	public static WebElement user2Chkbox; 

	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[8]//li[1]//input[1]")
	public static WebElement user3Chkbox; 

	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[9]//li[1]//input[1]")
	public static WebElement user4Chkbox; 

	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[10]//li[1]//input[1]")
	public static WebElement user5Chkbox; 

	@FindBy(xpath="//div[@id='WorkflowSettingDialog']/div[2]/div/div[3]/div/input[1]")
	public static WebElement OkBtn; 

	@FindBy(xpath="//div[@id='WorkflowSettingDialog']/div[2]/div/div[3]/div/input[2]")
	public static WebElement CancelBtn; 



	// Standard Rate

	@FindBy(xpath="//div[@id='id_left_mainmenu']/ul/li[1]/ul/li[6]/ul/li/a/span")
	public static List<WebElement> utilitiesSubMenusList;

	@FindBy (xpath="//*[@id='navigation_menu']/li[1]/ul/li[5]/ul/li")
	public static List<WebElement> utilitesList;

	@FindBy(xpath="//div[@id='id_left_mainmenu']/ul/li[1]/ul/li[6]")
	public static WebElement  utilities;

	@FindBy (xpath="//span[contains(text(),'Standard Rate')]")
	public static WebElement standardRate;

	@FindBy(xpath="//div[@id='navbarSupportedContent2']/ul/li[3]/a/i")
	public static WebElement closeBtnInStandardRate;

	@FindBy(xpath="//tbody[@id='OptProductId_table_data_body']/tr/td")
	public static List<WebElement> standardRateItemComboList;

	@FindBy(xpath="//input[@id='OptProductId']")
	public static WebElement standardRateItemCombo;

	@FindBy(xpath="//input[@id='txtEffectiveDate']")
	public static WebElement standardRateEffectiveDateTxt;

	@FindBy(xpath="//input[@id='txtrate']")
	public static WebElement standardRateRateTxt;

	@FindBy(xpath="//li[@id='id_EdtAdd']/a/i")
	public static WebElement standardRateAddBtn;

	@FindBy(xpath="//td[@id='id_StandardRate_col_1-2']")
	public static WebElement standardRateTableRow1EffectiveDate;

	@FindBy(xpath="//tbody[@id='id_StandardRate_body']/tr/td[4]")
	public static WebElement standardRateTableRow1Rate;

	@FindBy(xpath="//td[@id='id_StandardRate_col_2-2']")
	public static WebElement standardRateTableRow2EffectiveDate;

	@FindBy(xpath="//td[@id='id_StandardRate_col_2-3']")
	public static WebElement standardRateTableRow2Rate;

	@FindBy(xpath="//*[@id='btnEdit_Id_1']")
	public static WebElement standardRateEditRow1;

	@FindBy(xpath="//li[@id='id_Edt']/a/i")
	public static WebElement standardRateEditBtn;

	@FindBy(xpath="//*[@id='btnDelete_Id_1']")
	public static WebElement standardRateDeleteRow1;



	// Dashboard

	@FindBy(xpath="//ul[@class='dropdown-menu Select_dashboard']")
	public static WebElement selectDashboardComboBox ;

	@FindBy(xpath="//li[@id='Dashboard_AddDash']/a/i")
	public static WebElement newDashBoard;

	@FindBy(xpath="//li[@id='Dash_Edit']/a/i")
	public static WebElement editDashBoard;

	@FindBy(xpath="//li[@id='Dash_Delete']/a/i")
	public static WebElement deleteDashBoard;

	@FindBy(xpath="//input[@id='txtDashboardName']")
	public static WebElement dashBoardName;

	@FindBy(xpath="//*[@id='txtDashboardName_input_settings']/span")
	public static WebElement dashBoardSettingInCombobox;

	//Customize Display Columns Screen
	@FindBy(xpath="//*[@id='txtDashboardName_customize_popup_container']")
	public static WebElement dashBoardCustomizeDisplaysNameColumn;

	//Standard Fields Button
	@FindBy(xpath="//*[@id='txtDashboardName_customize_popup_footer']/div/div/input[1]")
	public static WebElement dashBoardPopCustomizeStandardFieldsBtn;

	@FindBy(xpath="//*[@id='txtDashboardName_customize_popup_standardfields_list']")
	public static WebElement dashBoardPopCustomizeStandardColumnAttributeFieldName;

	@FindBy(xpath="//*[@id='txtDashboardName_customize_popup_standardfields_header']")
	public static WebElement dashBoardPopCustomizeStandardColumnAttributeHeader;

	@FindBy(xpath="//*[@id='txtDashboardName_customize_popup_standardfields_alignment']")
	public static WebElement dashBoardPopCustomizeStandardColumnAttributeAlignment;

	@FindBy(xpath="//*[@id='txtDashboardName_customize_popup_standardfields_width']")
	public static WebElement dashBoardPopCustomizeStandardColumnAttributeWidth;

	@FindBy(xpath="//*[@id='txtDashboardName_customize_popup_footer']/div/div/input[3]")
	public static WebElement dashBoardPopCustomizeDisplayStandardColumnAttributeOkBtn;

	@FindBy(xpath="//*[@id='txtDashboardName_customize_popup_footer']/div/div/input[4]")
	public static WebElement dashBoardPopCustomizeDisplayStandardColumnAttributeCancelBtn;

	//Delete Column Button
	@FindBy(xpath="//*[@id='txtDashboardName_customize_popup_footer']/div/div/input[2]")
	public static WebElement dashBoardPopCustomizeDisplayColumnsDeleteColoumn;

	//Ok button
	@FindBy(xpath="//*[@id='txtDashboardName_customize_popup_footer']/div/div/input[3]")
	public static WebElement dashBoardPopCustomizeDisplayColumnsOkBtn;

	//Cancel button
	@FindBy(xpath="//*[@id='txtDashboardName_customize_popup_footer']/div/div/input[4]")
	public static WebElement dashBoardPopCustomizeDisplayColumnsCancelBtn;

	@FindBy(xpath="//*[@id='txtAreaDescription']")//*[@id="txtAreaDescription"]
	public static WebElement DashBoardDescription;				

	@FindBy(xpath="//input[@id='txtHeader']")
	public static WebElement dashBoardHeaderTextBox;

	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[2]/div[2]/div[2]/div[2]/div[1]/div[3]/div[1]/div[2]/div[1]/span[1]/button[1]")
	public static WebElement dashBoardHeaderSettingsBtn;

	@FindBy(xpath="//*[@id='fontFamily_ctrHeader']")
	public static WebElement dashBoardHeader_FontFamilyDropdown;

	@FindBy(xpath="//*[@id='fontStyle_ctrHeader']")
	public static WebElement dashBoardHeader_FontStyleDropdown;

	@FindBy(xpath="//*[@id='fontWeight_ctrHeader']")
	public static WebElement dashBoardHeader_FontWeightDropdown;

	@FindBy(xpath="//*[@id='fontSizes_ctrHeader']")
	public static WebElement dashBoardHeader_FontSizeDropdown;

	@FindBy(xpath="//*[@id='fontForeColor_ctrHeader']")
	public static WebElement dashBoardHeader_FontForeColorDropdown;

	@FindBy(xpath="//*[@id='fontBackColor_ctrHeader']")
	public static WebElement dashBoardHeader_FontBackColorDropdown;

	@FindBy(xpath="//*[@id='chkBaseline_ctrHeader']")
	public static WebElement dashBoardHeader_BaselineCheckbox;

	@FindBy(xpath="//*[@id='chkOverLine_ctrHeader']")
	public static WebElement dashBoardHeader_OverLineCheckbox;

	@FindBy(xpath="//*[@id='chkStrikeThrough_ctrHeader']")
	public static WebElement dashBoardHeader_StrikeThroughCheckbox;

	@FindBy(xpath="//*[@id='chkUnderline_ctrHeader']")
	public static WebElement dashBoardHeader_UnderlineCheckbox;

	@FindBy(xpath="//*[@id='previewfont_ctrHeader']")
	public static WebElement dashBoardHeader_FontPreviewDescription;

	@FindBy(xpath="//*[@id='chkEnableFont_ctrHeader']")
	public static WebElement dashBoardHeader_EnableFontCheckbox;

	@FindBy(xpath="//*[@id='btnSaveFonts1']")
	public static WebElement dashBoardHeaderSaveBtn;

	@FindBy(xpath="//*[@id='btnCloseFonts1']")
	public static WebElement dashBoardHeaderCLoseBtn;


	@FindBy(xpath="//*[@id='txtBody']")
	public static WebElement dashBoardBody;

	@FindBy(xpath="//*[@id='iBody']")
	public static WebElement DashBoardBodySettingBtn;

	@FindBy(xpath="//*[@id='fontFamily_ctrBody']")
	public static WebElement dashBoardBody_FontFamilyDropdown;

	@FindBy(xpath="//*[@id='fontStyle_ctrBody']")
	public static WebElement dashBoardBody_FontStyleDropdown;

	@FindBy(xpath="//*[@id='fontWeight_ctrBody']")
	public static WebElement dashBoardBody_FontWeightDropdown;

	@FindBy(xpath="//*[@id='fontSizes_ctrBody']")
	public static WebElement dashBoardBody_FontSizeDropdown;

	@FindBy(xpath="//*[@id='fontForeColor_ctrBody']")
	public static WebElement dashBoardBody_FontForeColorDropdown;

	@FindBy(xpath="//*[@id='fontBackColor_ctrBody']")
	public static WebElement dashBoardBody_FontBackColorDropdown;

	@FindBy(xpath="//*[@id='chkBaseline_ctrBody']")
	public static WebElement dashBoardBody_BaselineCheckbox;

	@FindBy(xpath="//*[@id='chkOverLine_ctrBody']")
	public static WebElement dashBoardBody_OverLineCheckbox;

	@FindBy(xpath="//*[@id='chkStrikeThrough_ctrBody']")
	public static WebElement dashBoardBody_StrikeThroughCheckbox;

	@FindBy(xpath="//*[@id='chkUnderline_ctrBody']")
	public static WebElement dashBoardBody_UnderlineCheckbox;

	@FindBy(xpath="//*[@id='previewfont_ctrBody']")
	public static WebElement dashBoardBody_FontPreviewDescription;

	@FindBy(xpath="//*[@id='chkEnableFont_ctrBody']")
	public static WebElement dashBoardBody_EnableFontCB;

	@FindBy(xpath="//*[@id='btnSaveFonts2']")
	public static WebElement dashBoardBodySaveBtn;

	@FindBy(xpath="//span[@id='btnCloseFonts2']")
	public static WebElement dashBoardBodyCLoseBtn;

	@FindBy(xpath="//*[@id='txtFooter']")
	public static WebElement dashBoardFooter;

	@FindBy(xpath="//*[@id='iFooter']")
	public static WebElement DashBoardFooterSettingBtn;

	@FindBy(xpath="//*[@id='fontFamily_ctrFooter']")
	public static WebElement dashBoardFooter_FontFamilyDropdown;

	@FindBy(xpath="//*[@id='fontStyle_ctrFooter']")
	public static WebElement dashBoardFooter_FontStyleDropdown;

	@FindBy(xpath="//*[@id='fontWeight_ctrFooter']")
	public static WebElement dashBoardFooter_FontWeightDropdown;

	@FindBy(xpath="//*[@id='fontSizes_ctrFooter']")
	public static WebElement dashBoardFooter_FontSizeDropdown;

	@FindBy(xpath="//*[@id='fontForeColor_ctrFooter']")
	public static WebElement dashBoardFooter_FontForeColorDropdown;

	@FindBy(xpath="//*[@id='fontBackColor_ctrFooter']")
	public static WebElement dashBoardFooter_FontBackColorDD;

	@FindBy(xpath="//*[@id='chkBaseline_ctrFooter']")
	public static WebElement dashBoardFooter_BaselineCheckbox;

	@FindBy(xpath="//*[@id='chkOverLine_ctrFooter']")
	public static WebElement dashBoardFooter_OverLineCheckbox;

	@FindBy(xpath="//*[@id='chkStrikeThrough_ctrFooter']")
	public static WebElement dashBoardFooter_StrikeThroughCheckbox;

	@FindBy(xpath="//*[@id='chkUnderline_ctrFooter']")
	public static WebElement dashBoardFooter_UnderlineCheckbox;

	@FindBy(xpath="//*[@id='previewfont_ctrFooter']")
	public static WebElement dashBoardFooter_FontPreviewDescription;

	@FindBy(xpath="//*[@id='chkEnableFont_ctrFooter']")
	public static WebElement dashBoardFooter_EnableFontCheckbox;

	@FindBy(xpath="//*[@id='btnSaveFonts3']")
	public static WebElement dashBoardFooterSaveBtn;

	@FindBy(xpath="//*[@id='btnCloseFonts3']")
	public static WebElement dashBoardFooterCLoseBtn;

	@FindBy(xpath="//div[@id='First-Body-Row']/div/div[2]/div[1]/div/label/span")
	public static WebElement dashBoardSetAsDefaultCheckbox;

	@FindBy(xpath="//div[@id='First-Body-Row']/div/div[2]/div[2]/div/label/span")
	public static WebElement dashBoardActiveCheckbox;

	//Users
	@FindBy(xpath="//*[@id='rdoUsers']")
	public static WebElement dashBoardUsersRadioBtn;

	@FindBy(xpath="//*[@id='widthDiv']")
	public static WebElement dashBoardUsersAvaliableForSharing;

	@FindBy(xpath="//*[@id='NinthFdiv']")
	public static WebElement dashBoardSharedUsers;

	//Roles
	@FindBy(xpath="//*[@id='rdoRoles']")
	public static WebElement dashBoardRolesRadioBtn;				 

	@FindBy(xpath="//*[@id='widthDiv']")
	public static WebElement dashBoarRolesAvaliableForSharing;

	@FindBy(xpath="//*[@id='NinthFdiv']")
	public static WebElement dashBoardSharedRoles;

	@FindBy(xpath="//*[@id='sNinthDiv']/i[2]")
	public static WebElement dashBoardRightSideArrowIcon;

	@FindBy(xpath="//*[@id='sNinthDiv']/i[1]")
	public static WebElement dashBoardLeftSideArrowIcon;

	@FindBy(xpath="//*[@id='Create_Dash_Save']/a/i")
	public static WebElement dashBoardSaveBtn;

	@FindBy(xpath="//*[@id='Create_Dash_Close']/a/i")
	public static WebElement dashBoardCancelBtn;

	@FindBy(xpath="//*[@id='lblAvailableUsers']")
	public static WebElement  UsersavailableforsharingLabel;

	@FindBy(xpath="//*[@id='lblSharedUsers']")
	public static WebElement SharedUsersLabel;

	@FindBy(xpath="//*[@id='lblAvailableRoles']")
	public static WebElement  Rolesavailableforsharinglabel;

	@FindBy(xpath="//*[@id='lblSharedRoles']")
	public static WebElement SharedRolesLabel;

	@FindBy(xpath="//div[@class='theme_color font-6']")
	public static WebElement dashboardValidationAndConfirmationMessage;

	//Dashboard Customization 
	@FindBy(xpath="//li[@id='Dashboard_Dash_Config']/a/i")
	public static WebElement dashBoardCustomizeBtn;

	@FindBy(xpath="//*[@id='ErrorMessage']")
	public static WebElement validationPleaseSelectanyDashboardMessage;

	@FindBy(xpath="//*[@id='dashIcons']/div/ul/li[1]/h4")
	public static WebElement selectdashBoardOptionText;

	@FindBy(xpath="//*[@id='Dashboard_ddlList']/li[2]/a/label")
	public static WebElement dashBoardSecond;

	@FindBy(xpath="//*[@id='Dashboard_ddlList']/li[3]/a/label")
	public static WebElement dashBoardThird;


	@FindBy(xpath="//*[@id='Dashboardconfirm']/div/div/div[1]")
	public static WebElement dashBoardConfirmationMessage;

	@FindBy(xpath="//div[@id='Dashboardconfirm']//button[@id='deletebtn']")
	public static WebElement dashBoardConfirmationMessageOkBtn;

	@FindBy(xpath="//div[@id='Dashboardconfirm']//button[@id='btnCancel']")
	public static WebElement dashBoardConfirmationMessageCancelBtn;

	@FindBy(xpath="//*[@id='ErrorMessage']")
	public static WebElement validationDashboardMessageAtBotton;


	//-----------------Create DashLet Elements

	//@FindBy(xpath="//div[@id='GraphList']/ul/li[1]/div/input")
	@FindBy(xpath="//*[@id='searchBox']")
	public static WebElement  dashBoardCustomizeSearch;

	@FindBy(xpath="//*[@id='GraphOption']")
	public static WebElement  dashBoardCustomizeGraph;

	//@FindBy(xpath="//a[@class='theme_icon-color']")
	@FindBy(xpath="//*[@id='Newgraph']/a")
	public static WebElement  dashBoardCustomizeGraphNewGraph; 

	@FindBy(xpath="//*[@id='id_Dashlet1']/a")
	public static WebElement  dashBoardCustomizeNewGraphCreated; 						

	@FindBy(xpath="//i[@class='icon-edit PANEl_EDiTE pull-right theme_icon-color']")
	public static WebElement  dashBoardCustomizeEditGraphCreated;

	//@FindBy(xpath="//i[@class='icon-delete PANEl_DELetE pull-right theme_icon-color']")
	@FindBy(xpath="(//i[@class='icon-delete PANEl_DELetE pull-right theme_icon-color'])[1]")
	public static WebElement  dashBoardCustomizeDeleteGraphCreated;

	@FindBy(xpath="//label[@id='lblWarnMessage']")
	public static WebElement  dashBoardCustomizeDeleteTheDashletMessage;

	@FindBy(xpath="//input[@id='btnDelete_Dashboard']")
	public static WebElement  dashBoardCustomizeDeleteDasletYesBtn;

	@FindBy(xpath="//button[contains(text(),'No')]")
	public static WebElement  dashBoardCustomizeDeleteDashletNoBtn;

	@FindBy(xpath="//li[@id='id_Dashlet0']//span[@class='fa fa-circle iconCircle theme_icon-color']")
	public static WebElement  dashBoardCreatedNewGraph;

	@FindBy(xpath="//*[@id='id_Dashlet1']/a")
	public static WebElement  dashBoardGraphNewFromDrag;	

	@FindBy(xpath="//*[@id='CustumizeDiv']")
	public static WebElement  dashBoardGraphNewToDrag;

	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[2]/div[2]/div[2]/div[4]")
	public static WebElement  dashBoardGraphDashlet;

	//---------------Graph customize Definition

	@FindBy(xpath="//*[@id='DashletTabs']/li[1]")
	public static WebElement  dashBoardCustomizeGraphDefTab;

	@FindBy(xpath="//input[@id='DashletName']")
	public static WebElement  dashBoardCustomizeGraphDefName;

	@FindBy(xpath="//*[@id='Type']")
	public static WebElement  dashBoardCustomizeGraphDefShowAs;

	@FindBy(xpath="//*[@id='IsDisplaySummary']")
	public static WebElement  dashBoardCustomizeGraphDefDisplaySummaryCheckBox;

	@FindBy(xpath="//*[@id='Include']")
	public static WebElement  dashBoardCustomizeGraphDefIncludeDropdown;

	@FindBy(xpath="//*[@id='IncludeValue']")
	public static WebElement  dashBoardCustomizeGraphDefIncludeValueTxtbox;

	@FindBy(xpath="//*[@id='IncludeType']")
	public static WebElement  dashBoardCustomizeGraphDefIncludeTypeDropdown;

	@FindBy(xpath="//*[@id='SeriesName1']")
	public static WebElement  dashBoardCustomizeGraphDefSeriesName;

	@FindBy(xpath="//*[@id='ddlDataSet1']")
	public static WebElement  dashBoardCustomizeGraphDefDataSetDropdown;

	@FindBy(xpath="//*[@id='yAxis_1']/div[1]/div/span")
	public static WebElement  dashBoardCustomizeGraphAddBtnImg;

	@FindBy(xpath="//*[@id='default_Graph_1']/div[1]/div[2]/div[2]/span")
	public static WebElement  dashBoardCustomizeGraphDeleteBtnImg;

	@FindBy(xpath="//*[@id='ddlValue1']")
	public static WebElement  dashBoardCustomizeGraphDefValueDropdown;

	@FindBy(xpath="//*[@id='ddlPeriod1']")
	public static WebElement  dashBoardCustomizeGraphDefPeriodDropdown;	



	// filter property 
	@FindBy(xpath="//*[@id='default_Graph_1']/div[2]/div[2]/span")
	public static WebElement  dashBoardCustomizeGraphDefFliterBtn;

	@FindBy(xpath="//*[@id='36_1_AdvancegraphFilter']/table/tbody/tr/td[1]/select")
	public static WebElement graphFilterConjuctionDropdown ;

	@FindBy(xpath="//*[@id='36_1_AdvancegraphFilter']/table/tbody/tr/td[2]/input")
	public static WebElement graphFilterSelectField ;

	@FindBy(xpath="//*[@id='36_1_AdvancegraphFilter']/table/tbody/tr/td[3]/select")
	public static WebElement graphFilterSelectOperatorDropdown ;

	@FindBy(xpath="//*[@id='36_1_AdvancegraphFilter']/table/tbody/tr/td[4]/select")
	public static WebElement graphFilterCompareWithDropdown ;

	@FindBy(xpath="//*[@id='36_1_AdvancegraphFilter']/table/tbody/tr/td[5]/input")
	public static WebElement graphFilterValurText ;

	@FindBy(xpath="//*[@id='36_1_AdvancegraphFilter']/table/tbody/tr/td[7]/span")
	public static WebElement graphFilterTOAddOneMoreRowBtn ;

	@FindBy(xpath="//*[@id='36_1_AdvancegraphFilter']/table/tbody/tr[1]/td[6]/span")
	public static WebElement graphFilterTODeleteFirstRowBtn ;

	@FindBy(xpath="//*[@id='36_1_AdvancegraphFilter']/table/tbody/tr[2]/td[6]/span")
	public static WebElement graphFilterTODeleteSecondtRowBtn ;	

	@FindBy(xpath="//*[@id='ddlSummary1']")
	public static WebElement  dashBoardCustomizeGraphDefSummary;

	@FindBy(xpath="//*[@id='XAxis']")
	public static WebElement  dashBoardCustomizeGraphDefXAxis;

	@FindBy(xpath="//*[@id='SaveDashlet_Icon']")
	public static WebElement  dashBoardCustomizeGraphDefSave;

	@FindBy(xpath="//li[@id='SaveDashlet_Icon']//span[text()='Save']")
	public static WebElement  dashBoardCustomizeReportSave;

	@FindBy(xpath="//*[@id='Newdashlet_Icons']/ul/li[2]")
	public static WebElement  dashBoardCustomizeGraphDefCancel;


	//------------------------Graph customize Properties

	@FindBy(xpath="//*[@id='DashletTabs']/li[2]")
	public static WebElement  dashBoardCustomizeGraphProperties;

	@FindBy(xpath="//*[@id='ddlGraphTypes']")
	public static WebElement  dashBoardCustomizeGraphPropertiesGraphTypeDropdown;

	@FindBy(xpath="//*[@id='Themes']")
	public static WebElement  dashBoardCustomizeGraphPropertiesThemeDropdown;

	@FindBy(xpath="//*[@id='LegendPosition']")
	public static WebElement  dashBoardCustomizeGraphPropertiesLegendPositionDropdown;

	@FindBy(xpath="//*[@id='IsZooming']")
	public static WebElement  dashBoardCustomizeGraphPropertiesZoomCheckbox;


	//----------------------Graph customize Maximized view

	@FindBy(xpath="//*[@id='DashletTabs']/li[3]")
	public static WebElement  dashBoardCustomizeGraphMaximizedView;

	@FindBy(xpath="//*[@id='DisplayType']")
	public static WebElement  dashBoardCustomizeGraphMaximizedViewDisplayCombobox;

	@FindBy(xpath="//*[@id='DisplayValue']")
	public static WebElement  dashBoardCustomizeGraphMaximizedViewDiplayitemSelectedCombobox;

	@FindBy(xpath="//*[@id='maxViewDiv']/div/div[1]/div/span[1]")
	public static WebElement  dashBoardCustomizeGraphMaximizedViewCustomizeIcon;

	@FindBy(xpath="//*[@id='maxViewDiv']/div/div[1]/div/span[2]")
	public static WebElement  dashBoardCustomizeGraphMaximizedViewCloseIcon;


	//Report*****************************************							

	@FindBy(xpath="//*[@id='ReportOption']")
	public static WebElement  dashBoardCustomizeReport;

	@FindBy(xpath="//a[contains(text(),'New Report')]")
	public static WebElement  dashBoardCustomizeReportNewReport;

	//Report customize Report

	@FindBy(xpath="//*[@id='DashletTabs']/li[1]")
	public static WebElement  dashBoardCustomizeReport_ReportDef;

	@FindBy(xpath="//select[@id='ddlReports']")
	public static WebElement  dashBoardCustomizeReportReportName;

	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[2]/div[2]/div[2]/div[5]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/input[3]")
	public static WebElement  dashBoardCustomizeReportCaptionText;

	@FindBy(xpath="//*[@id='chkReport']")
	public static WebElement  dashBoardCustomizeReportReportRadioBtn;

	@FindBy(xpath="//*[@id='chkGraph']")
	public static WebElement  dashBoardCustomizeReportGraphRadioBtn;

	@FindBy(xpath="//input[@id='cmbUserTypeMaster']")
	public static WebElement  inputMasterCombobox;

	@FindBy(xpath="//input[@id='txtsrch-term']")
	public static WebElement  searchMasterCombobox;

	@FindBy(xpath="//div[@id='dvReportInputs']")
	public static WebElement  inputReportsSettings;

	@FindBy(xpath="//td[contains(@class,'tdName padding0')]")
	public static WebElement  nameMasterSearchCombobox;

	@FindBy(xpath="//*[@id='LandingGridBody']/tr[3]/td[11]")
	public static WebElement  nameMasterInputCombobox;

	@FindBy(xpath="//*[@id='id_Dashlet2']/a/label")
	public static WebElement  newReportNameDisplayInCreateDashlet;

	@FindBy(xpath="//*[@id='SaveDashlet_Icon']")
	public static WebElement  receiptsEditCustomizeReportSettingsSave;

	@FindBy(xpath="//*[@id='Newdashlet_Icons']/ul/li[2]")
	public static WebElement  receiptsEditCustomizeReportSettingsCancel;

	@FindBy (xpath="//select[@id='DateOptions_dashletReport_551']")
	public static WebElement dashBoardCustomizeReportDateOptionDrpdwn;


	@FindBy (xpath="//input[@id='selectAllMasters_dashletReport_551']")
	public static WebElement dashboardCustomizeReportSelectAllItemsChkbox;



	//InfoPanel****************************************************************				

	@FindBy(xpath="//*[@id='InfopanelOption']")
	public static WebElement  dashBoardCustomizeInfoPanel;

	@FindBy(xpath="//*[@id='NewInfopanel']")
	public static WebElement  dashBoardCustomizeInfopanelNewInfopanel;

	//------------------------InfoPanel customize Def

	@FindBy(xpath="//*[@id='DashletTabs']/li[1]/a/i")
	public static WebElement  dashBoardCustomizeInfoPanelDef;

	@FindBy(xpath="//*[@id='DashletTabs']/li[1]")
	public static WebElement  dashBoardCustomizeInfoPanelLabel;

	@FindBy(xpath="//*[@id='DashletName']")
	public static WebElement  dashBoardCustomizeInfoPanelDefNameText;

	@FindBy(xpath="//*[@id='Type']")
	public static WebElement  dashBoardCustomizeInfoPanelDefInfoTypeDropdown;

	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[2]/div[2]/div[2]/div[5]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/span[1]")
	public static WebElement  dashBoardCustomizeInfoPaneDefAddIcon;

	@FindBy(xpath="//*[@id='Variable1']")
	public static WebElement  dashBoardCustomizeInfoPanelDefFirstGridVariable;

	@FindBy(xpath="//*[@id='txtStandardVarCaption1']")
	public static WebElement  dashBoardCustomizeInfoPanelDefFirstGridCaption;

	@FindBy(xpath="//*[@id='ddlVariablePeriod1']")
	public static WebElement  dashBoardCustomizeInfoPanelDefFirstGridPeriod;

	@FindBy(xpath="//*[@id='divNewVariable1']/div[1]/div/span")
	public static WebElement  dashBoardCustomizeInfoPanelDefFirstGridDeleteIcon;

	@FindBy(xpath="//*[@id='divNewVariable1']/div[3]/div[2]/div/span")
	public static WebElement  dashBoardCustomizeInfoPanelDefFirstGridFilterIcon;




	//Info type as Define		

	@FindBy(xpath="//*[@id='txtDefineCaption1']")
	public static WebElement infoPanelDef_InfoType_DefineCaption;		

	@FindBy(xpath="//*[@id='ddlDataSet1']")
	public static WebElement infoPanelDef_InfoType_DefineDataSetDropDown;

	@FindBy(xpath="//*[@id='ddlValue1']")
	public static WebElement infoPanelDef_InfoType_DefineValueDropdown;

	@FindBy(xpath="//*[@id='ddlSummary1']")
	public static WebElement infoPanelDef_InfoType_DefineSummaryDropdown;

	@FindBy(xpath="//*[@id='ddlPeriod1']")
	public static WebElement infoPanelDef_InfoType_DefinePeriodDropdown;


	//In Create Dashlets, after saving info panel is added


	@FindBy(xpath="//*[@id='id_Dashlet2']/a/label")
	public static WebElement infoPanel_SelectingFirst;				

	@FindBy(xpath="//*[@id='id_Dashlet2']/i[2]")
	public static WebElement infoPanel_SelectingFirst_EditBtn;	

	@FindBy(xpath="//*[@id='id_Dashlet2']/i[1]")
	public static WebElement infoPanel_SelectingFirst_DeleteBtn;									


	/////Selected Info Type Drop down is Query															

	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[2]/div[2]/div[2]/div[5]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/span[1]")
	public static WebElement infoPanelDef_InfoType_QueryAddBtn;	

	@FindBy(xpath="//div[@id='InfoPanel_QueryDiv1']//span[@class='icon-delete icon-font6']")
	public static WebElement infoPanelDef_InfoType_QueryDelete1Btn;	

	@FindBy(xpath="//input[@id='txtQueryCaption1']")
	public static WebElement infoPanelDef_InfoType_QueryCaption1Textbox;	

	@FindBy(xpath="//select[@id='ddlQueryPeriod1']")
	public static WebElement infoPanelDef_InfoType_QueryPeriod1Dropdown;	

	@FindBy(xpath="//textarea[@id='txtQuery1']")
	public static WebElement infoPanelDef_InfoType_QueryDecrption1TxtBox;	

	@FindBy(xpath="//select[@id='ddlQueryValue1']")
	public static WebElement infoPanelDef_InfoType_QueryValue1Dropdown;	


	////////////////////////// For 2 Row elements	



	@FindBy(xpath="//*[@id='InfoPanel_QueryDiv2']/div[1]/div/span")
	public static WebElement infoPanelDef_InfoType_QueryDelete2Btn;	

	@FindBy(xpath="//*[@id='txtQueryCaption2']")
	public static WebElement infoPanelDef_InfoType_QueryCaption2Textbox;	

	@FindBy(xpath="//*[@id='ddlQueryPeriod2']")
	public static WebElement infoPanelDef_InfoType_QueryPeriod2Dropdown;	

	@FindBy(xpath="//*[@id='txtQuery2']")
	public static WebElement infoPanelDef_InfoType_QueryDecrption2TxtBox;	

	@FindBy(xpath="//*[@id='ddlQueryValue2']")
	public static WebElement infoPanelDef_InfoType_QueryValue2Dropdown;	


	//------------------------InfoPanel customize------- Properties

	@FindBy(xpath="//*[@id='DashletTabs']/li[2]")
	public static WebElement  dashBoardCustomizeInfoPanelProperties;


	@FindBy(xpath="//*[@id='Dashlet_InfoPanelPropeties']/div[2]/div[1]/div/input")
	public static WebElement  dashBoardCustomizeInfoPanelPropertiesPalette;

	@FindBy(xpath="//*[@id='Dashlet_InfoPanelPropeties']/div[2]/div[2]/div/input")
	public static WebElement  dashBoardCustomizeInfoPanelPropertiesBackGround;





	//----------------------InfoPanel customize Maximized view		

	@FindBy(xpath="//*[@id='DashletTabs']/li[3]")
	public static WebElement  dashBoardCustomizeInfoPanelMaximizedView;

	@FindBy(xpath="//*[@id='DisplayType']")
	public static WebElement  dashBoardCustomizeInfoPanelMaximizedViewDisplay;

	@FindBy(xpath="//*[@id='DisplayValue']")
	public static WebElement  dashBoardCustomizeInfoPanelMaximizedViewDiplayitemSelected;

	@FindBy(xpath="//*[@id='maxViewDiv']/div/div[1]/div/span[1]")
	public static WebElement  dashBoardCustomizeInfoPanelMaximizedViewCustomizeIcon;

	@FindBy(xpath="//*[@id='maxViewDiv']/div/div[1]/div/span[2]")
	public static WebElement  dashBoardCustomizeInfoPanelMaximizedViewCloseIcon;

	@FindBy(xpath="//*[@id='SaveDashlet_Icon']")
	public static WebElement  dashBoardCustomizeInfoPanelMaximizedViewSave;

	@FindBy(xpath="//*[@id='Newdashlet_Icons']/ul/li[2]")
	public static WebElement  dashBoardCustomizeInfoPanelMaximizedViewCancel;

	@FindBy(xpath="//label[@class='col-xs-12 col-sm-12 col-md-12 col-lg-12 padding0 Flabel theme_color-inverse font-4']")
	public static WebElement  dashBoardInfoPanel;

	@FindBy(xpath="//div[@id='dashlate3']")
	public static WebElement  dashBoardInfoPanelToDelete;

	@FindBy(xpath="//*[@id='dashlate0']/div[1]/label")
	public static WebElement  dashlet;
	//*[@id="dashlate0"]

	@FindBy(xpath="/html/body/ul[1]/li")
	public static WebElement  dashboardInfoPanelDeleteOption;

	@FindBy(xpath="//*[@id='myModalLabel']")
	public static WebElement  dashboardDeleteOptionMessage;

	@FindBy(xpath="//div[@id='confirm']//button[@id='btnCancel']")
	public static WebElement  dashboardDeleteOptionMessageCancelBtn;

	//@FindBy(xpath="//div[@id='confirm']//button[@id='deletebtn']")
	@FindBy(xpath="//*[@id='confirm']/div[2]/div/div[3]/div/input[1]")
	public static WebElement  dashboardDeleteOptionMessageOkBtn;

	@FindBy(xpath="//*[@id='SaveDashlet_Icon']")
	public static WebElement  dashBoardCustomizeInfoPanelSaveBtn;

	@FindBy(xpath="//*[@id='Newdashlet_Icons']/ul/li[2]")
	public static WebElement  dashBoardCustomizeInfoPanelCancelBtn;

	//------------------------WorkFlow customize Def								
	@FindBy(xpath="//*[@id='DocumentInfoOption']")
	public static WebElement  dashBoardCustomizeInfoWorkFlow;

	@FindBy(xpath="//*[@id='NewWorkflow']")
	public static WebElement  dashBoardCustomizeWorkFlow_NewWorkflow;

	@FindBy(xpath="//*[@id='DashletTabs']/li[1]")
	public static WebElement  dashBoardCustomizeWorkFlowDef;

	@FindBy(xpath="//*[@id='tbWkflowName']")
	public static WebElement  dashBoardCustomizeWorkFlowDefName;

	@FindBy(xpath="//*[@id='ddlWkflow']")
	public static WebElement  dashBoardCustomizeWorkFlowDef_WorkFlowName;

	@FindBy(xpath="//*[@id='SaveDashlet_Icon']")
	public static WebElement  dashBoardCustomizeWorkFlowSave;

	@FindBy(xpath="//*[@id='Newdashlet_Icons']/ul/li[2]")
	public static WebElement  dashBoardCustomizeWorkFlowCancel;																	

	@FindBy(xpath="//*[@id='Dash_Save']")
	public static WebElement  dashBoardCustomizeDashSave;

	@FindBy(xpath="//*[@id='Dash_Close']")
	public static WebElement  dashBoardCustomizeDashCancel;

	@FindBy(xpath="//*[@id='lblErrMessage_Graph']")
	public static WebElement  dashBoardCustomizeMessageAtBottom;

	//After Saving Dashlets in the Dashboards
	@FindBy(xpath="//label[@class='theme_color-inverse']")
	public static WebElement  dashBoardCustomizeDashletGraph;

	@FindBy(xpath="//div[@class='theme_color-inverse']")
	public static WebElement  dashBoardCustomizeDashlNewReport;

	@FindBy(xpath="//span[@class='icon-reject2 theme_color']")
	public static WebElement ErrorMessageCloseBtn;

	@FindBy(xpath="//*[@id='mainHeader_MainLayout']/nav/div/ul/li[6]/a")
	public static WebElement signoutBtn;

	@FindBy(xpath="//ul[@id='Useravailable']/li[text()='UserAllOptionsST']")
	public static WebElement userAllOptionsBtn;

	//@FindBy (xpath="//div[@id='DBdashlate0']/div[1]/label")
	@FindBy (xpath="//div[@id='DBdashlate0']/div[1]")
	public static WebElement createdDashletInHomepage;

	@FindBy(xpath="//div[@id='dashlate0']/div[4]")
	public static WebElement dashlateDrag;

	@FindBy(xpath="//div[@id='dashlate1']/div[4]")
	public static WebElement dashlate1Drag;

	@FindBy(xpath="//*[@id='dashlate1']/div[1]")
	public static WebElement createdDashlet1;

	//DataManagement menu
	@FindBy (xpath="//*[@id='25']")
	public static WebElement dataMangementMenu;

	@FindBy (xpath="//*[@id='26']/span")
	public static WebElement backup;

	@FindBy (xpath="//*[@id='27']/span")
	public static WebElement restore;

	@FindBy (xpath="//div[@id='BackUpControls']/ul/li/div[1]")
	public static WebElement backupBtn;

	@FindBy(xpath="//input[@id='txtBackUpFile']")
	public static WebElement backupFileNameTxt;


	// Buyer and Seller Pricebook Validations

	/*@FindBy(xpath="//*[@id='2007']")
	public static WebElement  financialsTransactionsPurchaseMenu; */

	@FindBy(xpath="//*[@id='2008']")
	public static WebElement  purchaseVoucher;

	@FindBy(xpath="//*[@id='2009']/span")
	public static WebElement  purchaseReturnsVoucher;

	@FindBy(xpath="//*[@id='2050']/span")
	public static WebElement  purchaseVoucherN;

	/*@FindBy(xpath="//span[contains(text(),'Purchase Voucher VAT')]")
	public static WebElement  purchaseVouchersVat;*/

	@FindBy(xpath="//input[@id='id_header_4']")
	public static WebElement  vendorAccountTxt;

	@FindBy(xpath="//tbody[@id='id_header_4_table_body']/tr/td[2]")
	public static List<WebElement> vendorAccountComboList;

	@FindBy(xpath="//tbody[@id='id_header_4_table_body']/tr/td[2]")
	public static List<WebElement> customerAccountComboList;

	@FindBy(xpath="//tbody[@id='id_header_10_table_body']/tr/td[2]")
	public static List<WebElement> currencyComboList;

	@FindBy(xpath="//tbody[@id='id_header_268435459_table_body']/tr/td[2]")
	public static List<WebElement> departmentComboList;

	/*@FindBy(xpath="//tbody[@id='id_body_536870916_table_body']/tr/td[2]")
	public static List<WebElement> warehouseBodyComboList;

	@FindBy(xpath="//tbody[@id='id_header_86_table_body']/tr/td[2]")
	public static List<WebElement> st_warehouse1HeaderComboList;

	@FindBy(xpath="//tbody[@id='id_body_87_table_body']/tr/td[2]")
	public static List<WebElement> st_warehouse2BodyComboList;

	@FindBy(xpath="//input[@id='id_body_536870916']")
	public static WebElement enter_WarehouseTxt;

	@FindBy(xpath="//*[@id='id_body_536870916_table_body']/tr")
	public static List<WebElement> pvvGridWarehouseList*/;

	/*@FindBy(xpath="//*[@id='id_body_23_table_body']/tr")
	public static List<WebElement> pvvGridItemList;

	@FindBy(xpath="//input[@id='id_body_33554522']")
	public static WebElement enter_PvTaxable;*/

	@FindBy(xpath="//*[@id='id_body_33554534']")
//	@FindBy(xpath="//*[@id='id_body_33554542']")
	public static WebElement enter_PvTaxable;

	@FindBy(xpath="//tbody[@id='id_body_16777307_table_body']/tr/td[2]")
	public static List<WebElement> pvvGridTaxCodeList;

	@FindBy(xpath="//input[@id='id_body_16777323']")
	public static WebElement enter_PvTaxCode;

	@FindBy(xpath="//input[@id='id_body_33554490']")
	public static WebElement enter_PVDiscount;

	@FindBy(xpath = "//*[@id='id_Pick']/a/span")
	public static WebElement billRefPickIcon;

	@FindBy(xpath = "//*[@id='id_Ok']/a/span")
	public static WebElement billRefOkBtn;
	
	@FindBy(xpath="//*[@id='id_Pick']/a/span")
	public static WebElement pickBtn;
	
	@FindBy(xpath="//input[@id='txtNewReference']")
	public static WebElement newReferenceTxt;

	@FindBy(xpath="//*[@id='id_Ok']/a/span")
	public static WebElement Bill_OkBtn;

	@FindBy(xpath="//input[@id='id_header_3']")
	public static WebElement  purchaseAccountTxt;

	/*@FindBy(xpath="//tbody[@id='id_header_4_table_body']/tr/td[2]")
	public static List<WebElement> vendorAccountListCount; */

	@FindBy(xpath="//input[@id='id_header_6']")
	public static WebElement  voucherHeaderDueDate;

	@FindBy(xpath="//*[@id='id_header_6_input_container']/div/i[2]")
	public static WebElement  voucherHeaderDueDateCalenderBtn;

	@FindBy(xpath="//*[@id='id_header_6_day_today']/td/span[1]")
	public static WebElement  voucherHeaderDueDateCalenderTodayBtn;


	/*@FindBy(xpath="//tbody[@id='id_header_268435459_table_body']/tr/td[2]")
	public static List<WebElement> departmentListCount;*/

	@FindBy(xpath="//*[@id='id_header_268435470']")
	public static WebElement  placeOFSupplyTxt;

	@FindBy(xpath="//input[@id='id_header_268435471']")
	public static WebElement  jurisdictionTxt;


	@FindBy(xpath="//input[@id='id_body_12']")
	public static WebElement  enter_PurchaseAccountTxt;

	/*@FindBy(xpath="//input[@id='id_body_33554521']")
	public static WebElement enter_PvVat;*/

	@FindBy(xpath="//*[@id='id_body_33554533']")
	//@FindBy(xpath="//*[@id='id_body_33554541']")
	public static WebElement enter_PvVat;

	/*@FindBy(xpath="//input[@id='id_body_16777329']")
	public static WebElement  enter_TaxCode;*/

	@FindBy(xpath="//input[@id='id_body_16777331']")
	public static WebElement  enter_TaxCode;

	@FindBy (xpath="//tbody[@id='tblBodyTransRender']/tr/td[2]/input")
	public static List<WebElement> grid_CheckBoxList;

	@FindBy (xpath="//tbody[@id='tblBodyTransRender']/tr/td[5]")
	public static List<WebElement> grid_VoucherNoList;


	@FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr[1]/td[5]")
	public static WebElement openingStocksNewHomeRow1VoucherNo;

	@FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr[1]/td[6]")
	public static WebElement openingStocksNewHomeRow1CreatedBy;

	@FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr[1]/td[7]")
	public static WebElement openingStocksNewHomeRow1ModifiedBy;

	@FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr[1]/td[12]")
	public static WebElement openingStocksNewHomeRow1SuspendStatus;

	@FindBy(xpath="//tr[@id='trRender_1']//td//input")
	public static WebElement  grid_ChkBox1;

	@FindBy(xpath="//tr[@id='trRender_2']//td//input")
	public static WebElement  grid_ChkBox2;

	@FindBy(xpath="//tr[@id='trRender_3']//td//input")
	public static WebElement  grid_ChkBox3;

	@FindBy(xpath="//tr[@id='trRender_4']//td//input")
	public static WebElement  grid_ChkBox4;

	@FindBy(xpath="//*[@id='id_header_268435459']")
	public static WebElement  departmentValuetxt;

	@FindBy(xpath="//*[@id='tblBodyTransRender']")
	public static WebElement  voucherBodyGrid;

	@FindBy(xpath="//*[@id='id_transaction_viewcontainer']/li[1]/a")
	public static WebElement voucherAll;

	@FindBy(xpath="//*[@id='id_transaction_viewcontainer']/li[2]/a")
	public static WebElement pendingAuthorizationProcess;

	@FindBy(xpath="//a[contains(text(),'Pending Material Requisition')]")
	public static WebElement pendingMaterialRequistion;

	@FindBy(xpath="//*[@id='noofalerts']")
	public static WebElement noOfAlertsDisplay;

	@FindBy(xpath="//li[@class='header']")
	public static WebElement alertsMainHeader;

	@FindBy(xpath="//li[@class='treeview-menu']//a[@id='1']")
	public static WebElement alertTransactionsTreeView;

	@FindBy(xpath="//*[@id='1']/span")
	public static WebElement alertTransaction;

	@FindBy(xpath="//*[@id='7936']/span")
	public static WebElement transactionAuthorization;

	@FindBy(xpath="//*[@id='masterChild_7936']/a/span")
	public static WebElement transactionAuthorizationText;

	@FindBy(xpath="//*[@id='id_PlanName 1']/label")
	public static WebElement pendingBudgetInfo;

	@FindBy(xpath="//*[@id='Grid_ApproveBudgetMore_control_heading_ctrl_1']")
	public static WebElement checkboxInBudgetAuthorization;

	@FindBy(xpath="//*[@id='btnAuthorize']")
	public static WebElement btnAuthorizeInBudgetAuthorization;

	@FindBy(xpath="//*[@id='btnClose']")
	public static WebElement btnCloseInBudgetAuthorization;

	/*@FindBy(xpath="//tbody[@id='id_body_23_table_body']/tr/td[2]")
	public static List<WebElement> itemListCount;

	@FindBy(xpath="//tbody[@id='id_header_3_table_body']/tr/td[2]")
	public static List<WebElement> purchaseAccountListCount;*/ 

	@FindBy(xpath="//div[3]/div[2]/table/tbody/tr[2]/td/input[2]")
	public static WebElement purchaseAccountEnter;

	@FindBy(xpath="//input[@id='id_header_3']")
	public static WebElement  MRpurchaseAccountTxt;

	@FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr[1]/td[13]")
	public static WebElement openingStocksNewHomeRow1AuthorizationStatus;

	/*@FindBy(xpath="//input[@id='txtNewReference']")
	public static WebElement  billRefNewReferenceTxt;

	@FindBy(xpath="//input[@id='txtOnAccount']")
	public static WebElement  billRefTxtOnAccount;		

	@FindBy(xpath="//input[@id='Searchtxt']")
	public static WebElement  billRefSearchTxt;		

	@FindBy(xpath="//select[@id='cbmShowBillForAdjustment']")
	public static WebElement  billRefNewReferenceDropdown;		

	@FindBy(xpath="//select[@id='cbmShowBills']")
	public static WebElement  billRefShowBillsDropdown;		

	@FindBy(xpath="//th[@id='id_Adjustment_Grid_control_heading_11']//div[contains(text(),'Due Date')]")
	public static WebElement  billRefGridHeaderDueDate;		

	@FindBy(xpath="//th[@id='id_Adjustment_Grid_control_heading_16']//div[contains(text(),'Currency')]")
	public static WebElement  billRefGridHeaderCurrency;		

	@FindBy(xpath="//th[@id='id_Adjustment_Grid_control_heading_21']//div[contains(text(),'Original Amt')]")
	public static WebElement  billRefGridHeaderOrginalAmt;		

	@FindBy(xpath="//th[@id='id_Adjustment_Grid_control_heading_23']//div[contains(text(),'Balance Amount')]")
	public static WebElement  billRefGridHeaderBalanceAmt;		

	@FindBy(xpath="//th[@id='id_Adjustment_Grid_control_heading_33']//div[contains(text(),'Adjustment Amount')]")
	public static WebElement  billRefGridHeaderAdjustmentAmt;		

	@FindBy(xpath="//div[contains(text(),'Native Currency')]")
	public static WebElement  billRefGridHeaderNativeCurrency;		

	@FindBy(xpath="//th[@id='id_Adjustment_Grid_control_heading_41']//div[contains(text(),'Narration')]")
	public static WebElement  billRefGridHeaderNarration;		

	@FindBy(xpath="//th[@id='id_Adjustment_Grid_control_heading_50']//div[contains(text(),'Previous Adjustment Amount')]")
	public static WebElement  billRefGridHeaderPreviosAdjustmentAmt;		

	@FindBy(xpath="//td[@id='id_Adjustment_Grid_col_1-0']")
	public static WebElement  billRefGridFirstRow;	

	@FindBy(xpath="//td[@id='id_Adjustment_Grid_col_1-1']//input")
	public static WebElement  billRefGridFirstRowChkbox;		

	@FindBy(xpath="//td[@id='id_Adjustment_Grid_col_2-1']//input")
	public static WebElement  billRefGridSecondRowChkbox;		

	@FindBy(xpath="//td[@id='id_Adjustment_Grid_col_3-1']//input")
	public static WebElement  billRefGridThirdRowChkbox;		

	@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr[1]/td[13]")
	public static WebElement  billRefGridFirstRowAdjustmentAmtTxt;		

	@FindBy(xpath="//td[@id='id_Adjustment_Grid_col_2-33']")
	public static WebElement  billRefGridSecondRowAdjustmentAmtTxt;		

	@FindBy(xpath="//td[@id='id_Adjustment_Grid_col_3-33']")
	public static WebElement  billRefGridThirdRowAdjustmentAmtTxt;		

	@FindBy(xpath="//td[@id='id_Adjustment_Grid_col_1-41']")
	public static WebElement  billRefGridFirstNarrationTxt;		

	@FindBy(xpath="//td[@id='id_Adjustment_Grid_col_2-41']")
	public static WebElement  billRefGridSecondNarrationTxt;		

	@FindBy(xpath="//td[@id='id_Adjustment_Grid_col_3-41']")
	public static WebElement  billRefGridThirdNarrationTxt;		

	@FindBy(xpath="//input[@id='id_Narration']")
	public static WebElement  narrationToWrite;		

	@FindBy(xpath="//input[@id='id_limit']")
	public static WebElement  adjustAmtToWrite;		

	@FindBy(xpath="//div[@class='col-xs-6 navbar-header']//div[@class='navText']")
	public static WebElement  billRefPaymentsBillReferenceLabel;		

	@FindBy(xpath="//div[@class='col-sm-12 col-md-3']//input[@class='Fbutton']")
	public static WebElement  billRefAdjustOnFIFOBtn;		

	@FindBy(xpath="//*[@id='id_Cancel']/a/span")
	public static WebElement  billRefcancel;		

	@FindBy(xpath="//label[@id='id_transactionentry_footer_panel_summary_value_net']")
	public static WebElement  billRefSummatyValueNet;		

	@FindBy(xpath="//td[@id='AccountAmount_col_1-1']")
	public static WebElement  billRefGridVendorRow1;	

	@FindBy(xpath="//td[@id='AccountAmount_col_1-2']")
	public static WebElement  billRefGridVendorAmtRow1;	

	@FindBy(xpath="//td[@id='AccountAmount_col_2-1']")
	public static WebElement  billRefGridVendorRow2;	

	@FindBy(xpath="//td[@id='AccountAmount_col_2-2']")
	public static WebElement  billRefGridVendorAmtRow2;	

	@FindBy(xpath="//td[@id='AccountAmount_col_3-1']")
	public static WebElement  billRefGridVendorRow3;

	@FindBy(xpath="//label[@id='tbNewRefAmountAdjValue']")
	public static WebElement amtAdjustedAgainstNewREfinOtherVouchers;

	@FindBy(xpath="//label[@id='tbAmountToAdjustInTransCurrencyValue']")
	public static WebElement transactionCurency;

	@FindBy(xpath="//label[@id='tbAmountToAdjustInBaseCurrencyValue']")
	public static WebElement baseCurrency;

	@FindBy(xpath="//label[@id='tbAmountToAdjustInLocalCurrencyValue']")
	public static WebElement localCurrencyDhs;

	@FindBy(xpath="//label[@id='tbNewRefBalAmountValue']")
	public static WebElement balanceNewReferenceAmt;

	@FindBy(xpath="//li[@id='id_li_Adjustment_Tab_BreakUpByTag']//span[@class='font-5'][contains(text(),'Break Up by Tag')]")
	public static WebElement breakUpByTagTab;

	@FindBy(xpath="//span[@class='font-5'][contains(text(),'Adjustment')]")
	public static WebElement adjustmentTab;

	@FindBy(xpath="//td[@id='id_BreakUpByTag_Grid_col_1-9']")
	public static WebElement breakUpByTagDueDate;

	@FindBy(xpath="//td[@id='id_BreakUpByTag_Grid_col_1-10']")
	public static WebElement breakUpByTagDepartment;

	@FindBy(xpath="//td[@id='id_BreakUpByTag_Grid_col_1-18']")
	public static WebElement breakUpByTagBaseCurrency;

	@FindBy(xpath="//td[@id='id_BreakUpByTag_Grid_col_1-17']")
	public static WebElement breakUpByTagBillCurrency;

	@FindBy(xpath="//td[@id='id_BreakUpByTag_Grid_col_1-20']")
	public static WebElement breakUpByTagBaseConversionRate;

	@FindBy(xpath="//td[@id='id_BreakUpByTag_Grid_col_1-21']")
	public static WebElement breakUpByTagBaseLocalCurrency;

	@FindBy(xpath="//label[@id='txtblkAmountadjusted']")
	public static WebElement  billRefAdjustAmountInTransCurency;

	@FindBy(xpath="//label[@id='txtblkAmounttobeadjust']")
	public static WebElement  billRefBalanceAmountAdjustInTrnasCurrency;

	@FindBy(xpath="//label[@id='id_BillWise_IP_LocalConversionRateValue']")
	public static WebElement  billRefInfoBarLocalCurrencyRate;

	@FindBy(xpath="//label[@id='tbAmountToAdjustInTransCurrencyValue']")
	public static WebElement billRefTransactionCurency;

	@FindBy(xpath="//label[@id='tbAmountToAdjustInBaseCurrencyValue']")
	public static WebElement billRefBaseCurrency;

	@FindBy(xpath="(//*[@id='id_BillWise_IP_AccountName'])[1]")
	public static WebElement conversationRateBaseCurrencyRate;
	
	@FindBy(xpath = "//div[@id='billwise2']/div[1]/label[2]")
	public static WebElement conversationRateBaseCurrencyRate;
	
	@FindBy(xpath="//*[@id='id_BillWise_IP_LocalConversionRateValue']")
	public static WebElement conversationRateLocalCurrencyRate;
*/
	@FindBy(xpath="//*[@id='id_transactionentry_save']")
	public static WebElement openingBalancesSaveBtn;

	@FindBy(xpath="//input[@id='id_header_11']")
	public static WebElement  voucherHeaderExchangeRate;

	@FindBy(xpath="//input[@id='id_header_145']")
	public static WebElement  voucherHeaderLocalExchangeRate;

	@FindBy(xpath="//label[@id='id_transactionentry_footer_panel_summary_value_18']")
	public static WebElement vocFooterdebitAmount;
	
	@FindBy(xpath="//*[@id='id_transactionentry_summary']/div[1]/div[2]/div/table")
	public static WebElement vocFooterExpandBtn;

	@FindBy(xpath="//label[@id='id_transactionentry_footer_panel_summary_value_19']")
	public static WebElement vocFooterCreditAmount;

	@FindBy (xpath="//tbody[@id='tblBodyTransRender']/tr/td")
	public static List<WebElement> voucherHomeBodyList;

	@FindBy (xpath="//tbody[@id='tblBodyTransRender']/tr[2]/td[2]")
	public static WebElement voucherHomeBody2rdRow1stCol;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[1]")
	public static WebElement  select1stRow_0stColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[2]")
	public static WebElement  select1stRow_1stColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[3]")
	public static WebElement  select1stRow_2ndColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[4]")
	public static WebElement  select1stRow_3rdColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[5]")
	public static WebElement  select1stRow_4thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[6]")
	public static WebElement  select1stRow_5thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[7]")
	public static WebElement  select1stRow_6thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[8]")
	public static WebElement  select1stRow_7thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[9]")
	public static WebElement  select1stRow_8thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[10]")
	public static WebElement  select1stRow_9thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[11]")
	public static WebElement  select1stRow_10thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[12]")
	public static WebElement  select1stRow_11thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[13]")
	public static WebElement  select1stRow_12thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[14]")
	public static WebElement  select1stRow_13thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[15]")
	public static WebElement  select1stRow_14thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[16]")
	public static WebElement  select1stRow_15thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[17]")
	public static WebElement  select1stRow_16thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[18]")
	public static WebElement  select1stRow_17thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[19]")
	public static WebElement  select1stRow_18thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[20]")
	public static WebElement  select1stRow_19thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[21]")
	public static WebElement  select1stRow_20thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[22]")
	public static WebElement  select1stRow_21stColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[23]")
	public static WebElement  select1stRow_22ndColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[24]")
	public static WebElement  select1stRow_23rdColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[25]")
	public static WebElement  select1stRow_24thColumn;


	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[1]")
	public static WebElement  select2ndRow_0stColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[2]")
	public static WebElement  select2ndRow_1stColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[3]")
	public static WebElement  select2ndRow_2ndColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[4]")
	public static WebElement  select2ndRow_3rdColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[5]")
	public static WebElement  select2ndRow_4thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[6]")
	public static WebElement  select2ndRow_5thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[7]")
	public static WebElement  select2ndRow_6thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[8]")
	public static WebElement  select2ndRow_7thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[9]")
	public static WebElement  select2ndRow_8thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[10]")
	public static WebElement  select2ndRow_9thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[11]")
	public static WebElement  select2ndRow_10thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[12]")
	public static WebElement  select2ndRow_11thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[13]")
	public static WebElement  select2ndRow_12thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[14]")
	public static WebElement  select2ndRow_13thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[15]")
	public static WebElement  select2ndRow_14thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[16]")
	public static WebElement  select2ndRow_15thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[17]")
	public static WebElement  select2ndRow_16thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[18]")
	public static WebElement  select2ndRow_17thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[19]")
	public static WebElement  select2ndRow_18thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[20]")
	public static WebElement  select2ndRow_19thColumn;


	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[2]")
	public static WebElement  select3rdRow_1stColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[3]")
	public static WebElement  select3rdRow_2ndColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[4]")
	public static WebElement  select3rdRow_3rdColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[5]")
	public static WebElement  select3rdRow_4thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[6]")
	public static WebElement  select3rdRow_5thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[7]")
	public static WebElement  select3rdRow_6thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[8]")
	public static WebElement  select3rdRow_7thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[9]")
	public static WebElement  select3rdRow_8thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[10]")
	public static WebElement  select3rdRow_9thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[11]")
	public static WebElement  select3rdRow_10thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[12]")
	public static WebElement  select3rdRow_11thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[13]")
	public static WebElement  select3rdRow_12thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[14]")
	public static WebElement  select3rdRow_13thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[15]")
	public static WebElement  select3rdRow_14thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[16]")
	public static WebElement  select3rdRow_15thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[17]")
	public static WebElement  select3rdRow_16thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[18]")
	public static WebElement  select3rdRow_17thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[19]")
	public static WebElement  select3rdRow_18thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[20]")
	public static WebElement  select3rdRow_19thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[21]")
	public static WebElement  select3rdRow_20thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[22]")
	public static WebElement  select3rdRow_21stColumn;



	@FindBy(xpath="//input[@id='id_body_23']")
	public static WebElement  enter_ItemTxt;
	
	@FindBy(xpath="//*[@id='id_body_24']")
	public static WebElement enter_Units;

	@FindBy(xpath="//input[@id='id_body_12']")
	public static WebElement  enter_SalesAccountTxt;

	@FindBy(xpath="//input[@id='id_body_24']")
	public static WebElement  enter_UnitTxt;

	@FindBy(xpath="//input[@id='id_body_26']")
	public static WebElement  enter_Quantity;

	@FindBy(xpath="//input[@id='id_body_27']")
	public static WebElement  enter_Rate;

	@FindBy(xpath="//input[@id='id_body_28']")
	public static WebElement  enter_Gross;

	@FindBy (xpath="//input[@id='id_body_38']")
	public static WebElement enter_RMA;

	@FindBy(xpath="//input[@id='id_body_13']")
	public static WebElement  enter_Batch;

	@FindBy(xpath="//input[@id='id_body_12']")
	public static WebElement  enter_DebitACTxt;


	@FindBy(xpath="//input[@id='id_body_37']")
	public static WebElement enter_Expirydate;

	@FindBy(xpath="//input[@id='id_body_36']")
	public static WebElement enter_MfgDate;

	@FindBy(xpath="//input[@id='id_body_108']")
	public static WebElement enter_BatchRate;

	@FindBy(xpath="//input[@id='id_body_109']")
	public static WebElement enter_BatchRateValue1;

	@FindBy(xpath="//input[@id='id_body_110']")
	public static WebElement enter_BatchRateValue2;

	@FindBy(xpath="//*[@id='id_transactionentry_previous']")
	public static WebElement  previousBtn;

	@FindBy(xpath="//*[@id='id_transactionentry_next']")
	public static WebElement  nextBtn;

	@FindBy(xpath="//*[@id='id_transactionentry_print']")
	public static WebElement  new_PrintBtn;

	@FindBy(xpath="//*[@id='id_transactionentry_suspend']")
	public static WebElement  new_SuspendBtn;

	@FindBy(xpath="//*[@id='id_transactionentry_close']")
	public static WebElement  new_CloseBtn;

	@FindBy(xpath="//li[@id='id_transactionentry_delete']/a/span")
	public static WebElement  new_DeleteBtn;



	@FindBy(xpath="//span[@class='icon-scroll icon-font6']")
	public static WebElement  freeFlowBtn;

	@FindBy(xpath="//span[@class='icon-header icon-font6']")
	public static WebElement  jumpToHeaderSectionBtn;

	@FindBy(xpath="//span[@class='icon-panel icon-font6']")
	public static WebElement  jumpToBodySectionBtn;

	@FindBy(xpath="//span[@class='icon-footer icon-font6']")
	public static WebElement  jumpToFooterBtn;

	@FindBy(xpath="//*[@id='id_transactionentry_new']/a/span")
	public static WebElement  new_newBtn;

	@FindBy(xpath="//*[@id='id_transactionentry_save']")
	public static WebElement  voucherSaveBtn;

	@FindBy(xpath="(//*[@class='icon-close hiconright2'])[1]")  
	public static WebElement  voucherhomeCloseBtn;
	
	@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[15]/a/i")
	public static WebElement poVoucherHomeCloseBtn;

	@FindBy(xpath="//*[@id='2023']")
	public static WebElement  financialsTransactionsJournalsMenu;
	
	@FindBy(xpath="//*[@id='navigation_menu']/li[2]/ul/li[2]/ul/li[5]/ul/li")
	public static List<WebElement> JournalsSubMenusList;

	@FindBy(xpath="//*[@id='navigation_menu']/li[2]/ul/li[2]/ul/li[5]/ul/li[14]")
	public static WebElement OpeningBalanceMenu;
	
	@FindBy(xpath="//*[@id='2028']")
	public static WebElement  openingBalancesVoucher;

	@FindBy(xpath="//*[@id='id_transaction_homescreen_Delete']/i")
	public static WebElement  voucherHomeDeleteBtn;

	@FindBy(xpath="//a[@id='id_transaction_homescreen_convert']/i")
	public static WebElement  convertBtn;

	@FindBy(xpath="//span[@class='icon-export icon-font6']")
	public static WebElement  exportToXMLBtn;  

	@FindBy(xpath="//*[@id='id_header_1_input_image']/span")
	public static WebElement  documentNumberdropdown_ExpansionBtn;

	@FindBy(xpath="//td[@id='id_header_2_input_image']//span[@class='icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow']")
	public static WebElement  dateTxt_CalenderBtn;

	@FindBy(xpath="//tr[@id='id_header_2_day_today']//span[@class='theme_color-inverse'][contains(text(),'Today')]")
	public static WebElement  calender_TodayBtn;

	@FindBy(xpath="//input[@id='id_header_268435460']")
	public static WebElement  warehouseTxt;

	/*@FindBy(xpath="//tbody[@id='id_header_268435460_table_body']/tr/td[2]")
	public static List<WebElement> warehouseHeaderComboList;

	@FindBy(xpath="//tbody[@id='id_body_23_table_body']/tr/td[2]")
	public static List<WebElement> itemComboList;*/

	@FindBy(xpath="//a[@id='137']")
	public static WebElement  inventoryTransactionsMenu; 

	@FindBy(xpath="//*[@id='139']")
	public static WebElement  inventoryTransactionsPurchasesMenu; 

	@FindBy(xpath="//*[@id='2010']/span")
	public static WebElement  requestForQuoteVoucher;

	@FindBy(xpath="//*[@id='2011']/span")
	public static WebElement  purchasesQuotationsVoucher;

	@FindBy(xpath="//*[@id='2012']")
	public static WebElement  purchasesOrdersVoucher;

	@FindBy(xpath="//*[@id='2013']")
	public static WebElement  materialReceiptNotesVoucher;

	@FindBy(xpath="//*[@id='140']")
	public static WebElement  inventoryTransactionsSalesMenu; 

	@FindBy (xpath="//div[@id='Modal_Header']/div[1]")
	public static WebElement rmaScreenTitle;

	@FindBy (xpath="//input[@id='txtSerialNo']")
	public static WebElement rmaSerialNumberTxtField;

	@FindBy (xpath="//input[@id='txtQuantity']")
	public static WebElement rmaQuantityTxtField;

	@FindBy (xpath="//*[@id='Modal_Search_Body']/div/div[1]/div[3]/div/span/i")
	public static WebElement rmaAddBtn;

	@FindBy (xpath="//div[@id='Modal_Search_Body']/div[3]/label")
	public static WebElement rmaNumberofItemsLabel;

	@FindBy (xpath="//th[@id='RMA_Table_control_heading_1']/div")
	public static WebElement rmaTableHeadingRMA;

	@FindBy (xpath="//th[@id='RMA_Table_control_heading_2']/div")
	public static WebElement rmaTableHeadingDocumentNo;

	@FindBy (xpath="//th[@id='RMA_Table_control_heading_3']/div")
	public static WebElement rmaTableHeadingDocumentDate;

	//@FindBy (xpath="//div[@id='RMAModel_Bottom']/div[1]/label")
	@FindBy(xpath="//*[@id='div_Clear'][1]")
	public static WebElement rmaClearBtn;

	@FindBy (xpath="//div[@id='RMAModel_Bottom']/div[3]/label")
	public static WebElement rmaCancelBtn;

	@FindBy (xpath="//div[@id='Modal_Header']/div[2]/span/i")
	public static WebElement rmaCloseBtn;

	@FindBy (xpath="//td[@id='RMA_Table_col_1-1']")
	public static WebElement rmaTableRow1Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_2-1']")
	public static WebElement rmaTableRow2Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_3-1']")
	public static WebElement rmaTableRow3Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_4-1']")
	public static WebElement rmaTableRow4Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_5-1']")
	public static WebElement rmaTableRow5Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_6-1']")
	public static WebElement rmaTableRow6Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_7-1']")
	public static WebElement rmaTableRow7Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_8-1']")
	public static WebElement rmaTableRow8Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_9-1']")
	public static WebElement rmaTableRow9Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_10-1']")
	public static WebElement rmaTableRow10Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_11-1']")
	public static WebElement rmaTableRow11Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_12-1']")
	public static WebElement rmaTableRow12Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_13-1']")
	public static WebElement rmaTableRow13Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_14-1']")
	public static WebElement rmaTableRow14Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_15-1']")
	public static WebElement rmaTableRow15Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_16-1']")
	public static WebElement rmaTableRow16Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_17-1']")
	public static WebElement rmaTableRow17Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_18-1']")
	public static WebElement rmaTableRow18Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_19-1']")
	public static WebElement rmaTableRow19Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_20-1']")
	public static WebElement rmaTableRow20Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_1-2']")
	public static WebElement rmaTableRow1Column2;

	@FindBy (xpath="//td[@id='RMA_Table_col_1-3']")
	public static WebElement rmaTableRow1Column3;

	@FindBy (xpath="//td[@id='RMA_Table_col_1-1']")
	public static WebElement RMAPopupTableRow1Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_2-1']")
	public static WebElement RMAPopupTableRow2Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_3-1']")
	public static WebElement RMAPopupTableRow3Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_4-1']")
	public static WebElement RMAPopupTableRow4Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_5-1']")
	public static WebElement RMAPopupTableRow5Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_6-1']")
	public static WebElement RMAPopupTableRow6Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_7-1']")
	public static WebElement RMAPopupTableRow7Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_8-1']")
	public static WebElement RMAPopupTableRow8Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_9-1']")
	public static WebElement RMAPopupTableRow9Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_10-1']")
	public static WebElement RMAPopupTableRow10Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_11-1']")
	public static WebElement RMAPopupTableRow11Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_12-1']")
	public static WebElement RMAPopupTableRow12Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_13-1']")
	public static WebElement RMAPopupTableRow13Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_14-1']")
	public static WebElement RMAPopupTableRow14Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_15-1']")
	public static WebElement RMAPopupTableRow15Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_16-1']")
	public static WebElement RMAPopupTableRow16Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_17-1']")
	public static WebElement RMAPopupTableRow17Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_18-1']")
	public static WebElement RMAPopupTableRow18Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_19-1']")
	public static WebElement RMAPopupTableRow19Column1;

	@FindBy (xpath="//td[@id='RMA_Table_col_20-1']")
	public static WebElement RMAPopupTableRow20Column1;

	@FindBy (xpath="//td[@id='id_body_38_input_image']/img")
	public static WebElement rmaExpansionBtn;

	@FindBy (xpath="//div[@id='RMAModel_Bottom']/div[1]/label")
	public static WebElement RMAPopupClearBtn;

	@FindBy (xpath="//button[@id='div_OK']")
	public static WebElement RMAPopupOkBtn;

	@FindBy (xpath="(//button[@id='div_Clear'])[2]")
	public static WebElement RMAPopupCancelBtn;

	@FindBy (xpath="//div[@id='Modal_Header']/div[2]/span/i")
	public static WebElement RMAPopupCloseBtn;

	@FindBy (xpath="//input[@id='id_bins_balance']")
	public static WebElement binBalanceTxt;

	@FindBy (xpath="//*[@id='id_btnautoallocate']")
	public static WebElement binAutoAllocateBtn;

	@FindBy (xpath="//div[@id='TransactionNewBin']/div[2]/div/div[1]/span/i")
	public static WebElement binCancelBtn;

	@FindBy (xpath="//input[@id='srch_bin']")
	public static WebElement binSearchTxt;

	@FindBy (xpath="//*[@id='TransactionNewBin']/div[2]/div/div[2]/div/div/div[1]/div[3]/div/i")
	public static WebElement binSearchBtn; 

/*	@FindBy (xpath="//*[@id='115']")
	public static WebElement eraseAll;
*/
	@FindBy(xpath="//select[@id='cmbBehaviour']")
	public static WebElement editScreenBehaviourDropdown;

	@FindBy(xpath="//input[@id='FTextBox_13_2000']")
	public static WebElement widthTxt;

	@FindBy(xpath="//input[@id='id_body_33554518']")
	public static WebElement enter_purrt0Txt;

	@FindBy(xpath="//input[@id='id_body_33554519']")
	public static WebElement enter_purrt1Txt;

	@FindBy(xpath="//input[@id='id_body_33554520']")
	public static WebElement enter_purrtGrp0Txt;

	@FindBy(xpath="//input[@id='id_body_33554521']")
	public static WebElement enter_purrtGrp1Txt;

	@FindBy (xpath="//input[@id='editScreen_qtyFormula_textbox']")
	public static WebElement editScreenBehaviourFormulaTxt;
	
	@FindBy (xpath="//*[@id='editScreen_qtyPreLoaded_textbox']")
	public static WebElement editScreenBehaviourPreloadFormulaTxt;

	//@FindBy (xpath="//button[@id='editScreen_qtyFormula_Ok']")
	@FindBy(xpath="//*[@id='editScreen_qtyPreLoaded_Ok']")
	public static WebElement editScreenBehaviourFormulaOkBtn;

	@FindBy (xpath="//button[@id='editScreen_qtyFormula_Cancel']")
	public static WebElement editScreenBehaviourFormulaCancelBtn;

	@FindBy (xpath="//a[@id='btnCustomizeClose']/i")
	public static WebElement customizationCloseBtn;

	@FindBy(xpath="//label[contains(text(),'Posting details')]")
	public static WebElement postingDetailsBtn;

	@FindBy(xpath="//a[@id='id_transactionentry_more']/span")
	public static WebElement  toggleBtn;

	@FindBy(xpath="//*[@id='id_transactionentry_postingdetail_popup']/div/div/div[2]/div/div[2]/div[1]/div/div")
	public static List<WebElement>  postingDetailsDebitList; 

	@FindBy(xpath="//*[@id='id_transactionentry_postingdetail_popup']/div/div/div[2]/div/div[2]/div[1]/div/div[1]")
	public static List<WebElement>  postingDetailsDebitSecList_AccList;

	@FindBy(xpath="//*[@id='id_transactionentry_postingdetail_popup']/div/div/div[2]/div/div[2]/div[2]/div/div")
	public static List<WebElement>  postingDetailsCreditList; 

	@FindBy(xpath="//*[@id='id_transactionentry_postingdetail_popup']/div/div/div[2]/div/div[2]/div[2]/div/div[1]")
	public static List<WebElement>  postingDetailsCreditSec_AccList; 

	@FindBy(xpath="//*[@id='id_transactionentry_postingdetail_popup']/div/div/div[2]/div/div[3]/div[1]/div/div[2]")
	public static WebElement  postingDetailsDebitSum; 

	@FindBy(xpath="//*[@id='id_transactionentry_postingdetail_popup']/div/div/div[2]/div/div[3]/div[2]/div/div[2]")
	public static WebElement  postingDetailsCreditSum; 

	@FindBy(xpath="//*[@id='id_transactionentry_postingdetail_popup']/div[2]/div/div[1]/span/i")
	public static WebElement  postingDetailsCloseBtn; 

	@FindBy(xpath="//*[@id='opt_LayoutID']")
	public static WebElement LayoutNameTxt;

	@FindBy(xpath="//*[@id='cmd_LayoutOptions']")
	public static WebElement LayoutOptionsdropdown;

	@FindBy(xpath="//*[@id='btnPreview']")
	public static WebElement  PreviewBtn;

	@FindBy(xpath="//*[@id='btnPrint']")
	public static WebElement footerPrintBtn;

	@FindBy(xpath="//*[@id='btnPreview']/following-sibling::input[@value='Close']")
	public static WebElement CloseBtn;

	@FindBy (xpath="(//div[@id='dvCreateViewMain']/div[1])[2]/div/div[2]/span[1]/i")
	public static WebElement createLayoutBtn;



	@FindBy(xpath="//i[@class='icon-font6 icon-new']")
	public static WebElement  Newdropdown;

	@FindBy(xpath="//*[@id='forHf']/span[1]/ul/li/a[1]")
	public static WebElement  New_LayoutBtn;

	@FindBy(xpath="//*[@id='forHf']/span[1]/ul/li/a[2]")
	public static WebElement  New_PageBtn;

	@FindBy(xpath="//*[@id='3']")
	public static WebElement  OpenBtn_First;

	@FindBy(xpath="//*[@id='7']")
	public static WebElement  OpenBtn_Second;

	@FindBy(xpath="//button[@id='btn_PopUpOk']")
	public static WebElement  OpenBtn_OkBtn;

	@FindBy(xpath="//*[@id='id_PopUp']/div/div/div[3]/button[2]")
	public static WebElement  OpenBtn_CancelBtn;

	@FindBy(xpath="//*[@id='forHf']/span[2]/ul/li/a[2]")
	public static WebElement FromXMLBtn;


	@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[1]")
	public static WebElement  SaveBttn;

	@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[2]")
	public static WebElement  SaveAsBtn; //Clicking popup is same as savebtn popup

	@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[3]")
	public static WebElement  SaveAsImageBtn; //No Response

	@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[4]")
	public static WebElement  SaveAsXMLBtn; // Downloads xml file

	@FindBy(xpath="//*[@id='forHf']/span[23]/ul/li/a[5]")
	public static WebElement  SaveAsHTMLBtn; // Showing Error

	@FindBy(xpath="//*[@id='forHf']/span[4]")
	public static WebElement  Removedropdown;

	@FindBy(xpath="//*[@id='forHf']/span[4]/ul/li/a[1]")
	public static WebElement  RemoveBtn; //Raises popup

	@FindBy(xpath="//*[@id='forHf']/span[4]/ul/li/a[2]")
	public static WebElement  RemoveCurrentPageBtn; //Raises popup

	@FindBy(xpath="//*[@id='forHf']/span[4]/ul/li/a[3]")
	public static WebElement  RemoveControlBtn;				

	@FindBy(xpath="//*[@id='forHf']/span[5]")
	public static WebElement  PageNo;

	@FindBy(xpath="//*[@id='forHf']/span[6]")
	public static WebElement  Controldrpdown;

	@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[1]")
	public static WebElement  StaticTextBtn;

	@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[2]")
	public static WebElement  BitmapBtn;

	@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[3]")
	public static WebElement  BodyGridBtn;	

	@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[4]")
	public static WebElement  RectangleBtn;	

	@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[5]")
	public static WebElement  EcllipseBtn;	

	@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[6]")
	public static WebElement  LineBtn;	

	@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[7]")
	public static WebElement  AreaBtn;

	@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[8]")
	public static WebElement  TableBtn;	

	@FindBy(xpath="//*[@id='forHf']/span[7]")
	public static WebElement  Alignmentdropdown;

	@FindBy(xpath="//*[@id='forHf']/span[7]/ul/li/a[1]")
	public static WebElement  LeftAlignmentBtn;

	@FindBy(xpath="//*[@id='forHf']/span[7]/ul/li/a[2]")
	public static WebElement  RightAlignmentBtn;

	@FindBy(xpath="//*[@id='forHf']/span[7]/ul/li/a[3]")
	public static WebElement  TopAlignmentBtn;	

	@FindBy(xpath="//*[@id='forHf']/span[7]/ul/li/a[4]")
	public static WebElement  BottomAlignmentBtn;

	@FindBy(xpath="//*[@id='forHf']/span[8]")
	public static WebElement  Sizedropdown;

	@FindBy(xpath="//*[@id='forHf']/span[8]/ul/li/a[1]")
	public static WebElement  SameSizeBtn;

	@FindBy(xpath="//*[@id='forHf']/span[8]/ul/li/a[2]")
	public static WebElement  SameWidthBtn;

	@FindBy(xpath="//*[@id='forHf']/span[8]/ul/li/a[3]")
	public static WebElement  SameHeightBtn;

	@FindBy(xpath="//*[@id='forHf']/span[9]")
	public static WebElement  Printdropdown;

	@FindBy(xpath="//*[@id='forHf']/span[9]/ul/li/a[1]")
	public static WebElement  Print_PrintBtn;

	@FindBy(xpath="//*[@id='forHf']/span[9]/ul/li/a[2]")
	public static WebElement  Print_PreviewBtn;

	@FindBy(xpath="//*[@id='fit-button']")
	public static WebElement  Preview_FitToWidthBtn;

	@FindBy(xpath="//*[@id='zoom-in-button']")
	public static WebElement  Preview_ZoomInBtn;

	@FindBy(xpath="//*[@id='zoom-out-button']")
	public static WebElement  Preview_ZoomOutBtn;

	@FindBy(xpath="//*[@id='forHf']/span[9]/ul/li/a[3]")
	public static WebElement  PageSetUpBtn;

	@FindBy(xpath="//*[@id='id_PageSetUpPreview']")
	public static WebElement  PageSetupPagePreviewField;

	@FindBy(xpath="//*[@id='id_PageType']")
	public static WebElement  PageSetupPaperSizedropdown;

	@FindBy(xpath="//*[@id='id_PaperOrientation'][@value='0']")
	public static WebElement  PageSetupPotraitRadio;

	@FindBy(xpath="//*[@id='id_PaperOrientation'][@value='1']")
	public static WebElement  PageSetupLandScapeRadio;

	@FindBy(xpath="//*[@id='id_PaperUnit'][@value='0']")
	public static WebElement  PageSetupCentimetersRadio;

	@FindBy(xpath="//*[@id='id_PaperUnit'][@value='1']")
	public static WebElement  PageSetupInchesRadio;

	@FindBy(xpath="//*[@id='id_PageSizeLeftMarginDisplay']")
	public static WebElement  PageSetupLeftTxt;

	@FindBy(xpath="//*[@id='id_PageSizeRightMarginDisplay']")
	public static WebElement  PageSetupRightTxt;

	@FindBy(xpath="//*[@id='id_PageSizeTopMarginDisplay']")
	public static WebElement  PageSetupTopTxt;

	@FindBy(xpath="//*[@id='id_PageSizeTopMarginDisplay']")
	public static WebElement  PageSetupBottomTxt;

	@FindBy(xpath="//*[@id='id_PageSetupOkBtn']")
	public static WebElement  PageSetupOkBtn;

	@FindBy(xpath="//*[@id='id_PageSetupCancelBtn']")
	public static WebElement  PageSetUpCancelBtn;

	@FindBy(xpath="//li[@id='li_InvoicePageSetup']")
	public static WebElement  pagesetupTab;	

	@FindBy(xpath="//li[@id='li_InvoiceDesigner']")
	public static WebElement  designerTab;	

	@FindBy(xpath="//li[@id='li_InvoicePreview']")
	public static WebElement  previewTab;	


	@FindBy(xpath="//*[@id='opt_DocSetID']")
	public static WebElement  DocumentSet;

	@FindBy(xpath="//*[@id='opt_DocSetID_input_image']/span")
	public static WebElement DocumentSet_ExpandBtn;

	@FindBy(xpath="//*[@id='opt_DocSetID_input_settings']/span")
	public static WebElement DocumentSet_SettingsBtn;

	@FindBy(xpath="//*[@id='opt_DocSetID_customize_popup_container']")
	public static WebElement  DocumentSet_Settings_Container;

	@FindBy(xpath="//*[@id='opt_DocSetID_customize_popup_footer']/div/div/input[1]")
	public static WebElement  DocumentSet_Settings_StandardFieldsBtn;

	@FindBy(xpath="//*[@id='opt_DocSetID_customize_popup_standardfields_list']")
	public static WebElement  DocumentSet_Settings_StandardFields_Filedsdropdown;

	@FindBy(xpath="//*[@id='opt_DocSetID_customize_popup_standardfields_header']")
	public static WebElement  DocumentSet_Settings_StandardFields_HeaderTxt;

	@FindBy(xpath="//*[@id='opt_DocSetID_customize_popup_standardfields_alignment']")
	public static WebElement  DocumentSet_Settings_StandardFields_Allignmentdropdown;

	@FindBy(xpath="//*[@id='opt_DocSetID_customize_popup_standardfields_width']")
	public static WebElement  DocumentSet_Settings_StandardFields_WidthTcxt;

	@FindBy(xpath="//div[@id='id_popup_container']//input[3]")
	public static WebElement  DocumentSet_Settings_StandardFields_OkBtn;

	@FindBy(xpath="//div[@id='id_popup_container']//input[4]")
	public static WebElement  DocumentSet_Settings_StandardFields_CancelBtn;

	@FindBy(xpath="//*[@id='opt_DocSetID_customize_popup_footer']/div/div/input[2]")
	public static WebElement  DocumentSet_Settings_DeletecolumnBtn;

	@FindBy(xpath="//div[@id='id_popup_container']//input[3]")
	public static WebElement  DocumentSet_Settings_OkBtn;

	@FindBy(xpath="//div[@id='id_popup_container']//input[4]")
	public static WebElement  DocumentSett_Settings_CancelBtn;


	@FindBy(xpath="//div[@id='dvCreateViewMain']/div[2]/div/div[2]/span[1]/i")
	public static WebElement  DocumentSet_CreatelayoutBtn;

	@FindBy(xpath="//*[@id='myNavbar']/ul/li/span[1]")
	public static WebElement  DocumentSet_Createlayout_SaveBtn;

	@FindBy(xpath="//*[@id='deleteBtn']")
	public static WebElement  DocumentSet_Createlayout_DeleteBtn;

	@FindBy(xpath="//*[@id='myNavbar']/ul/li/span[3]")
	public static WebElement  DocumentSet_Createlayout_ResetBtn;

	@FindBy(xpath="//*[@id='myNavbar']/ul/li/span[4]")
	public static WebElement  DocumentSet_Createlayout_CancelBtn;

	@FindBy(xpath="//*[@id='txtDocSetName']")
	public static WebElement  DocumentSetTxt;

	@FindBy(xpath="//*[@id='BrowseId']")
	public static WebElement  DocumentSetBrowseBtn;

	@FindBy(xpath="//*[@id='btnPrint']")
	public static WebElement  Browse_OkBtn;

	@FindBy(xpath="//*[@id='footer']/button[2]")
	public static WebElement  Browse_CancelBtn;

	@FindBy(xpath="//*[@id='chkId']")
	public static WebElement  EjectPaperAfterEachFormatPrintChkBox;

	@FindBy(xpath="(//div[@id='dvCreateViewMain']/div[1])[2]/div/div[2]/span[2]/i")
	public static WebElement  LayoutName_Editlayout;

	@FindBy(xpath="//div[@id='dvCreateViewMain']/div[2]/div/div[2]/span[2]/i")
	public static WebElement  DocumentSet_Editlayout;

	@FindBy(xpath="//*[@id='btnPrint']")
	public static WebElement  PrintBtn;

	@FindBy(xpath="//span[contains(text(),'PurchaseAC')]//i[@class='icon-expand icon-font8']")
	public static WebElement Finance_Purchases_Header_PurchaseAcExpandBtn;

	@FindBy(xpath="//span[contains(text(),'Time')]//i[@class='icon-expand icon-font8']")
	public static WebElement  Finance_Purchases_Header_TimeExpandBtn;

	@FindBy(xpath="//li[@class='treeview active']//span[contains(text(),'Date')]//i[@class='icon-expand icon-font8']")
	public static WebElement  Finance_Purchases_Header_DateExpandBtn;

	@FindBy(xpath="//span[contains(text(),'VendorAC')]//i[@class='icon-expand icon-font8']")
	public static WebElement Finance_Purchases_Header_VendorACExpandBtn;




	@FindBy(xpath="//span[contains(text(),'UpdateStock')]")
	public static WebElement  Finance_Purchases_Header_UpdateStockBtn;

	@FindBy(xpath="//span[contains(text(),'RaiseReceipt')]")
	public static WebElement  Finance_Purchases_Header_RaiseReceiptBtn;

	@FindBy(xpath="//span[contains(text(),'Warehouse')]//i[@class='icon-expand icon-font8']")
	public static WebElement  Finance_Purchases_Header_WareHouseExpandBtn;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[8]/a")
	public static WebElement  Finance_Purchases_Header_DepartmentExpandBtn;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[8]/ul/li[1]/a")
	public static WebElement  Finance_Purchases_Department_Namedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[8]/ul/li[2]/a")
	public static WebElement  Finance_Purchases_Department_Codedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[8]/ul/li[3]/a")
	public static WebElement  Finance_Purchases_Department_Aliasdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[8]/ul/li[4]/a")
	public static WebElement  Finance_Purchases_Department_Currencydraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[8]/ul/li[5]/a")
	public static WebElement  Finance_Purchases_Department_GroupLevel1draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[8]/ul/li[6]/a")
	public static WebElement  Finance_Purchases_Department_GroupLevel1Codedraganddrop;

	@FindBy(xpath="//span[contains(text(),'Narration')]")
	public static WebElement  Finance_Purchases_Narrationdraganddrop;

	@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Test')]")
	public static WebElement  Finance_Purchases_Testdraganddrop;

	@FindBy(xpath="//span[contains(text(),'Formula')]")
	public static WebElement  Finance_Purchases_Formuladraganddrop;

	@FindBy(xpath="//span[contains(text(),'FD%')]")
	public static WebElement  Finance_Purchases_FDdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[10]/a")
	public static WebElement  Finance_Purchases_ChequeNodraganddrop;

	@FindBy(xpath="//li[@class='treeview active']//span[contains(text(),'Created by')]")
	public static WebElement  Finance_Purchases_CreatedBydraganddrop;

	@FindBy(xpath="//span[contains(text(),'Created date')]//i[@class='icon-expand icon-font8']")
	public static WebElement  Finance_Purchases_Header_CreatedDateExpandBtn;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[12]/ul/li[1]/a")
	public static WebElement  Finance_Purchases_CreatedDate_CreatedDatedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[12]/ul/li[2]/a")
	public static WebElement  Finance_Purchases_CreatedDate_Yeardraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[12]/ul/li[3]/a")
	public static WebElement  Finance_Purchases_CreatedDate_Monthdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[12]/ul/li[4]/a")
	public static WebElement  Finance_Purchases_CreatedDate_DayofYeardraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[12]/ul/li[5]/a")
	public static WebElement  Finance_Purchases_CreatedDate_Daydraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[12]/ul/li[6]/a")
	public static WebElement  Finance_Purchases_CreatedDate_Weekdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[12]/ul/li[7]/a")
	public static WebElement  Finance_Purchases_CreatedDate_WeekDaydraganddrop;


	@FindBy(xpath="//span[contains(text(),'Created time')]//i[@class='icon-expand icon-font8']")
	public static WebElement  Finance_Purchases_Header_CreatedTimeExpandBtn;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[13]/ul/li[1]/a")
	public static WebElement  Finance_Purchases_CreatedTime_CreatedTimedrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[13]/ul/li[2]/a")
	public static WebElement  Finance_Purchases_CreatedTime_Hourdrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[13]/ul/li[3]/a")
	public static WebElement  Finance_Purchases_CreatedTime_Minutedrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[13]/ul/li[4]/a")
	public static WebElement  Finance_Purchases_CreatedTime_Seconddrapanddrop;

	@FindBy(xpath="//li[@class='treeview active']//span[contains(text(),'Modified by')]")
	public static WebElement  Finance_Purchases_Header_ModifiedBydrapanddrop;

	@FindBy(xpath="//span[contains(text(),'Modified date')]//i[@class='icon-expand icon-font8']")
	public static WebElement  Finance_Purchases_Header_ModifiedDateExpandBtn;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[15]/ul/li[1]/a")
	public static WebElement  Finance_Purchases_ModifiedDate_ModifiedDatedrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[15]/ul/li[2]/a")
	public static WebElement  Finance_Purchases_ModifiedDate_Yeardrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[15]/ul/li[3]/a")
	public static WebElement  Finance_Purchases_ModifiedDate_monthdrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[15]/ul/li[4]/a")
	public static WebElement  Finance_Purchases_ModifiedDate_DayofYeardrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[15]/ul/li[5]/a")
	public static WebElement  Finance_Purchases_ModifiedDate_Daydrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[15]/ul/li[6]/a")
	public static WebElement  Finance_Purchases_ModifiedDate_Weekdrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[15]/ul/li[7]/a")
	public static WebElement  Finance_Purchases_ModifiedDate_WeekDaydrapanddrop;


	@FindBy(xpath="//span[contains(text(),'Modified time')]//i[@class='icon-expand icon-font8']")
	public static WebElement  Finance_Purchases_Header_ModifiedTimeExpandBtn;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[16]/ul/li[1]/a")
	public static WebElement  Finance_Purchases_ModifiedTime_ModifiedTimedrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[16]/ul/li[2]/a")
	public static WebElement  Finance_Purchases_ModifiedTime_Hourdrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[16]/ul/li[3]/a")
	public static WebElement  Finance_Purchases_ModifiedTime_Minutedrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[16]/ul/li[4]/a")
	public static WebElement  Finance_Purchases_ModifiedTime_Seconddrapanddrop;


	@FindBy(xpath="//span[contains(text(),'Print count')]")
	public static WebElement  Finance_Purchases_Header_PrintCountdrapanddrop;	

	@FindBy(xpath="//span[contains(text(),'Revision number')]")
	public static WebElement  Finance_Purchases_Header_RevisionNumberdrapanddrop;	

	@FindBy(xpath="//span[contains(text(),'Authorize status')]")
	public static WebElement  Finance_Purchases_Header_AuthorizeStatusdrapanddrop;

	@FindBy(xpath="//span[contains(text(),'Authorize date')]//i[@class='icon-expand icon-font8']")
	public static WebElement  Finance_Purchases_Header_AuthorizeDateExpandBn;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[20]/ul/li[1]/a")
	public static WebElement  Finance_Purchases_AuthorizeDate_AuthorizeDatedrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[20]/ul/li[2]/a")
	public static WebElement  Finance_Purchases_AuthorizeDate_Yeardrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[20]/ul/li[3]/a")
	public static WebElement  Finance_Purchases_AuthorizeDate_Monthdrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[20]/ul/li[4]/a")
	public static WebElement  Finance_Purchases_AuthorizeDate_DayofYeardrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[20]/ul/li[5]/a")
	public static WebElement  Finance_Purchases_AuthorizedDate_Daydrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[20]/ul/li[6]/a")
	public static WebElement  Finance_Purchases_AuthorizeDate_Weekdrapanddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[20]/ul/li[7]/a")
	public static WebElement  Finance_Purchases_AuthorizeDate_WeekDaydrapanddrop;

	@FindBy(xpath="//span[contains(text(),'Email count')]")
	public static WebElement  Finance_Purchases_Header_EMailCountdrapanddrop;

	@FindBy(xpath="//i[@class='icon-collepse icon-font8']")
	public static WebElement  FieldsMinimizeBtn;

	////// Finance_Purchases ****Body****
	// Named as PSS_Body {Common for Finance purchase,sales***Inventory Purchase,sales,stocks}

	@FindBy(xpath="//span[contains(text(),'Item')]//i[@class='icon-expand icon-font8']")
	public static WebElement  PSS_Body_ItemtExpansion;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[1]/a")
	public static WebElement  PSS_Item_Namedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[2]/a")
	public static WebElement  PSS_Item_Codedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[3]/a")
	public static WebElement  PSS_Item_Budgetdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[4]/a")
	public static WebElement  PSS_Item_Aliasdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[5]/a")
	public static WebElement  PSS_Item_ABCcodeCarryingCostdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[6]/a")
	public static WebElement  PSS_Item_ABCcodeMargindraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[7]/a")
	public static WebElement  PSS_Item_ABCcodeRevenuedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[8]/a")
	public static WebElement  PSS_Item_ABCcodeValuedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[9]/a")
	public static WebElement  PSS_Item_AlternateCategorydraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[10]/a")
	public static WebElement  PSS_Item_AlternateCategoryCodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[11]/a")
	public static WebElement  PSS_Item_BarCode1draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[12]/a")
	public static WebElement  PSS_Item_BarCode2draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[13]/a")
	public static WebElement  PSS_Item_BarCode3draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[14]/a")
	public static WebElement  PSS_Item_Bindraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[15]/a")
	public static WebElement  PSS_Item_BinCodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[16]/a")
	public static WebElement  PSS_Item_BinCapacitydraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[17]/a")
	public static WebElement  PSS_Item_BOMdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[18]/a")
	public static WebElement  PSS_Item_Categorydraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[19]/a")
	public static WebElement  PSS_Item_CategoryCodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[20]/a")
	public static WebElement  PSS_Item_CBMdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[21]/a")
	public static WebElement  PSS_CostOfExcessStockACItem_draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[22]/a")
	public static WebElement  PSS_Item_CostOfExcessStockACcodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[23]/a")
	public static WebElement  PSS_Item_CostOfIssueAccountdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[24]/a")
	public static WebElement  PSS_Item_CostOfIssueAccountCodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[25]/a")
	public static WebElement  PSS_Item_CostOfSalesReturnACdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[26]/a")
	public static WebElement  PSS_Item_CostOfSalesReturnACcodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[27]/a")
	public static WebElement  PSS_Item_CostOfShortageStockACdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[28]/a")
	public static WebElement  PSS_Item_CostOfShortageStockACcodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[29]/a")
	public static WebElement  PSS_Item_DefaultBaseUnitdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[30]/a")
	public static WebElement  PSS_Item_DefaultBaseUnitCodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[31]/a")
	public static WebElement  PSS_Item_DefaultPurchaseUnitdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[32]/a")
	public static WebElement  PSS_Item_DefaultPurchaseUnitcodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[33]/a")
	public static WebElement  PSS_Item_DefaultReplenishmentdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[34]/a")
	public static WebElement  PSS_Item_DefaultSalesUnitdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[35]/a")
	public static WebElement  PSS_Item_DefaultSalesUnitcodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[36]/a")
	public static WebElement  PSS_Item_Descriptiondraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[37]/a")
	public static WebElement  PSS_Item_DontShowProductsExpiringInDaysdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[38]/a")
	public static WebElement  PSS_Item_Flushingdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[39]/a")
	public static WebElement  PSS_Item_Heightdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[40]/a")
	public static WebElement  PSS_Item_IndirectCostdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[41]/a")
	public static WebElement  PSS_Item_ItemMakedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[42]/a")
	public static WebElement  PSS_Item_ItemTypedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[43]/a")
	public static WebElement  PSS_Item_Lengthdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[44]/a")
	public static WebElement  PSS_Item_ManufacturePolicydraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[45]/a")
	public static WebElement  PSS_Item_OverHeadCostdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[46]/a")
	public static WebElement  PSS_Item_PackingBOMdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[47]/a")
	public static WebElement  PSS_Item_PerishableItemdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[48]/a")
	public static WebElement  PSS_Item_Printerdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[49]/a")
	public static WebElement  PSS_Item_Profitdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[50]/a")
	public static WebElement  PSS_Item_PurchaseVarianceACdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[51]/a")
	public static WebElement  PSS_Item_PurchaseVarianceACcodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[52]/a")
	public static WebElement  PSS_Item_ReorderLeveldraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[53]/a")
	public static WebElement  PSS_Item_SalesAccountdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[54]/a")
	public static WebElement  PSS_Item_SalesAccountCodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[55]/a")
	public static WebElement  PSS_Item_StandardCostdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[56]/a")
	public static WebElement  PSS_Item_StocksAccountdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[57]/a")
	public static WebElement  PSS_Item_StocksAccountCodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[58]/a")
	public static WebElement  PSS_Item_TaxCodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[59]/a")
	public static WebElement  PSS_Item_TaxCodecodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[60]/a")
	public static WebElement  PSS_Item_ValuationMethoddraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[61]/a")
	public static WebElement  PSS_Item_Weightdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[62]/a")
	public static WebElement  PSS_Item_Widthdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[63]/a")
	public static WebElement  PSS_Item_WIPaccountdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[64]/a")
	public static WebElement  PSS_Item_WIPaccountCodedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[65]/a")
	public static WebElement  PSS_Item_GroupLevel1draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[66]/a")
	public static WebElement  PSS_Item_GroupLevel1Codedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[67]/a")
	public static WebElement  PSS_Item_Buyingratedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[68]/a")
	public static WebElement  PSS_Item_BuyingVAl1draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[69]/a")
	public static WebElement  PSS_Item_BuyingVal2draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[70]/a")
	public static WebElement  PSS_Item_BuyingVal3draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[71]/a")
	public static WebElement  PSS_Item_BuyingVal4draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[72]/a")
	public static WebElement  PSS_Item_BuyingVal5draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[73]/a")
	public static WebElement  PSS_Item_BuyingVal6draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[74]/a")
	public static WebElement  PSS_Item_BuyingVal7draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[75]/a")
	public static WebElement  PSS_Item_BuyingVal8draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[76]/a")
	public static WebElement  PSS_Item_BuyingVal9draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[77]/a")
	public static WebElement  PSS_Item_BuyingVal10draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[78]/a")
	public static WebElement  PSS_Item_BuyingVal11draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[79]/a")
	public static WebElement  PSS_Item_BuyingVal12draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[80]/a")
	public static WebElement  PSS_Item_BuyingVal13draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[81]/a")
	public static WebElement  PSS_Item_SellingRatedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[82]/a")
	public static WebElement  PSS_Item_SellingVal1draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[83]/a")
	public static WebElement  PSS_Item_SellingVal2draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[84]/a")
	public static WebElement  PSS_Item_SellingVal3draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[85]/a")
	public static WebElement  PSS_Item_SellingVal4draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[86]/a")
	public static WebElement  PSS_Item_SellingVal5draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[87]/a")
	public static WebElement  PSS_Item_SellingVal6draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[88]/a")
	public static WebElement  PSS_Item_SellingVal7draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[89]/a")
	public static WebElement  PSS_Item_SellingVal8draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[90]/a")
	public static WebElement  PSS_Item_SellingVal9draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[91]/a")
	public static WebElement  PSS_Item_SellingVal10draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[92]/a")
	public static WebElement  PSS_Item_SellingVal11draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[93]/a")
	public static WebElement  PSS_Item_SellingVal12draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[1]/ul/li[94]/a")
	public static WebElement  PSS_Item_SellingVal13draganddrop;


	@FindBy(xpath="//span[contains(text(),'Unit')]//i[@class='icon-expand icon-font8']")
	public static WebElement  PSS_Body_UnitExpandBtn;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[3]/a")
	public static WebElement  PSS_Body_RDdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[4]/a")
	public static WebElement  PSS_Body_AvgRate0draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[5]/a")
	public static WebElement  PSS_Body_Quantitydraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[6]/a")
	public static WebElement  PSS_Body_LPurchaseOrdersdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[7]/a")
	public static WebElement  PSS_Body_Ratedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[8]/a")
	public static WebElement  PSS_Body_Grossdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[9]/a")
	public static WebElement  PSS_Body_Batchdraganddrop;

	@FindBy(xpath="//li[@class='treeview active']//span[contains(text(),'Bins')]")
	public static WebElement  PSS_Body_BinsExpandBtn;

	@FindBy(xpath="//span[contains(text(),'ExpDate')]")
	public static WebElement  PSS_Body_ExpDatedraganddrop;

	@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'RMA')]")
	public static WebElement  PSS_Body_RMAdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[13]/a")
	public static WebElement  PSS_Body_Netdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[14]/a")
	public static WebElement  PSS_Body_NetBaseCurrencydraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/ul/li[15]/a")
	public static WebElement  PSS_Body_NetLocalCurrencydraganddrop;

	@FindBy(xpath="//span[contains(text(),'RD Value')]")
	public static WebElement  PSS_Body_RDValuedraganddrop;

	@FindBy(xpath="//span[contains(text(),'Avg Rate(O) Value')]")
	public static WebElement  PSS_Body_AvgRate0Valuedraganddrop;

	@FindBy(xpath="//input[@id='id_body_33554514']")
	public static WebElement enter_salrt0Txt;

	@FindBy(xpath="//input[@id='id_body_33554515']")
	public static WebElement enter_salrt1Txt;

	@FindBy(xpath="//input[@id='id_body_33554516']")
	public static WebElement enter_salrtGrp0Txt;

	@FindBy(xpath="//input[@id='id_body_33554517']")
	public static WebElement enter_salrtGrp1Txt;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_table_body']/tr[1]/td")
	public static List<WebElement> voucherBodyRow1List;

	@FindBy(xpath="//*[@id='LiDesignField']")
	public static WebElement  FieldsTab;

	@FindBy(xpath="//*[@id='LiProp']")
	public static WebElement   PropertiesTab;

	@FindBy(xpath="//span[contains(text(),'Header')]")
	public static WebElement  Finance_Purchases_HeaderExpandBtn;

	@FindBy(xpath="//span[contains(text(),'Body')]//i[@class='icon-expand icon-font8']")
	public static WebElement  PSS_BodyExpandbtn;

	@FindBy(xpath="//span[contains(text(),'User details')]//i[@class='icon-expand icon-font8']")
	public static WebElement  UserDetailsExpandbtn;

	@FindBy(xpath="//span[contains(text(),'Miscellaneous')]//i[@class='icon-expand icon-font8']")
	public static WebElement MiscellaneousExpandBtn;

	@FindBy(xpath="//span[contains(text(),'Company')]//i[@class='icon-expand icon-font8']")
	public static WebElement CompanyExpandBtn;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[1]/a")
	public static WebElement  UserDetails_UserNamedraganddrop;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[2]/a")
	public static WebElement  UserDetails_MobileNumberdraganddrop;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[3]/a")
	public static WebElement  UserDetails_PhoneNumberdraganddrop;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[4]/a")
	public static WebElement  UserDetails_EmailIddraganddrop;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[5]/a")
	public static WebElement  UserDetails_Signaturedraganddrop;	

	@FindBy(xpath="//span[contains(text(),'Employee')]//i[@class='icon-expand icon-font8']")
	public static WebElement  UserDetails_EmployeeExpandBtn;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[7]/a")
	public static WebElement  UserDetails_AuthoriseUserNamedraganddrop;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[8]/a")
	public static WebElement  UserDetails_AuthoriseUserMobileNumberdraganddrop;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[9]/a")
	public static WebElement  UserDetails_AuthoriseUserPhoneNumberdraganddrop;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[10]/a")
	public static WebElement  UserDetails_AuthoriseUserEmailIddraganddrop;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[11]/a")
	public static WebElement  UserDetails_AuthoriseUserSignaturedraganddrop;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[12]/a")
	public static WebElement  UserDetails_AuthoriseBydraganddrop;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[13]/a")
	public static WebElement  UserDetails_NextUserNamedraganddrop;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[14]/a")
	public static WebElement  UserDetails_NextUserMobileNumberdraganddrop;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[15]/a")
	public static WebElement  UserDetails_NextUserPhoneNumberdraganddrop;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[16]/a")
	public static WebElement  UserDetails_NextUserEmailIddraganddrop;	

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[17]/a")
	public static WebElement  UserDetails_NextUserSignaturedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[18]/a")
	public static WebElement  UserDetails_CreatedBySignaturedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[3]/ul/li[19]/a")
	public static WebElement  UserDetails_ModifiedBySignaturedraganddrop;

	@FindBy(xpath="//span[contains(text(),'Login name')]")
	public static WebElement  Miscellaneous_LoginNamedraganddrop;

	@FindBy(xpath="//span[contains(text(),'Page number')]")
	public static WebElement  Miscellaneous_PageNumberdraganddrop;

	@FindBy(xpath="//span[contains(text(),'Print date')]")
	public static WebElement  Miscellaneous_PrintDatedraganddrop;

	@FindBy(xpath="//span[contains(text(),'Print time')]")
	public static WebElement  Miscellaneous_PrintTimedraganddrop;

	@FindBy(xpath="//span[contains(text(),'Total page')]")
	public static WebElement  Miscellaneous_TotalPagedraganddrop;

	@FindBy(xpath="//span[contains(text(),'Sr no')]")
	public static WebElement  Miscellaneous_SrNodraganddrop;

	@FindBy(xpath="//span[contains(text(),'Page type')]")
	public static WebElement  Miscellaneous_PageTypedraganddrop;

	@FindBy(xpath="//span[contains(text(),'Ageing month-wise')]")
	public static WebElement  Miscellaneous_AgeingMonthWisedraganddrop;

	@FindBy(xpath="//span[contains(text(),'Ageing balance')]")
	public static WebElement  Miscellaneous_AgeingBalancedraganddrop;

	@FindBy(xpath="//span[contains(text(),'Ageing slab1')]")
	public static WebElement  Miscellaneous_AgeingSlab1draganddrop;

	@FindBy(xpath="//span[contains(text(),'Ageing slab2')]")
	public static WebElement  Miscellaneous_AgeingSlab2draganddrop;

	@FindBy(xpath="//span[contains(text(),'Ageing slab3')]")
	public static WebElement  Miscellaneous_AgeingSlab3draganddrop;

	@FindBy(xpath="//span[contains(text(),'Ageing slab4')]")
	public static WebElement  Miscellaneous_AgeingSlab4draganddrop;

	@FindBy(xpath="//span[contains(text(),'Ageing slab5')]")
	public static WebElement  Miscellaneous_AgeingSlab5draganddrop;

	@FindBy(xpath="//span[contains(text(),'Ageing slab6')]")
	public static WebElement  Miscellaneous_AgeingSlab6draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[4]/ul/li[16]/a")
	public static WebElement  Miscellaneous_AgeingSlab7draganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[4]/ul/li[17]/a")
	public static WebElement  Miscellaneous_LastSlabdraganddrop;

	@FindBy(xpath="//span[contains(text(),'Company Logo')]")
	public static WebElement  Company_CompanyLogodraganddrop;

	@FindBy(xpath="//span[contains(text(),'Company Name')]")
	public static WebElement  Company_CompanyNamedraganddrop;

	@FindBy(xpath="//span[contains(text(),'Accounting Date')]//i[@class='icon-expand icon-font8']")
	public static WebElement  Comapny_AccountingDateExpandbtn;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[5]/ul/li[3]/ul/li[1]/a")
	public static WebElement  AccountingDate_AccountingDatedraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[5]/ul/li[3]/ul/li[2]/a")
	public static WebElement  AccountingDate_Yeardraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[5]/ul/li[3]/ul/li[3]/a")
	public static WebElement  AccountingDate_Monthdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[5]/ul/li[3]/ul/li[4]/a")
	public static WebElement  AccountingDate_DayOfYeardraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[5]/ul/li[3]/ul/li[5]/a")
	public static WebElement  AccountingDate_Daydraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[5]/ul/li[3]/ul/li[6]/a")
	public static WebElement  AccountingDate_Weekdraganddrop;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[5]/ul/li[3]/ul/li[7]/a")
	public static WebElement  AccountingDate_WeekDaydraganddrop;

	@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Country')]")
	public static WebElement  Company_Countrydraganddrop;

	@FindBy(xpath="//li[@class='treeview active']//span[contains(text(),'Currency')]")
	public static WebElement  Company_Currencydraganddrop;

	@FindBy(xpath="//span[contains(text(),'DefaultLang')]")
	public static WebElement  Company_DefaultLangdraganddrop;

	@FindBy(xpath="//span[contains(text(),'SupportedLang')]")
	public static WebElement  Company_SupportedLangdraganddrop;

	@FindBy(xpath="//span[contains(text(),'DocNo')]")
	public static WebElement  Finance_Purchases_Header_DocNOdraganddrop;

	@FindBy(xpath="//span[contains(text(),'Header')]")
	public static WebElement  clickonHeader;

	@FindBy(xpath="//div[@id='dvCreateViewMain']/div/div[2]/span/i")
	public static WebElement headerPlus;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[1]/ul/li[1]/a/span/text()")
	public static WebElement clickOnDoc;

	@FindBy(xpath="//div[@id='InvoiceDesignFields']/ul/li/ul/li/a/span")
	public static WebElement docNo;

	@FindBy(xpath="//*[@id='id_LayoutDiv_0")
	public static WebElement docNoo;

	@FindBy(xpath="//div[@id='StaticText_0_2002']//p[contains(text(),'DocNo')]")
	public static WebElement draggedDocNo;

	@FindBy(xpath="//*[@id='InvoiceDesignFields']/ul/li[2]/a/span")
	public static WebElement  clickOnBody;

	@FindBy(xpath="//*[@id='id_PopuplayoutName']")
	public static WebElement  layoutNameTextField;

	@FindBy(xpath="//*[@id='id_layoutName']")
	public static WebElement  SavedlayoutNameTitle;

	@FindBy(xpath="//*[@id='BodyColumn_2000_0']")
	public static WebElement warehouseTxtField;

	@FindBy(xpath="//*[@id='BodyColumn_2000_1']")
	public static WebElement itemTxtField;

	@FindBy(xpath="//*[@id='BodyColumn_2000_2']")
	public static WebElement taxCodeTxtField;

	@FindBy(xpath="//*[@id='BodyColumn_2000_3']")
	public static WebElement purchaseAccountTxtfiled;

	@FindBy(xpath="//*[@id='BodyColumn_2000_4']")
	public static WebElement unitTxtField;

	@FindBy(xpath="//*[@id='BodyColumn_2000_5']")
	public static WebElement quantityTxtField;

	@FindBy(xpath="//*[@id='BodyColumn_2000_6']")
	public static WebElement rateTxtField;

	@FindBy(xpath="//*[@id='BodyColumn_2000_7']")
	public static WebElement grossTxtField;

	@FindBy(xpath="//*[@id='BodyColumn_2000_8']")
	public static WebElement vatTxtField;

	@FindBy(xpath="/html/body/ul/li[5]/span")
	public static WebElement PropertiesBtn;

	@FindBy(xpath="//*[@id='FTextBox_17_2000']")
	public static WebElement widthTxt2;

	@FindBy(xpath="//*[@id='Combobox_2_2000']")
	public static WebElement Txtalignment;

	@FindBy(xpath="//*[@id='BodyGrid_0_2000']")
	public static WebElement bodyGrid ;

	@FindBy(xpath="//*[@id='FTextBox_25_2000']")
	public static WebElement bodyGridTopTextFiled;

	@FindBy(xpath="//*[@id='opt_LayoutID_input_image']/span")
	public static WebElement layoutTab;

	@FindBy(xpath="//tbody[@id='opt_LayoutID_table_body']/tr/td[2]")
	public static List<WebElement> designLayoutList;

	@FindBy(xpath="//p[text()='ExchangeRate']")
	public static WebElement exchangeRateFiled;

	@FindBy(xpath="//p[text()='LocExchangeRate']")
	public static WebElement locexchangeRateFiled;

	@FindBy(xpath="//table[@id='TableGrid_0_2000']/thead/tr/th[text()='Discount']")
	public static WebElement discountFiledInbody;

	@FindBy(xpath="(//*[@id='navbarSupportedContent2'])[3]/ul/li[3]/div/a/i")
	public static WebElement  Opendropdown;

	@FindBy(xpath="(//*[@id='navbarSupportedContent2'])[3]/ul/li[3]/div/div/a[text()='Open']")
	public static WebElement  OpenBtn;

	@FindBy(xpath="(//*[@id='navbarSupportedContent2'])[3]/ul/li[3]/div/div/a[text()='from XML']")
	public static WebElement  fromXMLBtn;

	@FindBy(xpath="(//*[@id='navbarSupportedContent2'])[3]/ul/li[4]/div/a/i")
	public static WebElement  Savedropdown;

	@FindBy(xpath="(//*[@id='navbarSupportedContent2'])[3]/ul/li[4]/div/div/a[text()='Save']")
	public static WebElement  PrintSaveBttn;

	@FindBy(xpath="//input[@id='id_PopuplayoutName']")
	public static WebElement  SaveTxtField;

	@FindBy(xpath="//button[@id='btn_PopUpOk']")
	public static WebElement  Save_OkBtn;

	@FindBy(xpath="//*[@id='id_PopUp']/div/div/div[3]/button[2]")
	public static WebElement  Save_CancekBtn;

	@FindBy(xpath="//*[@id='id_InvoiceDesingCancel']/i")
	public static WebElement  ExitBtn;	

	//User Name and password in Credentials
	@FindBy(xpath="//*[@id='txtAdminUsername']")
	public static WebElement adminName;

	@FindBy(xpath="//*[@id='txtAdminPassword']")
	public static WebElement adminPassword;

	@FindBy(xpath="/html/body/section/div[8]/div/div/div[4]/button[1]")
	public static WebElement loginButton;

	@FindBy(xpath="/html/body/section/div[8]/div/div/div[4]/button[2]")
	public static WebElement cancelButton;

	//Restore Screen Options
	@FindBy(xpath="//div[contains(text(),'Restore Company')]")
	public static WebElement restoreCompanyBtnOfLoginPage;

	@FindBy(xpath="//input[@id='txtRestoreFile']")
	public static WebElement restoreInputTxt;

	@FindBy(xpath="//button[@class='btn Fbutton']")
	public static WebElement restoreBrowseBtn;

	@FindBy(xpath="//*[@id='RestoreCompanyModal']/div/div/div[2]/button[1]")
	public static WebElement restoreOverwriteYesBtn;

	@FindBy (xpath="//tbody[@id='Test_body']/tr[1]/td[3]")
	public static WebElement DocumentSet_Select1stRowFilter;

	@FindBy (xpath="//tbody[@id='Test_body']/tr[2]/td[3]")
	public static WebElement DocumentSet_Select2ndRowFilter;

	@FindBy (xpath="//tbody[@id='Test_body']/tr[1]/td[5]")
	public static WebElement DocumentSet_Select1stRowLayout;

	@FindBy (xpath="//tbody[@id='Test_body']/tr[2]/td[5]")
	public static WebElement DocumentSet_Select2ndRowLayout;

	@FindBy (xpath="//input[@id='btnFilterId']")
	public static WebElement DocumentSet_EnterFilter;

	@FindBy (xpath="//input[@id='opt_LayoutId']")
	public static WebElement DocumentSet_EnterLayout;

	@FindBy (xpath="//div[@id='dv_DocSetFilterMain']/div[1]/span[2]")
	public static WebElement DocumentSet_AdvanceFilterBtn;

	@FindBy (xpath="(//div[@id='0_2_AdvanceFilter_']/table/tbody/tr/td/select)[1]")
	public static WebElement DocumentSet_AdvanceFilterConjunctionDrpdwn;

	@FindBy (xpath="(//div[@id='0_2_AdvanceFilter_']/table/tbody/tr/td/select)[2]")
	public static WebElement DocumentSet_AdvanceFilterSelectOperatorDrpdwn;

	@FindBy (xpath="(//div[@id='0_2_AdvanceFilter_']/table/tbody/tr/td/select)[3]")
	public static WebElement DocumentSet_AdvanceFilterCompareWithDrpdwn;

	@FindBy (xpath="(//div[@id='0_2_AdvanceFilter_']/table/tbody/tr/td/input)[1]")
	public static WebElement DocumentSet_AdvanceFilterSelectFieldTxt;

	@FindBy (xpath="//*[@id='filterTree_0_2_AdvanceFilter_']/ul/a[4]")
	public static WebElement DocumentSet_AdvanceFilterSelectFieldVendorAC;

	@FindBy (xpath="//input[@id='advancefilter_master_0_2_']")
	public static WebElement DocumentSet_AdvanceFilterValueTxt;

	@FindBy (xpath="//button[@id='btnSetFilterVal']")
	public static WebElement DocumentSet_AdvanceFilterOkBtn;

	@FindBy (xpath="//*[@id='dv_DocSetFilter']/div/div/div[3]/div/button[2]")
	public static WebElement DocumentSet_AdvanceFilterCloseBtn;

	@FindBy (xpath="//*[@id='0_2_AdvanceFilter_']/table/tbody/tr/td[5]/input")
	public static WebElement DocumentSet_AdvanceFilterValueCol;

	@FindBy(xpath="//a[@id='2014']")
	public static WebElement  financialTransactionSalesMenu;

	@FindBy(xpath="//span[contains(text(),'Sales invoice VAT')]")
	public static WebElement  salesInvoiceVATVoucher;


	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[1]")
	public static WebElement  firstRowIndex;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[1]/span")
	public static WebElement  secondRowIndex;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[1]")
	public static WebElement  thirdRowIndex;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[1]")
	public static WebElement  fourthRowIndex;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[1]")
	public static WebElement  fifthRowIndex;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[1]")
	public static WebElement  sixthRowIndex;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[1]")
	public static WebElement  seventhRowIndex;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[1]")
	public static WebElement  eigthRowIndex;

	@FindBy(xpath="//*[@id='id_transentry_body_menu']/a[1]/label")
	public static WebElement  deleteRowBtn;


	@FindBy(xpath="//div[@id='id_transentry_body_menu']/a[3]/label")
	public static WebElement  selectRowBtn;

	@FindBy(xpath="//span[@class='icon-clone icon-font8']")
	public static WebElement  duplicateRowBtn;

	@FindBy(xpath="//input[@id='id_header_268435470']")
	public static WebElement  salesInvoiceVATPlaceOFSupply;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr")
	public static List<WebElement>  LinkRowCount;

	@FindBy(xpath="//input[@id='id_body_33554476']")
	public static WebElement enter_AQTxt;

	@FindBy(xpath="//input[@id='id_body_33554477']")
	public static WebElement enter_FQTxt;

	@FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr/td[5]")
	public static List<WebElement> voucherGridDocNo;

	@FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr/td[11]")
	public static List<WebElement> voucherGridSuspendStatus;

	@FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr/td[2]/input")
	public static List<WebElement> voucherGridIndexChkBox;

	@FindBy(xpath="//a[contains(text(),'Pending Sales Orders')]")
	public static WebElement  homepagePendingSalesOrders;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[21]")
	public static WebElement  select2ndRow_20thColumn;

	@FindBy(xpath="//input[@id='id_limit']")
	public static WebElement gridEnterAdjustAmtRow1;

	@FindBy(xpath="//*[@id='id_transactionentry_copydocument']")
	public static WebElement  copyDocumentOption;

	@FindBy(xpath="//tr[@id='tr_copydoc_voucher_1']//td//input")
	public static WebElement cd_FirstChkBox;

	@FindBy(xpath="//input[@id='id_transaction_copydoc_ok']")
	public static WebElement cd_OkBtn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr/td")
	public static List<WebElement> voucherGridBodyList;



	@FindBy(xpath="//*[@id='tblBodyTransRender']/tr/td[5]")
	public static List<WebElement> bodyVoucherNolist;

	@FindBy(xpath="//*[@id='tblBodyTransRender']/tr/td[2]/input")
	public static List<WebElement> bodyChkBoxlist;

	@FindBy(xpath="//li[@id='id_transactionentry_delete']/a/span")
	public static WebElement  voucherDeleteBtn;

	@FindBy(xpath="//input[@id='id_header_4']")
	public static WebElement  customerAccountTxt;

	@FindBy(xpath="//tbody[@id='id_header_3_table_body']/tr/td[2]")
	public static List<WebElement> salesAccountListCount; 

	/*@FindBy(xpath="//tbody[@id='id_header_4_table_body']/tr/td[2]")
	public static List<WebElement> customerAccountListCount;*/ 

	@FindBy(xpath="//tbody[@id='id_header_4_table_body']/tr/td[2]")
	public static List<WebElement> cnAccountListCount; 

	/*@FindBy(xpath="//tbody[@id='id_body_536870916_table_body']/tr/td[2]")
	public static List<WebElement> pvwareHouseListCount;*/

	@FindBy(xpath="//button[@id='div_PICKFIFO'][1]")
	public static WebElement batchPickOnFIFOIcon;

	@FindBy(xpath="//button[@id='div_Ok1']")
	public static WebElement batchOkIcon;

	@FindBy (xpath="(//button[@id='div_OK'])[2]")
	public static WebElement searchRMAOkBtn;

	@FindBy (xpath="//*[@id='SearchModal_Content']/div[3]/div/div[2]/label")
	public static WebElement searchRMACancelBtn;

	@FindBy (xpath="//*[@id='SearchModal_Content']/div[1]/div[2]/span/i")
	public static WebElement searchRMACloseBtn;

	@FindBy(xpath="//*[@id='RMASearchTable']/tbody/tr[1]")
	public static WebElement rmaSearchTableBodyIfBlankRows;

	@FindBy (xpath="//input[@id='Search_txtSearchBox']")
	public static WebElement searchRMASearchTxtField;

	@FindBy (xpath="//i[@id='btn_RMASsearch_Search']")
	public static WebElement searchRMASearchBtn;

	@FindBy (xpath="//input[@id='Search_txtQuantity']")
	public static WebElement searchRMAQuantityTxtField;

	@FindBy (xpath="//label[@id='Search_lblItems']")
	public static WebElement searchRMANoOfItemsTxt;

	@FindBy (xpath="(//button[@id='div_OK'])[1]")
	public static WebElement rmaoutwardOkBtn;

	@FindBy (xpath="//label[contains(text(),'Cancel')]")
	public static WebElement rmaoutwardCancelBtn;

	@FindBy (xpath="//*[@id='sp_search']")
	public static WebElement rmaoutwardExpandBtn;

	@FindBy(xpath = "//i[@id='sp_search']")
	public static WebElement RMAPopupSerialNoExpansionBtn;

	@FindBy(xpath = "//input[@id='chkSelectALL']")
	public static WebElement searchRMASelectAllChkBox;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[22]")
	public static WebElement  select2ndRow_21stColumn;

	@FindBy(xpath="//input[@id='id_EmailFromCust']")
	public static WebElement  printScreenPickEmailChkboxIsSelected;
	
	@FindBy(xpath="//input[@id='id_EmailFromCust']/following-sibling::span")
	public static WebElement  printScreenPickEmailChkbox;

	@FindBy(xpath="//input[@id='id_EmailTo']")
	public static WebElement  printScreenSentEmailTxt;

	@FindBy(xpath="//input[@id='id_EmailCC']")
	public static WebElement  printScreenCCTxt;

	@FindBy(xpath="//input[@id='id_EmailSubject']")
	public static WebElement  printSCreenSubjectTxt;

	@FindBy(xpath="//select[@id='cmd_LayoutOptions']")
	public static WebElement  printSCreenLayoutOption;

	@FindBy(xpath="//input[@id='btnPrint']")
	public static WebElement  emailSendBtn;

	@FindBy(xpath="//span[@id='id_InvoiceDesingCancel']")
	public static WebElement  printCLoseBtn;

	@FindBy(xpath="//input[@id='opt_LayoutID']")
	public static WebElement  printlayoutTxt;

	@FindBy(xpath="//input[@id='OutgoingPortNo']")
	public static WebElement  outgoingPortTxt;

	@FindBy(xpath="//input[@id='SSLSocketPort']")
	public static WebElement  sslPortTXt;

	@FindBy(xpath="//input[@id='EmailId']")
	public static WebElement  MailSettingEmailTxt;

	@FindBy(xpath="//input[@id='Password']")
	public static WebElement  MailSettingPasswordTxt;

	@FindBy(xpath="//div[text()='Attachments area']/parent::div/div[4]//span/div/div[1]")
	public static WebElement mailAttachmentDownloadBtn;

	@FindBy(xpath="//*[@id=':4']/div[2]/div[1]/div/div[2]/div[3]")
	public static WebElement mailDeleteButton;

	@FindBy(xpath="//*[@id='gb']/div[2]/div[3]/div[1]/div[2]/div/a/img")
	public static WebElement gmailUserBtn;

	@FindBy(xpath="//*[@id='gb_71']")
	public static WebElement gmailSignOutBtn;


	@FindBy(xpath="//div[contains(text(),'Remove an account')]")
	public static WebElement removeAccountBtn;

	@FindBy(xpath="//body/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/form[1]/span[1]/section[1]/div[1]/div[1]/div[1]/div[1]/ul[1]/li[1]/div[1]/div[2]/*[1]")
	public static WebElement removeDeleteBtn;


	@FindBy(xpath="//*[@id='yDmH0d']/div[5]/div/div[2]/div[3]/div[1]/span/span")
	public static WebElement yesRemoveBtn;

	@FindBy(xpath="//*[@id=':1']/div/div/div[8]/div/div[1]/div[3]/div/table/tbody/tr/td[4]/div[2]/span/span")
	public static List<WebElement> row1mailFromList;

	@FindBy(xpath="//input[@id='identifierId']")
	public static WebElement  userNameTxt;

	@FindBy(xpath="//span[contains(text(),'Next')]")
	public static WebElement NextBtn;

	@FindBy(xpath="//input[@name='password']")
	public static WebElement  PasswordTxt;

	@FindBy(xpath="//div[@id=':2i']")
	public static WebElement  row1;

	@FindBy(xpath="//span[@class='y2']")
	public static WebElement mailIDBody ;

	@FindBy(xpath="//*[@id=':2f']")
	public static WebElement mailIDBodyASBody;

	@FindBy(xpath="//img[@class='CToWUd a6T']")
	public static WebElement mailBodyTxt;


	/*@FindBy (xpath="//input[@id='idEraseTransaction']")
	public static WebElement eraseTranscationsRadio;

	@FindBy (xpath="//div[@id='navbarSupportedContent2']/ul/li[1]/a/i")
	public static WebElement eraseAllOkBtn;

	@FindBy (xpath="//div[@id='navbarSupportedContent2']/ul/li[2]/a/i")
	public static WebElement eraseAllCancelBtn;

*/
	// Master Preload

	@FindBy(xpath="//span[@class='font-5'][contains(text(),'Main')]")
	public static WebElement generalMainTab;

	//@FindBy(xpath="//div[@id='tabId_2']/div/ul/li[2]")
	@FindBy(xpath="//*[@id='tab_HeaderDetails_0']")
	public static WebElement generalHeaderDetailsTab;

	//@FindBy(xpath="//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Save')]")
	@FindBy(xpath="//*[@id='all_TabsField']/div[9]/nav/div/div/ul/li[1]/a/i")
	public static WebElement accountSaveBtn;

	//@FindBy(xpath="//*[@id='tabId_2']/div[1]/a[1]/i")
	@FindBy(xpath="//*[@id='btnAdd_Field_MasterCust']/a/i")
	public static WebElement accountAddBtn;

	@FindBy(xpath="//*[@id='tabId_2']/div[1]/a[5]/i")
	public static WebElement accountCloseBtn;

	@FindBy(xpath="//input[@id='ExtraField_FieldCaption']")
	public static WebElement extraFields_FieldDetailsCaption;   

	@FindBy(xpath="//div[@id='tab_FieldDetails_ExtraField']//div[1]//div[2]//div[1]//input[1]")
	public static WebElement extraFields_FieldDetailsName;  

	@FindBy(xpath="//select[@id='ddlDatatype_ExtraField']")
	public static WebElement extraFields_FieldDetailsDatTypeDropdown; 

	@FindBy (xpath="//input[@id='ExtraField_MasterToLink']")
	public static WebElement masterToLinkCombo;

	//@FindBy (xpath="//div[@id='extraFieldContainer_div']/div[1]/div[1]")
	
	@FindBy(xpath="//*[@id='btnSaveRule']/a/i")
	public static WebElement extraFieldOkBtn;

	@FindBy (xpath="(//*[@id='btnCloseMasterCutomizationFromMasterScreen']/a/i)[1]")
	public static WebElement customizeMasterCloseBtn;

	@FindBy(xpath="//div[@id='HeaderDetails_0']/div/div/table/tbody/tr/td[4]")
	public static List<WebElement> customizeMasterFieldCaptionList;

	@FindBy(xpath="//div[@id='HeaderDetails_0']/div/div/table/tbody/tr/td[6]")
	public static List<WebElement> customizeMasterVariableNameList;

	@FindBy(xpath="//input[@id='ItemCostCenter']")
	public static WebElement Item_CostCenterCombo;

	@FindBy (xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_MasterTolink']")
	public static WebElement editLayoutMasterToLinkTxt;

	@FindBy(xpath="//select[@id='inventoryOpt_ProductRate']")
	public static WebElement inventoryOptionsTabItemRateDrpdwn;

	@FindBy(xpath="//span[contains(text(),'Configure Transactions')]")
	public static WebElement  configureTransactionBtn;

	
	@FindBy(xpath="//a[@id='148']/span[contains(text(),'Preferences')]")
	public static WebElement  preferencesBtn;

	@FindBy(xpath="//a[@id='768']")
	public static WebElement  docCustPurchasesVoucher;

	@FindBy(xpath="//*[@id='doc_Searchbox']")
	public static WebElement  searchBox;

	@FindBy(xpath="//a[@id='60']")
	public static WebElement  finacinalsMenu;

	@FindBy(xpath="//a[@id='61']")
	public static WebElement  transactionsExpandBtn;

	@FindBy(xpath="//span[contains(text(),'Purchases Vouchers')]")
	public static WebElement  purchasesVoucherBtn;

	@FindBy(xpath="//div[@id='id_transaction_homescreen_edit']//span[@class='icon-edit icon-font6 ImagesinArabic']")
	public static WebElement  editBtn;

	@FindBy(xpath="//a[@id='btnTransHomePrint']/i")
	public static WebElement  printBtn;

	@FindBy(xpath="(//*[@id='navbarSupportedContent2']/ul/li[5])[1]")
	public static WebElement  authorizeBtn;

	@FindBy(xpath="//*[@id='id_transaction_homescreen_Suspend']/i")
	public static WebElement  suspendBtn;

	@FindBy(xpath="//span[@class='icon-reject2 icon-font6']")
	public static WebElement  rejectBtn;

	@FindBy(xpath="//span[@class='icon-cheque icon-font6']")
	public static WebElement  chequeReturnBtn;

	@FindBy(xpath="//span[@class='icon-printbarcode icon-font6']")
	public static WebElement  printBarCodeBtn;

	@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[12]")
	public static WebElement  settingsBtn;
	
	@FindBy(xpath="(//*[@id='navbarSupportedContent2']/ul/li[9]/a/i)[1]")
	public static WebElement  salesInvoiceSettingsBtn;
	
	@FindBy(xpath="//*[@id='panelsStayOpen-headingOne']/button")
	public static WebElement homePannelOpenBtn;
	
	@FindBy(xpath="//*[@id='btnMoreViews']")
	public static WebElement moreOptionsBtn;
	
	@FindBy(xpath = "//*[@id='btnCustomize']/i")
	public static WebElement customizeBtn;

	@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[13]/a/i")  
	public static WebElement  homeCloseBtn;

	@FindBy(xpath="//input[@id='id_header_1']")
	public static WebElement  documentNumberTxt;

	@FindBy(xpath="//input[@id='id_header_2']")
	public static WebElement  dateTxt;

	@FindBy(xpath="//input[@id='id_header_268435470']")
	public static WebElement purchaseVoucherVATPlaceOFSupply;

	/*@FindBy(xpath="//tbody[@id='id_header_268435470_table_body']/tr/td[2]")
	public static List<WebElement>  placeOFSupplyList;*/

	@FindBy(xpath="//input[@id='id_header_268435471']")
	public static WebElement  jurisdictionTxtt;


	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[22]")
	public static WebElement  select2ndRow_21thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[22]")
	public static WebElement  select3rdRow_21thColumn;


	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[26]")
	public static WebElement  select3rdRow_24thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[31]")
	public static WebElement  select1stRow_30thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[31]")
	public static WebElement  select2ndRow_30thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[31]")
	public static WebElement  select3rdRow_30thColumn;

	@FindBy(xpath="//*[@id='id_body_536870916']")
	public static WebElement  pvWareHouseTxt;

	@FindBy(xpath="//*[@id='id_body_33554540']")
	public static WebElement  enter_CaptionDefaultTxt;

	@FindBy(xpath="//*[@id='id_body_33554541']")
	public static WebElement  enter_addToNetDeductFromStock;


	@FindBy(xpath="//*[@id='id_body_33554542']")
	public static WebElement  enter_AddToNetNotApplicableToStock;

	@FindBy(xpath="//*[@id='id_body_33554543']")
	public static WebElement  enter_DeductAddToNetAddToStock;

	@FindBy(xpath="//*[@id='id_body_33554544']")
	public static WebElement  enter_NotApplicableAddToNetAddToStock;

	@FindBy(xpath="//*[@id='id_body_33554545']")
	public static WebElement  enter_DeductAddToNetDeductAddToStock;

	@FindBy(xpath="//*[@id='id_body_33554546']")
	public static WebElement  enter_NotApplicableAddToNetNotApplicableAddToStock;

	@FindBy(xpath="//*[@id='id_body_27']")
	public static WebElement  enter_RateTxt;

	@FindBy(xpath="//*[@id='id_body_28']")
	public static WebElement  enter_GrossTxt;


	@FindBy(xpath="//*[@id='id_body_33554521']")
	public static WebElement  enter_VATTxt;

	@FindBy(xpath="//*[@id='id_transactionentry_footer_panel_summary_value_net']/span[2]")
	public static WebElement  netvalue;







	@FindBy(xpath="//*[@id='551']/span")
	public static WebElement StockledgerBtn; 


	@FindBy (xpath="//*[@id='lblnchecked']/input")
	public static WebElement accountsFirtCheckbox;

	@FindBy(xpath="//*[@id='81']")
	public static WebElement reportsMenuInFinancials; 


	@FindBy(xpath="//*[@id='500']/span")
	public static WebElement ledgerBtn; 


	@FindBy(xpath="//*[@id='reportViewControls']/ul/li/span[4]/i")
	public static WebElement ledgerOkBtn; 

	@FindBy(xpath="//*[@id='rptheading']/div/span[2]")
	public static WebElement ledgerTitle;


	@FindBy(xpath="//*[@id='rptheading']/div/span[2]")
	public static WebElement stockLedgerTile;

	@FindBy(xpath="//table[@class='CommonReportTable']/tbody/tr[1]/td")
	public static List<WebElement> reportsRow1List;

	@FindBy(xpath="//table[@class='CommonReportTable']/tbody/tr[2]/td")
	public static List<WebElement> reportsRow2List;

	@FindBy(xpath="//table[@class='CommonReportTable']/tbody/tr[3]/td")
	public static List<WebElement> reportsRow3List;

	@FindBy(xpath="//table[@class='CommonReportTable']/tbody/tr[4]/td")
	public static List<WebElement> reportsRow4List;

	@FindBy(xpath="//td[@id='editScreen_bodyrow_9 _1']/span[@id='editIcon']")
	public static WebElement editBtn_9;

	@FindBy(xpath="//td[@id='editScreen_bodyrow_10 _1']/span[@id='editIcon']")
	public static WebElement editBtn_10;

	@FindBy(xpath="//td[@id='editScreen_bodyrow_11 _1']/span[@id='editIcon']")
	public static WebElement editBtn_11;

	@FindBy(xpath="//td[@id='editScreen_bodyrow_12 _1']/span[@id='editIcon']")
	public static WebElement editBtn_12;

	@FindBy(xpath="//td[@id='editScreen_bodyrow_13 _1']/span[@id='editIcon']")
	public static WebElement editBtn_13;

	@FindBy(xpath="//*[@id='editScreen_FieldsCustomization_chkHideNetFromSummary']")
	public static WebElement hideFromSummaryChkBox;





	@FindBy(xpath="//*[@id='id_transactionentry_summary']/div/div[2]/div/span")
	public static WebElement SummaryExpansionBtn;

	@FindBy(xpath="//div[@id='id_transactionentry_summary_static']/div/div[1]/label")
	public static List<WebElement> transactionSummaryList;

	@FindBy(xpath="//div[@id='id_transactionentry_summary_static']/div/div[2]/label")
	public static List<WebElement> transactionSummaryValuesList;

	@FindBy(xpath="//thead[@id='id_transaction_entry_detail_table_head']/tr/th/div[1]")
	public static List<WebElement> pvVATHeaderFieldList;


	@FindBy(xpath="(//span[@class='theme_button_color']/i[@class='icon-font6 icon-custamize'])[2]")
	public static WebElement ledgerCustomizeBtn;


	@FindBy(xpath="//*[@id='id_Trans_spanPlus']")
	public static WebElement ledgerTransactionExpandBtn;

	@FindBy(xpath="//*[@id='rd_customization_tree2']/span/span/i")
	public static WebElement ledgerExtraFieldExpandBtn;


	@FindBy(xpath="//ul[@data-text = 'show']//li")
	public static List<WebElement> extraFieldsList;

	/*@FindBy(xpath="//input[@id='id_header_10']")
	public static WebElement  voucherHeaderCurrency;
*/
	/*@FindBy(xpath="//tbody[@id='id_header_10_table_body']/tr/td[2]")
	public static List<WebElement> currencyListCount;*/	

	@FindBy(xpath="//input[@id='id_header_268435459']")
	public static WebElement  departmentTxt;

	@FindBy(xpath="//*[@id='id_header_268435459_table_body']/tr/td[2]")
	public static List<WebElement> openingBalDepartmentList;

	/*@FindBy(xpath="//input[@id='id_body_12']")
	public static WebElement  enter_AccountTxt;

	@FindBy(xpath="//*[@id='id_body_12_table_body']/tr")
	public static List<WebElement> openingBalAccountListInGrid;
*/
	@FindBy(xpath="//input[@id='id_body_18']")
	public static WebElement  enter_DebitTxt;

	@FindBy(xpath="//input[@id='id_body_19']")
	public static WebElement  enter_CreditTxt;

	/*@FindBy(xpath="//*[@id='lblAccount']")
	public static WebElement  billRefPartyName;

	@FindBy(xpath="//table[@id='id_Adjustment_Grid']")
	public static WebElement billRefAdjustBillsGrid;
*/
	@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr")
	public static List<WebElement> billRefAdjustBillsGridList;

	@FindBy(xpath="//input[@id='id_body_33554547']")
	public static WebElement enter_Behavior;

	@FindBy(xpath="//input[@id='id_body_33554548']")
	public static WebElement enter_Preload;

	@FindBy(xpath="//table[@id='editScreen_bodyTable']/tbody/tr/td[3]")
	public static List<WebElement> editScreenCaptionList;

	@FindBy(xpath="//table[@id='editScreen_bodyTable']/tbody/tr/td[1]/span[1]")
	public static List<WebElement> editScreenEditIconList;

	@FindBy(xpath="//tbody[@id='formulaGrideditScreen_FieldsCustomization_Formula_body']/tr/td[3]")
	public static List<WebElement>  editScreenFormulaNameList;

	@FindBy(xpath="//tbody[@id='formulaGrideditScreen_FieldsCustomization_Formula_body']/tr/td[4]")
	public static List<WebElement>  editScreenFormulaVariableList;

	@FindBy(xpath="//button[@id='editScreen_FieldsCustomization_Formula_Ok']")
	public static WebElement formulaokBtn;

	@FindBy(xpath="//span[@id='spnHeaderText']")
	public static WebElement costCenterTitle;	

	@FindBy (xpath="//*[@id='1108']/span")
	public static WebElement costcenter;

	@FindBy (xpath="//*[@id='1109']/span")
	public static WebElement location;

	@FindBy (xpath="//*[@id='1110']/span")
	public static WebElement region;

	@FindBy (xpath="//*[@id='1111']/span")
	public static WebElement country;

	@FindBy (xpath="//*[@id='1112']/span")
	public static WebElement state;

	@FindBy (xpath="//*[@id='1113']/span")
	public static WebElement city;

	@FindBy (xpath="//input[@id='iCountry']")
	public static WebElement countryComboBox;

	@FindBy (xpath="//tbody[@id='iCountry_table_body']/tr/td[2]")
	public static List<WebElement> countryComboBoxList;

	@FindBy (xpath="//input[@id='iState']")
	public static WebElement stateComboBox;

	@FindBy (xpath="//tbody[@id='iState_table_body']/tr/td[2]")
	public static List<WebElement> stateComboBoxList;

	@FindBy(xpath="//*[@id='GenerateNewMasterModel']/div/div[1]/div[2]/div/ul/li[2]/button[5]")
	public static WebElement newCloseBtn;

	@FindBy(xpath="//select[@id='EditLayout_FieldsCustomization_FieldDetails_DataType']")
	public static WebElement editLayoutDataTypeDropdown;
	
	@FindBy(xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_defaultValue']")
	public static WebElement editLayoutDefaultValueTxt;
	
	@FindBy(xpath="//*[@id='updateButton']/i")
	  public static WebElement updateBtnInPvVAT;

	@FindBy(xpath="//*[@id='drpdownSlider']")
	public static WebElement mastersSlider;
	

	@FindBy (xpath="//a[@id='id_showAllUnAuth']")
	public static WebElement showAllUnAuthBtn;

	@FindBy (xpath="//a[@id='btnAuthorize']/i")
	public static WebElement masterAuthorizeBtn;

	@FindBy (xpath="//input[@id='txt_authMessage']")
	public static WebElement reasonForAuthorizationTxt; 

	@FindBy (xpath="//input[@id='btnAuthorizationSave']")
	public static WebElement reasonPopupSaveBtn;



	@FindBy(xpath="//*[@id='idGlobalError']/div/div[2]")
	public static List<WebElement> errorMessageList;
	
	@FindBy(xpath="//*[@id='idGlobalError']/div/div[1]/button")
	public static List<WebElement> errorMessageCloseBtnList;




	// Opening Stocks
	
	@FindBy(xpath="//*[@id='id_transaction_viewcontainer']/li[1]/a")
	public static WebElement allVouchersInHomePage;

	/*@FindBy(xpath="//table[@id='id_rc_columnheadertable']/thead/tr/th/div/p")
	public static List<WebElement>  custBodyHeaderList;

	@FindBy(xpath="//*[@id='plnCustomizecolumn']/div/span[4]/a")
	public static WebElement  custRemoveBtn;

	@FindBy(xpath="//a[@id='Save']/i")
	public static WebElement  custSaveBtn;

	@FindBy(xpath="//a[@id='btnCustomizeClose']/i")
	public static WebElement  custCancelBtn;*/

	@FindBy(xpath="//a[@id='RearrangeFormula']/i")
	public static WebElement  custRearrangeFormulaBtn;
	
	@FindBy(xpath="//a[@id='2033']")
	public static WebElement  inventoryTransactionsStocksMenu; 
	
	@FindBy(xpath="//a[@id='2049']")
	public static WebElement  openingStocksNewVoucher;


	@FindBy(xpath="//a[@id='2037']//span[contains(text(),'Opening Stocks')]")
	public static WebElement  openingStocksVoucher;

	@FindBy (xpath="//div[@id='id_btnalternatecategorycheck']//div[@class='toolbar_button_image']")
	public static WebElement binAlternateCategoryCheckBtn;

	@FindBy (xpath="//*[@id='id_btnautoallocate']/following-sibling::input[@value='Pick']")
	public static WebElement binPickBtn; 

	@FindBy (xpath="//div[@id='id_btnautoallocategroup']//div[@class='toolbar_button_image']")
	public static WebElement binAutoAllocateWithinGroupBtn;
	
	@FindBy(xpath="//*[@id='id_body_84_input_container']/i")
	public static WebElement bin1Txt;

	@FindBy (xpath="//input[@id='id_bins_ok']")
	public static WebElement binOkBtn;

	@FindBy (xpath="//tfoot[@id='id_transaction_bins_grid_foot']//td[13]")
	public static WebElement binSumInward;

	@FindBy(xpath="//span[@class='icon-sum']")
	public static WebElement  binRowSum;

	@FindBy(xpath="//tfoot[@id='id_transaction_bins_grid_foot']//td[2]")
	public static WebElement  binRowSum_2ndColumn;

	@FindBy(xpath="//tfoot[@id='id_transaction_bins_grid_foot']//td[4]")
	public static WebElement  binSelect4htRow_3rdColumn;

	@FindBy(xpath="//tfoot[@id='id_transaction_bins_grid_foot']//td[6]")
	public static WebElement  binRowSum_4thColumn;

	@FindBy(xpath="//tfoot[@id='id_transaction_bins_grid_foot']//td[7]")
	public static WebElement  binRowSum_5thColumn;

	@FindBy(xpath="//tfoot[@id='id_transaction_bins_grid_foot']//td[8]")
	public static WebElement  binRowSum_6thColumn;

	@FindBy(xpath="//tfoot[@id='id_transaction_bins_grid_foot']//td[9]")
	public static WebElement  binRowSum_7thColumn;

	@FindBy(xpath="//tfoot[@id='id_transaction_bins_grid_foot']//td[12]")
	public static WebElement  binRowSum_8thColumn;

	@FindBy(xpath="//tfoot[@id='id_transaction_bins_grid_foot']//td[13]")
	public static WebElement  binRowTotalToBeAllocated;

	@FindBy (xpath="(//input[@id='id_bins_totalquantity'])[1]")
	public static WebElement binTotalQty;

	@FindBy (xpath="(//input[@id='id_bins_totalquantity'])[2]")
	public static WebElement binBaseUOM;

	@FindBy(xpath="//input[@id='id_transaction_bins_grid_control_heading_ctrl_12']")
	public static WebElement enter_ToBeAllocateTxt;

	@FindBy (xpath="//button[@class='btn Fbutton']")
	public static WebElement folderpathExpandBtn;

	@FindBy (xpath="//*[@id='navbarSupportedContent2']/ul/li[1]/a/i")
	public static WebElement restoreCompanyBtn;

	@FindBy (xpath="//*[@id='RestoreCompanyModal']/div/div/div[2]/button[1]")
	public static WebElement overRideYesBtn;

	@FindBy (xpath="//*[@id='RestoreCompanyModal']/div/div/div[2]/button[2]")
	public static WebElement overRideNoBtn;

	@FindBy (xpath="//*[@id='RestoreCompanyModal']/div/div/div[2]/button[3]")
	public static WebElement overRideCancelBtn;

	@FindBy (xpath="//tbody[@id='id_transaction_bins_grid_body']/tr/td[1]")
	public static List<WebElement> binPopupIndexList;

	@FindBy (xpath="//tbody[@id='id_transaction_bins_grid_body']/tr/td[2]")
	public static List<WebElement> binPopupBinsList;

	@FindBy (xpath="//tbody[@id='id_transaction_bins_grid_body']/tr/td[13]")
	public static List<WebElement> binPopupToBeAllocatedList;

	@FindBy (xpath="//tbody[@id='tblBodyTransRender']/tr[1]/td[12]")
	public static WebElement voucherHomeRow1SuspendedStatus;

	@FindBy(xpath="//ul[@id='id_transaction_viewcontainer']/li[1]/a")
	public static WebElement allVouchersLink;

	@FindBy(xpath="//a[@id='btnCreateView']/i")
	public static WebElement  createViewBtn;

	@FindBy(xpath="//*[@id='txtViewName']")
	public static WebElement  viewNameTxtInCreateView;

	@FindBy(xpath="//*[@id='cmbDataSet']")
	public static WebElement  dataSetInCreateView;

	@FindBy(xpath="//*[@id='cmbViewType']")
	public static WebElement  viewTypeInCreateView;

	@FindBy(xpath="//*[@id='cmbAuthorization']")
	public static WebElement  authorizationInCreateView;

	@FindBy(xpath="//*[@id='cmbCheckStatus']")
	public static WebElement  checkSatusInCreateView;

	@FindBy(xpath="//*[@id='cmbQC']")
	public static WebElement  qcInCreateView;

	@FindBy(xpath="//*[@id='Suspendstatus']")
	public static WebElement  suspendStatusInCreateView;

	@FindBy(xpath="//*[@id='idFilterCustomizeIcon']")
	public static WebElement  filterOptionInCreateView;

	@FindBy(xpath="//*[@id='dvCreateEditView']/div[2]/div/div[3]/div/input[1]")
	public static WebElement  saveOptionInCreateView;

	@FindBy(xpath="//*[contains(text(),'Suspended Vouchers')]")
	public static WebElement createViewNameDisplayInHomePage;
	
	@FindBy(xpath="//*[@id='id_rc_columnheadertable']/thead/tr/th[4]")
	public static WebElement  modifiedHeaderTab;

	@FindBy(xpath="//*[@id='id_rc_columnheadertable']/thead/tr/th[13]")
	public static WebElement  warehouseCodeHeaderTxt;

	@FindBy(xpath="//input[@id='HidethisColumn']")
	public static WebElement  hideThisColumnChkboxIsSelected;	
	
	@FindBy(xpath="//input[@id='HidethisColumn']/following-sibling::span")
	public static WebElement  hideThisColumnChkbox;

	@FindBy(xpath="//*[@id='tblHeadTransRender']/tr/th")
	public static List<WebElement>  homePageBodyHeaderList;
	
	@FindBy(xpath="//span[contains(text(),'Extra Fields')]")
	public static WebElement  custTransExtraFieldExpandBtn;

	@FindBy(xpath="//li[@id='rd_customization_tree3258']//span[contains(text(),'Warehouse')]")
	public static WebElement  warehouseExpandBtn;

	@FindBy(xpath="//*[@id='rd_customization_tree3259']")
	public static WebElement  warehouseName;

	@FindBy(xpath="//li[@id='rd_customization_tree3260']")
	public static WebElement  warehouseCode;

	@FindBy(xpath="//input[@id='ColumnHeading']")
	public static WebElement  columnHeadingTxt;

	@FindBy(xpath="//li[@data-fieldname='Department']//span[contains(text(),'Department')]")
	public static WebElement  cusDepartmentExpandBtn;

	@FindBy(xpath="//li[@data-fieldname='Department']/ul/li[1]")
	public static WebElement  cusDepName;

	@FindBy(xpath="//*[@id='id_rc_columnheadertable']/thead/tr[1]/th/div/p")
	public static List<WebElement> cusGridHeaderList;

	@FindBy(xpath="//*[@id='id_Default_spanPlus']")
	public static WebElement  custLeftPannelDefaultExpandBtn;

	@FindBy(xpath="//*[@id='id_prog_spanPlus']")
	public static WebElement  custLeftPannelProgExpandBtn;

	@FindBy(xpath="//*[@id='id_Trans_spanPlus']")
	public static WebElement  custLeftPannelTransExpandBtn;

	@FindBy(xpath="//*[@id='id_rc_complete_tree_container']/div[1]/ul/li/ul/li/ul/li/span")
	public static List<WebElement>  custLeftPannelInnerFieldList;

	@FindBy(xpath="//*[@id='ReportSpecific']/li")
	public static List<WebElement>  custLeftPannelDefaultFieldInnerFieldList;	

	@FindBy(xpath="//*[@id='ProgFileds']/li")
	public static List<WebElement>  custLeftPannelProgFiledsInnerFieldList;
	
	@FindBy(xpath="//*[@id='id_search_menu']/input")
	public static WebElement serachMenuTextHomePage;

	
	// Material Requistion
	
	@FindBy(xpath="//a[@id='2038']")
	public static WebElement materialRequisitionVoucher;
	
	
	// Purchase Orders
	
	@FindBy(xpath="//li[@id='id_transactionentry_authorize']/a/span")
	public static WebElement  POauthorizeBtn;
	
	@FindBy (xpath="//a[contains(text(),'Pending Purchases Orders')]")
	public static WebElement pendingPurchasesOrdersLink;	
	
	@FindBy(xpath="//a[@id='id_transactionentry_closelink']/label")
	public static WebElement  closeLinksOption;
	
	@FindBy (xpath="//tbody[@id='id_transaction_entry_detail_workflow_popup_body']/tr/td")
	public static List<WebElement> LPurchaseOrderPopupList;

	@FindBy (xpath="//tbody[@id='id_transaction_entry_detail_workflow_popup_body']/tr/td[7]")
	public static List<WebElement> LPurchaseOrderPopupVoucherNos;

	@FindBy (xpath="//tbody[@id='id_transaction_entry_detail_workflow_popup_body']/tr/td[2]/input")
//	@FindBy(xpath="(//tbody[@id='id_transaction_entry_detail_workflow_popup_body']/tr/td[2]/input[1])[1]")
	public static List<WebElement> LPurchaseOrderPopupCheckBoxs;

	@FindBy (xpath="//*[@id='id_transactionentry_workflow_popup']/div[2]/div/div[3]/div/input[4]")
	public static WebElement LPurchaseOrderPopupOkBtn;

	//@FindBy (xpath="//i[@id='id_transactionentry_workflow_popup_close']")
	@FindBy(xpath="//*[@id='id_transactionentry_workflow_popup']/div[2]/div/div[1]/span/i")
	public static WebElement LPurchaseOrderPopupCloseBtn;

	@FindBy (xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr[1]/td[11]")
	public static WebElement LpurchaseOrderPopupRow1SuspendCol;

	@FindBy (xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr[2]/td[11]")
	public static WebElement LpurchaseOrderPopupRow2SuspendCol;
	
	
	// Purchase voucher
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr/td[2]/input")
	public static WebElement workFlowChkBox1;

	@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_col_2-1']/input")
	public static WebElement workFlowChkBox2;

	@FindBy(xpath="//input[@id='id_transaction_entry_detail_workflow_control_heading_ctrl_1']")
	public static WebElement  workFlowHeaderChkBoxCL;

	@FindBy (xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr/td[2]/input")
	public static List<WebElement> workFlowChkBoxList;

	@FindBy (xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr/td[7]")
	public static List<WebElement> workFlowVoucherNosList;

	@FindBy(xpath="//*[@id='id_transactionentry_workflow_popup']/div[2]/div/div[3]/div/input[4]")
	public static WebElement  workFlowOkBtnCL;
	
	@FindBy (xpath="//input[@id='id_header_67108978']")
	public static WebElement purchaseVATNarrationTxt;
	
	@FindBy (xpath="//*[@id='id_transaction_entry_detail_table_control_heading_2']/div[1]")
	public static WebElement purchaseVATTableHeaderColumn2;

	
	
	// Payments Page
	
	/*@FindBy(xpath="//label[@id='id_BillWise_IP_AccountName']")
	public static WebElement breakUpDetailsAccount;

	@FindBy(xpath="//label[@id='id_BillWise_IP_ProductName']")
	public static WebElement breakUpDetailsItem;

	@FindBy(xpath="//label[@id='id_BillWise_IP_TagName']")
	public static WebElement breakUpDetailsDepartment;

	@FindBy(xpath="//label[@id='id_infoPanel_lblAEDTranAmountValue']")
	public static WebElement asOnEntryDateTransAmt;
*/
	/*@FindBy(xpath="//label[@id='id_infoPanel_lblAEDBaseConversionValue']")
	public static WebElement asOnEntryDateBaseConcersationRate;

	@FindBy(xpath="//label[@id='id_infoPanel_lblAEDBaseAmountValue']")
	public static WebElement asOnEntryDateBaseAmount;

	@FindBy(xpath="//label[@id='id_infoPanel_lblAEDLocalConversionValue']")
	public static WebElement asOnEntryDateLocConversationRate;

	@FindBy(xpath="//label[@id='id_infoPanel_lblAEDLocalAmountValue']")
	public static WebElement asOnEntryDateAmt;

	@FindBy(xpath="//label[@id='id_infoPanel_lblAADTranAmountValue']")
	public static WebElement balOnAdjstDateTransAmt;*/

	/*@FindBy(xpath="//label[@id='id_infoPanel_lblAADBaseConversionValue']")
	public static WebElement  balOnAdjstDateBasrConversionRate;

	@FindBy(xpath="//label[@id='id_infoPanel_lblAADBaseAmountValue']")
	public static WebElement  balOnAdjstDateBaseAmount;

	@FindBy(xpath="//label[@id='id_infoPanel_lblAADLocalConversionValue']")
	public static WebElement  balOnAdjstDateLocalConversionRate;

	@FindBy(xpath="//label[@id='id_infoPanel_lblAADLocalAmountValue']")
	public static WebElement  balOnAdjstDateAmt;

	@FindBy(xpath="//label[@id='id_infoPanel_lblAdjTranAmountValue']")
	public static WebElement adjustmentsAmount1;

	@FindBy(xpath="//label[@id='id_infoPanel_lblAdjBaseAmountValue']")
	public static WebElement adjustmentsAmount2;

	@FindBy(xpath="//label[@id='id_infoPanel_lblAdjLocalAmountValue']")
	public static WebElement adjustmentsAmount3;

	@FindBy(xpath="//label[@id='id_infoPanel_lblNativeCurrencyValue']")
	public static WebElement adjustmentsAmount4;

	@FindBy(xpath="//label[@id='id_infoPanel_lblExDiffTranAmountValue']")
	public static WebElement exchangeGainLossForBaseCurrency;

	@FindBy(xpath="//label[@id='id_infoPanel_lblExDiffLocalAmountValue']")
	public static WebElement exchangeGainLossForLocalCurrency;
	
	/*@FindBy(xpath="//*[@id='id_header_4_table_body']/tr/td[2]")
	public static List<WebElement>  cashAndBAnkAccountList;*/

	/*@FindBy(xpath="//input[@id='id_header_4']")
	public static WebElement newCashBankAccountTxt;
*/
	/*@FindBy(xpath="//*[@id='id_body_12_table_body']/tr/td[2]")
	public static List<WebElement> bodyAccountListInGrid;*/

	/*@FindBy(xpath="//*[@id='id_body_39_table_body']/tr/td[2]")
	public static List<WebElement> bodyCreditAccountListInGrid;

	@FindBy(xpath="//input[@id='id_body_16777307']")
	public static WebElement  enterpayVATTaxCode;*/

	@FindBy(xpath="//input[@id='id_body_16']")
	public static WebElement  enter_Amount;

	@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr")
	public static List<WebElement> billRefAdjustBillsList;

	/*@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr[1]/td[13]")
	public static WebElement gridAdjustmentAmtRow1;

	@FindBy(xpath="//*[@id='id_transactionentry_save']/a/span")
	public static WebElement  MRsaveBtn;
	*/
	@FindBy(xpath="//a[@id='2003']//span[contains(text(),'Payments')]")
	public static WebElement  paymentsVoucher;
	
	@FindBy(xpath="//span[contains(text(),'Payments VAT')]")
	public static WebElement  paymentsVATVoucher;
	
	@FindBy(xpath="//a[@id='2001']")
	public static WebElement  cashAndBankMenu; 

	@FindBy(xpath="//a[@id='2002']//span[contains(text(),'Receipts')]")
	public static WebElement  receiptsVoucher;

	@FindBy(xpath="//span[contains(text(),'Receipts VAT')]")
	public static WebElement  recepitsVATVoucher;

	@FindBy(xpath="//span[contains(text(),'JV VAT View')]")
	public static WebElement  JVVATViewVoucher;
	
	
	
	// Excesses In Stocks
	
/*	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[7]")
	public static WebElement  select4thRow_6thColumn;*/
	
	@FindBy(xpath="//span[contains(text(),'Excesses in Stocks')]")
	public static WebElement  excessesInStocksVoucher;
	
	
	
	// Stock Transfer
	
	@FindBy (xpath="//input[@id='chkSelectALL']")
	public static WebElement searchRMAHeaderChkBox;
	
	@FindBy(xpath="//span[contains(text(),'Stock Transfers')]")
	public static WebElement  stockTransfersVoucher;
	
	/*@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[2]")
	public static WebElement  select4thRow_1stColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[12]")
	public static WebElement  select4thRow_11thColumn;
*/
	@FindBy(xpath="//input[@id='id_header_86']")
	public static WebElement  stockTransferheaderWarehouse1Txt;

	/*@FindBy(xpath="//tbody[@id='id_header_86_table_body']/tr/td[2]")
	public static List<WebElement> wareHouseHeaderListCount;*/

	@FindBy(xpath="//input[@id='id_body_87']")
	public static WebElement  enter_BodyWarehouse2Txt;

	/*@FindBy(xpath="//tbody[@id='id_body_87_table_body']/tr/td[2]")
	public static List<WebElement> wareHouseBodyListCount;*/

	@FindBy(xpath="//*[@id='id_transactionentry_workflow_popup_heading']/following-sibling::span/i")
	public static WebElement  workFlowCloseBtn;

	@FindBy (xpath="//*[@id='Id_MRPBatch_Grid_body']/tr/td[2]")
	public static List<WebElement> batchPopupNameList;

	@FindBy (xpath="//*[@id='Id_MRPBatch_Grid_body']/tr/td[1]")
	public static List<WebElement> batchPopupIndexList;

	/*@FindBy (xpath="//*[@id='id_transaction_copydocpopup']/div/div/div[1]/button/i")
	
	public static WebElement cd_CloseBtn;
*/
	@FindBy(xpath="//*[@id='id_transactionentry_copytoclipboard']")
	public static WebElement  copyToClipBoardOption;

	@FindBy(xpath="//*[@id='id_transactionentry_pastefromclipboard']")
	public static WebElement  pasteFromClipBoardOption;
	
	// Sales Order
	
	@FindBy(xpath="//*[@id='cmbBehaviour']")
	public static WebElement behaviourDropdown;
	
	@FindBy(xpath="//*[@id='editScreen_qtyFormula_textbox']")
	public static WebElement editScreenFormulaTxt;
	
	@FindBy(xpath="//*[@id='editScreen_qtyPreLoaded_textbox']")
	public static WebElement editScreenQuantityPreloadedTxt;
	
	@FindBy(xpath="//input[@id='editScreen_qtyPreLoaded_formulaText']")
	public static WebElement qtyPreloadedformulaTxt;

	@FindBy(xpath="//button[@id='editScreen_qtyPreLoaded_Ok'][text()='Ok']")
	public static WebElement qtyPreloadedFormulaokBtn;
	
	@FindBy(xpath="//select[@id='inventoryOpt_ReservationType']")
	public static WebElement reservationTypeDropdown;
	
	@FindBy(xpath="//*[@id='2017']/span")
	public static WebElement  salesOrdersVoucher;
	
	@FindBy(xpath="//input[@id='id_body_33554537']")
	//@FindBy(xpath="//*[@id='id_body_33554548']")
	public static WebElement so_enter_AQTxt;

	@FindBy(xpath="//input[@id='id_body_33554538']")
//	@FindBy(xpath="//*[@id='id_body_33554549']")
	public static WebElement so_enter_FQTxt;
	
	@FindBy(xpath="//*[@id='Id_BatchReservation_FGrid_body']/tr/td")
	public static List<WebElement> batchReservationPopGridBody;

	@FindBy(xpath="//*[@id='Id_BatchReservation_FGrid_body']/tr/td[4]")
	public static List<WebElement> batchReservationPopGridBatch;

	@FindBy(xpath="//*[@id='Id_BatchReservation_FGrid_body']/tr/td[3]")
	public static List<WebElement> batchReservationPopGridQtyToRelease;

	@FindBy(xpath="//input[@id='txtQtyToReserve']")
	public static WebElement enter_QtyToReserve;

	@FindBy(xpath="//div[@id='id_Reservation_Modal_SalesOrder']/div[2]/div/div[3]/div/ul[4]/li[1]/div/div[2]")
	public static WebElement res_PickBtn;

	@FindBy(xpath="//div[@id='id_Reservation_Modal_SalesOrder']/div[2]/div/div[3]/div/ul[4]/li[2]/div/div[2]")
	public static WebElement res_DiscardBtn;

	@FindBy(xpath="//div[@id='id_Reservation_Modal_SalesOrder']/div[2]/div/div[3]/div/ul[4]/li[3]/div/div[2]")
	public static WebElement res_OkBtn;

	@FindBy(xpath="//div[@id='id_Reservation_Modal_SalesOrder']/div[2]/div/div[3]/div/ul[4]/li[4]/div/div[2]")
	public static WebElement res_CloseBtn;
	
	@FindBy(xpath="//*[@id='id_transactionentry_settings']")
	public static WebElement  settingsOption;

	//@FindBy(xpath="//*[@id='id=']/tbody/tr/td/span[2]")
	@FindBy(xpath="//*[@id='deleteIcon']")
	public static List<WebElement> editScreenbodyDeleteBtn;

//	@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']/div[3]/div/div/table/tbody/tr/td[3]")
	@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']/div[2]/div[2]/div//table//tr//td[3]")
	public static List<WebElement>  editScreeAtyPreLoadFormulaList;

	//@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']/div[3]/div/div/table/tbody/tr/td[4]")
	@FindBy(xpath="//*[@id='formulaGrideditScreen_qtyPreLoaded_body']/tr/td[4]")
	public static List<WebElement>  editScreeAtyPreLoadFormulaVariableList;

	@FindBy(xpath="//a[@id='btnCustomizeClose']/i")
	public static WebElement settings_closeBtn;

	@FindBy(xpath="//*[@id='Id_StockReservation_FGrid_row_1']/td")
	public  static List<WebElement> stockReservationPopGridBody;

	@FindBy(xpath="//td[@id='Id_StockReservation_FGrid_col_1-4']")
	public static WebElement stockRes_QtyToReserveRow1;

	@FindBy(xpath="//*[@id='Id_BinReservation_FGrid_body']/tr/td")
	public  static List<WebElement> binReservationPopGridBody;

	@FindBy(xpath="//*[@id='Id_BinReservation_FGrid_body']/tr/td[5]")
	public  static List<WebElement> binReservationPopGridBin;

	@FindBy(xpath="//*[@id='Id_BinReservation_FGrid_body']/tr/td[3]")
	public  static List<WebElement> binReservationPopGridQtyToRelease;
	
	@FindBy(xpath="//tbody[@id='Id_RMAReservation_FGrid_body']/tr/td[3]")
	public static List<WebElement> RmaWarehouseList;
	
	@FindBy(xpath="//tbody[@id='Id_RMAReservation_FGrid_body']/tr/td[4]")
	public static List<WebElement> RmaRmaNoList;

	@FindBy(xpath="//tbody[@id='Id_RMAReservation_FGrid_body']/tr/td[5]/input")
	public static List<WebElement> RmaSelectedList;

	@FindBy(xpath="//*[@id='Id_RMAReservation_FGrid_body']/tr/td[4]")
	public static List<WebElement> ReservePopUpRmaNoList; 

	@FindBy(xpath="//*[@id='Id_RMAReservation_FGrid_body']/tr/td[5]/input")
	public static List<WebElement> ReservePopUpRmaNoChkBox; 
	
	
	// Sales Invoice
	
	@FindBy(xpath="//span[contains(text(),'Sales Returns')]")
	public static WebElement  salesReturnsVoucher;

	@FindBy(xpath="//ul[@id='id_transaction_viewcontainer']//li[2]")
	public static WebElement  pendingSalesInvoicesView;
	
	/*@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[6]")
	public static WebElement  select4thRow_5thColumn;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[9]")
	public static WebElement  select4thRow_8thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[15]")
	public static WebElement  select4thRow_14thColumn;*/
	
	@FindBy(xpath="//select[@id='id_transactionentry_workflow_dropdown']")
	public static WebElement  workFlowDropdown;

	@FindBy(xpath="//select[@id='id_transactionentry_workflow_searchon']")
	public static WebElement  searchOnDropdownCL;

	@FindBy(xpath="//select[@id='id_transactionentry_workflow_searchtype']")
	public static WebElement  containingDropdownCL;

	@FindBy(xpath="//input[@id='id_transactionentry_workflow_search_input']")
	public static WebElement  workFlowSearchTxtCL;

	@FindBy(xpath="//div[@id='id_transaction_entry_workflow_control_area2']//td[4]")
	public static WebElement  workFlowCustomizeBtnCL;



	//1 Row CRTL + L


	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[1]/td[7]")
	public static WebElement  workFlowRow1Column1CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[1]/td[8]")
	public static WebElement  workFlowRow1Column2CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[1]/td[9]")
	public static WebElement  workFlowRow1Column3CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[1]/td[10]")
	public static WebElement  workFlowRow1Column4CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[1]/td[11]")
	public static WebElement  workFlowRow1Column5CL;


	//2 Row CRTL + L

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_body1]/tr[2]/td[2]/input")
	public static WebElement workFlowRow2ChkBoxCL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[2]/td[7]")
	public static WebElement  workFlowRow2Column1CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[2]/td[8]")
	public static WebElement  workFlowRow2Column2CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[2]/td[9]")
	public static WebElement  workFlowRow2Column3CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[2]/td[10]")
	public static WebElement  workFlowRow2Column4CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[2]/td[11]")
	public static WebElement  workFlowRow2Column5CL;


	//3 Row CRTL + L

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_body1]/tr[3]/td[2]/input")
	public static WebElement workFlowRow3ChkBoxCL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[3]/td[7]")
	public static WebElement  workFlowRow3Column1CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[3]/td[8]")
	public static WebElement  workFlowRow3Column2CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[3]/td[9]")
	public static WebElement  workFlowRow3Column3CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[3]/td[10]")
	public static WebElement  workFlowRow3Column4CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[3]/td[11]")
	public static WebElement  workFlowRow3Column5CL;


	//4 Row CRTL + L

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_body1]/tr[4]/td[2]/input")
	public static WebElement workFlowRow4ChkBoxCL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[4]/td[7]")
	public static WebElement  workFlowRow4Column1CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[4]/td[8]")
	public static WebElement  workFlowRow4Column2CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[4]/td[9]")
	public static WebElement  workFlowRow4Column3CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[4]/td[10]")
	public static WebElement  workFlowRow4Column4CL;

	@FindBy(xpath="//tbody[@id='id_transaction_entry_detail_workflow_body']/tr[4]/td[11]")
	public static WebElement  workFlowRow4Column5CL;

	@FindBy(xpath="//div[@id='id_transactionentry_header']//span[@class='col-xs-6 icon-collepse icon-font6 no_padding_left_right theme_color-inverse']")
	public static WebElement  new_HeaderMinimizeBtn;

	@FindBy(xpath="//div[@class='col-xs-12']//span[@class='col-xs-6 icon-font6 no_padding_left_right theme_color-inverse icon-expand']")
	public static WebElement  new_HeaderExpandBtn;

	@FindBy(xpath="//div[@id='id_transaction_entry_container']//span[contains(text(),'Purchase Vouchers N')]")
	public static WebElement  purchaseVoucherNLink;





	//WorkFlow Elements


	@FindBy(xpath="//select[@id='id_transactionentry_workflow_popup_searchon']")
	public static WebElement  searchOnDropdown;

	@FindBy(xpath="//select[@id='id_transactionentry_workflow_popup_searchtype']")
	public static WebElement  containingDropdown;

	@FindBy(xpath="//input[@id='id_transactionentry_workflow_popup_search']")
	public static WebElement  workFlowSearchTxt;


	@FindBy(xpath="//*[@id='id_transactionentry_workflow_popup']/div[2]/div/div[3]/div/input[1]")
	public static WebElement  workFlowCustomizeBtn;

	//@FindBy(xpath="//*[@id='id_transactionentry_workflow_popup']/div[2]/div/div[3]/div/input[2]")
	@FindBy(xpath="//*[@id='id_transactionentry_workflow_popup']/div[2]/div/div[3]/div/input[4]")
	public static WebElement  workFlowOkBtn;

	@FindBy(xpath="//div[@class='toolbar_button_text font-7'][contains(text(),'Ok')]")
	public static WebElement  workFlowOkBtnInSalesVoucher;


	@FindBy(xpath="//span[@id='id_transaction_entry_workflow_togglebutton']")
	public static WebElement  workFlowMinimizeBtn;



	//WorkFlow Screen with Line wise
	@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_1-1']//input")
	public static WebElement  workFlowHeaderChkBox;

	@FindBy(xpath="//th[@id='id_transaction_entry_detail_workflow_popup_control_heading_6']//div[contains(text(),'Voucher No')]")
	public static WebElement  workFlowHeaderVoucherNo;

	@FindBy(xpath="//th[@id='id_transaction_entry_detail_workflow_popup_control_heading_7']//div[contains(text(),'Date')]")
	public static WebElement  workFlowHeaderDate;

	@FindBy(xpath="//div[contains(text(),'Name')]")
	public static WebElement  workFlowHeaderName;

	@FindBy(xpath="//th[@id='id_transaction_entry_detail_workflow_popup_control_heading_9']//div[contains(text(),'Quantity')]")
	public static WebElement  workFlowHeaderQty;

	@FindBy(xpath="//div[contains(text(),'Suspended')]")
	public static WebElement  workFlowHeaderSuspend;


	//WorkFlow with Line wise
	// 1 Row

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr[1]/td[2]/input")
	public static WebElement workFlowRadioBtnRowOne;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr[2]/td[2]/input")
	public static WebElement workFlowRadioBtnRowTwo;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr[1]/td[7]")
	public static WebElement  workFlowRow1Column1;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr[1]/td[8]")
	public static WebElement  workFlowRow1Column2;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr[1]/td[9]")
	public static WebElement  workFlowRow1Column3;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr[1]/td[10]")
	public static WebElement  workFlowRow1Column4;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr[1]/td[11]")
	public static WebElement  workFlowRow1Column5;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr[2]/td[7]")
	public static WebElement  workFlowRow2Column1;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr[2]/td[8]")
	public static WebElement  workFlowRow2Column2;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr[2]/td[9]")
	public static WebElement  workFlowRow2Column3;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr[2]/td[10]")
	public static WebElement  workFlowRow2Column4;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr[2]/td[11]")
	public static WebElement  workFlowRow2Column5;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_body']/tr")
	public static List<WebElement>  LinkRowCountCL;	
	
	@FindBy (xpath="//input[@id='id_body_35_11']")
	public static WebElement enter_LSalesOrders;
	
	
	
	// Receipts 
	
	@FindBy(xpath="//*[@id='id_body_16777305']")
	public static WebElement  enterReceiptsVATTaxCode;

	@FindBy(xpath="//input[@id='id_header_268435470']")
	public static WebElement  placeofSupplyTxt;

	/*@FindBy(xpath="//tbody[@id='id_header_268435470_table_body']/tr/td[2]")
	public static List<WebElement>  placeofSupplyList;

	@FindBy(xpath="//tbody[@id='id_header_268435471_table_body']/tr/td[2]")
	public static List<WebElement>  jurisdictionList;*/

	@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr/td[3]")
	public static List<WebElement>  billwiseAdjustBillsDocList;

	@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr/td[13]")
	public static List<WebElement>  billwiseAdjustBillsAdjustColumn;

	@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr/td[2]/input")
	public static List<WebElement>  billwiseAdjustBillsChkBoxList;

	@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr[1]/td[2]/input")
	public static WebElement  billrefAdjuBills1stChkbox;
	
	
	// Credit Notes
	
	@FindBy(xpath="//*[@id='2023']/span")
	public static WebElement  finTransJournalsMenu;

	@FindBy(xpath="//*[@id='doc_TagsTable_body']/tr/td[2]")
	public static List<WebElement>  masterList;

	@FindBy(xpath="//*[@id='doc_TagsTable_body']/tr/td[1]")
	public static List<WebElement>  masterIndexList;

	@FindBy(xpath="//a[@id='navigationtab2']")
	public static WebElement settingMiscellaneousTab;
	
	@FindBy(xpath="(//*[@id='panelsStayOpen-headingThree'])[2]")
	public static WebElement miscCurrencyTabExpBtn;

	@FindBy(xpath="//select[@id='misc_CurrencyAddCurrencyIn']")
	public static WebElement miscAddCurrencyDropdown;

//	@FindBy(xpath="//input[@id='misc_currencyInputExchangeRate']")
	@FindBy(xpath="//*[@id='panelsStayOpen-collapseThree']/label[2]/span")
	public static WebElement miscInputExchangeRateChkbox;

//	@FindBy(xpath="//input[@id='misc_currencyInputLocalExchangeRate']")
	@FindBy(xpath="//*[@id='id_misc_inputlocalexchangerate_section']/label/span")
	public static WebElement miscInputLocalExchangeRateChkbox;

	@FindBy(xpath="//select[@id='misc_arapDueDate']")
	public static WebElement miscDueDateDropdown;

	@FindBy(xpath="//span[contains(text(),'Credit Notes VAT')]")
	public static WebElement  creditNotesVATMenu; 

	@FindBy(xpath="//span[contains(text(),'Debit Notes VAT')]")
	public static WebElement  debitNotesVatMenu;

	@FindBy(xpath="//span[@id='Authorizedetails']")
	public static WebElement  debitNotesVatAuthorizedetailsBtn;

	@FindBy(xpath="//label[contains(text(),'Setting')]")
	public static WebElement settingBtn;

	@FindBy(xpath="//*[@id='updateButton']/i")
	public static WebElement settingUpdateIcon;

	//@FindBy(xpath="//i[@class='icon-close icon-font6']")
	@FindBy(xpath="//*[@id='btnCustomizeClose']/i")
	public static WebElement settingCloseIcon;
	
	@FindBy(xpath="//input[@id='id_body_16777316']")
	public static WebElement  enterDebitVATTaxCode;

	@FindBy(xpath="//input[@id='id_body_16777318']")
	public static WebElement  enterCreditVATTaxCode;

	@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr")
	public static List<WebElement> cn_billRefAdjustBillsGrid;
	
	/*@FindBy(xpath="//*[@id='id_body_12_table_body']/tr/td[2]")
	public static List<WebElement> accountListCount;*/
	
	
	
	// JV VAT
	
	@FindBy(xpath="//input[@id='id_body_39']")
	public static WebElement  enter_CreditACTxt;

	@FindBy(xpath="//input[@id='id_body_16777320']")
	public static WebElement  enterJVVATTaxCode;
	
	
	
	// PDC 
	
	
	@FindBy(xpath = "//*[@id='chkPostDatedCheques']/following-sibling::span")
	public static WebElement postDatedChequeChkbox;

	@FindBy(xpath = "//input[@id='chkAutoCheckMaturedPDC']/following-sibling::span")
	public static WebElement pdcAutoCheckforMaturedChkbox;

	@FindBy(xpath = "//input[@id='chkReverseACPostPDCConversion']/following-sibling::span")
	public static WebElement pdcReseverseAccountpostingOnPDCConversionChkbox;
	
	@FindBy(xpath = "//input[@id='chkReverseACPostPDCConversion']")
	public static WebElement pdcReseverseAccountpostingOnPDCConversionChkboxIsSelected;
	
	@FindBy(xpath = "//input[@id='PDCDiscountedAccount']")
	public static WebElement pdcCheckDiscountingTxt;

	@FindBy(xpath = "//input[@id='PDCDiscountedVoucherType']")
	public static WebElement pdcChequeDiscountingVoucherType;

	
	@FindBy(xpath="//*[@id='preferenceUL']/li[2]/ul/li[4]/a")
	public static WebElement  PDCBtn;

	@FindBy(xpath = "//input[@id='gphDisplayInLedger']")
	public static WebElement displayLedgerAndBalanceChkboxIsSelected;
	
	@FindBy(xpath = "//*[@id='dvPostDatedCheques']/div/label[2]/span")
	public static WebElement displayLedgerAndBalanceChkbox;

	@FindBy(xpath = "//input[@id='chkPostDatedCheques']")
	public static WebElement postDatedChequeChkboxIsSelected;
	
	@FindBy(xpath="//input[@id='id_body_16777309']")
	public static WebElement  enterTaxcode;

	@FindBy(xpath="//input[@id='id_body_16777311']")
	public static WebElement  enterPVPVATTaxcode;
	
	@FindBy(xpath="//input[@id='id_header_67108866']")
	public static WebElement  receipts_ChequeNoTxt;

	@FindBy(xpath="//input[@id='id_header_67108869']")
	public static WebElement  payments_ChequeNoTxt;
	
	@FindBy(xpath="//input[@id='id_header_268435471']")
	public static WebElement  PDRVAT_JuridictionTxt;
	
	@FindBy(xpath="//*[@id='id_header_6']")
	public static WebElement  maturityDateTxt;
	
	@FindBy(xpath="//input[@id='id_header_268435470']")
	public static WebElement  PDRVATPlaceOfSupplyTXt;
	
	@FindBy(xpath="//label[contains(text(),'Copy to Clipboard')]")
	public static WebElement copytoClipboardBtn;

	@FindBy(xpath="//label[contains(text(),'Paste from Clipboard')]")
	public static WebElement pastefromClipboardBtn;
	
	@FindBy(xpath="//input[@id='id_header_67108981']")
	public static WebElement voucherHeaderPDCNOTxt; 
	
	@FindBy(xpath="//*[@id='90']/span")
	public static WebElement  convertMaturedPDCs;

	@FindBy(xpath="//div[@class='navText']//span[contains(text(),'Convert Matured PDCs')]")
	public static WebElement  convertMaturedPDCsLabel;

	@FindBy(xpath="//span[contains(text(),'Auto Postings')]")
	public static WebElement  autoPostingMenu;

	@FindBy(xpath="//*[@id='90']//span[contains(text(),'Convert Matured PDCs')]")
	public static WebElement  convertMaturedPDCSMenu;

	@FindBy(xpath="//input[@id='start_date']")
	public static WebElement asOnDateTxt;

	@FindBy(xpath="//input[@id='chkReassign']")
	public static WebElement reassignChkbox;

	@FindBy(xpath="//input[@id='chkPostOnDate']")
	public static WebElement postOnDateChkbox;

	@FindBy(xpath="//input[@id='post_date']")
	public static WebElement postOnDateTxt;

	@FindBy(xpath="//input[@id='chkSelectAll']")
	public static WebElement gridSelectAllOption;

	@FindBy(xpath="//input[@id='chkColumn']")
	public static WebElement gridSelectAllCOl;
	
	@FindBy(xpath="//i[@class='icon-font6 icon-clear']")
	public static WebElement convertMaturedPDCsClearIcon;

	@FindBy(xpath="//*[@id='btnOK']/a/i")
	public static WebElement convertMaturedPDCsOkIcon;
	
	@FindBy(xpath = "//*[@id='btnReportcancel']/a/i")
	public static WebElement convertMaturedPDCsCancelIcon;

	@FindBy(xpath = "//div[contains(text(),'PDC Doc')]")
	public static WebElement gridPDCTxt;

	@FindBy(xpath = "//*[@id='btnOkToConvert']/a/i")
	public static WebElement pdcVoucherOkIcon;

	@FindBy(xpath = "//*[@id='btnReportcancel']")
	public static WebElement pdcVoucherCancelIcon;


	@FindBy(xpath="//input[@id='id_header_4']")
	public static WebElement  cashBankAccountTxt;

	@FindBy(xpath="//*[@id='chkColumn_1']")
	public static WebElement pdcGridRow1Chkbox;

	@FindBy(xpath="//td[@id='mainTable_col_1-3']")
	public static WebElement pdcGridRow1Col3;

	@FindBy(xpath="//td[@id='mainTable_col_1-4']")
	public static WebElement pdcGridRow1Col4;

	@FindBy(xpath="//td[@id='mainTable_col_1-5']")
	public static WebElement pdcGridRow1Col5;
	@FindBy(xpath="//td[@id='mainTable_col_1-6']")
	public static WebElement pdcGridRow1Col6;

	@FindBy(xpath="//td[@id='mainTable_col_1-7']")
	public static WebElement pdcGridRow1Col7;

	@FindBy(xpath="//td[@id='mainTable_col_1-8']")
	public static WebElement pdcGridRow1Col8 ;

	@FindBy(xpath="//td[@id='mainTable_col_1-10']")
	public static WebElement pdcGridRow1Col9 ;

	@FindBy(xpath="//span[contains(text(),'PDR VAT')]")
	public static WebElement PDRVAT ;
	
	/*@FindBy(xpath="//*[@id='id_header_268435470_table_body']/tr/td[2]")
	public static List<WebElement> placeOfSupplyList;

	@FindBy(xpath="//*[@id='id_header_268435471_table_body']/tr/td[2]")
	public static List<WebElement> jurdictionList;*/

	/*@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr[1]/td[8]")
	public static WebElement gridOrginalAmtRow1;
	*/
	/*@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr[1]/td[10]")
	public static WebElement gridBalanceAmtRow1;*/
	
	@FindBy(xpath="//input[@id='id_header_67108933']")
	public static WebElement  PDRVATChequeNoTxt;
	
	@FindBy(xpath="//*[@id='mytable_body']/tr/td[3]")
	public static List<WebElement> pdcCheckBoxlist;

	@FindBy(xpath="//*[@id='mytable_body']/tr/td[4]")
	public static List<WebElement> pdcDOClist;

	@FindBy(xpath="//*[@id='mytable_body']/tr/td[6]")
	public static List<WebElement> pdcBanklist;

	@FindBy(xpath="//a[@id='25']//span[contains(text(),'Data Management')]")
	public static WebElement  dataManagementMenu;

	@FindBy(xpath="//span[contains(text(),'Reindex')]")
	public static WebElement  reindexMenu;

	@FindBy(xpath="//input[@id='chk_fullReindex']")
	public static WebElement  fullReindexChkBox;

	@FindBy(xpath="//span[@class='icon-ok icon-font6']")
	public static WebElement  reindexOkBtn;

	@FindBy(xpath="//div[@class='col-sm-12 btnheader_img']//span[@class='icon-close icon-font6']")
	public static WebElement  reindexCancelBtn;

	@FindBy(xpath="//tbody//input[@placeholder='Value']")
	public static WebElement pdcVoucherFilterDateTxt;

	@FindBy(xpath="//input[@id='advancefilter_date_90_0']")
	public static WebElement enter_PDCFilterDate;

	@FindBy(xpath="//button[@id='btnFilterPDCDetails']")
	public static WebElement pdcVoucherFilterBtn;
	
	@FindBy(xpath="//*[@id='id_transactionentry_footer_panel_summary_value_net']/span[2]")
	public static WebElement recepitsFooterAmt;
	
	@FindBy(xpath="//span[contains(text(),'Auto Postings')]")
	public static WebElement  financialsTransactionsAutoPostingsMenu;
	
	@FindBy(xpath="//span[contains(text(),'Configure Transactions')]")
	public static WebElement  settingsConfigureTransactions;

	//@FindBy(xpath="//*[@id='salesnav']/li[2]/ul/li[4]/a")
	@FindBy(xpath="//*[@id='preferenceUL']/li[2]/ul/li[4]/a")
	public static WebElement  PDCMenu;

	@FindBy(xpath="(//*[@id='btnMasterSaveClick']/i)[2]")
	public static WebElement popSaveBtn;

	@FindBy (xpath="//a[@id='1104']//span[contains(text(),'Account')]")
	public static WebElement accountsMenu;

	@FindBy(xpath="//input[@id='sName']")
	public static WebElement accountNewNameTxt;

	@FindBy(xpath="//*[@id='btnMasterSaveClick']")
	public static WebElement accountNewSaveBtn;

	/*@FindBy(xpath="//*[@id='id_header_4_table_body']/tr/td[2]")
	public static List<WebElement> accountBankList;*/

	@FindBy(xpath="//*[@id='id_header_4_table_data_body']/tr/td[1]")
	public static List<WebElement> accountBankListwithCreditLimit;
	
	@FindBy(xpath="//*[@id='90_0_AdvanceFilter_']/table/tbody/tr/td[1]/select")
	public static WebElement pdcFilterConjuctionDropdown;

	@FindBy(xpath="//tbody//input[@placeholder='Select Field']")
	public static WebElement pdcFilterStartdateField;

	@FindBy(xpath="//*[@id='90_0_AdvanceFilter_']/table/tbody/tr/td[3]/select")
	public static WebElement pdcFilterOperatorDropdown;

	@FindBy(xpath="//*[@id='90_0_AdvanceFilter_']/table/tbody/tr/td[4]/select")
	public static WebElement pdcFilterCompareWithDropdown;;

	@FindBy(xpath="//td[@id='mytable_col_1-5']")
	public static WebElement gridRow1SelectBankCol;

	@FindBy(xpath="//td[@id='mytable_col_3-5']")
	public static WebElement gridRow3SelectBankCol;

	@FindBy(xpath="//td[@id='mytable_col_4-5']")
	public static WebElement gridRow4SelectBankCol;

	@FindBy(xpath="//td[@id='mytable_col_2-5']")
	public static WebElement gridRow2SelectBankCol;

	@FindBy(xpath="//input[@id='mytableoptReversePosting']")
	public static WebElement gridEnterBank;

	@FindBy(xpath="//div[@id='filterTree_90_0_AdvanceFilter_']//a[@id='4']")
	public static WebElement pdcFilterStartDate;

	@FindBy(xpath="//div[@id='filterTree_90_0_AdvanceFilter_']//a[@id='1']")
	public static WebElement pdcFilterChequenumber;

	@FindBy(xpath="//div[@id='filterTree_90_0_AdvanceFilter_']//a[@id='2']")
	public static WebElement pdcFilterCustomer;

	@FindBy(xpath="//div[@id='filterTree_90_0_AdvanceFilter_']//a[@id='3']")
	public static WebElement pdcFilterBank;

	@FindBy(xpath="//tbody/tr[1]/td[5]/div[1]/ul[1]/a[1]")
	public static WebElement pdcFilterValueStartDate;

	/*@FindBy(xpath="//div[@id='filterTree_90_0_AdvanceFilter_']//a[@id='4']")
				public static WebElement pdcFilterValueStartDate;*/

	@FindBy(xpath="//div[@id='filterTree_90_0_AdvanceFilter_']//a[@id='1']")
	public static WebElement pdcFilterValueChequenumber;

	@FindBy(xpath="//div[@id='filterTree_90_0_AdvanceFilter_']//a[@id='2']")
	public static WebElement pdcFilterValueCustomer;

	@FindBy(xpath="//div[@id='filterTree_90_0_AdvanceFilter_']//a[@id='3']")
	public static WebElement pdcFilterValueBank;
	
	@FindBy(xpath="//span[contains(text(),'PDP VAT')]")
	public static WebElement PDPVAT ;
	
	@FindBy(xpath="//input[@id='id_header_67108938']")
	public static WebElement pvpVATChequeTXT;

	
	
	
	@FindBy(xpath="//*[@id='editScreen_FieldsCustomization_DefaultValue']")
	public static WebElement editScreenDefaultValue;

	@FindBy(xpath="//tbody[@id='formulaGrideditScreen_FieldsCustomization_PreLoad_body']/tr/td[3]")
	public static List<WebElement>  editScreenPreloadNameList;

	@FindBy(xpath="//tbody[@id='formulaGrideditScreen_FieldsCustomization_PreLoad_body']/tr/td[4]")
	public static List<WebElement>  editScreenPreloadVariableList;

	@FindBy(xpath="//button[@id='editScreen_FieldsCustomization_PreLoad_Ok']")
	public static WebElement preloadokBtn;

	@FindBy(xpath="//a[contains(text(),'Properties')]")
	public static WebElement editScreenPropertiesTab; 
	
	@FindBy(xpath="//*[@id='editScreen_FieldsCustomization_chkReadOnly']/following-sibling::span")
	public static WebElement editScreenPropertiesReadOnly;  
	
	@FindBy(xpath="//*[@id='editScreen_FieldsCustomization_Formula_formulaText']")
	public static WebElement editScreenFormulaSearchTxt;
	
	@FindBy(xpath="//*[@id='editScreen_FieldsCustomization_PreLoad_formulaText']")
	public static WebElement editScreenPreLoadSearchTxt;
	
	@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/div[1]/span[1]/i")
	public static WebElement  editLayoutAddBtn;

	@FindBy(xpath="//i[@class='icon-copyfields icon-font7']")
	public static WebElement  editLayoutLoadBtn;

	@FindBy(xpath="//*[@id='editScreen_CustomizeButtons']/span[3]")
	public static WebElement  editLayoutCustomizeBtn;

	@FindBy(xpath="//*[@id='editIcon']")
	public static WebElement  editLayoutEditBtn;

	@FindBy(xpath="//*[@id='deleteIcon']")
	public static WebElement  editLayoutDeleteBtn; 

	@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[4]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr/td[1]/span[2]")
	public static WebElement  editscreenDeleteBtn; 

	@FindBy(xpath="//*[@id='editLayout_headerrow_1 _1']/span[1]")
	public static WebElement editLayouEditBtn;
	
	
	
	
	@FindBy(xpath = "//input[@id='txtGeneralMsg_RuleIfMsg_DocCustRules']")
	public static WebElement rules_MessageTabGeneralTXT;

	@FindBy(xpath = "//input[@id='txtCaptionFormat_IF_DocCustRules']")
	public static WebElement rulesTabIFGridEnterCaption;

	@FindBy(xpath = "//input[@id='txtCaptionFormat_Else_DocCustRules']")
	public static WebElement rulesTabElseGridEnterCaption;


	@FindBy(xpath = "//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-10']")
	public static WebElement rulesTabGrid1st_10thCol;

	@FindBy(xpath = "//input[@id='chkMandatoryFormat_IF_DocCustRules']")
	public static WebElement rulesTabGridMandatoryChkbox;		

	//@FindBy(xpath = "//input[@id='chkNewRecord_Rule_DocCustRules']")
	@FindBy(xpath="//*[@id='chkListApplyOn_DocCustRules']/div/div[2]/div/label/span")
	public static WebElement rulesTabNewRecordChkbox;

//	@FindBy(xpath = "//input[@id='chkEdit_Rule_DocCustRules']")
	@FindBy(xpath="//*[@id='chkListApplyOn_DocCustRules']/div/div[3]/div/label/span")
	public static WebElement rulesTabEditChkbox;

	@FindBy(xpath = "//input[@id='chkLoad_Rule_DocCustRules']")
	public static WebElement rulesTabLoadChkbox;

//	@FindBy(xpath = "//input[@id='chkOnLeave_Rule_DocCustRules']")
	@FindBy(xpath="//*[@id='chkListEvaluateOn_DocCustRules']/div/div[2]/div/label/span")
	public static WebElement rulesTabOnLeaveChkbox;

	@FindBy(xpath = "//input[@id='chkOnEnter_Rule_DocCustRules']")
	public static WebElement rulesTabOnEnterChkbox;

	@FindBy(xpath = "//input[@id='chkBeforeDelete_Rule_DocCustRules']")
	public static WebElement rulesTabBeforeDeleteChkbox;

	//@FindBy(xpath = "//input[@id='chkIsRuleActive_Rule_DocCustRules']")
	@FindBy(xpath="//*[@id='chkListStatus_DocCustRules']/div/div/div/label/span")
	public static WebElement rulesTabActiveChkbox;


	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-1']")
	public static WebElement rulesGrid1stRow_1stcol ;

	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_2-1']")
	public static WebElement rulesGrid2ndRow_1stcol ;

	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_2-4']")
	public static WebElement rulesGrid2ndRow_4thcol ;

	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_2-5']")
	public static WebElement rulesGrid2ndRow_5thcol ;	    

	@FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-1']")
	public static WebElement rulesGridElse1stRow_1stcol ;

	@FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_2-1']")
	public static WebElement rulesGridElse2ndRow_1stcol ;

	@FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_2-4']")
	public static WebElement rulesGridElse2ndRow_4thcol ;

	@FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_2-5']")
	public static WebElement rulesGridElse2ndRow_5thcol ;

	@FindBy(xpath="//select[@id='ddlFieldsDiv_Rule_IF_DocCustRules']")
	public static WebElement rulesIFEnterFiled ;

	@FindBy(xpath="//select[@id='ddlFieldsDiv_Rule_ELSE_DocCustRules']")
	public static WebElement rulesElseEnterFiled ;

	@FindBy(xpath="//select[@id='ddlChangeValueFormat_IF_DocCustRules']")
	public static WebElement rulesIFEnterChangeValueDrpdwn;

	@FindBy(xpath="//select[@id='ddlChangeValueFormat_ELSE_DocCustRules']")
	public static WebElement rulesElseEnterChangeValueDrpdwn;

	@FindBy(xpath="//input[@id='txtTextField']")
	public static WebElement rulesIFEnterValue;
	
	@FindBy(xpath="//*[@id='triggers_allVoucherDropDown']")
	public static WebElement editLayoutTriggersAllVouchers;  	

	@FindBy(xpath="//*[@id='17_1_AdvanceFilter_']/table/tbody/tr/td[1]/select")
	public static WebElement editLayoutTriggerValue;
	

	@FindBy(xpath="//button[@id='editScreen_FieldsCustomization_Formula_Ok']")
	public static WebElement  editScreenAddBehaviourFormulaokBtn;

	@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_PreLoad_textbox']")
	public static WebElement editScreenPreloadTxt;
	
	@FindBy(xpath="//*[@id='doc_baseDocument']")
	public static WebElement  baseDocumentTxtInDocCust;
	
	@FindBy(xpath="//a[contains(text(),'Properties')]")
	public static WebElement editLayoutPropertiesTab;  
	
	@FindBy(xpath="//input[@id='editLayout_FieldsCustomization_Properties_chkreadOnly']/following-sibling::span")
	public static WebElement readOnlyChkBox;
	
	@FindBy(xpath="//td[@id='editLayout_headerrow_1 _4']")
	public static WebElement  editLayoutHeader1stRowCaption;
	
	@FindBy(xpath="//*[@id='editScreen_FieldsCustomization_chkHidden']")
	public static WebElement editScreenPropertiesHidden;  
	
	@FindBy(xpath="//span[@class='icon-reject2 theme_color']")
	public static WebElement errorMessagecloseBtn;
	
	@FindBy(xpath="//*[@id='views_FieldsTable']/tbody/tr/td[3]")
	public static List<WebElement>  viewsGridFiledsList;

	@FindBy(xpath="//*[@id='views_FieldsTable']/tbody/tr/td[1]")
	public static List<WebElement>  viewsGridFiledsEditList;
	
	@FindBy(xpath="//*[@id='17_1_AdvanceFilter_']/table/tbody/tr/td[5]/input")
	public static WebElement  triggersValueTxtGetValue;
	
	@FindBy(xpath = "//div[@id='views_tabContent']//tr[2]//td[1]//span[1]")
	public static WebElement viewGridRow2EditBtn;
	
	@FindBy(xpath = "//*[@id='EditLayout_FieldsCustomization_FieldDetails_DataType']")
	public static WebElement fieldDetails_DataTypeDrpdwn;
	
	@FindBy(xpath = "//*[@id='dvCustomizeEditLayoutField']/div[2]/div/div[3]/input[1]")
	public static WebElement fieldDetails_ApplyBtn;
	
	@FindBy(xpath = "//*[@id='dvCustomizeEditLayoutField']/div[2]/div/div[3]/input[2]")
	public static WebElement fieldDetails_CloseBtn;
	
	@FindBy(xpath="(//*[@id='4']/span)[1]")
	public static WebElement  triggerVendorACExpansionBtn;

	//@FindBy(xpath="//div[@class='content-wrapper']//li[11]//ul[1]//li[1]//a[1]")
	@FindBy(xpath="(//*[@id='5002'])[2]")
	public static WebElement  triggerVendorAcName;

	@FindBy(xpath="//*[@id='4']")
	public static WebElement  triggerVendorAc;

	@FindBy(xpath="//*[@id='5002']")
	public static WebElement  triggerNameUnderVendor;
	
	
	@FindBy(xpath = "//*[@id='id_reportmenudisplay']/ul/li[9]")
	public static WebElement report_FilterExpandBtn;
	
	@FindBy(xpath = "//*[@id='id_reportmenudisplay']/ul/li[9]/ul//li[1]")
	public static WebElement report_FilterBtn;

	@FindBy(xpath = "//*[@id='id_reportmenudisplay']/ul/li[9]/ul/li[2]")
	public static WebElement report_RefineBtn;
	
	@FindBy(xpath = "//*[@id='idFilterCustomizeIcon']")
	public static WebElement report_FilterCustomizeBtn;

	@FindBy(xpath = "//span[@id='a']")
	public static WebElement report_Filter_FilterBtn;

	@FindBy(xpath = "//*[@id='551_0_AdvanceFilter']/table/tbody/tr/td[1]/select")
	public static WebElement reportsAdvanced_WhereDropDown;

	@FindBy(xpath = "//tbody//input[@placeholder='Select Field']")
	public static WebElement reportsAdvanced_SelectFielddropdown;

	@FindBy(xpath = "//table/tbody/tr/td[3]/select")
	public static WebElement reportsAdvanced_SelectOperatorDropDown;

	@FindBy(xpath = "//table/tbody/tr/td[4]/select")
	public static WebElement reportsAdvanced_ComparewithDropDown;

	@FindBy(xpath = "//table/tbody/tr/td[5]/input")
	public static WebElement reportsAdvanced_ValueTxt;

	@FindBy(xpath = "//input[@id='advancefilter_master_551_0']")
	public static WebElement reportsAdvanced_EnterValue;

	@FindBy(xpath = "//table/tbody/tr/td[7]/span")
	public static WebElement reportsAdvanced_AddIconBtn;

	@FindBy(xpath = "//table/tbody/tr/td[6]/span")
	public static WebElement reportsAdvanced_DeleteIcon1Btn;

	@FindBy(xpath = "//table/tbody/tr[2]/td[6]/span")
	public static WebElement reportsAdvanced__DeleteIcon2Btn;

	@FindBy(xpath = "//div[@id='filterTree_551_0_AdvanceFilter']//a[@id='536870916']")
	public static WebElement reports_AdvancedFilter_SelectField_WH_ExpandBtn;

	@FindBy(xpath = "//div[@id='filterTree_551_0_AdvanceFilter']//a[@id='5058']")
	public static WebElement reports_AdvancedFilter_SelectField_WH_NameBtn;

	@FindBy(xpath = "//input[@id='chkSaveFilter']")
	public static WebElement report_FilterSaveFilterPermanentlyChkBox;

	@FindBy(xpath = "//span[@id='filterRefresh']")
	public static WebElement report_Filter_RefreshBtn;

	@FindBy(xpath = "//tbody/tr[1]/td[6]/span[1]")
	public static WebElement report_Filter_RefreshConditionBtn;

	@FindBy(xpath = "//i[contains(@class,'icon icon-ok')]")
	public static WebElement report_FilterOkBtn;

	@FindBy(xpath = "//i[contains(@class,'icon icon-close')]")
	public static WebElement report_FilterCancelBtn;

	@FindBy(xpath = "//div[@id='dvfilter_']//span[contains(@class,'icon-down-arrow icon-font8 dropdown-toggle theme_button_color')]")
	public static WebElement report_FilterDownBtn;

	@FindBy(xpath = "//i[contains(@class,'icon-font6 icon-analyze')]")
	public static WebElement report_AnalyzeBtn;

	@FindBy(xpath = "//i[contains(@class,'icon-font6 icon-cross-reference')]")
	public static WebElement report_CrossReferenceBtn;

	@FindBy(xpath = "//i[contains(@class,'icon-font6 icon-options')]")
	public static WebElement report_OptionsBtn;

	@FindBy(xpath = "//div[@id='dvoptions_']//span[contains(@class,'icon-down-arrow icon-font8 dropdown-toggle theme_button_color')]")
	public static WebElement report_OptionsDownBtn;

	
	@FindBy(xpath = "//a[contains(@class,'clsMasterName')][contains(text(),'ITEMS GROUP')]")
	public static WebElement sl_ItemGroupsBtn;

	@FindBy(xpath = "//input[@id='liSelectAllMasters']")
	public static WebElement sl_HeaderSelectChkBox;

	@FindBy(xpath = "//i[contains(@class,'icon-font6 icon-schedule')]")
	public static WebElement sl_ScheduleBtn;

	@FindBy(xpath = "//i[contains(@class,'icon-font6 icon-custamize')]")
	public static WebElement sl_CustomizeBtn;

	@FindBy(xpath = "//i[contains(@class,'icon-font6 icon-filter')]")
	public static WebElement sl_FilterBtn;

	@FindBy(xpath = "//*[@class='icon-ok hiconright2']")
	public static WebElement sl_OkBtn;

	@FindBy(xpath = "//*[@id='id_mainreportmenuheadings']/ul/li[11]/a/i")
	public static WebElement sl_CloseBtn;
	
	@FindBy(xpath = "//*[@id='id_mainreportmenuheadings']/ul/li[10]/a/i")
	public static WebElement sl_OkBtn1;

	@FindBy(xpath = "//*[@id='id_mainreportmenuheadings']/ul/li[10]/a/i")
	public static WebElement sl_CloseBtn1;
	

	@FindBy(xpath="//*[@class='adminprofile']/a")
	public static WebElement userNameDisplayLogo;

	@FindBy(xpath = "//select[@id='DateOptions_']")
	public static WebElement sl_DateOptionDropdown;

	@FindBy(xpath = "//input[@id='id_starting_date_']")
	public static WebElement sl_StartDateTxt;

	@FindBy(xpath = "//td[@id='id_starting_date__input_image']//span[contains(@class,'icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow')]")
	public static WebElement sl_StartDateCalanderBtn;

	@FindBy(xpath = "//input[@id='id_ending_date_']")
	public static WebElement sl_EndDateTxt;

	@FindBy(xpath = "//td[@id='id_ending_date__input_image']//span[contains(@class,'icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow')]")
	public static WebElement sl_EndDateCalanderBtn;

	@FindBy(xpath = "//input[@id='RITCheckbox__3']")
	public static WebElement sl_IncludeServiceTypeItemChkBox;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[8]/label/span")
	public static WebElement sl_IncludeOpeBalChkBox;
	
	
	
	@FindBy(xpath = "//select[@id='RITLayout_']")
	public static WebElement sl_LayoutDropdown;

	@FindBy(xpath = "//select[@id='RITOutput_']")
	public static WebElement sl_OutputDropdown;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[1]/div/label/span")
	public static WebElement sl_SelectAllItemsChkBox;

	@FindBy(xpath = "//div[contains(@class,'first')]//input[contains(@class,'btn btn-xs Fbutton')]")
	public static WebElement sl_FirstBtn;

	@FindBy(xpath = "//div[contains(@class,'prev')]//input[contains(@class,'btn btn-xs Fbutton')]")
	public static WebElement sl_PreviousBtn;

	@FindBy(xpath = "//input[@id='btn1']")
	public static WebElement sl_PageBtn;

	@FindBy(xpath = "//div[contains(@class,'next')]//input[contains(@class,'btn btn-xs Fbutton')]")
	public static WebElement sl_NextBtn;

	@FindBy(xpath = "//div[contains(@class,'End')]//input[contains(@class,'btn btn-xs Fbutton')]")
	public static WebElement sl_EndBtn;

	@FindBy(xpath = "//tr[1]//td[8]//div[1]//label[1]")
	public static WebElement sl_1stRowChkBox;

	@FindBy(xpath = "//tr[2]//td[8]//div[1]//label[1]")
	public static WebElement sl_2ndRowChkBox;

	////////

	@FindBy(xpath = "//span[@id='reportSort']")
	public static WebElement report_sortingBtn;

	@FindBy(xpath = "//span[@id='reportRefresh']")
	public static WebElement report_RefreshBtn;

	@FindBy(xpath = "//span[@id='print_report_']")
	public static WebElement report_PrintBtn;

	@FindBy(xpath = "//*[@id='reportRenderControls']/ul/li/span[4]")
	public static WebElement report_ExportBtn;

	
	
	@FindBy(xpath = "//*[@id='trRender_0']/td[1]")
	public static WebElement sl_1stRow1stCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[2]")
	public static WebElement sl_1stRow2ndCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[3]")
	public static WebElement sl_1stRow3rdCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[4]")
	public static WebElement sl_1stRow4thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[5]")
	public static WebElement sl_1stRow5thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[6]")
	public static WebElement sl_1stRow6thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[7]")
	public static WebElement sl_1stRow7thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[8]")
	public static WebElement sl_1stRow8thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[9]")
	public static WebElement sl_1stRow9thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[10]")
	public static WebElement sl_1stRow10thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[11]")
	public static WebElement sl_1stRow11thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[12]")
	public static WebElement sl_1stRow12thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[13]")
	public static WebElement sl_1stRow13thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[14]")
	public static WebElement sl_1stRow14thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[15]")
	public static WebElement sl_1stRow15thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[16]")
	public static WebElement sl_1stRow16thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[17]")
	public static WebElement sl_1stRow17thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[18]")
	public static WebElement sl_1stRow18thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[19]")
	public static WebElement sl_1stRow19thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[20]")
	public static WebElement sl_1stRow20thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[21]")
	public static WebElement sl_1stRow21thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[22]")
	public static WebElement sl_1stRow22thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[23]")
	public static WebElement sl_1stRow23thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[24]")
	public static WebElement sl_1stRow24thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[25]")
	public static WebElement sl_1stRow25thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[26]")
	public static WebElement sl_1stRow26thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[27]")
	public static WebElement sl_1stRow27thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[28]")
	public static WebElement sl_1stRow28thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[29]")
	public static WebElement sl_1stRow29thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[30]")
	public static WebElement sl_1stRow30thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[31]")
	public static WebElement sl_1stRow31thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[32]")
	public static WebElement sl_1stRow32thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[33]")
	public static WebElement sl_1stRow33thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[34]")
	public static WebElement sl_1stRow34thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[35]")
	public static WebElement sl_1stRow35thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[36]")
	public static WebElement sl_1stRow36thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[37]")
	public static WebElement sl_1stRow37thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[38]")
	public static WebElement sl_1stRow38thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[39]")
	public static WebElement sl_1stRow39thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[40]")
	public static WebElement sl_1stRow40thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[41]")
	public static WebElement sl_1stRow41thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[42]")
	public static WebElement sl_1stRow42thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[43]")
	public static WebElement sl_1stRow43thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[44]")
	public static WebElement sl_1stRow44thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[45]")
	public static WebElement sl_1stRow45thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[46]")
	public static WebElement sl_1stRow46thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[47]")
	public static WebElement sl_1stRow47thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[48]")
	public static WebElement sl_1stRow48thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[49]")
	public static WebElement sl_1stRow49thCol;

	@FindBy(xpath = "//*[@id='trRender_0']/td[50]")
	public static WebElement sl_1stRow50thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[1]")
	public static WebElement sl_2ndRow1stCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[2]")
	public static WebElement sl_2ndRow2ndCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[3]")
	public static WebElement sl_2ndRow3rdCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[4]")
	public static WebElement sl_2ndRow4thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[5]")
	public static WebElement sl_2ndRow5thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[6]")
	public static WebElement sl_2ndRow6thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[7]")
	public static WebElement sl_2ndRow7thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[8]")
	public static WebElement sl_2ndRow8thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[9]")
	public static WebElement sl_2ndRow9thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[10]")
	public static WebElement sl_2ndRow10thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[11]")
	public static WebElement sl_2ndRow11thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[12]")
	public static WebElement sl_2ndRow12thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[13]")
	public static WebElement sl_2ndRow13thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[14]")
	public static WebElement sl_2ndRow14thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[15]")
	public static WebElement sl_2ndRow15thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[16]")
	public static WebElement sl_2ndRow16thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[17]")
	public static WebElement sl_2ndRow17thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[18]")
	public static WebElement sl_2ndRow18thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[19]")
	public static WebElement sl_2ndRow19thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[20]")
	public static WebElement sl_2ndRow20thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[21]")
	public static WebElement sl_2ndRow21thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[22]")
	public static WebElement sl_2ndRow22thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[23]")
	public static WebElement sl_2ndRow23thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[24]")
	public static WebElement sl_2ndRow24thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[25]")
	public static WebElement sl_2ndRow25thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[26]")
	public static WebElement sl_2ndRow26thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[27]")
	public static WebElement sl_2ndRow27thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[28]")
	public static WebElement sl_2ndRow28thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[29]")
	public static WebElement sl_2ndRow29thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[30]")
	public static WebElement sl_2ndRow30thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[31]")
	public static WebElement sl_2ndRow31thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[32]")
	public static WebElement sl_2ndRow32thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[33]")
	public static WebElement sl_2ndRow33thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[34]")
	public static WebElement sl_2ndRow34thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[35]")
	public static WebElement sl_2ndRow35thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[36]")
	public static WebElement sl_2ndRow36thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[37]")
	public static WebElement sl_2ndRow37thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[38]")
	public static WebElement sl_2ndRow38thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[39]")
	public static WebElement sl_2ndRow39thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[40]")
	public static WebElement sl_2ndRow40thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[41]")
	public static WebElement sl_2ndRow41thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[42]")
	public static WebElement sl_2ndRow42thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[43]")
	public static WebElement sl_2ndRow43thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[44]")
	public static WebElement sl_2ndRow44thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[45]")
	public static WebElement sl_2ndRow45thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[46]")
	public static WebElement sl_2ndRow46thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[47]")
	public static WebElement sl_2ndRow47thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[48]")
	public static WebElement sl_2ndRow48thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[49]")
	public static WebElement sl_2ndRow49thCol;

	@FindBy(xpath = "//*[@id='trRender_1']/td[50]")
	public static WebElement sl_2ndRow50thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[1]")
	public static WebElement sl_3rdRow1stCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[2]")
	public static WebElement sl_3rdRow2ndCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[3]")
	public static WebElement sl_3rdRow3rdCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[4]")
	public static WebElement sl_3rdRow4thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[5]")
	public static WebElement sl_3rdRow5thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[6]")
	public static WebElement sl_3rdRow6thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[7]")
	public static WebElement sl_3rdRow7thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[8]")
	public static WebElement sl_3rdRow8thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[9]")
	public static WebElement sl_3rdRow9thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[10]")
	public static WebElement sl_3rdRow10thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[11]")
	public static WebElement sl_3rdRow11thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[12]")
	public static WebElement sl_3rdRow12thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[13]")
	public static WebElement sl_3rdRow13thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[14]")
	public static WebElement sl_3rdRow14thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[15]")
	public static WebElement sl_3rdRow15thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[16]")
	public static WebElement sl_3rdRow16thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[17]")
	public static WebElement sl_3rdRow17thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[18]")
	public static WebElement sl_3rdRow18thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[19]")
	public static WebElement sl_3rdRow19thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[20]")
	public static WebElement sl_3rdRow20thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[21]")
	public static WebElement sl_3rdRow21thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[22]")
	public static WebElement sl_3rdRow22thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[23]")
	public static WebElement sl_3rdRow23thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[24]")
	public static WebElement sl_3rdRow24thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[25]")
	public static WebElement sl_3rdRow25thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[26]")
	public static WebElement sl_3rdRow26thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[27]")
	public static WebElement sl_3rdRow27thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[28]")
	public static WebElement sl_3rdRow28thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[29]")
	public static WebElement sl_3rdRow29thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[30]")
	public static WebElement sl_3rdRow30thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[31]")
	public static WebElement sl_3rdRow31thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[32]")
	public static WebElement sl_3rdRow32thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[33]")
	public static WebElement sl_3rdRow33thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[34]")
	public static WebElement sl_3rdRow34thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[35]")
	public static WebElement sl_3rdRow35thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[36]")
	public static WebElement sl_3rdRow36thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[37]")
	public static WebElement sl_3rdRow37thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[38]")
	public static WebElement sl_3rdRow38thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[39]")
	public static WebElement sl_3rdRow39thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[40]")
	public static WebElement sl_3rdRow40thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[41]")
	public static WebElement sl_3rdRow41thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[42]")
	public static WebElement sl_3rdRow42thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[43]")
	public static WebElement sl_3rdRow43thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[44]")
	public static WebElement sl_3rdRow44thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[45]")
	public static WebElement sl_3rdRow45thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[46]")
	public static WebElement sl_3rdRow46thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[47]")
	public static WebElement sl_3rdRow47thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[48]")
	public static WebElement sl_3rdRow48thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[49]")
	public static WebElement sl_3rdRow49thCol;

	@FindBy(xpath = "//*[@id='trRender_2']/td[50]")
	public static WebElement sl_3rdRow50thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[1]")
	public static WebElement sl_4thRow1stCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[2]")
	public static WebElement sl_4thRow2ndCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[3]")
	public static WebElement sl_4thRow3rdCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[4]")
	public static WebElement sl_4thRow4thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[5]")
	public static WebElement sl_4thRow5thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[6]")
	public static WebElement sl_4thRow6thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[7]")
	public static WebElement sl_4thRow7thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[8]")
	public static WebElement sl_4thRow8thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[9]")
	public static WebElement sl_4thRow9thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[10]")
	public static WebElement sl_4thRow10thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[11]")
	public static WebElement sl_4thRow11thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[12]")
	public static WebElement sl_4thRow12thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[13]")
	public static WebElement sl_4thRow13thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[14]")
	public static WebElement sl_4thRow14thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[15]")
	public static WebElement sl_4thRow15thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[16]")
	public static WebElement sl_4thRow16thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[17]")
	public static WebElement sl_4thRow17thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[18]")
	public static WebElement sl_4thRow18thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[19]")
	public static WebElement sl_4thRow19thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[20]")
	public static WebElement sl_4thRow20thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[21]")
	public static WebElement sl_4thRow21thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[22]")
	public static WebElement sl_4thRow22thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[23]")
	public static WebElement sl_4thRow23thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[24]")
	public static WebElement sl_4thRow24thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[25]")
	public static WebElement sl_4thRow25thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[26]")
	public static WebElement sl_4thRow26thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[27]")
	public static WebElement sl_4thRow27thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[28]")
	public static WebElement sl_4thRow28thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[29]")
	public static WebElement sl_4thRow29thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[30]")
	public static WebElement sl_4thRow30thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[31]")
	public static WebElement sl_4thRow31thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[32]")
	public static WebElement sl_4thRow32thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[33]")
	public static WebElement sl_4thRow33thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[34]")
	public static WebElement sl_4thRow34thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[35]")
	public static WebElement sl_4thRow35thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[36]")
	public static WebElement sl_4thRow36thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[37]")
	public static WebElement sl_4thRow37thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[38]")
	public static WebElement sl_4thRow38thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[39]")
	public static WebElement sl_4thRow39thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[40]")
	public static WebElement sl_4thRow40thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[41]")
	public static WebElement sl_4thRow41thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[42]")
	public static WebElement sl_4thRow42thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[43]")
	public static WebElement sl_4thRow43thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[44]")
	public static WebElement sl_4thRow44thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[45]")
	public static WebElement sl_4thRow45thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[46]")
	public static WebElement sl_4thRow46thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[47]")
	public static WebElement sl_4thRow47thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[48]")
	public static WebElement sl_4thRow48thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[49]")
	public static WebElement sl_4thRow49thCol;

	@FindBy(xpath = "//*[@id='trRender_3']/td[50]")
	public static WebElement sl_4thRow50thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[1]")
	public static WebElement sl_5thRow1stCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[2]")
	public static WebElement sl_5thRow2ndCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[3]")
	public static WebElement sl_5thRow3rdCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[4]")
	public static WebElement sl_5thRow4thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[5]")
	public static WebElement sl_5thRow5thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[6]")
	public static WebElement sl_5thRow6thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[7]")
	public static WebElement sl_5thRow7thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[8]")
	public static WebElement sl_5thRow8thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[9]")
	public static WebElement sl_5thRow9thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[10]")
	public static WebElement sl_5thRow10thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[11]")
	public static WebElement sl_5thRow11thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[12]")
	public static WebElement sl_5thRow12thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[13]")
	public static WebElement sl_5thRow13thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[14]")
	public static WebElement sl_5thRow14thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[15]")
	public static WebElement sl_5thRow15thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[16]")
	public static WebElement sl_5thRow16thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[17]")
	public static WebElement sl_5thRow17thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[18]")
	public static WebElement sl_5thRow18thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[19]")
	public static WebElement sl_5thRow19thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[20]")
	public static WebElement sl_5thRow20thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[21]")
	public static WebElement sl_5thRow21thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[22]")
	public static WebElement sl_5thRow22thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[23]")
	public static WebElement sl_5thRow23thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[24]")
	public static WebElement sl_5thRow24thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[25]")
	public static WebElement sl_5thRow25thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[26]")
	public static WebElement sl_5thRow26thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[27]")
	public static WebElement sl_5thRow27thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[28]")
	public static WebElement sl_5thRow28thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[29]")
	public static WebElement sl_5thRow29thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[30]")
	public static WebElement sl_5thRow30thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[31]")
	public static WebElement sl_5thRow31thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[32]")
	public static WebElement sl_5thRow32thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[33]")
	public static WebElement sl_5thRow33thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[34]")
	public static WebElement sl_5thRow34thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[35]")
	public static WebElement sl_5thRow35thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[36]")
	public static WebElement sl_5thRow36thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[37]")
	public static WebElement sl_5thRow37thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[38]")
	public static WebElement sl_5thRow38thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[39]")
	public static WebElement sl_5thRow39thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[40]")
	public static WebElement sl_5thRow40thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[41]")
	public static WebElement sl_5thRow41thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[42]")
	public static WebElement sl_5thRow42thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[43]")
	public static WebElement sl_5thRow43thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[44]")
	public static WebElement sl_5thRow44thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[45]")
	public static WebElement sl_5thRow45thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[46]")
	public static WebElement sl_5thRow46thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[47]")
	public static WebElement sl_5thRow47thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[48]")
	public static WebElement sl_5thRow48thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[49]")
	public static WebElement sl_5thRow49thCol;

	@FindBy(xpath = "//*[@id='trRender_4']/td[50]")
	public static WebElement sl_5thRow50thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[1]")
	public static WebElement sl_6thRow1stCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[2]")
	public static WebElement sl_6thRow2ndCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[3]")
	public static WebElement sl_6thRow3rdCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[4]")
	public static WebElement sl_6thRow4thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[5]")
	public static WebElement sl_6thRow5thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[6]")
	public static WebElement sl_6thRow6thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[7]")
	public static WebElement sl_6thRow7thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[8]")
	public static WebElement sl_6thRow8thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[9]")
	public static WebElement sl_6thRow9thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[10]")
	public static WebElement sl_6thRow10thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[11]")
	public static WebElement sl_6thRow11thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[12]")
	public static WebElement sl_6thRow12thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[13]")
	public static WebElement sl_6thRow13thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[14]")
	public static WebElement sl_6thRow14thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[15]")
	public static WebElement sl_6thRow15thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[16]")
	public static WebElement sl_6thRow16thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[17]")
	public static WebElement sl_6thRow17thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[18]")
	public static WebElement sl_6thRow18thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[19]")
	public static WebElement sl_6thRow19thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[20]")
	public static WebElement sl_6thRow20thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[21]")
	public static WebElement sl_6thRow21thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[22]")
	public static WebElement sl_6thRow22thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[23]")
	public static WebElement sl_6thRow23thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[24]")
	public static WebElement sl_6thRow24thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[25]")
	public static WebElement sl_6thRow25thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[26]")
	public static WebElement sl_6thRow26thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[27]")
	public static WebElement sl_6thRow27thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[28]")
	public static WebElement sl_6thRow28thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[29]")
	public static WebElement sl_6thRow29thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[30]")
	public static WebElement sl_6thRow30thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[31]")
	public static WebElement sl_6thRow31thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[32]")
	public static WebElement sl_6thRow32thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[33]")
	public static WebElement sl_6thRow33thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[34]")
	public static WebElement sl_6thRow34thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[35]")
	public static WebElement sl_6thRow35thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[36]")
	public static WebElement sl_6thRow36thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[37]")
	public static WebElement sl_6thRow37thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[38]")
	public static WebElement sl_6thRow38thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[39]")
	public static WebElement sl_6thRow39thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[40]")
	public static WebElement sl_6thRow40thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[41]")
	public static WebElement sl_6thRow41thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[42]")
	public static WebElement sl_6thRow42thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[43]")
	public static WebElement sl_6thRow43thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[44]")
	public static WebElement sl_6thRow44thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[45]")
	public static WebElement sl_6thRow45thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[46]")
	public static WebElement sl_6thRow46thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[47]")
	public static WebElement sl_6thRow47thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[48]")
	public static WebElement sl_6thRow48thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[49]")
	public static WebElement sl_6thRow49thCol;

	@FindBy(xpath = "//*[@id='trRender_5']/td[50]")
	public static WebElement sl_6thRow50thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[1]")
	public static WebElement sl_7thRow1stCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[2]")
	public static WebElement sl_7thRow2ndCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[3]")
	public static WebElement sl_7thRow3rdCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[4]")
	public static WebElement sl_7thRow4thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[5]")
	public static WebElement sl_7thRow5thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[6]")
	public static WebElement sl_7thRow6thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[7]")
	public static WebElement sl_7thRow7thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[8]")
	public static WebElement sl_7thRow8thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[9]")
	public static WebElement sl_7thRow9thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[10]")
	public static WebElement sl_7thRow10thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[11]")
	public static WebElement sl_7thRow11thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[12]")
	public static WebElement sl_7thRow12thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[13]")
	public static WebElement sl_7thRow13thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[14]")
	public static WebElement sl_7thRow14thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[15]")
	public static WebElement sl_7thRow15thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[16]")
	public static WebElement sl_7thRow16thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[17]")
	public static WebElement sl_7thRow17thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[18]")
	public static WebElement sl_7thRow18thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[19]")
	public static WebElement sl_7thRow19thCol;

	@FindBy(xpath = "//*[@id='trRender_6']/td[20]")
	public static WebElement sl_7thRow20thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[1]")
	public static WebElement sl_8thRow1stCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[2]")
	public static WebElement sl_8thRow2ndCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[3]")
	public static WebElement sl_8thRow3rdCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[4]")
	public static WebElement sl_8thRow4thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[5]")
	public static WebElement sl_8thRow5thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[6]")
	public static WebElement sl_8thRow6thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[7]")
	public static WebElement sl_8thRow7thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[8]")
	public static WebElement sl_8thRow8thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[9]")
	public static WebElement sl_8thRow9thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[10]")
	public static WebElement sl_8thRow10thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[11]")
	public static WebElement sl_8thRow11thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[12]")
	public static WebElement sl_8thRow12thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[13]")
	public static WebElement sl_8thRow13thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[14]")
	public static WebElement sl_8thRow14thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[15]")
	public static WebElement sl_8thRow15thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[16]")
	public static WebElement sl_8thRow16thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[17]")
	public static WebElement sl_8thRow17thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[18]")
	public static WebElement sl_8thRow18thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[19]")
	public static WebElement sl_8thRow19thCol;

	@FindBy(xpath = "//*[@id='trRender_7']/td[20]")
	public static WebElement sl_8thRow20thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[1]")
	public static WebElement sl_9thRow1stCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[2]")
	public static WebElement sl_9thRow2ndCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[3]")
	public static WebElement sl_9thRow3rdCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[4]")
	public static WebElement sl_9thRow4thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[5]")
	public static WebElement sl_9thRow5thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[6]")
	public static WebElement sl_9thRow6thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[7]")
	public static WebElement sl_9thRow7thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[8]")
	public static WebElement sl_9thRow8thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[9]")
	public static WebElement sl_9thRow9thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[10]")
	public static WebElement sl_9thRow10thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[11]")
	public static WebElement sl_9thRow11thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[12]")
	public static WebElement sl_9thRow12thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[13]")
	public static WebElement sl_9thRow13thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[14]")
	public static WebElement sl_9thRow14thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[15]")
	public static WebElement sl_9thRow15thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[16]")
	public static WebElement sl_9thRow16thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[17]")
	public static WebElement sl_9thRow17thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[18]")
	public static WebElement sl_9thRow18thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[19]")
	public static WebElement sl_9thRow19thCol;

	@FindBy(xpath = "//*[@id='trRender_8']/td[20]")
	public static WebElement sl_9thRow20thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[1]")
	public static WebElement sl_10thRow1stCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[2]")
	public static WebElement sl_10thRow2ndCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[3]")
	public static WebElement sl_10thRow3rdCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[4]")
	public static WebElement sl_10thRow4thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[5]")
	public static WebElement sl_10thRow5thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[6]")
	public static WebElement sl_10thRow6thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[7]")
	public static WebElement sl_10thRow7thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[8]")
	public static WebElement sl_10thRow8thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[9]")
	public static WebElement sl_10thRow9thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[10]")
	public static WebElement sl_10thRow10thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[11]")
	public static WebElement sl_10thRow11thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[12]")
	public static WebElement sl_10thRow12thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[13]")
	public static WebElement sl_10thRow13thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[14]")
	public static WebElement sl_10thRow14thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[15]")
	public static WebElement sl_10thRow15thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[16]")
	public static WebElement sl_10thRow16thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[17]")
	public static WebElement sl_10thRow17thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[18]")
	public static WebElement sl_10thRow18thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[19]")
	public static WebElement sl_10thRow19thCol;

	@FindBy(xpath = "//*[@id='trRender_9']/td[20]")
	public static WebElement sl_10thRow20thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[1]")
	public static WebElement sl_11thRow1stCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[2]")
	public static WebElement sl_11thRow2ndCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[3]")
	public static WebElement sl_11thRow3rdCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[4]")
	public static WebElement sl_11thRow4thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[5]")
	public static WebElement sl_11thRow5thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[6]")
	public static WebElement sl_11thRow6thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[7]")
	public static WebElement sl_11thRow7thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[8]")
	public static WebElement sl_11thRow8thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[9]")
	public static WebElement sl_11thRow9thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[10]")
	public static WebElement sl_11thRow10thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[11]")
	public static WebElement sl_11thRow11thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[12]")
	public static WebElement sl_11thRow12thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[13]")
	public static WebElement sl_11thRow13thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[14]")
	public static WebElement sl_11thRow14thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[15]")
	public static WebElement sl_11thRow15thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[16]")
	public static WebElement sl_11thRow16thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[17]")
	public static WebElement sl_11thRow17thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[18]")
	public static WebElement sl_11thRow18thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[19]")
	public static WebElement sl_11thRow19thCol;

	@FindBy(xpath = "//*[@id='trRender_10']/td[20]")
	public static WebElement sl_11thRow20thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[1]")
	public static WebElement sl_12thRow1stCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[2]")
	public static WebElement sl_12thRow2ndCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[3]")
	public static WebElement sl_12thRow3rdCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[4]")
	public static WebElement sl_12thRow4thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[5]")
	public static WebElement sl_12thRow5thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[6]")
	public static WebElement sl_12thRow6thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[7]")
	public static WebElement sl_12thRow7thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[8]")
	public static WebElement sl_12thRow8thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[9]")
	public static WebElement sl_12thRow9thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[10]")
	public static WebElement sl_12thRow10thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[11]")
	public static WebElement sl_12thRow11thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[12]")
	public static WebElement sl_12thRow12thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[13]")
	public static WebElement sl_12thRow13thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[14]")
	public static WebElement sl_12thRow14thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[15]")
	public static WebElement sl_12thRow15thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[16]")
	public static WebElement sl_12thRow16thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[17]")
	public static WebElement sl_12thRow17thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[18]")
	public static WebElement sl_12thRow18thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[19]")
	public static WebElement sl_12thRow19thCol;

	@FindBy(xpath = "//*[@id='trRender_11']/td[20]")
	public static WebElement sl_12thRow20thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[1]")
	public static WebElement sl_13thRow1stCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[2]")
	public static WebElement sl_13thRow2ndCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[3]")
	public static WebElement sl_13thRow3rdCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[4]")
	public static WebElement sl_13thRow4thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[5]")
	public static WebElement sl_13thRow5thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[6]")
	public static WebElement sl_13thRow6thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[7]")
	public static WebElement sl_13thRow7thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[8]")
	public static WebElement sl_13thRow8thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[9]")
	public static WebElement sl_13thRow9thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[10]")
	public static WebElement sl_13thRow10thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[11]")
	public static WebElement sl_13thRow11thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[12]")
	public static WebElement sl_13thRow12thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[13]")
	public static WebElement sl_13thRow13thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[14]")
	public static WebElement sl_13thRow14thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[15]")
	public static WebElement sl_13thRow15thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[16]")
	public static WebElement sl_13thRow16thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[17]")
	public static WebElement sl_13thRow17thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[18]")
	public static WebElement sl_13thRow18thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[19]")
	public static WebElement sl_13thRow19thCol;

	@FindBy(xpath = "//*[@id='trRender_12']/td[20]")
	public static WebElement sl_13thRow20thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[1]")
	public static WebElement sl_14thRow1stCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[2]")
	public static WebElement sl_14thRow2ndCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[3]")
	public static WebElement sl_14thRow3rdCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[4]")
	public static WebElement sl_14thRow4thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[5]")
	public static WebElement sl_14thRow5thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[6]")
	public static WebElement sl_14thRow6thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[7]")
	public static WebElement sl_14thRow7thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[8]")
	public static WebElement sl_14thRow8thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[9]")
	public static WebElement sl_14thRow9thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[10]")
	public static WebElement sl_14thRow10thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[11]")
	public static WebElement sl_14thRow11thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[12]")
	public static WebElement sl_14thRow12thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[13]")
	public static WebElement sl_14thRow13thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[14]")
	public static WebElement sl_14thRow14thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[15]")
	public static WebElement sl_14thRow15thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[16]")
	public static WebElement sl_14thRow16thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[17]")
	public static WebElement sl_14thRow17thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[18]")
	public static WebElement sl_14thRow18thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[19]")
	public static WebElement sl_14thRow19thCol;

	@FindBy(xpath = "//*[@id='trRender_13']/td[20]")
	public static WebElement sl_14thRow20thCol;

	@FindBy(xpath = "//*[@id='trRender_14']/td[1]")
	public static WebElement sl_15thRow1stCol;

	@FindBy(xpath = "//*[@id='trRender_14']/td[2]")
	public static WebElement sl_15thRow2ndCol;

	@FindBy(xpath = "//*[@id='trRender_14']/td[3]")
	public static WebElement sl_15thRow3rdCol;

	@FindBy(xpath = "//*[@id='trRender_14']/td[4]")
	public static WebElement sl_15thRow4thCol;

	@FindBy(xpath = "//*[@id='trRender_14']/td[5]")
	public static WebElement sl_15thRow5thCol;

	@FindBy(xpath = "//*[@id='trRender_14']/td[6]")
	public static WebElement sl_15thRow6thCol;

	@FindBy(xpath = "//*[@id='trRender_14']/td[7]")
	public static WebElement sl_15thRow7thCol;

	@FindBy(xpath = "//*[@id='trRender_14']/td[8]")
	public static WebElement sl_15thRow8thCol;

	@FindBy(xpath = "//*[@id='trRender_14']/td[9]")
	public static WebElement sl_15thRow9thCol;

	@FindBy(xpath = "//*[@id='trRender_14']/td[10]")
	public static WebElement sl_15thRow10thCol;

	@FindBy(xpath = "//*[@id='trRender_14']/td[11]")
	public static WebElement sl_15thRow11thCol;

	@FindBy(xpath = "//*[@id='trRender_14']/td[12]")
	public static WebElement sl_15thRow12thCol;

	public static String xlSheetName = "SmokeFAReports";
	private int bankRecRow1ListCount;

	@FindBy(xpath = "//*[@id='navigation_menu']/li[2]/ul/li[3]/ul/li/a/span")
	public static List<WebElement> financialReportList;

	
	@FindBy(xpath = "//*[@id='id_header_3_table_body']/tr")
	public static List<WebElement> salesAccList;
	
	
	
	/*@FindBy(xpath = "//div[@id='idGlobalError']")
	public static WebElement validationConfirmationMessage;*/

	@FindBy(xpath = "//tbody[@id='LandingGridBody']/tr")
	public static List<WebElement> stockLedgerHometableRowCount;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr/td[3]")
	public static List<WebElement> reportVocGridList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr")
	public static List<WebElement> reportRowList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr[1]/td")
	public static List<WebElement> report1stRowList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr[2]/td")
	public static List<WebElement> report2ndRowList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr[3]/td")
	public static List<WebElement> report3rdRowList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr[4]/td")
	public static List<WebElement> report4thRowList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr[5]/td")
	public static List<WebElement> report5thRowList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr[6]/td")
	public static List<WebElement> report6thRowList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr[7]/td")
	public static List<WebElement> report7thRowList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr[8]/td")
	public static List<WebElement> report8thRowList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr[9]/td")
	public static List<WebElement> report9thRowList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr[10]/td")
	public static List<WebElement> report10thRowList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr[11]/td")
	public static List<WebElement> report11thRowList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr[12]/td")
	public static List<WebElement> report12thRowList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr[13]/td")
	public static List<WebElement> report13thRowList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr[14]/td")
	public static List<WebElement> report14thRowList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr[15]/td")
	public static List<WebElement> report15thRowList;

	
	@FindBy(xpath = "//*[@id='id_reportmenudisplay']/ul/li[13]")
	public static WebElement report_CloseBtn;

	
	@FindBy(xpath = "//div[@id='tblFooterReportRender']/div[2]/button[4]")
	public static WebElement report_NextBtn;
	
	@FindBy(xpath = "//div[@id='tblFooterReportRender']/div[2]/button[2]")
	public static WebElement report_PreviousBtn;

	@FindBy(xpath = "//div[@id='tblFooterReportRender']/div[2]/button[3]")
	public static WebElement report_PageNoBtn;


	
	
	// Cheque Discounting

	/*@FindBy(xpath = "//input[@id='optnCtrlCheqDisBank']")
	public static WebElement bankAccountTxt;*/

	@FindBy(xpath = "//input[@id='optctrlCheqDisCustomer']")
	public static WebElement CD_customerAccountTxt;

	@FindBy(xpath = "//input[@id='DepositingBank']")
	public static WebElement discountingBankTxt;

	@FindBy(xpath = "//i[@class='icon-reset icon-font6']")
	public static WebElement cd_LoadBtn;

	@FindBy(xpath = "//input[@id='txtMarginPercentageId']")
	public static WebElement cd_MarginTxt;

	@FindBy(xpath = "//input[@id='btnDepositingBank']")
	public static WebElement cd_ApplyBtn;

	@FindBy(xpath = "//i[@class='icon-clear icon-font6']")
	public static WebElement cd_clearBtn;

	@FindBy(xpath = "//*[@id='btnSave']/a/i")
	public static WebElement cd_saveBtn;

	@FindBy(xpath = "//*[@id='btnClose']/i")
	public static WebElement cd_CloseBtn;

	@FindBy(xpath = "//thead[@id='ChequeDiscountingTable_head']/tr/th/div")
	public static List<WebElement> cd_tableHeadeList;

	@FindBy(xpath = "//td[@id='ChequeDiscountingTable_col_1-2']")
	public static WebElement disCustomerRow1;

	@FindBy(xpath = "//td[@id='ChequeDiscountingTable_col_1-3']")
	public static WebElement disVoucherRow1;

	@FindBy(xpath = "//td[@id='ChequeDiscountingTable_col_1-4']")
	public static WebElement disMaturityDateRow1;

	@FindBy(xpath = "//td[@id='ChequeDiscountingTable_col_1-5']")
	public static WebElement disChequeNumberRow1;

	@FindBy(xpath = "//td[@id='ChequeDiscountingTable_col_1-6']")
	public static WebElement disAmountRow1;

	@FindBy(xpath = "//td[@id='ChequeDiscountingTable_col_1-7']")
	public static WebElement disDiscountlimitRow1;

	@FindBy(xpath = "//td[@id='ChequeDiscountingTable_col_1-8']")
	public static WebElement disDiscountAmtRow1;

	@FindBy(xpath = "//td[@id='ChequeDiscountingTable_col_1-9']")
	public static WebElement disBankRow1;

	@FindBy(xpath = "//td[@id='ChequeDiscountingTable_col_1-10']")
	public static WebElement disMarginRow1;

	@FindBy(xpath = "//td[@id='ChequeDiscountingTable_col_1-11']")
	public static WebElement disPostOnDateRow1;

	@FindBy(xpath = "//*[@id='optnCtrlCheqDisBank_table_body']/tr/td[2]")
	public static List<WebElement> chequeDisList;

	@FindBy(xpath = "//*[@id='ChequeDiscountingTable_body']/tr[1]/td")
	public static List<WebElement> chequeDisBodyGridList;

	@FindBy(xpath = "//*[@id='ChequeDiscountingTable_body']")
	public static WebElement chequeDisBodyGrid;
	
	
	@FindBy(xpath = "//select[@id='RITCombobox__3']")
	public static WebElement considerDatesBasedOnDropDown;
	
	
	@FindBy(xpath = "//input[@id='FOption_517_0_DefaultFilter_0']")
	public static WebElement enterJEDefaultItemTxt;

	@FindBy(xpath = "/html/body/section/div[1]/div/div/div[2]/ul/li[4]/a/i")
	public static WebElement filteRITEMExpandBtn;

	@FindBy(xpath = "//*[@id='FilterFields_209_0']/li[3]/a")
	public static WebElement filterSalesItemExpandBtn;

	@FindBy(xpath = "/html/body/section/div[1]/div/div/div[2]/ul/ul[4]/li[1]/div/label/span")
	public static WebElement filterItemNameChkbox;

	@FindBy(xpath = "//*[@id='dvReportDetails']/div/table/tbody/tr/td[2]")
	public static List<WebElement> reportDateList;
	
	@FindBy(xpath = "//input[@id='FOption_674_0_DefaultFilter_0']")
	public static WebElement enterJEDefaultAccTxt;

	@FindBy(xpath = "//input[@id='MasterGroup__101']")
	public static WebElement reportaccountTxt;

	@FindBy(xpath = "//*[@id='MasterGroup__101_table_body']/tr/td[2]")
	public static List<WebElement> reportaccountTxtList;

	@FindBy(xpath = "//select[@id='RITCombobox__1']")
	public static WebElement includePdcDropDown;

	@FindBy(xpath = "//select[@id='RITCombobox__5']")
	public static WebElement sortingDropDown;

	@FindBy(xpath = "//select[@id='RITCombobox__10']")
	public static WebElement includeNonAuthorizedDataDropDown;

	@FindBy(xpath = "//input[@id='RITCheckbox__6']")
	public static WebElement displayUnRealizedLossorGainChkBox;

	@FindBy(xpath = "//input[@id='RITCheckbox__4']")
	public static WebElement ignoreConvertedPdcChkBox;

	@FindBy(xpath = "//input[@id='FOption_531_0_DefaultFilter_0']")
	public static WebElement trailBalFilterAccTxt;

	@FindBy(xpath = "(//a[contains(text(),'Account')]/i[1])[1]")
	public static WebElement filterAccountExpandBtn;

	@FindBy(xpath = "//*[@id='FilterFields_674_0']/ul[1]/li[1]/div/label/span")
	public static WebElement filterAccNameChkbox;

	@FindBy(xpath = "//*[@id='FilterFields_534_0']/ul[1]/li[1]/div/label/span")
	public static WebElement tradingAccProfitAndLossNameChkbox;

	@FindBy(xpath = "(//*[@class='icon-close hiconright2'])[1]")
	public static WebElement homepageCloseBtn;

	
	

	@FindBy(xpath = "//*[@id='FilterFields_531_0']/ul[1]/li[1]/div/label/span")
	public static WebElement TrailBal_filterAccNameChkbox;

	
	
	@FindBy(xpath = "/html/body/section/div[1]/div/div/div[3]/input[1]")
	public static WebElement filter_FilterOkButton;

	@FindBy(xpath = "//*[@id='FOption_513_0_DefaultFilter_0']")
	public static WebElement enterDefaultAccTxt;

	@FindBy(xpath = "//input[@id='FOption_659_0_DefaultFilter_0']")
	public static WebElement enterVATDefaultAccTxt;

	@FindBy(xpath = "//*[@id='filter_Okbtn_']")
	public static WebElement filterOkButton;

	@FindBy(xpath = "//*[@id='531']")
	public static WebElement trialBalanceReport;
	
	@FindBy(xpath = "//*[@id='532']")
	public static WebElement profitandLossReport;

	@FindBy(xpath = "//*[@id='533']")
	public static WebElement tradingAccountReport;

	@FindBy(xpath = "//*[@id='534']")
	public static WebElement tradingandProfitAndLossReport;

	@FindBy(xpath = "//*[@id='535']")
	public static WebElement balanceSheetReport;

	@FindBy(xpath = "//*[@id='592']")
	public static WebElement finalAccountSchedulesReport;

	@FindBy(xpath = "//*[@id='593']")
	public static WebElement fundFlowReport;

	@FindBy(xpath = "//*[@id='594']")
	public static WebElement cashFlowReport;

	@FindBy(xpath = "//*[@id='252']")
	public static WebElement cashFlowCalendarReport;

	@FindBy(xpath = "//*[@id='595']")
	public static WebElement cashFlowAnalysisReport;

	@FindBy(xpath = "//*[@id='632']")
	public static WebElement advanceCashFlowReport;

	@FindBy(xpath = "//*[@id='630']")
	public static WebElement incomeExpenseTrendsReport;

	@FindBy(xpath = "//*[@id='655']")
	public static WebElement receivablePayableReport;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[4]/label/span")
	public static WebElement report5chkbox;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[8]/label/span")
	public static WebElement report9chkbox;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[9]/label/span")
	public static WebElement report10chkbox;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[10]/label/span")
	public static WebElement report11chkbox;
	
	@FindBy(xpath = "//*[@id='dvReportInputs']/div[11]/label/span")
	public static WebElement report12chkbox;
	
	
	@FindBy(xpath = "//*[@id='dvReportInputs']/div[12]/label/span")
	public static WebElement report13chkbox;
	
	@FindBy(xpath = "//*[@id='dvReportInputs']/div[13]/label/span")
	public static WebElement report14chkbox;
	
	@FindBy(xpath = "//*[@id='dvReportInputs']/div[14]/label/span")
	public static WebElement report15chkbox;
	
	

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[7]/label/span")
	public static WebElement report8chkbox;

	@FindBy(xpath = "//input[@id='FOption_533_0_DefaultFilter_0']")
	public static WebElement tradingAccFilterAccTxt;

	@FindBy(xpath = "//input[@id='FOption_534_0_DefaultFilter_0']")
	public static WebElement tradingAccProfitAndLossCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_535_0_DefaultFilter_0']")
	public static WebElement balSheetCusAccTxt;
	
	
	@FindBy(xpath = "//*[@id='id_rc_complete_tree_container']/div[2]/ul/li/span")
	public static WebElement defaultExpandBtn;

	@FindBy(xpath = "//span[contains(text(),'Particulars')]")
	public static WebElement cusParticularBtn;

	@FindBy(xpath = "//*[@id='TCol1']/span")
	public static WebElement cusAccField;

	@FindBy(xpath = "//div[@id='id_focus_msgbox_main']")
	public static WebElement ServerErrorPopup;

	@FindBy(xpath = "//div[@id='id_focus_msgbox_title']/div[2]/span")
	public static WebElement ServerErrorPopupCloseBtn;

	@FindBy(xpath = "//*[@id='id_focus_msgbox_detail']/span/h2/i")
	public static WebElement ServerErrorPopupTxt;
	
	
	@FindBy(xpath = "//input[@id='FOption_592_0_DefaultFilter_0']")
	public static WebElement finalaAccScheduleAccTxt;

	
	@FindBy(xpath = "//input[@id='FOption_593_0_DefaultFilter_0']")
	public static WebElement fundFlowCUsAccTxt;

	@FindBy(xpath = "//input[@id='FOption_594_0_DefaultFilter_0']")
	public static WebElement cashFlowCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_595_0_DefaultFilter_0']")
	public static WebElement cashFlowAnalysisCusAccTxt;

	@FindBy(xpath = "//a[@id='33']")
	public static WebElement budgetPlanningReport;

	@FindBy(xpath = "//input[@id='FOption_655_0_DefaultFilter_0']")
	public static WebElement recAndPayCusAccTxt;

	// Receivable and Payable Analysis Reports Starts from Here

	@FindBy(xpath = "//a[@id='558']//span[contains(text(),'Receivable and Payable Analysis')]")
	public static WebElement receivableAndPayableAnalysisMenu;

	@FindBy(xpath = "//a[@id='605']//span[contains(text(),'Customer Detail')]")
	public static WebElement customerDetailMenu;

	@FindBy(xpath = "//span[contains(text(),'Customer Summary')]")
	public static WebElement customerSummaryMenu;

	@FindBy(xpath = "//a[@id='606']//span[contains(text(),'Vendor Detail')]")
	public static WebElement vendorDetailMenu;

	@FindBy(xpath = "//a[@id='608']//span[contains(text(),'Vendor Summary')]")
	public static WebElement vendorSummaryMenu;

	@FindBy(xpath = "//a[@id='75']//span[contains(text(),'Letter of Credit')]")
	public static WebElement letterOfCreditReport;

	@FindBy(xpath = "//span[contains(text(),'Release Letter of Credit')]")
	public static WebElement releaseLetterOfCreditReport;

	@FindBy(xpath = "//span[contains(text(),'Customer Listing of Outstanding Bills')]")
	public static WebElement customerDetailsCustomerListingOfOutstandingBillsReport;

	@FindBy(xpath = "//span[contains(text(),'Customer Statements')]")
	public static WebElement customerDetailsCustomerStatementsReport;

	@FindBy(xpath = "//span[contains(text(),'Customer Due Date Analysis')]")
	public static WebElement customerDetailsCustomerDueDateAnalysisReport;

	@FindBy(xpath = "//a[@id='563']//span[contains(text(),'Ageing Details')]")
	public static WebElement customerDetailsCustomerAgeingDetailsReport;

	@FindBy(xpath = "//span[contains(text(),'Customer Detail Ageing by Due Date')]")
	public static WebElement customerDetailsCustomerDetailAgeingByDueDateReport;

	@FindBy(xpath = "//a[@id='566']//span[contains(text(),'Overdue Analysis')]")
	public static WebElement customerDetailsCustomerOverdueAnalysisReport;

	@FindBy(xpath = "//span[contains(text(),'Customer Ageing Summary')]")
	public static WebElement customerSummaryCustomerAgeingSummaryReport;

	@FindBy(xpath = "//a[@id='565']//span[contains(text(),'Summary Ageing by Due Date')]")
	public static WebElement customerSummaryAgeingByDueDateReport;

	@FindBy(xpath = "//a[@id='567']//span[contains(text(),'Overdue Summary')]")
	public static WebElement customerSummaryCustomerOverDueSummeryReport;

	@FindBy(xpath = "//span[contains(text(),'Customer bill-wise summary')]")
	public static WebElement customerSummaryCustomerBillWiseSummeryReport;

	@FindBy(xpath = "//span[contains(text(),'Vendor Listing of Outstanding Bills')]")
	public static WebElement vendorDetailsVendorListingOfOutstandingBillsReport;

	@FindBy(xpath = "//span[contains(text(),'Vendor Statements')]")
	public static WebElement vendorDetailsVendorStatementsReport;

	@FindBy(xpath = "//span[contains(text(),'Vendor Due Date Analysis')]")
	public static WebElement vendorDetailsVendorDueDateAnalysisReport;

	@FindBy(xpath = "//span[contains(text(),'Vendor Ageing Details')]")
	public static WebElement vendorDetailsVendorAgeingDetailsReport;

	@FindBy(xpath = "//span[contains(text(),'Vendor Detail Ageing by Due Date')]")
	public static WebElement vendorDetailsVendorDetailsAgeingByDueDateReport;

	@FindBy(xpath = "//span[contains(text(),'Vendor Overdue Analysis')]")
	public static WebElement vendorDetailsVendorOverdueAnalysisReport;

	@FindBy(xpath = "//span[contains(text(),'Vendor Ageing Summary')]")
	public static WebElement vendorSummeryVendorAgeingSummaryReport;

	@FindBy(xpath = "//span[contains(text(),'Vendor Summary Ageing by Due Date')]")
	public static WebElement vendorSummeryVendorSummaryAgeingByDueDateReport;

	@FindBy(xpath = "//span[contains(text(),'Vendor Overdue Summary')]")
	public static WebElement vendorSummeryVendorOverdueSummaryReport;

	@FindBy(xpath = "//span[contains(text(),'Vendor bill-wise summary')]")
	public static WebElement vendorSummeryVendorBillWiseSummaryReport;
	
	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody")
	public static WebElement reportsTable;

	@FindBy(xpath = "//div[@id='REPORTRENDERNEWControls']/ul/li/span[6]")
	public static WebElement sl_BackTrackBtn;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr/td[3]")
	public static List<WebElement> reportTableColumn2RowsList;

	@FindBy(xpath = "//div[@id='dvReportRenderBacktracking']/div[1]/div[1]/div[2]/table/tbody/tr/td")
	public static List<WebElement> backTrackItemDetailsMonthsTableList;

	@FindBy(xpath = "//div[@id='dvReportRenderBacktracking']/div[1]/div[1]/div[2]/table")
	public static WebElement backTrackItemDetailsMonthsTable;

	@FindBy(xpath = "//div[@id='dvReportRenderBacktracking']/div[1]/div[1]/div[2]/table/tbody/tr[1]/td[1]")
	public static WebElement backTrackItemDetailsMonths1stRow1stCol;

	@FindBy(xpath = "//div[@id='dvReportRenderBacktracking']/div[1]/div[1]/div[2]/table/tbody/tr[2]/td[1]")
	public static WebElement backTrackItemDetailsMonths2ndRow1stCol;

	@FindBy(xpath = "//div[@id='dvReportRenderBacktracking']/div[2]/div[1]/div[2]/table/tbody/tr/td")
	public static List<WebElement> backTrackItemDetailsDaysTableList;

	@FindBy(xpath = "//div[@id='dvReportRenderBacktracking']/div[2]/div[1]/div[2]/table")
	public static WebElement backTrackItemDetailsDaysTable;

	@FindBy(xpath = "//div[@id='dvReportRenderBacktracking']/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[1]")
	public static WebElement backTrackItemDetailsDays1stRow1stCol;

	@FindBy(xpath = "//div[@id='dvReportRenderBacktracking']/div[3]/div[1]/div[2]/table/tbody/tr/td")
	public static List<WebElement> backTrackItemDetailsVouchersTableList;

	@FindBy(xpath = "//div[@id='dvReportRenderBacktracking']/div[3]/div[1]/div[2]/table")
	public static WebElement backTrackItemDetailsVouchersTable;

	@FindBy(xpath = "//div[@id='dvReportRenderBacktracking']/div[3]/div[1]/div[2]/table/tbody/tr[1]/td[1]")
	public static WebElement backTrackItemDetailsVouchers1stRow1stCol;

	// Voucher Entry Page Header Fields

	@FindBy(xpath = "//input[@id='id_header_4']")
	public static WebElement vcustomerAccountTxt;

	/*@FindBy(xpath = "//input[@id='id_header_268435460']")
	public static WebElement warehouseTxt;*/

	@FindBy(xpath = "//tbody[@id='id_header_268435460_table_body']/tr/td[2]")
	public static List<WebElement> warehouseHeaderComboList;

	@FindBy(xpath = "//tbody[@id='id_body_23_table_body']/tr/td[2]")
	public static List<WebElement> itemComboList;

	/*@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[18]")
	public static WebElement select1stRow_17thColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[19]")
	public static WebElement select1stRow_18thColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[20]")
	public static WebElement select1stRow_19thColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[21]")
	public static WebElement select1stRow_20thColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[22]")
	public static WebElement select1stRow_21stColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[23]")
	public static WebElement select1stRow_22ndColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[24]")
	public static WebElement select1stRow_23rdColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[25]")
	public static WebElement select1stRow_24thColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[14]")
	public static WebElement select2ndRow_13thColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[18]")
	public static WebElement select2ndRow_17thColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[19]")
	public static WebElement select2ndRow_18thColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[20]")
	public static WebElement select2ndRow_19thColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[18]")
	public static WebElement select3rdRow_17thColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[19]")
	public static WebElement select3rdRow_18thColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[20]")
	public static WebElement select3rdRow_19thColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[21]")
	public static WebElement select3rdRow_20thColumn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[22]")
	public static WebElement select3rdRow_21stColumn;
*/
	@FindBy(xpath = "//input[@id='FOption_559_0_DefaultFilter_0']")
	public static WebElement cusOutStandingBillsCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_560_0_DefaultFilter_0']")
	public static WebElement cusStatementCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_561_0_DefaultFilter_0']")
	public static WebElement custDuDateAnalysisCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_563_0_DefaultFilter_0']")
	public static WebElement custAgeingDetailsAnalysisCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_564_0_DefaultFilter_0']")
	public static WebElement custAgeingDetailsDueDateCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_566_0_DefaultFilter_0']")
	public static WebElement custOverDueDateCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_541_0_DefaultFilter_0']")
	public static WebElement venListingOutstandBillCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_542_0_DefaultFilter_0']")
	public static WebElement vendorStatementCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_543_0_DefaultFilter_0']")
	public static WebElement vendorDueDateCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_545_0_DefaultFilter_0']")
	public static WebElement vendorAgeingDetailAnalysisCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_546_0_DefaultFilter_0']")
	public static WebElement vendorAgeingDetailByDueDateCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_548_0_DefaultFilter_0']")
	public static WebElement vendorOverDueCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_544_0_DefaultFilter_0']")
	public static WebElement vendorAheingSummaryAnalysisCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_547_0_DefaultFilter_0']")
	public static WebElement vendorAgeingSummaryByDueDateCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_549_0_DefaultFilter_0']")
	public static WebElement vendorSummeryOverDueCusAccTxt;

	@FindBy(xpath = "//input[@id='FOption_634_0_DefaultFilter_0']")
	public static WebElement vendorSummeryBillwiseCusAccTxt;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[6]/label/span")
	public static WebElement report7chkbox;


	

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[15]/label/span")
	public static WebElement report16chkbox;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[20]/label/span")
	public static WebElement report21chkbox;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[16]/label/span")
	public static WebElement report17chkbox;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[17]/label/span")
	public static WebElement report18chkbox;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[1]/label/span")
	public static WebElement osr_includeBillsDropdown;


	@FindBy(xpath = "//input[@id='txtBackUpFile']")
	public static WebElement backUpEnterFileNameTxt;

	@FindBy(xpath = "//*[@id='BackUpControls']/ul/li/div[1]/div[2]")
	public static WebElement backUpScreenBackUpBtn;

	/*@FindBy(xpath = "//*[@id='26']/span")
	public static WebElement backup;*/
	
	//Manual Adjustmnet Screen starts from here 

	@FindBy(xpath="//select[@id='id_SelectARAP']")
	public static WebElement MA_ARAPDrpdwn;

	@FindBy(xpath="//input[@id='optTag']")
	public static WebElement MA_DepDrpDwn;

	@FindBy(xpath="//select[@id='id_SelectType']")
	public static WebElement MA_SelectTypeDrpdwn;

	@FindBy(xpath="//input[@id='optAccount']")
	public static WebElement MA_AccountDrpdwn;

	@FindBy(xpath="//input[@id='optCurrency']")
	public static WebElement MA_CurrencyDrpDwn;

	@FindBy(xpath="//span[@id='btnLoad']")
	public static WebElement MA_LoadBtn;

	@FindBy(xpath="//span[@id='btnOk']")
	public static WebElement MA_OkBtn;

	@FindBy(xpath="//span[@id='btnCancel']")
	public static WebElement MA_CanelBtn;

	@FindBy(xpath="//*[@id='id_DrAdj_Grid_body']/tr")
	public static WebElement MA_DebitRowList;

	@FindBy(xpath="//*[@id='id_CrAdj_Grid_body']/tr")
	public static WebElement MA_CreditRowList;

	@FindBy(xpath="//label[@id='id_lblLeftGridTotal']")
	public static WebElement MA_DebitTotalSum;

	@FindBy(xpath="//label[@id='id_lblRightGridTotal']")
	public static WebElement MA_CreditTotalSum;


	@FindBy(xpath="//*[@id='20']/span")
	public static WebElement utilitesMenu;

	@FindBy(xpath="//*[@id='3316']/span")
	public static WebElement ManualAdjustemntMenu;

	@FindBy(xpath="//*[@id='id_CrAdj_Grid_col_1-1']/input")
	public static WebElement MA_CreditSide1stRowChkBox;

	@FindBy(xpath="//*[@id='id_DrAdj_Grid_col_1-1']/input")
	public static WebElement MA_DebitSide1stRowChkBox;

	@FindBy(xpath="//*[@id='id_DrAdj_Grid_body']/tr[1]/td")
	public static List<WebElement> MA_DebitSideRow1list;

	@FindBy(xpath="//*[@id='id_CrAdj_Grid_body']/tr[1]/td")
	public static List<WebElement>MA_CreditSideRow1list;

	
	@FindBy(xpath="//*[@id='id_transactionentry_infopanel_container']/div[1]/div[3]/span")
	public static WebElement infosideMinimiseBtn;

	// DataManagement menu
	/*@FindBy(xpath = "//*[contains(text(),' Data Management ')]")
	public static WebElement dataMangementMenu;
*/
	@FindBy(xpath = "//*[@id='115']")
	public static WebElement eraseAll;

	@FindBy(xpath = "//*[@id='navbarSupportedContent2']/ul/li[2]/a/i")
	public static WebElement eraseAllCancelBtn;

	
	
	@FindBy(xpath = "//input[@id='idEraseTransaction']")
	public static WebElement eraseTranscationsRadio;

	@FindBy(xpath = "//*[@id='navbarSupportedContent2']/ul/li[1]/a/i")
	public static WebElement eraseAllOkBtn;
	
	@FindBy(xpath = "//button[contains(text(),'Ok')]")
	public static WebElement loginRefreshOkBtn;

	/*@FindBy(xpath = "//ul[@id='navigation_menu']/li/a/div/div")
	public static List<WebElement> menusList;
*/
	@FindBy(xpath = "//ul[@id='navigation_menu']/li[3]/ul/li/a/span")
	public static List<WebElement> inventoryMenusList;

	@FindBy(xpath = "//ul[@id='navigation_menu']/li[3]/ul/li[1]/ul/li/a/span")
	public static List<WebElement> inventoryTransactionsMenusList;

	@FindBy(xpath = "//ul[@id='navigation_menu']/li[3]/ul/li[1]/ul/li[3]/ul/li/a/span")
	public static List<WebElement> inventoryTransactionsStocksMenusList;

	@FindBy(xpath = "//*[@id='dv_TranHomeHeading']/nav/div[2]/div/div/div[2]")
	public static List<WebElement> openingStocksNewToolBarButtons;

	@FindBy(xpath = "//ul[@id='navigation_menu']/li[2]/ul/li/a/span")
	public static List<WebElement> financialsMenusList;

	@FindBy(xpath = "//*[@id='62']/span")
	public static WebElement financialsCreditManagementMenu;

	/*@FindBy(xpath = "//*[@id='81']")
	public static WebElement financialsReportsMenu;

	@FindBy(xpath = "//label[contains(text(),'Financials')]")
	public static WebElement financialsMenu;*/

	@FindBy(xpath = "//*[@id='530']")
	public static WebElement financialsFinalAccountsMenu;

	@FindBy(xpath = "//*[@id='558']")
	public static WebElement financialsReceivableAndPayableAnalysisMenu;

	@FindBy(xpath = "//*[@id='3301']")
	public static WebElement financialsBudgetMenu;

	@FindBy(xpath = "//*[@id='3311']")
	public static WebElement financialsOnlinePaymentsMenu;

	public static int cSize;

	@FindBy(xpath = "//table[@id='id_rc_columnheadertable']/thead/tr/th/div/p")
	public static List<WebElement> custBodyHeaderList;

	@FindBy(xpath = "//*[@id='plnCustomizecolumn']/div/span[4]/a")
	public static WebElement custRemoveBtn;

	@FindBy(xpath = "//*[@id='Save']/i")
	public static WebElement custSaveBtn;

	@FindBy(xpath = "//*[@id='btnCustomizeClose']/div[1]/span")
	public static WebElement custCancelBtn;

	@FindBy(xpath = "//*[@id='id_header_4_table_body']/tr")
	public static List<WebElement> cashAndBAnkAccountList;

	@FindBy(xpath = "//input[@id='id_header_4']")
	public static WebElement newCashBankAccountTxt;

	/*@FindBy(xpath = "//*[@id='id_header_6']")
	public static WebElement voucherHeaderDueDate;
*/
	@FindBy(xpath = "//*[@id='id_header_10']")
	public static WebElement voucherHeaderCurrency;

	/*@FindBy(xpath = "//*[@id='id_header_145']")
	public static WebElement voucherHeaderLocalExchangeRate;
*/
	@FindBy(xpath = "//table[@id='id_TranTable']/tbody") 
	public static WebElement voucherBodyGridRowCount;

	@FindBy(xpath = "//*[@id='id_body_12_table_body']/tr")
	public static List<WebElement> bodyAccountListInGrid;

	@FindBy(xpath = "//*[@id='id_body_39_table_body']/tr")
	public static List<WebElement> bodyCreditAccountListInGrid;

	// Bill Reference Screen

	@FindBy(xpath = "//input[@id='txtNewReference']")
	public static WebElement billRefNewReferenceTxt;

	@FindBy(xpath = "//input[@id='txtOnAccount']")
	public static WebElement billRefTxtOnAccount;

	@FindBy(xpath = "//input[@id='Searchtxt']")
	public static WebElement billRefSearchTxt;

	@FindBy(xpath = "//select[@id='cbmShowBillForAdjustment']")
	public static WebElement billRefNewReferenceDropdown;

	@FindBy(xpath = "//select[@id='cbmShowBills']")
	public static WebElement billRefShowBillsDropdown;

	@FindBy(xpath = "//th[@id='id_Adjustment_Grid_control_heading_11']//div[contains(text(),'Due Date')]")
	public static WebElement billRefGridHeaderDueDate;

	@FindBy(xpath = "//th[@id='id_Adjustment_Grid_control_heading_16']//div[contains(text(),'Currency')]")
	public static WebElement billRefGridHeaderCurrency;

	@FindBy(xpath = "//th[@id='id_Adjustment_Grid_control_heading_21']//div[contains(text(),'Original Amt')]")
	public static WebElement billRefGridHeaderOrginalAmt;

	@FindBy(xpath = "//th[@id='id_Adjustment_Grid_control_heading_23']//div[contains(text(),'Balance Amount')]")
	public static WebElement billRefGridHeaderBalanceAmt;

	@FindBy(xpath = "//th[@id='id_Adjustment_Grid_control_heading_33']//div[contains(text(),'Adjustment Amount')]")
	public static WebElement billRefGridHeaderAdjustmentAmt;

	@FindBy(xpath = "//div[contains(text(),'Native Currency')]")
	public static WebElement billRefGridHeaderNativeCurrency;

	@FindBy(xpath = "//th[@id='id_Adjustment_Grid_control_heading_41']//div[contains(text(),'Narration')]")
	public static WebElement billRefGridHeaderNarration;

	@FindBy(xpath = "//th[@id='id_Adjustment_Grid_control_heading_50']//div[contains(text(),'Previous Adjustment Amount')]")
	public static WebElement billRefGridHeaderPreviosAdjustmentAmt;

	@FindBy(xpath = "//td[@id='id_Adjustment_Grid_col_1-0']")
	public static WebElement billRefGridFirstRow;

	@FindBy(xpath = "//td[@id='id_Adjustment_Grid_col_1-1']//input")
	public static WebElement billRefGridFirstRowChkbox;

	@FindBy(xpath = "//td[@id='id_Adjustment_Grid_col_2-1']//input")
	public static WebElement billRefGridSecondRowChkbox;

	@FindBy(xpath = "//td[@id='id_Adjustment_Grid_col_3-1']//input")
	public static WebElement billRefGridThirdRowChkbox;

	/*
	 * @FindBy(xpath="//tbody[@id='id_Adjustment_Grid_body']//td[13]") private
	 * static WebElement billRefGridFirstRowAdjustmentAmtTxt;
	 */
	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr/td[13]")
	public static WebElement billRefGridFirstRowAdjustmentAmtTxt;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[2]/td[13]")
	public static WebElement billRefGridSecondRowAdjustmentAmtTxt;

	@FindBy(xpath = "//td[@id='id_Adjustment_Grid_col_3-33']")
	public static WebElement billRefGridThirdRowAdjustmentAmtTxt;

	@FindBy(xpath = "//td[@id='id_Adjustment_Grid_col_1-41']")
	public static WebElement billRefGridFirstNarrationTxt;

	@FindBy(xpath = "//td[@id='id_Adjustment_Grid_col_2-41']")
	public static WebElement billRefGridSecondNarrationTxt;

	@FindBy(xpath = "//td[@id='id_Adjustment_Grid_col_3-41']")
	public static WebElement billRefGridThirdNarrationTxt;

	@FindBy(xpath = "//input[@id='id_Narration']")
	public static WebElement narrationToWrite;

	@FindBy(xpath = "//input[@id='id_limit']")
	public static WebElement adjustAmtToWrite;

	@FindBy(xpath = "//div[@class='col-xs-6 navbar-header']//div[@class='navText']")
	public static WebElement billRefPaymentsBillReferenceLabel;

	@FindBy(xpath = "//div[@class='col-sm-12 col-md-3']//input[@class='Fbutton']")
	public static WebElement billRefAdjustOnFIFOBtn;

/*	@FindBy(xpath = "//*[@id='id_Pick']/a/span")
	public static WebElement billRefPickIcon;

	@FindBy(xpath = "//*[@id='id_Ok']/a/span")
	public static WebElement billRefOkBtn;
*/
	@FindBy(xpath = "//*[@id='id_Cancel']/a/span")
	public static WebElement billRefcancel;

	@FindBy(xpath = "//label[@id='id_transactionentry_footer_panel_summary_value_net']")
	public static WebElement billRefSummatyValueNet;

	@FindBy(xpath = "//td[@id='AccountAmount_col_1-1']")
	public static WebElement billRefGridVendorRow1;

	@FindBy(xpath = "//td[@id='AccountAmount_col_1-2']")
	public static WebElement billRefGridVendorAmtRow1;

	@FindBy(xpath = "//td[@id='AccountAmount_col_2-1']")
	public static WebElement billRefGridVendorRow2;

	@FindBy(xpath = "//td[@id='AccountAmount_col_2-2']")
	public static WebElement billRefGridVendorAmtRow2;

	@FindBy(xpath = "//td[@id='AccountAmount_col_3-1']")
	public static WebElement billRefGridVendorRow3;

	@FindBy(xpath = "//td[@id='AccountAmount_col_4-1']")
	public static WebElement billRefGridVendorRow4;

	@FindBy(xpath = "//label[@id='tbNewRefAmountAdjValue']")
	public static WebElement amtAdjustedAgainstNewREfinOtherVouchers;

	@FindBy(xpath = "//label[@id='tbAmountToAdjustInTransCurrencyValue']")
	public static WebElement transactionCurency;

	@FindBy(xpath = "//label[@id='tbAmountToAdjustInBaseCurrencyValue']")
	public static WebElement baseCurrency;

	@FindBy(xpath = "//label[@id='tbAmountToAdjustInLocalCurrencyValue']")
	public static WebElement localCurrencyDhs;

	@FindBy(xpath = "//label[@id='tbNewRefBalAmountValue']")
	public static WebElement balanceNewReferenceAmt;

	@FindBy(xpath = "//li[@id='id_li_Adjustment_Tab_BreakUpByTag']//span[@class='font-5'][contains(text(),'Break Up by Tag')]")
	public static WebElement breakUpByTagTab;

	@FindBy(xpath = "//span[@class='font-5'][contains(text(),'Adjustment')]")
	public static WebElement adjustmentTab;

	@FindBy(xpath = "//td[@id='id_BreakUpByTag_Grid_col_1-9']")
	public static WebElement breakUpByTagDueDate;

	@FindBy(xpath = "//td[@id='id_BreakUpByTag_Grid_col_1-10']")
	public static WebElement breakUpByTagDepartment;

	@FindBy(xpath = "//td[@id='id_BreakUpByTag_Grid_col_1-18']")
	public static WebElement breakUpByTagBaseCurrency;

	@FindBy(xpath = "//td[@id='id_BreakUpByTag_Grid_col_1-17']")
	public static WebElement breakUpByTagBillCurrency;

	@FindBy(xpath = "//td[@id='id_BreakUpByTag_Grid_col_1-20']")
	public static WebElement breakUpByTagBaseConversionRate;

	@FindBy(xpath = "//td[@id='id_BreakUpByTag_Grid_col_1-21']")
	public static WebElement breakUpByTagBaseLocalCurrency;

	@FindBy(xpath = "//label[@id='txtblkAmountadjusted']")
	public static WebElement billRefAdjustAmountInTransCurency;

	@FindBy(xpath = "//label[@id='txtblkAmounttobeadjust']")
	public static WebElement billRefBalanceAmountAdjustInTrnasCurrency;

	@FindBy(xpath = "//label[@id='id_BillWise_IP_LocalConversionRateValue']")
	public static WebElement billRefInfoBarLocalCurrencyRate;

	@FindBy(xpath = "//label[@id='tbAmountToAdjustInTransCurrencyValue']")
	public static WebElement billRefTransactionCurency;

	@FindBy(xpath = "//label[@id='tbAmountToAdjustInBaseCurrencyValue']")
	public static WebElement billRefBaseCurrency;

	@FindBy(xpath = "//label[@id='id_BillWise_IP_AccountName']")
	public static WebElement breakUpDetailsAccount;

	@FindBy(xpath = "//label[@id='id_BillWise_IP_ProductName']")
	public static WebElement breakUpDetailsItem;

	@FindBy(xpath = "//label[@id='id_BillWise_IP_TagName']")
	public static WebElement breakUpDetailsDepartment;

	@FindBy(xpath = "//*[@id='id_BillWise_IP_AccountName']")
	public static WebElement conversationRateBaseCurrencyRate;

	@FindBy(xpath = "//*[@id='id_BillWise_IP_LocalConversionRateValue']")
	public static WebElement conversationRateLocalCurrencyRate;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblAEDTranAmountValue']")
	public static WebElement asOnEntryDateTransAmt;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblAEDBaseConversionValue']")
	public static WebElement asOnEntryDateBaseConcersationRate;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblAEDBaseAmountValue']")
	public static WebElement asOnEntryDateBaseAmount;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblAEDLocalConversionValue']")
	public static WebElement asOnEntryDateLocConversationRate;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblAEDLocalAmountValue']")
	public static WebElement asOnEntryDateAmt;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblAADTranAmountValue']")
	public static WebElement balOnAdjstDateTransAmt;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblAADBaseConversionValue']")
	public static WebElement balOnAdjstDateBasrConversionRate;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblAADBaseAmountValue']")
	public static WebElement balOnAdjstDateBaseAmount;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblAADLocalConversionValue']")
	public static WebElement balOnAdjstDateLocalConversionRate;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblAADLocalAmountValue']")
	public static WebElement balOnAdjstDateAmt;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblAdjTranAmountValue']")
	public static WebElement adjustmentsAmount1;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblAdjBaseAmountValue']")
	public static WebElement adjustmentsAmount2;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblAdjLocalAmountValue']")
	public static WebElement adjustmentsAmount3;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblNativeCurrencyValue']")
	public static WebElement adjustmentsAmount4;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblExDiffTranAmountValue']")
	public static WebElement exchangeGainLossForBaseCurrency;

	@FindBy(xpath = "//label[@id='id_infoPanel_lblExDiffLocalAmountValue']")
	public static WebElement exchangeGainLossForLocalCurrency;

	/*@FindBy(xpath = "//input[@id='id_limit']")
	public static WebElement gridEnterAdjustAmtRow1;
*/
	// Billreference Body Grid Elements

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[1]/td[8]")
	public static WebElement gridOrginalAmtRow1;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[2]/td[8]")
	public static WebElement gridOrginalAmtRow2;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[3]/td[8]")
	public static WebElement gridOrginalAmtRow3;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[4]/td[8]")
	public static WebElement gridOrginalAmtRow4;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[5]/td[8]")
	public static WebElement gridOrginalAmtRow5;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[1]/td[10]")
	public static WebElement gridBalanceAmtRow1;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[2]/td[10]")
	public static WebElement gridBalanceAmtRow2;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[3]/td[10]")
	public static WebElement gridBalanceAmtRow3;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[4]/td[10]")
	public static WebElement gridBalanceAmtRow4;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[5]/td[10]")
	public static WebElement gridBalanceAmtRow5;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[1]/td[13]")
	public static WebElement gridAdjustmentAmtRow1;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[2]/td[13]")
	public static WebElement gridAdjustmentAmtRow2;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[3]/td[13]")
	public static WebElement gridAdjustmentAmtRow3;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[4]/td[13]")
	public static WebElement gridAdjustmentAmtRow4;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[5]/td[13]")
	public static WebElement gridAdjustmentAmtRow5;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']r")
	public static WebElement billRefAdjustBills;

	//@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr")
	@FindBy(xpath = "//table[@id='id_Adjustment_Grid']")
	public static WebElement billRefAdjustBillsGrid;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[1]/td[3]")
	public static WebElement billRefAdjustBillsRow1DocNo;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[2]/td[3]")
	public static WebElement billRefAdjustBillsRow2DocNo;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[3]/td[3]")
	public static WebElement billRefAdjustBillsRow3DocNo;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[4]/td[3]")
	public static WebElement billRefAdjustBillsRow4DocNo;

	@FindBy(xpath = "//*[@id='id_Adjustment_Grid_body']/tr[5]/td[3]")
	public static WebElement billRefAdjustBillsRow5DocNo;

	@FindBy(xpath = "//*[@id='InfPnlAdjGrd_body']/tr/td[2]")
	public static WebElement InfoSideBarDocRow1;

/*	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[5]")
	public static WebElement  select1stRow_4thColumn;*/

	
	@FindBy(xpath="//*[@class='icon-arrow hiconright']")
	public static WebElement userNameImageDisplay;
	
	@FindBy(xpath = "//*[@id='InfPnlAdjGrd_body']/tr/td[3]")
	public static WebElement InfoSideBarAmtRow1;

	@FindBy(xpath = "//tbody[@id='InfPnlAdjGrd_body']//td[4]")
	public static WebElement InfoSideBarBaseAmtRow1;

	// Row 2

	@FindBy(xpath = "//*[@id='InfPnlAdjGrd_body']/tr[2]/td[2]")
	public static WebElement InfoSideBarDocRow2;

	@FindBy(xpath = "//*[@id='InfPnlAdjGrd_body']/tr[2]/td[3]")
	public static WebElement InfoSideBarAmtRow2;

	@FindBy(xpath = "//tbody[@id='InfPnlAdjGrd_body']//tr[2]//td[4]")
	public static WebElement InfoSideBarBaseAmtRow2;

	// Row 3

	@FindBy(xpath = "//*[@id='InfPnlAdjGrd_body']/tr[3]/td[2]")
	public static WebElement InfoSideBarDocRow3;

	@FindBy(xpath = "//*[@id='InfPnlAdjGrd_body']/tr[3]/td[3]")
	public static WebElement InfoSideBarAmtRow3;

	@FindBy(xpath = "//tbody[@id='InfPnlAdjGrd_body']//tr[3]//td[4]")
	public static WebElement InfoSideBarBaseAmtRow3;

	// Row 4

	@FindBy(xpath = "//*[@id='InfPnlAdjGrd_body']/tr[4]/td[2]")
	public static WebElement InfoSideBarDocRow4;

	@FindBy(xpath = "//*[@id='InfPnlAdjGrd_body']/tr[4]/td[3]")
	public static WebElement InfoSideBarAmtRow4;

	@FindBy(xpath = "//tbody[@id='InfPnlAdjGrd_body']//tr[4]//td[4]")
	public static WebElement InfoSideBarBaseAmtRow4;

	// row 5

	@FindBy(xpath = "//*[@id='InfPnlAdjGrd_body']/tr[5]/td[2]")
	public static WebElement InfoSideBarDocRow5;

	@FindBy(xpath = "//*[@id='InfPnlAdjGrd_body']/tr[5]/td[3]")
	public static WebElement InfoSideBarAmtRow5;

	@FindBy(xpath = "//tbody[@id='InfPnlAdjGrd_body']//tr[5]//td[4]")
	public static WebElement InfoSideBarBaseAmtRow5;

	@FindBy(xpath = "//*[@id='id_body_12_table_body']/tr")
	public static List<WebElement> openingBalAccountListInGrid;

	@FindBy(xpath = "//input[@id='id_body_19']")
	public static WebElement enter_OBCreditACTxt;

	@FindBy(xpath = "//*[@id='lblAccount']")
	public static WebElement billRefPartyName;

	// Purchase Voucher VAT
	@FindBy(xpath = "//*[@id='2007']")
	public static WebElement financialsTransactionsPurchaseMenu;

	//@FindBy(xpath = "//*[@id='2057']")
	@FindBy(xpath = "//*[@id='2064']")
	public static WebElement purchaseVouchersVat;
	
	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr/td")
	public static List<WebElement> reportBodyList;

	@FindBy(xpath = "//input[@id='RITCheckbox__2']")
	public static WebElement movedAccOnlyChkBox;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[6]/label/span")
	public static WebElement showallConsildateAmtChkbox;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[5]/label/span")
	public static WebElement includePurchaseReturnChkbox;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[7]/label/span")
	public static WebElement printAccIndexChkbox;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[8]/label/span")
	public static WebElement freshPageAccChkbox;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[9]/label/span")
	public static WebElement printAsStatementAccChkbox;

	
	@FindBy(xpath = "//*[@id='dvReportInputs']/div[10]/label/span")
	public static WebElement ignoreTransBalProfitAndLossChkbox;

	
	@FindBy(xpath = "//*[@id='dvReportInputs']/div[11]/label/span")
	public static WebElement displayTotalAtEndOfGroupChkbox;
	
	@FindBy(xpath = "//*[@id='dvReportInputs']/div[12]/label/span")
	public static WebElement dontMoveAccBasedOnBalChkbox;

	
	@FindBy(xpath = "//input[@id='RITCheckbox__6']")
	public static WebElement displayUnReliziedchkbox;

	@FindBy(xpath = "//input[@id='RITCheckbox__4']")
	public static WebElement ignoreConvertedPDCChkbox;

	@FindBy(xpath = "//input[@id='RITCheckbox__7']")
	public static WebElement displayedMaturedPDCChkbox;
	
	
	@FindBy(xpath = "//a[@id='207']")
	public static WebElement vatReportMenu;

	@FindBy(xpath = "//*[@id='navigation_menu']/li[2]/ul/li[3]/ul/li[7]/ul/li/a/span")
	public static List<WebElement> vatReportList;

	@FindBy(xpath = "//span[contains(text(),'Purchase Account Report')]")
	public static WebElement purchaseAccountReport;

	@FindBy(xpath = "//span[contains(text(),'VAT detailed report')]")
	public static WebElement vatDetailedReport;

	@FindBy(xpath = "//span[contains(text(),'Sales Account Report')]")
	public static WebElement salesAccountReport;

	@FindBy(xpath = "//span[contains(text(),'Sales by Customer Report')]")
	public static WebElement salesByCustomerReport;

	@FindBy(xpath = "//span[contains(text(),'VAT summary report')]")
	public static WebElement vatSummaryReport;

	@FindBy(xpath = "//span[contains(text(),'VAT audit file')]")
	public static WebElement vatAuditFileReport;

	@FindBy(xpath = "//span[contains(text(),'VAT return report')]")
	public static WebElement vatReturnReport;

	@FindBy(xpath = "//span[contains(text(),'Sales advance VAT report')]")
	public static WebElement salesAdvanceVatReport;

	@FindBy(xpath = "//select[@id='RITCombobox__1']")
	public static WebElement transactionTypeDropDown;
	
	@FindBy(xpath = "//select[@id='RITCombobox__2']")
	public static WebElement voucherTypeDropDown;

	@FindBy(xpath = "//select[@id='RITCombobox__3']")
	public static WebElement selectDealersDropDown;
	
	// Bank Reconciliation

	@FindBy(xpath = "//li[@id='btnSaveDisableImmediate']//div[@class='toolbar_button_image']")
	public static WebElement BRsaveBtn;

	@FindBy(xpath = "//div[contains(text(),'Backtrack')]")
	public static WebElement backTrackBtn;

	@FindBy(xpath = "//div[contains(text(),'Customize')]")
	public static WebElement BRcustomizeBtn;

	@FindBy(xpath = "//div[contains(text(),'Cancel')]")
	public static WebElement BRcancelBtn;

	@FindBy(xpath = "//input[@id='OptCtrlBank']")
	public static WebElement reportbankTxt;

	@FindBy(xpath = "//*[@id='OptCtrlBank_table_body']/tr/td[2]")
	public static List<WebElement> reportbankList;

	@FindBy(xpath = "//select[@id='sortOrder']")
	public static WebElement sortOrderDropDown;

	@FindBy(xpath = "//select[@id='selectStatus']")
	public static WebElement selectStatusDropDown;

	@FindBy(xpath = "//select[@id='DatePeriod']")
	public static WebElement datePeriodDropDown;

	@FindBy(xpath = "//input[@id='chkShow']")
	public static WebElement showConsolidatedAmountsForBankDeposits;

	@FindBy(xpath = "//select[@id='SelectDebit']")
	public static WebElement selectDrCrDropDown;

	@FindBy(xpath = "//div[@id='btnAdvFilterText']")
	public static WebElement advanceFilterBtn;

	/*@FindBy(xpath = "//span[@id='btnClear']")
	public static WebElement clearBtn;*/

	@FindBy(xpath = "//*[@id='selectContainer']/div[5]/div[2]/div/input[3]")
	public static WebElement LoadBtn;

	@FindBy(xpath = "//button[@id='btnPendingBills']")
	public static WebElement BRpendingBillsBtn;

	@FindBy(xpath = "//button[@id='btnRaiseReceipt']")
	public static WebElement raiseReceiptsBtn;

	@FindBy(xpath = "//button[@id='btnRaisePayment']")
	public static WebElement raisePaymentsBtn;

	@FindBy(xpath = "//thead[@id='BRTable_head']/tr/th/div")
	public static List<WebElement> reporttableHeadeList;

	@FindBy(xpath = "//*[@id='BRTable_body']/tr[1]/td")
	public static List<WebElement> bankRecRow1List;

	@FindBy(xpath = "//*[@id='BRTable_body']/tr[2]/td")
	public static List<WebElement> bankRecRow2List;

	@FindBy(xpath = "//*[@id='BRTable_body']/tr[3]/td")
	public static List<WebElement> bankRecRow3List;

	@FindBy(xpath = "//*[@id='BRTable_body']/tr[4]/td")
	public static List<WebElement> bankRecRow4List;

	@FindBy(xpath = "//*[@id='BRTable_body']/tr[5]/td")
	public static List<WebElement> bankRecRow5List;

	@FindBy(xpath = "//*[@id='BRTable_body']/tr/td[7]")
	public static List<WebElement> bankRecRow6List;

	@FindBy(xpath = "//*[@id='BRTable_body']/tr/td[8]")
	public static List<WebElement> bankRecRow7List;

	@FindBy(xpath = "//*[@id='BRTable_body']/tr/td[9]")
	public static List<WebElement> bankRecRow8List;

	@FindBy(xpath = "//label[@id='bookBal']")
	public static WebElement bankRecBookBal;

	@FindBy(xpath = "//*[@id='CancelCreateUser']/i")
	public static WebElement bankRecCloseBn;
	
	
	@FindBy(xpath = "//label[@id='outDebits']")
	public static WebElement bankRecOutDebits;

	@FindBy(xpath = "//label[@id='outCredits']")
	public static WebElement bankRecOutCredits;

	@FindBy(xpath = "//label[@id='clearedBal']")
	public static WebElement bankRecClearedBal;

	@FindBy(xpath = "//label[@id='OpeningBalance']")
	public static WebElement bankRecOpenBal;

	@FindBy(xpath = "//label[@id='debitCounts']")
	public static WebElement bankRecDebitCounts;

	@FindBy(xpath = "//label[@id='creditCounts']")
	public static WebElement bankRecCreditCounts;

	@FindBy(xpath = "//input[@id='bankBal']")
	public static WebElement bankRecBankBal;

	@FindBy(xpath = "//label[@id='diff']")
	public static WebElement bankRecDif;

	@FindBy(xpath = "//*[@id='OptCtrlBank_table_body']/tr/td[2]")
	public static List<WebElement> bankList;


	// Bank reconciliation statement

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[6]/label/span")
	public static WebElement brsshowConsolidatedAmountsChkBox;

	@FindBy(xpath = "//*[@id='dvReportInputs']/div[5]/label/span")
	public static WebElement brsIncludePdcChkBox;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr/td[2]")
	public static List<WebElement> report1stColList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr/td[3]")
	public static List<WebElement> report2ndColList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr/td[4]")
	public static List<WebElement> report3rdColList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr/td[5]")
	public static List<WebElement> report4thColList;

	@FindBy(xpath = "//div[@id='dvReportDetails']/div/table/tbody/tr/td[6]")
	public static List<WebElement> report5thColList;
	
	
	@FindBy(xpath = "//tbody[@id='id_header_4_table_body']/tr/td[2]")
	public static List<WebElement> vendorAccountListCount;

	@FindBy(xpath = "//input[@id='id_body_536870916']")
	public static WebElement enter_WarehouseTxt;

	@FindBy(xpath = "//*[@id='id_body_536870916_table_body']/tr")
	public static List<WebElement> pvvGridWarehouseList;

	@FindBy(xpath = "//*[@id='id_body_23_table_body']/tr")
	public static List<WebElement> pvvGridItemList;

	@FindBy(xpath = "//tbody[@id='id_body_536870916_table_body']/tr/td[2]")
	public static List<WebElement> warehouseBodyComboList;

	@FindBy(xpath = "//tbody[@id='id_header_86_table_body']/tr/td[2]")
	public static List<WebElement> st_warehouse1HeaderComboList;

	@FindBy(xpath = "//tbody[@id='id_body_87_table_body']/tr/td[2]")
	public static List<WebElement> st_warehouse2BodyComboList;

	@FindBy(xpath = "//tbody[@id='id_header_3_table_body']/tr/td[2]")
	public static List<WebElement> purchaseAccountListCount;
	
	@FindBy(xpath = "//*[@id='dvReportInputs']/div[4]/label/span")
	public static WebElement clubCashSalesForTheDayChkBox;

	@FindBy(xpath = "//*[@id='722']/span")
	public static WebElement virtualBankReport;

	
	


@FindBy(xpath="//select[@id='misc_arapARAPOptions']")
public static WebElement arapDrpDwn;

@FindBy(xpath="//select[@id='misc_arapOptions']")
public static WebElement settingMisOptionsDrpDwn;

@FindBy(xpath="//input[@id='misc_arapchkSelectbillsbeforeamt']")
public static WebElement settingMisSelectBillBeforeAmtChkbox;


@FindBy(xpath="//input[@id='misc_arapchkSelectbillsbeforeamt']/following-sibling::span")
public static WebElement settingMisSelectBillBeforeAmtChkboxISSelected;



//input[@id='misc_arapchkSelectbillsbeforeamt']/following-sibling::span

@FindBy(xpath="//input[@id='id_misc_adjustbillinline']")
public static WebElement settingMisAdjustBillInLineChkbox;

@FindBy(xpath="//input[@id='misc_arapchkPostdiscountbasedoncrdays']")
public static WebElement settingMisPostDiscountBasedOncreditDaysChkbox;


@FindBy(xpath="//select[@id='misc_arapNarration']")
public static WebElement settingMisPostNarrationdrpdwn;



@FindBy(xpath="//span[contains(text(),'Payments FIFO')]")
public static WebElement  paymentsFIFOVoucher;



/////////////////////////////////


@FindBy(xpath="//*[@id='2083']/span")
public static WebElement salesNewRefMenu;


/*@FindBy(xpath="//input[@id='id_header_4']")
public static WebElement  customerAccountTxt;*/





@FindBy(xpath="//tbody[@id='id_header_268435459_table_body']/tr/td[2]")
public static List<WebElement> departmentListCount;




@FindBy(xpath="//tbody[@id='id_header_10_table_body']/tr/td[2]")
public static List<WebElement> currencyListCount;	 




/*@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[2]")
public static WebElement  select1stRow_1stColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[3]")
public static WebElement  select1stRow_2ndColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[4]")
public static WebElement  select1stRow_3rdColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[5]")
public static WebElement  select1stRow_4thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[6]")
public static WebElement  select1stRow_5thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[7]")
public static WebElement  select1stRow_6thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[8]")
public static WebElement  select1stRow_7thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[9]")
public static WebElement  select1stRow_8thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[10]")
public static WebElement  select1stRow_9thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[11]")
public static WebElement  select1stRow_10thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[12]")
public static WebElement  select1stRow_11thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[13]")
public static WebElement  select1stRow_12thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[14]")
public static WebElement  select1stRow_13thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[15]")
public static WebElement  select1stRow_14thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[16]")
public static WebElement  select1stRow_15thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[17]")
public static WebElement  select1stRow_16thColumn;





@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[2]")
public static WebElement  select2ndRow_1stColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[3]")
public static WebElement  select2ndRow_2ndColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[4]")
public static WebElement  select2ndRow_3rdColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[5]")
public static WebElement  select2ndRow_4thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[6]")
public static WebElement  select2ndRow_5thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[7]")
public static WebElement  select2ndRow_6thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[8]")
public static WebElement  select2ndRow_7thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[9]")
public static WebElement  select2ndRow_8thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[10]")
public static WebElement  select2ndRow_9thColumn;
*/
/*@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[12]")
public static WebElement  select2ndRow_11thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[13]")
public static WebElement  select2ndRow_12thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[11]")
public static WebElement  select2ndRow_10thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[15]")
public static WebElement  select2ndRow_14thColumn;*/

/*@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[16]")
public static WebElement  select2ndRow_15thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[17]")
public static WebElement  select2ndRow_16thColumn;




@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[2]")
public static WebElement  select3rdRow_1stColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[3]")
public static WebElement  select3rdRow_2ndColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[4]")
public static WebElement  select3rdRow_3rdColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[5]")
public static WebElement  select3rdRow_4thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[6]")
public static WebElement  select3rdRow_5thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[7]")
public static WebElement  select3rdRow_6thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[8]")
public static WebElement  select3rdRow_7thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[9]")
public static WebElement  select3rdRow_8thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[10]")
public static WebElement  select3rdRow_9thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[11]")
public static WebElement  select3rdRow_10thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[12]")
public static WebElement  select3rdRow_11thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[13]")
public static WebElement  select3rdRow_12thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[14]")
public static WebElement  select3rdRow_13thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[15]")
public static WebElement  select3rdRow_14thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[16]")
public static WebElement  select3rdRow_15thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[17]")
public static WebElement  select3rdRow_16thColumn;*/




@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[2]")
public static WebElement  select4thRow_1stColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[3]")
public static WebElement  select4thRow_2ndColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[4]")
public static WebElement  select4thRow_3rdColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[5]")
public static WebElement  select4thRow_4thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[6]")
public static WebElement  select4thRow_5thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[7]")
public static WebElement  select4thRow_6thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[8]")
public static WebElement  select4thRow_7thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[9]")
public static WebElement  select4thRow_8thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[10]")
public static WebElement  select4thRow_9thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[11]")
public static WebElement  select4thRow_10thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[12]")
public static WebElement  select4thRow_11thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[13]")
public static WebElement  select4thRow_12thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[14]")
public static WebElement  select4thRow_13thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[15]")
public static WebElement  select4thRow_14thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[16]")
public static WebElement  select4thRow_15thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[17]")
public static WebElement  select4thRow_16thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[19]")
public static WebElement  select4thRow_18thColumn;



@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[2]")
public static WebElement  select5thRow_1stColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[3]")
public static WebElement  select5thRow_2ndColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[4]")
public static WebElement  select5thRow_3rdColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[5]")
public static WebElement  select5thRow_4thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[6]")
public static WebElement  select5thRow_5thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[7]")
public static WebElement  select5thRow_6thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[8]")
public static WebElement  select5thRow_7thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[9]")
public static WebElement  select5thRow_8thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[10]")
public static WebElement  select5thRow_9thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[11]")
public static WebElement  select5thRow_10thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[12]")
public static WebElement  select5thRow_11thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[13]")
public static WebElement  select5thRow_12thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[14]")
public static WebElement  select5thRow_13thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[15]")
public static WebElement  select5thRow_14thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[16]")
public static WebElement  select5thRow_15thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[17]")
public static WebElement  select5thRow_16thColumn;




@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[2]")
public static WebElement  select6thRow_1stColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[3]")
public static WebElement  select6thRow_2ndColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[4]")
public static WebElement  select6thRow_3rdColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[5]")
public static WebElement  select6thRow_4thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[6]")
public static WebElement  select6thRow_5thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[7]")
public static WebElement  select6thRow_6thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[8]")
public static WebElement  select6thRow_7thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[9]")
public static WebElement  select6thRow_8thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[10]")
public static WebElement  select6thRow_9thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[11]")
public static WebElement  select6thRow_10thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[12]")
public static WebElement  select6thRow_11thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[13]")
public static WebElement  select6thRow_12thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[14]")
public static WebElement  select6thRow_13thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[15]")
public static WebElement  select6thRow_14thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[16]")
public static WebElement  select6thRow_15thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[17]")
public static WebElement  select6thRow_16thColumn;



@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[2]")
public static WebElement  select7thRow_1stColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[3]")
public static WebElement  select7thRow_2ndColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[4]")
public static WebElement  select7thRow_3rdColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[5]")
public static WebElement  select7thRow_4thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[6]")
public static WebElement  select7thRow_5thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[7]")
public static WebElement  select7thRow_6thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[8]")
public static WebElement  select7thRow_7thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[9]")
public static WebElement  select7thRow_8thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[10]")
public static WebElement  select7thRow_9thColumn;




@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[2]")
public static WebElement  select8thRow_1stColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[3]")
public static WebElement  select8thRow_2ndColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[4]")
public static WebElement  select8thRow_3rdColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[5]")
public static WebElement  select8thRow_4thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[6]")
public static WebElement  select8thRow_5thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[7]")
public static WebElement  select8thRow_6thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[8]")
public static WebElement  select8thRow_7thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[9]")
public static WebElement  select8thRow_8thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[10]")
public static WebElement  select8thRow_9thColumn;



@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[2]")
public static WebElement  select9thRow_1stColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[3]")
public static WebElement  select9thRow_2ndColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[4]")
public static WebElement  select9thRow_3rdColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[5]")
public static WebElement  select9thRow_4thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[6]")
public static WebElement  select9thRow_5thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[7]")
public static WebElement  select9thRow_6thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[8]")
public static WebElement  select9thRow_7thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[9]")
public static WebElement  select9thRow_8thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[10]")
public static WebElement  select9thRow_9thColumn;




@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[2]")
public static WebElement  select10thRow_1stColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[3]")
public static WebElement  select10thRow_2ndColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[4]")
public static WebElement  select10thRow_3rdColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[5]")
public static WebElement  select10thRow_4thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[6]")
public static WebElement  select10thRow_5thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[7]")
public static WebElement  select10thRow_6thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[8]")
public static WebElement  select10thRow_7thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[9]")
public static WebElement  select10thRow_8thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[10]")
public static WebElement  select10thRow_9thColumn;



@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[2]")
public static WebElement  select11thRow_1stColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[3]")
public static WebElement  select11thRow_2ndColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[4]")
public static WebElement  select11thRow_3rdColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[5]")
public static WebElement  select11thRow_4thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[6]")
public static WebElement  select11thRow_5thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[7]")
public static WebElement  select11thRow_6thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[8]")
public static WebElement  select11thRow_7thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[9]")
public static WebElement  select11thRow_8thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[10]")
public static WebElement  select11thRow_9thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[2]")
public static WebElement  select12thRow_1stColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[3]")
public static WebElement  select12thRow_2ndColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[4]")
public static WebElement  select12thRow_3rdColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[5]")
public static WebElement  select12thRow_4thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[6]")
public static WebElement  select12thRow_5thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[7]")
public static WebElement  select12thRow_6thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[8]")
public static WebElement  select12thRow_7thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[9]")
public static WebElement  select12thRow_8thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[10]")
public static WebElement  select12thRow_9thColumn;

@FindBy(xpath="//input[@id='id_body_12']")
public static WebElement  enter_AccountTxt;

//@FindBy(xpath="//input[@id='id_body_16777332']")
@FindBy(xpath="//input[@id='id_body_16777307']")
public static WebElement  enterpayVATTaxCode;

/*@FindBy(xpath="//input[@id='id_body_16777306']")
public static WebElement  enterpayVATTaxCode;*/

/*@FindBy(xpath="//input[@id='id_body_16777330']")
public static WebElement  enterReceiptsVATTaxCode;*/




@FindBy(xpath="//a[@id='2075']")
public static WebElement salesReturnFIFO;



@FindBy(xpath="//a[@id='2055']")
public static WebElement  recepitsFIFOmenu;


  
@FindBy(xpath="//*[@id='tblBodyTransRender']/tr")
public static List<WebElement>  voucherBodyGridRowCountList;

	
@FindBy(xpath="//a[@id='2076']")
public static WebElement openingBalFIFOMenu;

	


@FindBy(xpath="//a[@id='2077']")
public static WebElement  JVFIFOMenu;

@FindBy(xpath="//a[@id='2032']")
public static WebElement  journalEntriesMenu;



/*@FindBy(xpath="//*[@id='id_transactionentry_postingdetail_popup']/div/div/div[2]/div/div[2]/div[1]/div/div")
public static List<WebElement>  postingDetailsDebitList; 

@FindBy(xpath="//*[@id='id_transactionentry_postingdetail_popup']/div/div/div[2]/div/div[2]/div[1]/div/div[1]")
public static List<WebElement>  postingDetailsDebitSecList_AccList;

@FindBy(xpath="//*[@id='id_transactionentry_postingdetail_popup']/div/div/div[2]/div/div[2]/div[2]/div/div")
public static List<WebElement>  postingDetailsCreditList; 

@FindBy(xpath="//*[@id='id_transactionentry_postingdetail_popup']/div/div/div[2]/div/div[2]/div[2]/div/div[1]")
public static List<WebElement>  postingDetailsCreditSec_AccList; 

@FindBy(xpath="//*[@id='id_transactionentry_postingdetail_popup']/div/div/div[2]/div/div[3]/div[1]/div/div[2]")
public static WebElement  postingDetailsDebitSum; 
	  	
@FindBy(xpath="//*[@id='id_transactionentry_postingdetail_popup']/div/div/div[2]/div/div[3]/div[2]/div/div[2]")
public static WebElement  postingDetailsCreditSum; */

/*@FindBy(xpath="//*[@id='id_transactionentry_postingdetail_popup']/div[2]/div/div[1]/span/i")
public static WebElement  postingDetailsCloseBtn; */
	
@FindBy(xpath="//a[@id='2079']")
public static WebElement PDPFIFOMenu;


@FindBy(xpath="//*[@id='id_transactionentry_save']/a/span")
public static WebElement  MRsaveBtn;

@FindBy(xpath="//*[@id='InfPnlAdjGrd_body']/tr/td")
public static List<WebElement>  baseAmtList;
	
@FindBy(xpath="//span[contains(text(),'PDR FIFO')]")
public static WebElement PDRFIFO;
	
@FindBy(xpath="//a[@id='2004']")
public static WebElement pettyCash;

@FindBy(xpath="//*[@id='2026']/span")
public static WebElement debitnoteMenu;




@FindBy(xpath="//*[@id='2080']/span")
public static WebElement pettyCashFIFO;


@FindBy(xpath="//span[contains(text(),'Inter Dep FIFO')]")
public static WebElement interDepFifo;



@FindBy(xpath="//a[@id='2082']")
public static WebElement NonJVFIFO;

@FindBy(xpath="//a[@id='2015']")
public static WebElement salesInvoiceVoucher;

;@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr/td[3]")
public static List<WebElement> entryPageAdjBillsDocList;


@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr/td[2]")
public static List<WebElement> entryPageAdjBillsChkboxList;

@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr/td[13]")
public static List<WebElement> entryPageAdjBillsAdjustAmtList;

@FindBy(xpath="//*[@id='id_transactionentry_infopanel_container']/div[1]/div[3]/span")
public static WebElement infosideBarBtn;

@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr/td[2]/input")
public static WebElement entrypageAdjsutmentBillsrow1Chkbox;


@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr/td")
public static List<WebElement> entrypageAdjsutmentBillsList;


@FindBy(xpath="(//*[@id='id_transaction_entry_workflow_togglebutton'])[2]")
public static WebElement entryPageBillwiseExpandBtn;


@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr/td[13]")
public static WebElement entryPageBillwiseAdjustMentTxt;

@FindBy(xpath="//input[@id='id_limit']")
public static WebElement entryPageBillwise_EnterAdjustMentTxt;


@FindBy(xpath="//*[@id='id_Adjustment_Grid_head']/tr[1]/th")
public static List<WebElement> adjustmentbillsHeaderList;

@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr/td[4]")
public static WebElement adjustmentbillsBillNo;



@FindBy(xpath="//span[@id='editLayout_FieldsCustomization_Close']//i[@class='icon-cancel icon-font7']")
public static WebElement editLayoutcloseBtn;



/*@FindBy(xpath="//i[@class='icon-close icon-font6']")
public static WebElement  closeBtn;*/

@FindBy(xpath="//*[@id='id_transactionentry_header1_section']/div/div[9]/div[2]/input")
public static WebElement pvVATBillNoTxt;





@FindBy(xpath="//div[@id='dvReportDetails']/div/table/tbody/tr[16]/td")
public static List<WebElement> report16thRowList;

@FindBy(xpath="//div[@id='dvReportDetails']/div/table/tbody/tr[17]/td")
public static List<WebElement> report17thRowList;

@FindBy(xpath="//div[@id='dvReportDetails']/div/table/tbody/tr[18]/td")
public static List<WebElement> report18thRowList;

@FindBy(xpath="//div[@id='dvReportDetails']/div/table/tbody/tr[19]/td")
public static List<WebElement> report19thRowList;

@FindBy(xpath="//*[@id='navigation_menu']/li[2]/ul/li[3]/ul/li[8]/ul/li/a/span")
public static List<WebElement> registerList;



@FindBy(xpath="//input[@id='id_body_33554464']")
public static WebElement enter_SALESINVOICEFQTxt;



@FindBy(xpath="//div[contains(text(),'Header/Footer')]")
public static WebElement sl_cusHeaderAndFooter;


@FindBy(xpath="//*[@id='forHf']/span[3]/span/i[1]")
public static WebElement sl_cusHeaderAndFooterSaveBtn;


@FindBy(xpath="//a[contains(text(),'Save')]")
public static WebElement HFsave_SaveBtn;

@FindBy(xpath="//input[@id='id_PopuplayoutName']")
public static WebElement HFsavePop_SaveTxt;

@FindBy(xpath="//button[@id='btn_PopUpOk']")
public static WebElement HFsavePop_OKBtn;


@FindBy(xpath="//*[@id='REPORTRENDERNEWControls']/ul/li/span[7]")
public static WebElement reportScreenCustomizeBtn;


@FindBy(xpath="//span[@class='icon-filter icon-font7']")
public static WebElement vat_filterBtn;

@FindBy(xpath="//span[@class='icon-ok icon-font7']")
public static WebElement vat_OkBtn;

@FindBy(xpath="//span[@class='icon-close icon-font7']")
public static WebElement vat_CloseBtn;

@FindBy(xpath="//select[@id='ddlVATHomeRepDateOpt']")
public static WebElement vat_DateOptionDropdown;

@FindBy(xpath="//input[@id='VATHomeReport_StartDate']")
public static WebElement vat_StartDateTxt;

@FindBy(xpath="//td[@id='VATHomeReport_StartDate_input_image']//span[@class='icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow']")
public static WebElement vat_StartDateCalanderBtn;

@FindBy(xpath="//input[@id='VATHomeReport_EndDate']")
public static WebElement vat_EndDateTxt;

@FindBy(xpath="//td[@id='VATHomeReport_EndDate_input_image']//span[@class='icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow']")
public static WebElement vat_EndDateCalanderBtn;

@FindBy(xpath="//select[@id='ddlVATHomeRepOutPut']")
public static WebElement vat_OutputDropdown;

@FindBy(xpath="//*[@id='2089']/span")
public static WebElement nonJVNewReferenceBtn;

@FindBy(xpath="//span[contains(text(),'Local amount based on filter')]")
public static WebElement localAmountBasedOnFilterChkBox;

@FindBy(xpath="//input[@id='FOption_569_0_DefaultFilter_0']")
public static WebElement  filterEnterPeakAndLowTxt;

@FindBy(xpath="//input[@id='FOption_538_0_DefaultFilter_0']")
public static WebElement  filterMasterInfoTxt;


@FindBy(xpath="//input[@id='FOption_677_0_DefaultFilter_0']")
public static WebElement  enterVATSalesDefaultItemTxt;

@FindBy(xpath="//input[@id='FOption_209_0_DefaultFilter_0']")
public static WebElement  enterVATDefaultItemTxt;

@FindBy(xpath="//input[@id='RITCheckbox__1']")
public static WebElement displayMonthlyTotalChkBox;

@FindBy(xpath="//input[@id='RITCheckbox__1']")
public static WebElement yearWiseComparsionChkbox;

@FindBy(xpath="//span[contains(text(),'Include Sales Return voucher')]")
public static WebElement includeSalesReturnsVoucherChkBox;





@FindBy(xpath="//*[@id='tblsummary']/tr/td")
public static List<WebElement> vatSummaryList;

@FindBy(xpath="//div[@id='myNavbarVATSummaryRep']//span[@class='icon-close icon-font7']")
public static WebElement vatSummaryExistBtn;

@FindBy(xpath="//table[@id='tblCompanyInfo']//td[9]")
public static WebElement  periodEndsTxt; 

@FindBy(xpath="//table[@id='tblCompanyInfo']//td[9]")
public static WebElement  FAFCrestionDateTxt; 

@FindBy(xpath="//*[@id='tblCompanyInfo']/tbody/tr/td[11]")
public static WebElement productVersionTxt ; 

@FindBy(xpath="//*[@id='tblPurListing']/tbody/tr/td[1]")
public static List<WebElement> purchaseSupplierName ; 

@FindBy(xpath="//*[@id='tblPurListing']/tbody/tr/td[4]")
public static List<WebElement> purchaseInvoiceNoList ; 

@FindBy(xpath="//*[@id='tblPurListing']/tbody/tr/td[7]")
public static List<WebElement> purchaseProductDescList ; 

@FindBy(xpath="//*[@id='tblPurListing']/tbody/tr/td[8]")
public static List<WebElement> purchaseValueAED ; 

@FindBy(xpath="//*[@id='tblPurListing']/tbody/tr/td[9]")
public static List<WebElement> purchaseVATVAlueAED ; 

@FindBy(xpath="//*[@id='tblPurListing']/tbody/tr/td[10]")
public static List<WebElement> purchaseTAxcode ; 


@FindBy(xpath="//*[@id='VATAuditFileBody']/div[2]/div/div[3]/label[4]")
public static WebElement  purchaseTransCountTotal; 

@FindBy(xpath="//*[@id='VATAuditFileBody']/div[2]/div/div[3]/label[3]")
public static WebElement  purchaseVatTotalAED;

@FindBy(xpath="//*[@id='VATAuditFileBody']/div[2]/div/div[3]/label[2]")
public static WebElement  purchaseTotalAED;



//Sales Grid





@FindBy(xpath="//*[@id='tblSupplySalListing']/tbody/tr/td[1]")
public static List<WebElement> CustomerNameList ; 

@FindBy(xpath="//*[@id='tblSupplySalListing']/tbody/tr/td[4]")
public static List<WebElement> salesInvoiceNoList ; 

@FindBy(xpath="//*[@id='tblSupplySalListing']/tbody/tr/td[6]")
public static List<WebElement> salesProductDescList ; 

@FindBy(xpath="//*[@id='tblSupplySalListing']/tbody/tr/td[7]")
public static List<WebElement> salesSupplyValueAED ; 

@FindBy(xpath="//*[@id='tblSupplySalListing']/tbody/tr/td[8]")
public static List<WebElement> salesVatValueAED ; 

@FindBy(xpath="//*[@id='tblSupplySalListing']/tbody/tr/td[9]")
public static List<WebElement> salesTaxcode ; 


@FindBy(xpath="//*[@id='VATAuditFileBody']/div[3]/div/div[3]/label[4]")
public static WebElement  salesTransCountTotal; 

@FindBy(xpath="//*[@id='VATAuditFileBody']/div[3]/div/div[3]/label[3]")
public static WebElement  salesVatTotalAED;

@FindBy(xpath="//*[@id='VATAuditFileBody']/div[3]/div/div[3]/label[2]")
public static WebElement  salesTotalAED;


//Ledger


@FindBy(xpath="//*[@id='tblGeneralLedger']/tbody/tr/td[2]")
public static List<WebElement> ledgerAccountIDList ; 

@FindBy(xpath="//*[@id='tblGeneralLedger']/tbody/tr/td[3]")
public static List<WebElement> ledgerAccountNameList ; 

@FindBy(xpath="//*[@id='tblGeneralLedger']/tbody/tr/td[4]")
public static List<WebElement> ledgerTransDescList ; 

@FindBy(xpath="//*[@id='tblGeneralLedger']/tbody/tr/td[5]")
public static List<WebElement> ledgerNameList ;

@FindBy(xpath="//*[@id='tblGeneralLedger']/tbody/tr/td[7]")
public static List<WebElement> ledgeSourceDocIDList ;

@FindBy(xpath="//*[@id='tblGeneralLedger']/tbody/tr/td[8]")
public static List<WebElement> ledgeSourceTypeList ;

@FindBy(xpath="//*[@id='tblGeneralLedger']/tbody/tr/td[9]")
public static List<WebElement> ledgeDebitList ;

@FindBy(xpath="//*[@id='tblGeneralLedger']/tbody/tr/td[10]")
public static List<WebElement> ledgeCreditList ;

@FindBy(xpath="//*[@id='tblGeneralLedger']/tbody/tr/td[11]")
public static List<WebElement> ledgeBalList ;


@FindBy(xpath="//*[@id='VATAuditFileBody']/div[4]/div/div[3]/label[5]")
public static WebElement  GLTCurrency; 

@FindBy(xpath="//*[@id='VATAuditFileBody']/div[4]/div/div[3]/label[4]")
public static WebElement  TransCountTotal; 

@FindBy(xpath="//*[@id='VATAuditFileBody']/div[4]/div/div[3]/label[3]")
public static WebElement  totalcredit;

@FindBy(xpath="//*[@id='VATAuditFileBody']/div[4]/div/div[3]/label[2]")
public static WebElement  totalDebit;


@FindBy(xpath="//div[contains(text(),'Exit')]")
public static WebElement vat_ExitBtn;

@FindBy(xpath="//*[@id='VATMonthlyReturnRepKSABody']//tr/td")
public static List<WebElement>  vatMonthlyreturnReportList; 


@FindBy(xpath="//div[contains(text(),'Back')]")
public static WebElement vat_BackBtn;




@FindBy(xpath="//*[@id='navigation_menu']/li[2]/ul/li[3]/ul/li[10]/ul/li/a/span")
private static List<WebElement> miscellaneousAnalysisReportList;



@FindBy(xpath="//div[@id='REPORTRENDERNEWControls']/ul/li/span[1]")
private static WebElement reportSortingBtn;	


@FindBy(xpath="//div[@id='REPORTRENDERNEWControls']/ul/li/ul/li[1]")
private static WebElement sortingAscendingBtn;	

@FindBy(xpath="//div[@id='REPORTRENDERNEWControls']/ul/li/ul/li[2]")
private static WebElement sortingDecendingBtn;	





@FindBy(xpath="//*[@id='navigation_menu']/li[2]/ul/li[3]/ul/li[11]/ul/li/a/span")
private static List<WebElement> mastersInformationReportList;



@FindBy(xpath="//input[@id='opt_DocumentTypeID']")
private static WebElement documentTxt;

@FindBy(xpath="//input[@id='opt_LayoutID']")
private static WebElement layoutTxt;

@FindBy(xpath="//input[@id='idrdDocument']")
private static WebElement documentRadioBtn;

@FindBy(xpath="//input[@id='idrdProduct']")
private static WebElement itemRadioBtn;

@FindBy(xpath="//i[@class='icon-add-preview icon-font6']")
private static WebElement pl_previewBtn;

@FindBy(xpath="//i[@class='icon-print icon-font6']")
private static WebElement pl_printBtn;

@FindBy(xpath="//i[@class='icon-close icon-font6']")
private static WebElement pl_CloseBtn;

//Skid Definition


@FindBy(xpath="//input[@id='skidId']")
private static WebElement skid;

@FindBy(xpath="//input[@id='txtprefix']")
private static WebElement prefixTxt;

@FindBy(xpath="//input[@id='txtStartingNumber']")
private static WebElement startingNoTxt;

@FindBy(xpath="//input[@id='chkSkid']")
private static WebElement enableReprintChkBox;

@FindBy(xpath="//input[@id='txtQuantity']")
private static WebElement quantityTxt;

@FindBy(xpath="//span[@class='icon-analyze icon-font6']")
private static WebElement sd_PreviewBtn;

@FindBy(xpath="//span[@class='icon-print icon-font6']")
private static WebElement sd_PrintBtn;

@FindBy(xpath="//span[@class='icon-close']")
private static WebElement sd_CloseBtn;


@FindBy(xpath="//*[@id='navigation_menu']/li[2]/ul/li[3]/ul/li[12]/ul/li/a/span")
private static List<WebElement> auditTrailList;

@FindBy(xpath="//span[contains(text(),'Excluded eleted transactions')]")
private static WebElement excludedElectedTransactionsChkBox;

@FindBy(xpath="//span[contains(text(),'Exclude modified transactions')]")
private static WebElement excludeModifiedTransactionsChkBox;

@FindBy(xpath="//span[contains(text(),'Include manual closed link')]")
private static WebElement includeManualClosedLinkChkBox;

@FindBy(xpath="//select[@id='RITCombobox__5']")
private static WebElement displayOrderDropDown;

@FindBy(xpath="//input[@id='RITTable__0']")
private static WebElement reportVoucherTxt;

@FindBy(xpath="//*[@id='RITTable__0_table_body']/tr/td[2]")
private static List<WebElement> reportVoucherTxtList;


@FindBy(xpath="//input[@id='RITTable__0']")
private static WebElement reportUsersTxt;

@FindBy(xpath="//*[@id='RITTable__0_table_body']/tr/td[2]")
private static List<WebElement> reportUsersTxtList;




//Account Query

@FindBy(xpath="//input[@id='accountQuery']")
private static WebElement accountQueryTxt;

@FindBy(xpath="//button[@class='btn btn-default btn-sm']")
private static WebElement getAccountDetailsBtn;

@FindBy(xpath="//label[contains(text(),'Opening Balance :')]")
private static WebElement aq_OpeningBalanceLable;

@FindBy(xpath="//label[contains(text(),'Credit :')]")
private static WebElement aq_CreditLable;

@FindBy(xpath="//label[contains(text(),'Debit :')]")
private static WebElement aq_DebitLable;

@FindBy(xpath="//div[5]//label[1]")
private static WebElement aq_BalanceLable;

@FindBy(xpath="//table[@id='creditTransTable']//th[contains(text(),'Date')]")
private static WebElement ltc_Date;

@FindBy(xpath="//table[@id='creditTransTable']//th[contains(text(),'Voucher')]")
private static WebElement ltc_Vocuher;

@FindBy(xpath="//table[@id='creditTransTable']//th[contains(text(),'Account')]")
private static WebElement ltc_Account;

@FindBy(xpath="//th[contains(text(),'Credit')]")
private static WebElement ltc_Credit;

@FindBy(xpath="//table[@id='debitTransTable']//th[contains(text(),'Date')]")
private static WebElement ltd_Date;

@FindBy(xpath="//table[@id='debitTransTable']//th[contains(text(),'Voucher')]")
private static WebElement ltd_Vocuher;

@FindBy(xpath="//table[@id='debitTransTable']//th[contains(text(),'Account')]")
private static WebElement ltd_Account;

@FindBy(xpath="//th[contains(text(),'Debit')]")
private static WebElement ltd_Dedit;

@FindBy(xpath="//span[@class='icon-panel icon-font7 pull-right']")
private static WebElement deptWiseBalnceExpandBtn;

@FindBy(xpath="//i[@class='icon-clear icon-font6']")
private static WebElement aq_ClearBtn;

@FindBy(xpath="//i[@class='icon-close icon-font6']")
private static WebElement aq_CloseBtn;

@FindBy(xpath="//span[@id='openingBal']")
private static WebElement  AQ_OpnBal;

@FindBy(xpath="//span[@id='credit']")
private static WebElement  AQ_Credit;

@FindBy(xpath="//span[@id='debit']")
private static WebElement  AQ_Debit;

@FindBy(xpath="//span[@id='balance']")
private static WebElement  AQ_Balance;

@FindBy(xpath="//*[@id='creditTransTable']/tbody/tr/td[2]")
private static List<WebElement>  AQ_CreditTransVouList;

@FindBy(xpath="//*[@id='creditTransTable']/tbody/tr/td[3]")
private static List<WebElement>  AQ_CreditTransAccountList;

@FindBy(xpath="//*[@id='creditTransTable']/tbody/tr/td[4]")
private static List<WebElement>  AQ_CreditTransCreditList;


@FindBy(xpath="//*[@id='debitTransTable']/tbody/tr/td[2]")
private static List<WebElement>  AQ_DebitTransVouList;

@FindBy(xpath="//*[@id='debitTransTable']/tbody/tr/td[3]")
private static List<WebElement>  AQ_DebitTransAccountList;

@FindBy(xpath="//*[@id='debitTransTable']/tbody/tr/td[4]")
private static List<WebElement>  AQ_DebitTransDebitList;

@FindBy(xpath="//*[@id='tagWiseTransData']/tbody/tr/td")
private static List<WebElement>  AQ_DepartmentList;


/////////////////////////



@FindBy(xpath="//div[@id='dv_TranHomeHeading']//span[@id='dv_sVoucherName']")
private static WebElement  openingStocksLabel;

@FindBy(xpath="//div[@id='dv_TranHomeHeading']//span[@id='dv_sVoucherName']")
private static WebElement  stocksTransferLabel;



	
/*
@FindBy(xpath="//input[@id='id_body_536870916']")
private static WebElement  pvWareHouseTxt;

@FindBy(xpath="//input[@id='id_header_268435460']")
private static WebElement  wareHouseTxt;*/


@FindBy(xpath="//*[@id='id_header_87']")
private static WebElement  stockTransferWareHouse2InHeaderTxt;
/*           


@FindBy(xpath="//*[@id='id_transactionentry_printbarcode']")
private static WebElement  printBarCodeOption;
*/

@FindBy(xpath="//span[contains(text(),'Inventory Options')]")
private static WebElement  inventoryTab;


@FindBy(xpath="//input[@id='inventoryOpt_chkStocktransferInput']")
private static WebElement  warehouse2HeaderChkBox;

	
@FindBy(xpath = "//a[@id='200']//span[contains(text(),'Reports')]")
public static WebElement inventoryReportsMenu;
	
@FindBy(xpath = "//*[@id='2034']/span")
public static WebElement shortagesInStockVoucher;

@FindBy(xpath="//*[@id='cmbColumn']")
public static WebElement ld_RefineLookinDropdown;

@FindBy(xpath="//*[@id='SelectOperator']")
public static WebElement ld_RefineOperatorDropdown;

@FindBy(xpath="//*[@id='idDisplay']//option")
public static List<WebElement> ld_RefineData;

@FindBy(xpath="//*[@id='idDisplay']")
public static WebElement ld_SelectRefineData;

//Final Accounts

@FindBy(xpath="//*[@id='navigation_menu']/li[2]/ul/li[5]")
public static WebElement finalAccountsMenu;

@FindBy(xpath="//*[@id='navigation_menu']/li[2]/ul/li[5]/ul/li[2]")
public static WebElement trailBalanceMenu;

@FindBy(xpath="//*[@id='id_mainreportmenuheadings']/ul/li[10]")
public static WebElement fa_OkBtn;


//Receivable and Payable

@FindBy(xpath="//*[@id='navigation_menu']/li[2]/ul/li[6]")
public static WebElement receivablePayableMenu;


@FindBy(xpath="//*[@id='navigation_menu']/li[2]/ul/li[6]/ul/li[2]")
public static WebElement rp_CustomerDetailMenu;

@FindBy(xpath="//*[@id='navigation_menu']/li[2]/ul/li[6]/ul/li[2]/ul/li[3]")
public static WebElement rp_CustomerStatements;


@FindBy(xpath="//*[@id='navigation_menu']/li[2]/ul/li[6]/ul/li[5]")
public static WebElement rp_VendorSummaryMenu;

@FindBy(xpath="//*[@id='navigation_menu']/li[2]/ul/li[6]/ul/li[5]/ul/li[2]")
public static WebElement rp_VendorAgeingSummary;


}

	