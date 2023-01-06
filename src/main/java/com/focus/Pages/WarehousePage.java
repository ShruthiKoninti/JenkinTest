package com.focus.Pages;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
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

public class WarehousePage extends BaseEngine 
{
	//Home Menu
	@FindBy (xpath="//*[@id='1']")
    public static WebElement homeMenu;
	
        //Masters Menu	
		@FindBy (xpath="//a[@id='1000']")
	    public static WebElement mastersMenu;
		
		    //Accounts 
			@FindBy (xpath="//a[@id='1104']//span[contains(text(),'Account')]")
			public static WebElement accounts;
			
				//Accounts Title
				@FindBy (xpath="//span[@id='spnHeaderText']")
				public static WebElement accountsTitle;
		
				//Master Main Header Fields		
				@FindBy(xpath="//i[@class='icon-new hiconright2']")
				public static WebElement masterNewBtn;
					
				@FindBy(xpath="//i[@class='icon-addgroup hiconright2']")
				public static WebElement masterAddGroupBtn; 
				 
				@FindBy(xpath="//i[@class='icon-edit hiconright2']")
				public static WebElement masterEditBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-clone']")
				public static WebElement masterCloneBtn; 
		
				@FindBy(xpath="//i[@class='icon-properties icon-font6']")
				public static WebElement masterPropertiesBtn; 
		
				@FindBy(xpath="//*[@class='icon-delete hiconright2']")
				public static WebElement masterDeleteBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-group']")
				public static WebElement masterGroupBtn; 
				
				@FindBy(xpath="//*[@id='toggle_ribbon']")
				public static WebElement masterRibbonToExpandOptions; 
		
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
		
				@FindBy(xpath="//a[@class='lSNext']")
				public static WebElement masterRibbonControlNextBtn; 
				
				@FindBy(xpath="//*[@id='btnXMLImport']")
				public static WebElement masterImportFromXmlBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-xmlexport']")
				public static WebElement masterExportFormatToXmlBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-custamize']")
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
		




		//Master Left panel----TreeID-----Elements

		@FindBy(xpath="//*[@id='iTreeId']")
		public static WebElement masterTreeIdDropdown; 

		@FindBy(xpath="//*[@id='spnhiererchy']/div/div[1]/div/div/button/span")
		public static WebElement masterTreeIDToggleBtn; 
		                        
				@FindBy(xpath="//*[@id='spnhiererchy']/div/div[1]/div/div/ul/li[1]/a")
				public static WebElement masterTreeIDToggle_CreateTreeBtn; 
				
							@FindBy(xpath="//h4[contains(text(),'Create Tree')]")
							public static WebElement treeID_CreateTreeLabel; 
							
							@FindBy(xpath="//*[@id='btnCancel']")
							public static WebElement treeID_CloseBtn; 
							
							@FindBy(xpath="//input[@id='txtTreeName']")
							public static WebElement treeID_TreeTxt; 
							
							@FindBy(xpath="//input[@id='rbtManual']")
							public static WebElement treeID_ManualRadioBtn; 
							
							@FindBy(xpath="//input[@id='rbtAuto']")
							public static WebElement treeID_AutoRadioBtn; 
							
							@FindBy(xpath="//select[@id='cmbTreeList']")
							public static WebElement treeID_SelectTreeDropdown; 
							
							@FindBy(xpath="//input[@id='chkIncludeGroup']")
							public static WebElement treeID_IncludeGroupsCheckBox; 
							
							@FindBy(xpath="//input[@id='chkAddLeaveAtZeroLevel']")
							public static WebElement treeID_AddLeafAlwaysAtTheZerolevelOnCreationCheckbox; 
							
							@FindBy(xpath="//*[@id='MasterTreeTable_SelectFields']")
							public static WebElement treeID_GridFirstSelectFields; 
							
							@FindBy(xpath="//*[@id='MasterTreeTable_col_2-1']")
							public static WebElement treeID_GridSecondSelectFields; 
							
							@FindBy(xpath="//input[@id='chkAutoGroup']")
							public static WebElement treeID_AlwaysAutoGroupOnCreationCheckbox; 
							
							@FindBy(xpath="//input[@id='chkAlwaysSorted']")
							public static WebElement treeID_AlwaysSortedCheckbox; 
							
							@FindBy(xpath="//*[@id='btnOk']")
							public static WebElement treeID_PopOkBtn; 
							
							@FindBy(xpath="//*[@id='btnCancel']")
							public static WebElement treeID_PopCancelBtn; 
				
				@FindBy(xpath="//*[@id='spnhiererchy']/div/div[1]/div/div/ul/li[2]")
				public static WebElement masterTreeIDToggle_EditBtn; 
				
				@FindBy(xpath="//*[@id='spnhiererchy']/div/div[1]/div/div/ul/li[3]")
				public static WebElement masterTreeIDToggle_DeleteBtn; 
				
				@FindBy(xpath="//*[@id='spnhiererchy']/div/div[1]/div/div/ul/li[4]")
				public static WebElement masterTreeIDToggle_setDefaultBtn;


		//TreeViewID----->Elements
				
		@FindBy(xpath="//select[@id='iTreeViewId']")
		public static WebElement masterTreeViewIDDropdown; 

		@FindBy(xpath="//*[@id='spnhiererchy']/div/div[2]/div/div/button/span")
		public static WebElement masterTreeViewIDToggleBtn;

				@FindBy(xpath="//*[@id='spnhiererchy']/div/div[2]/div/div/ul/li[1]")
				public static WebElement masterTreeViewIdCreateViewBtn; 
				
								@FindBy(xpath="//*[@id='createViewDiv']/div/div[1]/div[1]/h4")
								public static WebElement CreateView_Label; 
				
				
								@FindBy(xpath="//*[@id='btnCancel']/i")
								public static WebElement CreateView_CloseBtn; 
				
								@FindBy(xpath="//*[@id='ctrlCreateMasterViewH']")
								public static WebElement CreateView_CreateViewTxt; 
				
									@FindBy(xpath="//*[@id='ctrlCreateMasterViewH_input_image']/span")
									public static WebElement CreateView_ExpansionBtn; 
					
									@FindBy(xpath="//*[@id='ctrlCreateMasterViewH_input_settings']/span")
									public static WebElement CreateView_SettingBtn; 
					
									@FindBy(xpath="//input[@id='chkAvailbleonMobileApp']")
									public static WebElement CreateView_AvailbleonMobileAppCheckbox; 
					
									@FindBy(xpath="//*[@id='1104_0_AdvanceFilter']/table/tbody/tr/td[1]/select")
									public static WebElement CreateView_Conjuction1Dropdown; 
					
									@FindBy(xpath="//*[@id='1104_0_AdvanceFilter']/table/tbody/tr/td[2]/select")
									public static WebElement CreateView_AccountType1Dropdown; 
									                      
									@FindBy(xpath="//*[@id='1104_0_AdvanceFilter']/table/tbody/tr/td[3]/select")
									public static WebElement CreateView_EqaulTo1Dropdown; 
									
									@FindBy(xpath="//*[@id='1104_0_AdvanceFilter']/table/tbody/tr/td[4]/select")
									public static WebElement CreateView_Value1Dropdown; 
					
									@FindBy(xpath="//*[@id='1104_0_AdvanceFilter']/table/tbody/tr/td[5]/input")
									public static WebElement CreateView_Customer1Dropdown; 
									
									@FindBy(xpath="//*[@id='1104_0_AdvanceFilter']/table/tbody/tr/td[6]/span")
									public static WebElement CreateView_CloseFirstrow; 
												
									@FindBy(xpath="//*[@id='1104_0_AdvanceFilter']/table/tbody/tr/td[7]/span")
									public static WebElement CreateView_AddRowBtn; 				
									
									////Creation For Second Row
									
									@FindBy(xpath="//*[@id='1104_0_AdvanceFilter']/table/tbody/tr/td[1]/select")
									public static WebElement CreateView_Conjuction2Dropdown; 
					
									@FindBy(xpath="//*[@id='1104_0_AdvanceFilter']/table/tbody/tr[2]/td[2]/select")
									public static WebElement CreateView_AccountType2Dropdown; 
									                      
									@FindBy(xpath="//*[@id='1104_0_AdvanceFilter']/table/tbody/tr[2]/td[3]/select")
									public static WebElement CreateView_EqaulTo2Dropdown; 
									
									@FindBy(xpath="//*[@id='1104_0_AdvanceFilter']/table/tbody/tr[2]/td[4]/select")
									public static WebElement CreateView_Value2Dropdown; 
					
									@FindBy(xpath="//*[@id='1104_0_AdvanceFilter']/table/tbody/tr[2]/td[5]/input")
									public static WebElement CreateView_Customer2Dropdown; 
									
									@FindBy(xpath="//*[@id='1104_0_AdvanceFilter']/table/tbody/tr[2]/td[6]/span")
									public static WebElement CreateView_CloseSecondrow; 
					
								@FindBy(xpath="//button[@id='btnMasterViewDelete']")
								public static WebElement CreateView_DeleteBtn; 
				
								@FindBy(xpath="//button[@id='btnSave']")
								public static WebElement CreateView_SaveBtn; 
				
								@FindBy(xpath="//*[@id='btnCancel']")
								public static WebElement CreateView_CancelBtn; 
				
				@FindBy(xpath="//*[@id='spnhiererchy']/div/div[2]/div/div/ul/li[2]/a")
				public static WebElement masterTreeeViewIdEditBtn; 
				
				@FindBy(xpath="//*[@id='spnhiererchy']/div/div[2]/div/div/ul/li[3]/a")
				public static WebElement masterTreeViewIdDeleteBtn; 
				
			@FindBy(xpath="//div[@id='divTreeMenu']")
			public static WebElement masterTreeDisplayMenuAccountGroups; 		
				
		@FindBy(xpath="//*[@id='oncheckaccorderan']")
		public static WebElement masterOptionToHideAccountsGroupTree; /////To hide TreeId and TreeViewId Elements From HomeScreen

		@FindBy(xpath="//*[@id='0']/a")
		private static WebElement account;
		
		@FindBy(xpath="//*[@id='18']/a[1]")
		private static WebElement accountAssestsGroup;
		
		@FindBy(xpath="//*[@id='17']/a[2]")
		private static WebElement accountFixedAssestsGroup;
		
		@FindBy(xpath="//*[@id='23']/a[2]")
		private static WebElement accountCurrentAssetsGroup;

		@FindBy(xpath="//*[@id='8']/a[2]")
		private static WebElement accountCAshAndBankGroup;
		
		@FindBy(xpath="//*[@id='13']/a[2]")
		private static WebElement accountInventoriesGroup;
		
		@FindBy(xpath="//*[@id='22']/a[2]")
		private static WebElement accountAccountsReceivableGroup;
		
		@FindBy(xpath="//a[@class='clsMasterName'][contains(text(),'Customer Update Group')]")
		private static WebElement accountAccountsReceivableCustomerUpdateGroup;
		
		@FindBy(xpath="//*[@id='72']/a[2]")
		private static WebElement accountInvestmentsGroup;
		
		@FindBy(xpath="//*[@id='25']/a[2]")
		private static WebElement accountExpensesGroup;
		
		@FindBy(xpath="//*[@id='24']/a[2]")
		private static WebElement accountDirectExpensesGroup;
		
		@FindBy(xpath="//*[@id='73']/a[2]")
		private static WebElement accountIndirectExpensesGroup;
		
		@FindBy(xpath="//*[@id='54']/a[2]")
		private static WebElement accountAdministrativeExpensesGroup;
		
		@FindBy(xpath="//*[@id='60']/a[2]")
		private static WebElement accountEmployeeBenefitsGroup;
		
		@FindBy(xpath="//*[@id='64']/a[2]")
		private static WebElement accountFinancialChargesGroup;
		
		@FindBy(xpath="//*[@id='69']/a[2]")
		private static WebElement accountGainAndLossGroup;
		
		@FindBy(xpath="//*[@id='26']/a[2]")
		private static WebElement accountControlAccountsGroup;
		
		@FindBy(xpath="//*[@id='29']/a[2]")
		private static WebElement accountRevenueGroup;

		@FindBy(xpath="//*[@id='28']/a[2]")
		private static WebElement accountSalesGroup;
		
		@FindBy(xpath="//*[@id='33']/a[2]")
		private static WebElement accountEqitiesGroup;
		
		@FindBy(xpath="//*[@id='32']/a[2]")
		private static WebElement accountCapitalGroup;
		
		@FindBy(xpath="//*[@id='39']/a[2]")
		private static WebElement accountLiabilitiesGroup;
		
		@FindBy(xpath="//*[@id='38']/a[2]")
		private static WebElement accountLoanAndBorrowingGroup;
		
		@FindBy(xpath="//*[@id='37']/a[2]")
		private static WebElement accountLoansGroup;
		
		@FindBy(xpath="//*[@id='42']/a[2]")
		private static WebElement accountAccruedLiabilitiesGroup;
		
		@FindBy(xpath="//*[@id='45']/a[2]")
		private static WebElement accountTradePayableGroup;
		
		@FindBy(xpath="//*[@id='48']/a[2]")
		private static WebElement accountProvisionsGroup;
		
		//Header Elements of Master MainLanding

		@FindBy(xpath="//*[@id='txtsrch-term']")
		public static WebElement masterSearchTxt; 

		@FindBy(xpath="//*[@id='btnSearchAcc1']/i")
		public static WebElement masterSearchBtn; 

			@FindBy(xpath="//a[contains(text(),'Search on')]")
			public static WebElement masterSearchTxtBtn_SearchOnBtn; 

				@FindBy(xpath="//label[@class='form-group']")
				public static WebElement searchOn_SelectFilterLabel; 
				
				@FindBy(xpath="//select[@id='ddlSelectFilter']")
				public static WebElement searchOn_SelectFilterDropdown; 
				
				@FindBy(xpath="//input[@id='chkSearchSelectAll']")
				public static WebElement searchOn_SelectAllCheckBox; 
				
				@FindBy(xpath="//*[@id='0']")
				public static WebElement searchOn_NameCheckbox; 
				
				@FindBy(xpath="//input[@id='1']")
				public static WebElement searchOn_AliasCheckbox; 
				
				@FindBy(xpath="//input[@id='2']")
				public static WebElement searchOn_COdeCheckbox; 
				
				@FindBy(xpath="//input[@id='3']")
				public static WebElement searchOn_AccountTypeCheckbox; 
				
				@FindBy(xpath="//input[@id='4']")
				public static WebElement searchOn_CreditLimitCheckbox; 
				
				@FindBy(xpath="//input[@id='5']")
				public static WebElement searchOn_CreditDaysCheckbox; 
				
				@FindBy(xpath="//input[@id='6']")
				public static WebElement searchOn_CreatedByCheckbox; 
				
				@FindBy(xpath="//input[@id='7']")
				public static WebElement searchOn_ModifiedByCheckbox; 
				
				@FindBy(xpath="//input[@id='8']")
				public static WebElement searchOn_CreatedDateCheckbox; 
				
				@FindBy(xpath="//input[@id='9']")
				public static WebElement searchOn_ModifiedDate; 
				
				@FindBy(xpath="//input[@id='10']")
				public static WebElement searchOn_AllowOtherCompaniesToViewTheRecordCheckbox; 
				
				@FindBy(xpath="//input[@id='11']")
				public static WebElement searchOn_SyscReceiverDateCheckbox; 
				
				@FindBy(xpath="//input[@id='12']")
				public static WebElement searchOn_EditingLocationCheckbox; 
				
				@FindBy(xpath="//input[@id='13']")
				public static WebElement searchOn_ChequeDiscountLimitCheckbox; 
				
				@FindBy(xpath="//input[@id='14']")
				public static WebElement searchOn_RateOfInterestCheckbox; 
				
				@FindBy(xpath="//input[@id='15']")
				public static WebElement searchOn_BankACCheckbox; 
				
				@FindBy(xpath="//input[@id='16']")
				public static WebElement searchOn_PDCDIscountedAC; 
				
				@FindBy(xpath="//input[@id='17']")
				public static WebElement searchOn_DebitCreditProposalCheckbox; 
				
				@FindBy(xpath="//input[@id='18']")
				public static WebElement searchOn_DebitCreditRequiredCheckbox; 
				
				@FindBy(xpath="//input[@id='19']")
				public static WebElement searchOn_ExchangeAdjustmentGainACCheckbox; 
				
				@FindBy(xpath="//input[@id='20']")
				public static WebElement searchOn_ExchangeAdjustmentLossACCheckbox; 
				
				@FindBy(xpath="//input[@id='21']")
				public static WebElement searchOn_PrimaryAccountCheckbox; 
				
				@FindBy(xpath="//input[@id='22']")
				public static WebElement searchOn_DefaultCurrencyCheckbox; 
				
				@FindBy(xpath="//input[@id='23']")
				public static WebElement searchOn_ConsolidationMethodCheckbox; 
				
				@FindBy(xpath="//input[@id='24']")
				public static WebElement searchOn_PaymentTermsCheckbox; 
				
				@FindBy(xpath="//input[@id='25']")
				public static WebElement searchOn_RemainderTermsCheckbox; 
				
				@FindBy(xpath="//input[@id='26']")
				public static WebElement searchOn_FinanceChargeTermsCheckbox; 
				
				@FindBy(xpath="//input[@id='27']")
				public static WebElement searchOn_AddressCheckbox; 
				
				@FindBy(xpath="//input[@id='28']")
				public static WebElement searchOn_City1Checkbox; 
				
				@FindBy(xpath="//input[@id='29']")
				public static WebElement searchOn_Pin1Checkbox; 
				
				@FindBy(xpath="//input[@id='30']")
				public static WebElement searchOn_DeliveryAddreddCheckbox; 
				
				@FindBy(xpath="//input[@id='31']")
				public static WebElement searchOn_City2Checkbox; 
				
				@FindBy(xpath="//input[@id='32']")
				public static WebElement searchOn_Pin2Checkbox; 
				
				@FindBy(xpath="//input[@id='33']")
				public static WebElement searchOn_SendEmailToCustomerCheckbox; 
				
				@FindBy(xpath="//input[@id='34']")
				public static WebElement searchOn_AllowCustomerPortalCheckbox; 
				
				@FindBy(xpath="//input[@id='35']")
				public static WebElement searchOn_EmailCheckbox; 
				
				@FindBy(xpath="//input[@id='36']")
				public static WebElement searchOn_PasswordCheckbox; 
				
				@FindBy(xpath="//input[@id='37']")
				public static WebElement searchOn_TelNOCheckbox; 
				
				@FindBy(xpath="//input[@id='38']")
				public static WebElement searchOn_FaxNoCheckbox; 
				
				@FindBy(xpath="//input[@id='39']")
				public static WebElement searchOn_BankAccountNameCheckbox; 
				
				@FindBy(xpath="//input[@id='40']")
				public static WebElement searchOn_BankAccountNumberCheckbox; 
				
				@FindBy(xpath="//input[@id='41']")
				public static WebElement searchOn_IFSCCodeCheckbox; 
				
				@FindBy(xpath="//input[@id='42']")
				public static WebElement searchOn_PaymentTypeCheckbox; 
				
				@FindBy(xpath="//input[@id='43']")
				public static WebElement searchOn_FinanceEmailCheckbox; 
				
				@FindBy(xpath="//input[@id='44']")
				public static WebElement searchOn_PortalEmailCheckbox; 
				
				@FindBy(xpath="//input[@id='45']")
				public static WebElement searchOn_NetbalanceCheckbox; 
				
				@FindBy(xpath="//input[@id='46']")
				public static WebElement searchOn_DRBalanceCheckbox; 
				
				@FindBy(xpath="//input[@id='47']")
				public static WebElement searchOn_CRbalanceCheckbox; 
				
				@FindBy(xpath="//*[@id='divSearchMenu']/div[4]/button")
				public static WebElement searchOn_OkBtn;
						
			@FindBy(xpath="//a[contains(text(),'Advance Search')]")
			public static WebElement masterSearchTxtBtn_AdvanceSearchBtn; 

  
	            //First Row Fields

				@FindBy(xpath="//*[@id='1104_1_AdvanceFilter']/table/tbody/tr[1]/td[1]/select")
				public static WebElement advanceSearch_Conjuction1Dropdown;
				
				@FindBy(xpath="//tbody//input[@placeholder='Select Field']")
				public static WebElement advanceSearch_SelectField1Dropdown;
				
				@FindBy(xpath="//*[@id='1104_1_AdvanceFilter']/table/tbody/tr[1]/td[3]/select")
				public static WebElement advanceSearch_SelectOperator1Dropdown;
				
				@FindBy(xpath="//*[@id='1104_1_AdvanceFilter']/table/tbody/tr[1]/td[4]/select")
				public static WebElement advanceSearch_CompareWith1Dropdown;
				
				@FindBy(xpath="//*[@id='1104_1_AdvanceFilter']/table/tbody/tr[1]/td[5]/input")
				public static WebElement advanceSearch_Value1Txt;
				
				@FindBy(xpath="//*[@id='1104_1_AdvanceFilter']/table/tbody/tr[1]/td[6]/span")
				public static WebElement advanceSearch_CloseRow1Btn;
				
				@FindBy(xpath="//*[@id='1104_1_AdvanceFilter']/table/tbody/tr[1]/td[7]/span")
				public static WebElement advanceSearch_AddRowBtn;
				
			    //Second Row Fields
				@FindBy(xpath="//*[@id='1104_1_AdvanceFilter']/table/tbody/tr[2]/td[1]/select")
				public static WebElement advanceSearch_ConjuctionDropdown;
				
				@FindBy(xpath="//tbody//input[@placeholder='Select Field']")
				public static WebElement advanceSearch_SelectFieldDropdown;
				
				@FindBy(xpath="//*[@id='1104_1_AdvanceFilter']/table/tbody/tr[2]/td[3]/select")
				public static WebElement advanceSearch_SelectOperatorDropdown;
				
				@FindBy(xpath="//*[@id='1104_1_AdvanceFilter']/table/tbody/tr[2]/td[4]/select")
				public static WebElement advanceSearch_CompareWithDropdown;
				
				@FindBy(xpath="//*[@id='1104_1_AdvanceFilter']/table/tbody/tr[2]/td[5]/input")
				public static WebElement advanceSearch_ValueTxt;
				
				@FindBy(xpath="//*[@id='1104_1_AdvanceFilter']/table/tbody/tr[2]/td[6]/span")
				public static WebElement advanceSearch_CloseRow2Btn;
						
				@FindBy(xpath="//*[@id='divCreateNewTreeModalBody']/div[2]/div[2]/input")
				public static WebElement advanceSearch_GenerateQueryBtn;
				
				@FindBy(xpath="//div[@id='divStrQuery']")
				public static WebElement advanceSearch_QueryDescriptionTxt;
				
				@FindBy(xpath="//button[@id='btnOk']")
				public static WebElement advanceSearch_OkBtn;
				
				@FindBy(xpath="//*[@id='btnCancel']")
				public static WebElement advanceSearch_CancelBtn;
				
				@FindBy(xpath="//*[@id='btnCancel']/i")///Both Has Same ID But But Different in Fields 
				public static WebElement advanceSearch_CloseBtn;							

		@FindBy(xpath="//input[@id='cmbUserTypeMaster']")
		public static WebElement mastercmbMasterTxt; 

		@FindBy(xpath="//*[@id='cmbUserTypeMaster_input_image']/span")
		public static WebElement masterCmbMasterExpansionBtn; 

			@FindBy(xpath="//*[@id='cmbUserTypeMaster_input_settings']/span")
			public static WebElement masterCumMasterSettingBtn; 
						
				@FindBy(xpath="//*[@id='cmbUserTypeMaster_customize_popup_container']")
				public static WebElement masterCumMaster_Pop_sNameTxt;
				
				@FindBy(xpath="//*[@id='cmbUserTypeMaster_customize_popup_footer']/div/div/input[1]")
				public static WebElement masterCumMaster_StandardFieldsBtn;
					
					@FindBy(xpath="//*[@id='cmbUserTypeMaster_customize_popup_heading']")
					public static WebElement masterCumMaster_PopCustomizeDispalyColumnsLabel; 		
					
					@FindBy(xpath="//*[@id='cmbUserTypeMaster_customize_popup_standardfields_list']")
					public static WebElement masterCumMaster_Pop_StandardFields_Filedsdropdown;
					
					@FindBy(xpath="//*[@id='cmbUserTypeMaster_customize_popup_standardfields_header']")
					public static WebElement masterCumMaster_Pop_StandardFields_HeaderTxt;
					
					@FindBy(xpath="//*[@id='cmbUserTypeMaster_customize_popup_standardfields_alignment']")
					public static WebElement masterCumMaster_Pop_StandardFields_Allignmentdropdown;
					
					@FindBy(xpath="//*[@id='cmbUserTypeMaster_customize_popup_standardfields_width']")
					public static WebElement masterCumMaster_Pop_StandardFields_WidthTxt;
					
					@FindBy(xpath="//*[@id='cmbUserTypeMaster_customize_popup_footer']/div/div/input[3]")
					public static WebElement masterCumMaster_Pop_StandardFields_OkBtn;
					
					@FindBy(xpath="//*[@id='cmbUserTypeMaster_customize_popup_footer']/div/div/input[4]")
					public static WebElement masterCumMaster_Pop_StandardFields_CancelBtn;
					
			@FindBy(xpath="//*[@id='cmbUserTypeMaster_customize_popup_footer']/div/div/input[2]")
			public static WebElement masterCumMaster_Pop_DeleteColumnBtn;
			
			@FindBy(xpath="//*[@id='cmbUserTypeMaster_customize_popup_footer']/div/div/input[3]")
			public static WebElement masterCumMaster_Pop_OkBtn;
			
			@FindBy(xpath="//*[@id='cmbUserTypeMaster_customize_popup_footer']/div/div/input[4]")
			public static WebElement masterCumMaster_Pop_CancelBtn;
					

		@FindBy(xpath="//*[@id='li_moveup']/div/div[1]/a[1]/span")
		public static WebElement masterMoveTheSelectionToNextBookMarkBtn; 

		@FindBy(xpath="//span[@class='icon-font6 icon-bookmarkdown theme_icon-color']")
		public static WebElement masterMoveTheSelectionToPrevoiusBookmarkBtn; 

		@FindBy(xpath="//span[@class='theme_icon-color icon-font6 icon-show-all-records']")
		public static WebElement masterShowAllRecordsBtn; 

		@FindBy(xpath="//span[@class='icon-font6 icon-missmatched-entries theme_icon-color']")
		public static WebElement masterShowAllUnauthorisedRecords; 


		@FindBy(xpath="//span[@class='icon-font6 icon-closed-records theme_icon-color']")
		public static WebElement masterShowAllClosedRecordsBtn; 

		@FindBy(xpath="//span[@class='icon-font6 icon-external-module theme_icon-color']")
		public static WebElement masterAutoAdjustColoumnsBtn; 

		@FindBy(xpath="//input[@id='chkRetainSelection']")
		public static WebElement masterRetainSelectionCheckBox; 

		@FindBy(xpath="//*[@id='btnsideBar']")
		public static WebElement masterSideBarBtn; 

		//InfoPanel
		

		@FindBy(xpath="//label[contains(text(),'Info Panel')]")
		public static WebElement masterInfoPanelLabel; 

		@FindBy(xpath="//i[@class='icon-custamize theme_icon-color']")
		public static WebElement masterInfoPanelCustammizeBtn;


			@FindBy(xpath="//h4[contains(text(),'Customize InfoPanel')]")
			public static WebElement infoPanel_CustomizeInfoPanelLabel;
			
			@FindBy(xpath="//*[@id='Customize_AddNew_InfoPanel']/div[1]/span")
			public static WebElement infoPanel_PopAddBtn;
			
			@FindBy(xpath="//*[@id='Customize_RemoveOne_InfoPanel']/div[1]/span")
			public static WebElement infoPanel_PopRemoveBtn;
			
			@FindBy(xpath="//*[@id='btnShowPanelNames_CustomizeInfoPanel']/div/span")
			public static WebElement infoPanel_PopAllIcon;
			
			@FindBy(xpath="//*[@id='btnShowPanelDetails_CustomizeInfoPanel']/div/span")
			public static WebElement infoPanel_PopDetailsIcon;
			
			@FindBy(xpath="//*[@id='tbl_infoPanel_names']/tbody/tr[1]/td[4]/i")
			public static WebElement infoPane_PopGridCreditHistoryAddBtn;
			
			@FindBy(xpath="//*[@id='tbl_infoPanel_names']/tbody/tr[2]/td[4]/i")
			public static WebElement infoPane_PopGridAgingAnalysisAddBtn;
			
			@FindBy(xpath="//*[@id='tbl_infoPanel_names']/tbody/tr[3]/td[4]/i")
			public static WebElement infoPane_PopGridAgingAnlaysisBaseAddBtn;
			
			@FindBy(xpath="//*[@id='tbl_infoPanel_names']/tbody/tr[4]/td[4]/i")
			public static WebElement infoPane_PopGridAuthorizationAddBtn;
			
			@FindBy(xpath="//a[@class='Fbutton previous_link']")
			public static WebElement infoPane_PopDetailsPreviousBtn;
			
			@FindBy(xpath="//*[@id='page_navigation']/a[2]")
			public static WebElement infoPane_PopDeailsOneBtn;
			
			@FindBy(xpath="//*[@id='customize_InfoPanel_btnSave']")
			public static WebElement infoPane_PopDeatilsSaveBtn;
			
			@FindBy(xpath="//*[@id='customize_InfoPanel_btnCancel']")
			public static WebElement infoPane_PopDetailsCloseBtn;
			
			@FindBy(xpath="//*[@id='btnCancel']/i")
			public static WebElement infoPane_PopCloseBtn;


		@FindBy(xpath="//i[@id='id_InfoPanelEdit']")
		public static WebElement masterInfoPanelEditBtn;

		@FindBy(xpath="//i[@id='infopanelCloseButton']")
		public static WebElement masterInfoPanelCloseBtn;


		//////Grid Elements ---------->Body Field******************

		@FindBy(xpath="//th[2]//span[1]")
		public static WebElement masterGrid_Header_SelectTxt;

		@FindBy(xpath="//input[@id='liSelectAllMasters']")
		public static WebElement masterGridHeader_SelectAllRowsCheckBox;

		@FindBy(xpath="//section[@id='page_Content']//th[3]")
		public static WebElement masterGridHeader_MasterId;

		@FindBy(xpath="//section[@id='page_Content']//th[4]]")
		public static WebElement masterGridHeader_Name;
	
		@FindBy(xpath="//td[contains(text(),'ASSETS')]")
		public static WebElement masterAssets;
		
		@FindBy(xpath="//td[contains(text(),'EXPENSES')]")
		public static WebElement masterExpenses;
		
		@FindBy(xpath="//td[contains(text(),'CONTROL ACCOUNTS')]")
		public static WebElement masterControlAccounts;
		
		@FindBy(xpath="//td[contains(text(),'REVENUE')]")
		public static WebElement masterRevenue;
		
		@FindBy(xpath="//td[contains(text(),'EQUITIES')]")
		public static WebElement masterEquities;
	
		@FindBy(xpath="//td[contains(text(),'LIABILITIES')]")
		public static WebElement masterLiabilities;
		

		@FindBy(xpath="//section[@id='page_Content']//th[6]")
		public static WebElement masterGridHeader_ReorderLevel;

		@FindBy(xpath="//section[@id='page_Content']//th[7]")
		public static WebElement masterGridHeader_BinCapacity;

		@FindBy(xpath="//section[@id='page_Content']//th[8]")
		public static WebElement masterGridHeader_ISAttribute;

		@FindBy(xpath="//section[@id='page_Content']//th[9]")
		public static WebElement masterGridHeader_ProducyType;

		@FindBy(xpath="//th[contains(text(),'Valuation Method')]")
		public static WebElement masterGridHeader_ValuationMethod;

		@FindBy(xpath="//th[@class='OverFlowHidden text-center']//span[contains(text(),'No of decimals')]")
		public static WebElement masterGridHeader_NoOfDecimals;

		@FindBy(xpath="//section[@id='page_Content']//th[7]")
		public static WebElement masterGridHeader_RoundingType;
		
		@FindBy(xpath="//td[contains(text(),'Customer A')]")
		public static WebElement masterCustomerA;
		
		@FindBy(xpath="//td[contains(text(),'Customer B')]")
		public static WebElement masterCustomerB;
		
		@FindBy(xpath="//td[contains(text(),'Customer B')]")
		public static WebElement masterCustomerC;
		
		
		
		

		//Row Selecting


		@FindBy(xpath="//*[@id='lblnchecked']/input")
		public static WebElement masterGrid_SelectFirstRow;

		//Footer Fields
		@FindBy(xpath="//*[@id='id_PageButtonsMasterLanding']/div/div[1]/div/input")
		public static WebElement masterFirstBtn;

		@FindBy(xpath="//*[@id='id_PageButtonsMasterLanding']/div/div[2]/div/input")
		public static WebElement masterPreviousBtn;

		@FindBy(xpath="//*[@id='btn1']")
		public static WebElement masterOneBtn;

		@FindBy(xpath="//*[@id='id_PageButtonsMasterLanding']/div/div[4]/div/input")
		public static WebElement masterNextBtn;

		@FindBy(xpath="//*[@id='id_PageButtonsMasterLanding']/div/div[5]/div/input")
		public static WebElement masterEndBtn;
		
		
		//Accounts Creation Page 
		
		//Header tab Fields
		
		@FindBy(xpath="//*[@id='divBtnGroup0']/div/a")
		public static WebElement headerGeneralTab;
		
		@FindBy(xpath="//*[@id='divBtnGroup1']/div/a")
		public static WebElement headerAccountSettingTab;
		
		@FindBy(xpath="//*[@id='divBtnGroup2']/div/a")
		public static WebElement headerDetailsTab;
		
		@FindBy(xpath="//*[@id='divBtnGroup3']/div/a")
		public static WebElement headerPrintLayoutTab;
		
		@FindBy(xpath="//*[@id='divBtnGroup1']/div/a")
		public static WebElement headerUnitsTab;
		
		@FindBy(xpath="//*[@id='divBtnGroup2']/div/a")
		public static WebElement headerItemSettingsTab;
		
		@FindBy(xpath="//*[@id='divBtnGroup3']/div/a")
		public static WebElement headerClassificationTab;
		
		@FindBy(xpath="//*[@id='divBtnGroup4']/div/a")
		public static WebElement headerOtherDetailsTab;
		
		@FindBy(xpath="//*[@id='divBtnGroup5']/div/a")
		public static WebElement headerReplenishmentTab;
		
		@FindBy(xpath="//*[@id='divBtnGroup6']/div/a")
		public static WebElement headerOutletTab;
		
		
		
	    //Header Section Fields

		@FindBy(xpath="(//*[@id='btnMasterSaveClick'])[1]")
		public static WebElement saveBtn;
		
		@FindBy(xpath="//i[@class='icon-close-1 hiconright2']")
		public static WebElement closeBtn;
		
	    //Edit option Fields
		
		@FindBy(xpath="//*[@id='GenerateNewMasterModel']/div/div[1]/div[2]/div/ul/li[2]/button[6]")
		public static WebElement editcloseBtn;
		
		@FindBy(xpath="//i[@class='icon-font6 icon-paste-clipboard']")
		public static WebElement pasteBtn;
		
		@FindBy(xpath="//button[@id='btnMasterClone']")
		public static WebElement editcloneBtn;
		
		@FindBy(xpath="//*[@id='btnMasterCopy']")
		public static WebElement copyBtn;
		
		//New General Elements
		//Name
		@FindBy(xpath="//input[@id='sName']")
		public static WebElement nameTxt;
		
		
		//Code
		@FindBy(xpath="//input[@id='sCode']")
		public static WebElement codeTxt;
		
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

		
	//Setting Elements 
		
		//DebitCreditProposal
		@FindBy(xpath="//select[@id='iDebitCreditProposal']")
		public static WebElement debitCreditProposalDropdown;
		
		
		//DebitCreditRequired
		@FindBy(xpath="//select[@id='iDebitCreditRequired']")
		public static WebElement debitCreditRequiredDropdown;
		
		
		//ExchangeAdjustmentGainAC
		@FindBy(xpath="//input[@id='iExchangeAdjustmentGainAC']")
		public static WebElement exchangeAdjustmentGainACTxt;
		
		@FindBy(xpath="//*[@id='iExchangeAdjustmentGainAC_input_image']/span")
		public static WebElement exchangeAdjustmentGainAC_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iExchangeAdjustmentGainAC_input_settings']/span")
		public static WebElement exchangeAdjustmentGainAC_SettingBtn;
		
		
		//ExchangeAdjustmentLossAC
		@FindBy(xpath="//input[@id='iExchangeAdjustmentLossAC']")
		public static WebElement exchangeAdjustmentLossACTxt;
		
		@FindBy(xpath="//*[@id='iExchangeAdjustmentLossAC_input_image']/span")
		public static WebElement exchangeAdjustmentLossAC_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iExchangeAdjustmentLossAC_input_settings']/span")
		public static WebElement exchangeAdjustmentLossAC_SettingBtn;
		
		
		///PrimaryAccount
		@FindBy(xpath="//input[@id='iPrimaryAccount']")
		public static WebElement primaryAccountTxt;
		
		@FindBy(xpath="//*[@id='iPrimaryAccount_input_image']/span")
		public static WebElement primaryAccount_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iPrimaryAccount_input_settings']/span")
		public static WebElement primaryAccount_SettingBtn;
		
		
		////DefaultCurrency
		@FindBy(xpath="//input[@id='iDefaultCurrency']")
		public static WebElement defaultCurrencyTxt;
		
		@FindBy(xpath="//*[@id='iDefaultCurrency_input_image']/span")
		public static WebElement defaultCurrency_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iDefaultCurrency_input_settings']/span")
		public static WebElement defaultCurrency_SettingBtn;
		
		
		
		///ConsolidationMethod
		@FindBy(xpath="//select[@id='iConsolidationMethod']")
		public static WebElement consolidationMethodDropdown;
		
		
		///PaymentTerms
		@FindBy(xpath="//input[@id='iPaymentTerms']")
		public static WebElement paymentTermstxt;
		
		@FindBy(xpath="//*[@id='iPaymentTerms_input_image']/span")
		public static WebElement paymentTerms_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iPaymentTerms_input_settings']/span")
		public static WebElement paymentTerms_SettingBtn;
		
		
		//ReminderTerms
		@FindBy(xpath="//input[@id='iReminderTerms']")
		public static WebElement reminderTermsTxt;
		
		@FindBy(xpath="//*[@id='iReminderTerms_input_image']/span")
		public static WebElement reminderTerms_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iReminderTerms_input_settings']/span")
		public static WebElement reminderTerms_SettingBtn;
		
		
		///FinanceChargeTerms
		@FindBy(xpath="//*[@id='iFinanceChargeTerms']")
		public static WebElement financeChargeTermsTxt;
		
		@FindBy(xpath="//*[@id='iFinanceChargeTerms_input_image']/span")
		public static WebElement financeChargeTerms_Expansion;
		
		@FindBy(xpath="//*[@id='iFinanceChargeTerms_input_settings']/span")
		public static WebElement financeChargeTerms_SettingBtn;
	
		
		//Details Elements


		//Address
		@FindBy(xpath="//textarea[@id='sAddress']")
		public static WebElement addressTxt;
		
		
		//City
		@FindBy(xpath="//input[@id='iCity']")
		public static WebElement cityTxt;
		
		@FindBy(xpath="//*[@id='iCity_input_image']/span")
		public static WebElement city_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iCity_input_settings']/span")
		public static WebElement city_SettingBtn;
	
		//DeliveryAddress
		@FindBy(xpath="//textarea[@id='sDeliveryAddress']")
		public static WebElement deliveryAddressTxt;
		
		//Pin
		@FindBy(xpath="//input[@id='sPin']")
		public static WebElement pinTxt;
		
		
		//City2 DeliveryCity (City As Second TextBox)
		@FindBy(xpath="//input[@id='iDeliveryCity']")
		public static WebElement city2Text;
		
		@FindBy(xpath="//*[@id='iDeliveryCity_input_image']/span")
		public static WebElement city2_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iDeliveryCity_input_settings']/span")
		public static WebElement city2_SettingBtn;
		
		
		
		////Pin---->Delivery pin***()Second Pin TxtBox
		@FindBy(xpath="//input[@id='sDeliveryPin']")
		public static WebElement pin2Txt;
		
		
		//Send Email
		@FindBy(xpath="//input[@id='bSendEmailtocustomer']")
		public static WebElement sendEmailCheckBox;
		
		//AllowCustomerPortal
		@FindBy(xpath="//input[@id='bAllowCustomerPortal']")
		public static WebElement allowCustomerPortalCheckBox;
		
		//SendEmail
		@FindBy(xpath="//input[@id='sEMail']")
		public static WebElement sendEmailTxt;
		
		//Password
		@FindBy(xpath="//input[@id='sPassword']")
		public static WebElement passwordTxt;
		
		///TelphoneNumber
		@FindBy(xpath="//input[@id='sTelNo']")
		public static WebElement telphoneNumberTxt;
		
		//FaxNo
		@FindBy(xpath="//input[@id='sFaxNo']")
		public static WebElement faxNoTxt;
		
		
		//PortalEmail
		@FindBy(xpath="//input[@id='sPortalEmail']")
		public static WebElement portalEmailTxt;
		
		
		/////Print Layout Elements	
		//Voucher Type
		@FindBy(xpath="//label[contains(text(),'Voucher Type')]")
		public static WebElement gridVoucherTypeTxt;
		
		@FindBy(xpath="//*[@id='iVoucherType_input_image']/span")
		public static WebElement gridVoucher_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iVoucherType_input_settings']/span")
		public static WebElement gridVoucher_SettingBtn;

		
		///////Print layout
		@FindBy(xpath="//label[contains(text(),'Print Layout')]")
		public static WebElement gridPrintLayoutTxt;

		@FindBy(xpath="//*[@id='iPrintLayout_input_image']/span")
		public static WebElement gridPrintLayout_ExpansionBtn;

		@FindBy(xpath="//*[@id='iPrintLayout_input_settings']/span")
		public static WebElement gridPrintLayout_SettingBtn;
		
		
		
		
		//Item General Elements
		
		
		@FindBy(xpath="//select[@id='iProductType']")
		public static WebElement itemTypeDropdown;
		
		@FindBy(xpath="//input[@id='fReorderLevel']")
		public static WebElement reorderLevelTxt;
		
		@FindBy(xpath="//input[@id='iCapacity']")
		public static WebElement binCapacityTxt;
		
		@FindBy(xpath="//input[@id='iColNo']")
		public static WebElement binColNoTxt;
		
		@FindBy(xpath="//select[@id='iValuationMethod']")
		public static WebElement valuationMethodDropdown;
		
		@FindBy(xpath="//input[@id='iCategory']")
		public static WebElement categoryTxt;
		
		@FindBy(xpath="//*[@id=iCategory_input_image']/span")
		public static WebElement category_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iCategory_input_settings']/span")
		public static WebElement category_SettingBtn;
		
		@FindBy(xpath="//select[@id='iProductMake']")
		public static WebElement itemMakeDropdown;
		
		@FindBy(xpath="//input[@id='Printer']")
		public static WebElement printerTxt;
		
		@FindBy(xpath="//input[@id='iAlternateCategory']")
		public static WebElement alternateCategorytxt;
		
		@FindBy(xpath="//*[@id='iAlternateCategory_input_image']/span")
		public static WebElement alternateCategory_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iAlternateCategory_input_settings']/span")
		public static WebElement alternate_SettingBtn;
		
		/*@FindBy(xpath="//textarea[@id='sDescription']")
		public static WebElement descriptionTxt;*/
		
		@FindBy(xpath="//input[@id='bPerishableItem']")
		public static WebElement perishableItemCheckbox;
		
		@FindBy(xpath="//input[@id='iBin']")
		public static WebElement binTxt;
		
		@FindBy(xpath="//*[@id='iBin_input_image']/span")
		public static WebElement bin_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iBin_input_settings']/span")
		public static WebElement bin_SettingBtn;
		
		
		
		/////Item---New---Units
		
		@FindBy(xpath="//input[@id='iDefaultBaseUnit']")
		public static WebElement defaultBaseUnitTxt;
		
		@FindBy(xpath="//*[@id='iDefaultBaseUnit_input_image']/span")
		public static WebElement defaultBaseUnit_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iDefaultBaseUnit_input_settings']/span")
		public static WebElement defaultBaseUnit_SettingBtn;
		
		@FindBy(xpath="//input[@id='iDefaultSalesUnit']")
		public static WebElement defaultSalesUnit_Txt;
		
		@FindBy(xpath="//*[@id='iDefaultSalesUnit_input_image']/span")
		public static WebElement defaultSalesUnit_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iDefaultSalesUnit_input_settings']/span")
		public static WebElement defaultSalesUnit_SettingBtn;
		
		@FindBy(xpath="//input[@id='iDefaultPurchaseUnit']")
		public static WebElement defaultPurchaseUnitTxt;
		
		@FindBy(xpath="//*[@id='iDefaultPurchaseUnit_input_image']/span")
		public static WebElement defaultPurchaseUnit_ExpansioinBtn;
		
		@FindBy(xpath="//*[@id='iDefaultPurchaseUnit_input_settings']/span")
		public static WebElement defaultPurchaseUnit_SettingBtn;
		
		@FindBy(xpath="//input[@id='fLength']")
		public static WebElement lengthTxt;
		
		@FindBy(xpath="//input[@id='fWidth']")
		public static WebElement widthTxt;
		
		@FindBy(xpath="//input[@id='fHeight']")
		public static WebElement heightTxt;
		
		@FindBy(xpath="//input[@id='fCBM']")
		public static WebElement CBMTxt;
		
		
		/////Master--Item--Item--New--Settings
		
		@FindBy(xpath="//input[@id='fStandardCost']")
		public static WebElement standardCostTxt;
		
		@FindBy(xpath="//input[@id='fOverheadCost']")
		public static WebElement overheadCostTxt;
		
		@FindBy(xpath="//input[@id='fIndirectCost']")
		public static WebElement indirectCostTxt;
		
		@FindBy(xpath="//input[@id='fProfit']")
		public static WebElement profitTxt;
		
		@FindBy(xpath="//input[@id='iDontshowproductexpireddays']")
		public static WebElement dontshowproductexpiringindaysTxt;
		
		@FindBy(xpath="//input[@id='iTaxCode']")
		public static WebElement taxCodeTxt;
		
		@FindBy(xpath="//*[@id='iTaxCode_input_image']/span")
		public static WebElement taxCode_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iTaxCode_input_settings']/span")
		public static WebElement taxCode_SettingBtn;
		
		
		
		/////////Master--Item--Item--New--Classification
		
			
		@FindBy(xpath="//select[@id='iABCCodeValue']")
		public static WebElement ABCCodeValueDropdown;
		
		@FindBy(xpath="//select[@id='iABCCodeMargin']")
		public static WebElement ABCCodeMarginDropdown;
		
		@FindBy(xpath="//select[@id='iABCCodeRevenue']")
		public static WebElement ABCCodeRevenueDropdown;
		
		@FindBy(xpath="//select[@id='iABCCodeCarryingCost']")
		public static WebElement ABCCodeCarryingCostDropdown;
		
		@FindBy(xpath="//*[@id='sBarcode1']")
		public static WebElement barcode1Txt;
		
		@FindBy(xpath="//input[@id='sBarcode2']")
		public static WebElement barcode2Txt;
		
		@FindBy(xpath="//input[@id='sBarcode3']")
		public static WebElement barcode3Txt;
		
		
		///Units ---Grid---First Row
		@FindBy(xpath="//input[@id='iUnit1']")
		public static WebElement gridUnitTxt;
		
		@FindBy(xpath="//*[@id='iUnit_input_image']/span")
		public static WebElement gridUnit_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iUnit_input_settings']/span")
		public static WebElement gridUnit_SettingBtn;
		
		@FindBy(xpath="//tr[@class='tableRow']//input[@id='sBarcode1']")
		public static WebElement barcodeTxt;
		
		@FindBy(xpath="//input[@id='sBatch1']")
		public static WebElement batchTxt;
		
		
		
		
	    //Master--Item--Item--New--Other Details
		
		/*@FindBy(xpath="//input[@id='iCostOfIssueAccount']")
		public static WebElement costOfIssueAccountTxt;*/
		
		@FindBy(xpath="//*[@id='iCostOfIssueAccount_input_image']/span")
		public static WebElement costOfIssueAccount_ExpansiionBtn;
		
		@FindBy(xpath="//*[@id='iCostOfIssueAccount_input_settings']/span")
		public static WebElement costOfIssueAccount_SettingBtn;
		
		@FindBy(xpath="//input[@id='iStocksAccount']")
		public static WebElement stocksAccountTxt;
		
		@FindBy(xpath="//*[@id='iStocksAccount_input_image']/span")
		public static WebElement stocksAccount_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iStocksAccount_input_settings']/span")
		public static WebElement stocksAccount_SettingBtn;
		
		@FindBy(xpath="//input[@id='iSalesAccount']")
		public static WebElement salesAccount_Txt;
		
		@FindBy(xpath="//*[@id='iSalesAccount_input_image']/span")
		public static WebElement salesAccount_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iSalesAccount_input_settings']/span")
		public static WebElement salesAccount_SettingBtn;
		
		@FindBy(xpath="//input[@id='iWIPAccount']")
		public static WebElement WIPAccountTxt;
		
		@FindBy(xpath="//*[@id='iWIPAccount_input_image']/span")
		public static WebElement WIPAccoun_tExpansionBtn;
		
		@FindBy(xpath="//*[@id='iWIPAccount_input_settings']/span")
		public static WebElement WIPAccount_SettingBtn;
		
		@FindBy(xpath="//input[@id='iCostofShortageStockAC']")
		public static WebElement costofShortageStockACTxt;
		
		@FindBy(xpath="//*[@id='iCostofShortageStockAC_input_image']/span")
		public static WebElement costofShortageStockAC_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iCostofShortageStockAC_input_settings']/span")
		public static WebElement costofShortageStockAC_SettingBtn;
		
		@FindBy(xpath="//input[@id='iCostofExcessStockAC']")
		public static WebElement otherDetails_CostofExcessStockAC;
		
		@FindBy(xpath="//*[@id='iCostofExcessStockAC_input_image']/span")
		public static WebElement costofExcessStockAC_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='iCostofExcessStockAC_input_settings']/span")
		public static WebElement costofExcessStockAC_SettingBtn;
		
		
		
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
		
		@FindBy(xpath="//*[@id='ucCancel']/div")
		public static WebElement unitConversionCancelBtn;
		
		@FindBy(xpath="//*[@id='btnDeleteUnitConversion']/div")
		public static WebElement unitConversionDeleteBtn;
		
		
		
		
		
		//Grid---OtherDetails---1 row elements
		
		@FindBy(xpath="//input[@id='ioWarehouse']")
		public static WebElement otherDetailsGrid_WarehouseTxt;
		
		@FindBy(xpath="//*[@id='ioWarehouse_input_image']/span")
		public static WebElement otherDetailsGrid_Warehouse_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='ioWarehouse_input_settings']/span")
		public static WebElement otherDetailsGrid_Warehouse_SettingBtn;
		
		@FindBy(xpath="//input[@id='iSupplier1']")
		public static WebElement grid_Supplier1Txt;
		
		@FindBy(xpath="//input[@id='sSupCode1']")
		public static WebElement grid_SupCode1Txt;
		
		@FindBy(xpath="//input[@id='iStartDate1']")
		public static WebElement grids_StartDate1Txt;
		
		//Master----Item---New---Replenishment

		@FindBy(xpath="//select[@id='iDefaultReplenishment']")
		public static WebElement defaultReplenishmentDropDown;

		@FindBy(xpath="//select[@id='iManufacturePolicy']")
		public static WebElement manufacturePolicyDropDown;

		@FindBy(xpath="//input[@id='iPackingBOM']")
		public static WebElement packingBOMTxt;

		@FindBy(xpath="//*[@id='iPackingBOM_input_image']/span")
		public static WebElement packingBOMExpansionBtn;

		@FindBy(xpath="//*[@id='iPackingBOM_input_settings']/span")
		public static WebElement packingBOMSettingBtn;

		@FindBy(xpath="//input[@id='iBOM']")
		public static WebElement BOMTxt;

		@FindBy(xpath="//*[@id='iBOM_input_image']/span")
		public static WebElement BOMExpansionBtn;

		@FindBy(xpath="//*[@id='iBOM_input_settings']/span")
		public static WebElement BOMSettingBtn;

		@FindBy(xpath="//select[@id='iFlushing']")
		public static WebElement flushingDropdown;

		@FindBy(xpath="//input[@id='irEnddate']")
		public static WebElement grid_EndingDate;

		@FindBy(xpath="//input[@id='iReplenishqty1']")
		public static WebElement grid_Replenishqty1;

		@FindBy(xpath="//input[@id='irWarehouse']")
		public static WebElement replenishmentGrid_WarehouseTxt;

		@FindBy(xpath="//*[@id='irWarehouse_input_image']/span")
		public static WebElement replenishmentGrid_Warehouse_ExpansionBtn;

		@FindBy(xpath="//*[@id='irWarehouse_input_settings']/span")
		public static WebElement replenishmentGrid_Warehouse_SettingBtn;

		@FindBy(xpath="//input[@id='irStartdate1']")
		public static WebElement grid_StartDateTxt;

		
		
		//OutLet Elements
		
		
		@FindBy(xpath="//input[@id='txtsrch-term0']")
		public static WebElement searchBox;

		@FindBy(xpath="//*[@id='liSearchOn']/a")
		public static WebElement search_SearchOnBtn;

		@FindBy(xpath="//*[@id='liAdvanceSearch']/a")
		public static WebElement search_AdvanceSearch ;

		@FindBy(xpath="//input[@id='cmbUserTypeMaster']")
		public static WebElement cmbSearchTxt;

		@FindBy(xpath="//*[@id='cmbUserTypeMaster_input_image']/span")
		public static WebElement cmbSearchExpansionBtn;

		@FindBy(xpath="//*[@id='cmbUserTypeMaster_input_settings']/span")
		public static WebElement cmbSearchSettingBtn;

		@FindBy(xpath="//a[@id='chkResizeGrid0']")
		public static WebElement resizeGridBtn ;

		@FindBy(xpath="//input[@id='chkRetainSelection']")
		public static WebElement retainSelectionCheckbox;

		@FindBy(xpath="//label[@id='oncheckaccorderan0']")
		public static WebElement checkAccorderanBtn;
		
		@FindBy(xpath="//*[@id='iTreeId0']")
		public static WebElement treeIdDropdown;
		
		@FindBy(xpath="//*[@id='iTreeViewId0']")
		public static WebElement treeViewIdDropdown;
		
		
		
		
		
		
		
		//WareHouse
		
		
		@FindBy(xpath="//input[@id='iBins']")
		public static WebElement binsText;

		@FindBy(xpath="//*[@id='iBins_input_image']/span")
		public static WebElement bins_ExpansionBtn;

		@FindBy(xpath="//*[@id='iBins_input_settings']/span")
		public static WebElement bins_SettingBtn;

		@FindBy(xpath="//input[@id='iPhysInventoryCountingFrequency']")
		public static WebElement physInventoryCountingFrequencyTxt;

		@FindBy(xpath="//select[@id='iWarehouseType']")
		public static WebElement warehouseTypeDropdown;

		@FindBy(xpath="//input[@id='bDontMaintainBin']")
		public static WebElement dontMaintainBinCheckBox;


		// State

		@FindBy(xpath="//input[@id='iCountry']")
		public static WebElement countryTxt;

		@FindBy(xpath="//*[@id='iCountry_input_image']/span")
		public static WebElement country_ExpansionBtn;

		@FindBy(xpath="//*[@id='iCountry_input_settings']/span")
		public static WebElement country_SettingBtn;


		// City

		@FindBy(xpath="//input[@id='iState']")
		public static WebElement stateTxt;

		@FindBy(xpath="//*[@id='iState_input_image']/span")
		public static WebElement state_ExpansionBtn;

		@FindBy(xpath="//*[@id='iState_input_settings']/span")
		public static WebElement state_SettingBtn;



	   //Master------>Delete Field
		
		
		
		@FindBy(xpath="//*[@id='btnDelete']/i")
		public static WebElement deleteBtn;
		
		@FindBy(xpath="//div[@id='idGlobalError']")
		public static WebElement validationConfirmationMessage;

		@FindBy(xpath = "//*[@id='idGlobalError']/div/div[2]")
		public static WebElement errorMessage;
		
		@FindBy(xpath = "//*[@id='idGlobalError']/div/div[1]/button")
		public static WebElement errorMessageCloseBtn;
		
		/*@FindBy(xpath="//div[@id='idGlobalError']//td[3]")
		public static WebElement closeValidationConfirmationMessage;*/
		
		@FindBy(xpath="//*[@id='LandingGridBody']/tr[1]/td[11]")
		public static WebElement accountCreation;
		
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[11]")
		public static WebElement accountNewCreationName;
		
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[2]/td[11]")
		public static WebElement accountNewCreationNameOfLane;               
		
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[8]/div[1]/label[1]/input[1]")
		public static WebElement accountFirstCheckBoxToSelection;

		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/ol[1]/li[5]/span[1]")
		public static WebElement accountGroupTitleDisplay;
		
		
		//Master Customization
		
		/*@FindBy(xpath="//i[@class='icon-expand']")
		private static WebElement plusBtn;

		@FindBy(xpath="//i[@class='icon-collepse']")
		private static WebElement minusBtn;*/

		@FindBy(xpath="//span[contains(text(),'Master Fields')]")
		private static WebElement masterFieldsOption;
		
			@FindBy(xpath="//*[@id='masterFields_list']/li[1]/span")
			private static WebElement generalTab;
			
				@FindBy(xpath="//span[@class='font-5'][contains(text(),'Main')]")
				private static WebElement generalMainTab;
				
					@FindBy(xpath="//div[@id='mainTab_0']//tbody//tr[1]//td[4]")
					private static WebElement accountCustomizeName;

					@FindBy(xpath="//div[@id='mainTab_0']//tr[2]//td[4]")
					private static WebElement accountCustomizeCode;

					@FindBy(xpath="//div[@id='mainTab_0']//tr[3]//td[4]")
					private static WebElement accountCustomizeAccountType;

					@FindBy(xpath="//div[@id='mainTab_0']//tr[4]//td[4]")
					private static WebElement accountCustomizeCreditLimit;

					@FindBy(xpath="//div[@id='mainTab_0']//tr[5]//td[4]")
					private static WebElement accountCustomizeCreditDays;				
				
				@FindBy(xpath="//div[@id='tabId_1']//span[@class='font-5'][contains(text(),'Header Details')]")
				private static WebElement generalHeaderDetailsTab;
						
					@FindBy(xpath="//div[@id='tabId_1']//tbody[contains(@class,'ui-sortable')]//tr[1]//td[4]")
					private static WebElement accountCustomizeChequeDiscountLimit;

					@FindBy(xpath="//div[@id='HeaderDetails_0']//tr[2]//td[4]")
					private static WebElement accountCustomizeRateOfInterest;

					@FindBy(xpath="//*[@id='HeaderDetails_0']/div/div/table/tbody/tr[3]/td[4]")
					private static WebElement accountCustomizeBankAc;

					@FindBy(xpath="//*[@id='HeaderDetails_0']/div/div/table/tbody/tr[4]/td[4]")
					private static WebElement accountCustomizePDCDiscountedAccount;
					
					@FindBy(xpath="//*[@id='btnMove_Field_MasterCust']/div[1]/span")
					private static WebElement accountHeaderMoveFiledBtn;

					@FindBy(xpath="//div[contains(text(),'Preview')]")
					private static WebElement accountHeaderPreview;
					
					@FindBy(xpath="//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Save')]")
					private static WebElement accountSaveBtn;
								     
					@FindBy(xpath="//*[@id='btnAdd_Field_MasterCust']/div[1]/span")
					private static WebElement accountAddBtn;
											     
					@FindBy(xpath="//*[@id='btnCloseMasterCutomizationFromMasterScreen']/div[1]/span")
					private static WebElement accountCloseBtn;
	
			@FindBy(xpath="//span[@class='Flabel theme_icon-color'][contains(text(),'Settings')]")
			private static WebElement settingOption;

					@FindBy(xpath="//div[@id='tabId_69']//span[@class='font-5'][contains(text(),'Header Details')]")
					private static WebElement settingsHeaderDetailsTab;
					
					/*@FindBy(xpath="//div[@id='tabId_68']//span[@class='font-5'][contains(text(),'Header Details')]")
					private static WebElement settingsHeaderDetailsTab;*/
			
							  @FindBy(xpath="//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[1]/td[1]/i[1]")
							  private static WebElement accountSettingsHeaderEditFirstRowBtn;
							                  
							  @FindBy(xpath="//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[2]/td[1]/i[1]")
							  private static WebElement accountSettingsHeaderEditSecondRowBtn;
					
							  @FindBy(xpath="//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[3]/td[1]/i[1]")
							  private static WebElement accountSettingsHeaderEditThirdRowBtn;
					
							  @FindBy(xpath="//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[4]/td[1]/i[1]")
							  private static WebElement accountSettingsHeaderEditFourthRowBtn;
					
							  @FindBy(xpath="//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[5]/td[1]/i[1]")
							  private static WebElement accountSettingsHeaderEditFifthRowBtn;
					
							  @FindBy(xpath="//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[6]/td[1]/i[1]")
							  private static WebElement accountSettingsHeaderEditSixthRowBtn;
					
							  @FindBy(xpath="//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[7]/td[1]/i[1]")
							  private static WebElement accountSettingsHeaderEditSeventhRowBtn;
					
							  @FindBy(xpath="//*[@id='HeaderDetails_1']/div/div/table/tbody/tr[8]/td[1]/i[1]")
							  private static WebElement accountSettingsHeaderEditEighthRowBtn;
							  
							  @FindBy(xpath="//div[contains(text(),'Move Field')]")
							  private static WebElement accountSettingsHeaderMoveFiledBtn;

							  @FindBy(xpath="//div[contains(text(),'Preview')]")
							  private static WebElement accountSettingsHeaderPreview;
							  						
							  @FindBy(xpath="//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Save')]")
							  private static WebElement accountSettingsSaveBtn;
							  									    
							  @FindBy(xpath="//div[@id='btnAdd_Field_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Add')]")
							  private static WebElement accountSettingsAddBtn;						

							  @FindBy(xpath="//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
							  private static WebElement accountSettingsCloseBtn;
							  
					
					/*@FindBy(xpath="//div[@id='tabId_68']//span[@class='font-5'][contains(text(),'Body Details')]")
					private static WebElement settingsBodyDetailsTab;*/
					
					@FindBy(xpath="//div[@id='tabId_69']//span[@class='font-5'][contains(text(),'Body Details')]")
					private static WebElement settingsBodyDetailsTab;
					
					                @FindBy(xpath="//div[@id='BodyDetails_1']//div[@class='scrollable tbl-emptyfields-mastercust']")
					                private static WebElement settingsBodyDetailsEmptyTab;
			

			@FindBy(xpath="//span[@class='Flabel theme_icon-color'][contains(text(),'Details')]")
			private static WebElement accountMasterFieldDetailsOption;
			
				/*@FindBy(xpath="//div[@id='tabId_69']//span[@class='font-5'][contains(text(),'Header Details')]")
		        private static WebElement detailsHeaderDetailsTab;
		        
		         @FindBy(xpath="//div[@id='tabId_69']//tbody[@class='ui-sortable']//tr[1]//td[4]")
								  private static WebElement detailsHeaderAddress;
								  
								  @FindBy(xpath="//div[@id='tabId_69']//tr[2]//td[4]")
								  private static WebElement detailsHeaderTelNo;
								  									    
								  @FindBy(xpath="//div[@id='tabId_69']//tr[3]//td[4]")
								  private static WebElement detailsHeaderFaxNo;		
								                  
								  @FindBy(xpath="//div[@id='tabId_69']//tr[4]//td[4]")
								  private static WebElement detailsHeaderCity;
						
								  @FindBy(xpath="//div[@id='tabId_69']//tr[5]//td[4]")
								  private static WebElement detailsHeaderPin;
						
								  @FindBy(xpath="//div[@id='tabId_69']//tr[6]//td[4]")
								  private static WebElement detailsHeaderDeliveryAddress;
						
								  @FindBy(xpath="//div[@id='tabId_69']//tr[7]//td[4]")
								  private static WebElement detailsHeaderCityOne;
						
								  @FindBy(xpath="//div[@id='tabId_69']//tr[8]//td[4]")
								  private static WebElement detailsHeaderPinOne;
						
								  @FindBy(xpath="//div[@id='tabId_69']//tr[9]//td[4]")
								  private static WebElement detailsHeaderSendEmailToCustomer;
						
								  @FindBy(xpath="//div[@id='tabId_69']//tr[10]//td[4]")
								  private static WebElement detailsHeaderAllowCustomerPortal;
								  
								  @FindBy(xpath="//div[@id='tabId_69']//tr[11]//td[4]")
								  private static WebElement detailsHeaderEmail;

								  @FindBy(xpath="//div[@id='tabId_69']//tr[12]//td[4]")
								  private static WebElement detailsHeaderPassword;

								  @FindBy(xpath="//div[@id='tabId_69']//tr[13]//td[4]")
								  private static WebElement detailsHeaderBankAccountName;

								  @FindBy(xpath="//div[@id='tabId_69']//tr[14]//td[4]")
								  private static WebElement detailsHeaderBankAccountNumber;

								  @FindBy(xpath="//div[@id='tabId_69']//tr[15]//td[4]")
								  private static WebElement detailsHeaderIFSCCode;

								  @FindBy(xpath="//div[@id='tabId_69']//tr[16]//td[4]")
								  private static WebElement detailsHeaderPaymentType;

								  @FindBy(xpath="//div[@id='tabId_69']//tr[17]//td[4]")
								  private static WebElement detailsHeaderFinanceEmail;							  
						
	                              @FindBy(xpath="//div[@id='tabId_69']//tr[18]//td[4]")
								  private static WebElement detailsHeaderPortalEmail;	

	                              @FindBy(xpath="//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Save')]")
								  private static WebElement detailsHeaderBankSaveBtn;

								  @FindBy(xpath="//div[contains(text(),'Move Field')]")
								  private static WebElement detailsHeaderMoveFieldBtn;

								  @FindBy(xpath="//div[contains(text(),'Add')]")
								  private static WebElement detailsHeaderAddBtn;

								  @FindBy(xpath="//div[contains(text(),'Preview')]")
								  private static WebElement detailsHeaderPreviewBtn;

								  @FindBy(xpath="//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
								  private static WebElement detailsHeaderCloseBtn;	*/
		        
		        
				
				@FindBy(xpath="//div[@id='tabId_70']//span[@class='font-5'][contains(text(),'Header Details')]")
		        private static WebElement detailsHeaderDetailsTab;
							
								  @FindBy(xpath="//div[@id='tabId_70']//tbody[@class='ui-sortable']//tr[1]//td[4]")
								  private static WebElement detailsHeaderAddress;
								  
								  @FindBy(xpath="//div[@id='tabId_70']//tr[2]//td[4]")
								  private static WebElement detailsHeaderTelNo;
								  									    
								  @FindBy(xpath="//div[@id='tabId_70']//tr[3]//td[4]")
								  private static WebElement detailsHeaderFaxNo;		
								                  
								  @FindBy(xpath="//div[@id='tabId_70']//tr[4]//td[4]")
								  private static WebElement detailsHeaderCity;
						
								  @FindBy(xpath="//div[@id='tabId_70']//tr[5]//td[4]")
								  private static WebElement detailsHeaderPin;
						
								  @FindBy(xpath="//div[@id='tabId_70']//tr[6]//td[4]")
								  private static WebElement detailsHeaderDeliveryAddress;
						
								  @FindBy(xpath="//div[@id='tabId_70']//tr[7]//td[4]")
								  private static WebElement detailsHeaderCityOne;
						
								  @FindBy(xpath="//div[@id='tabId_70']//tr[8]//td[4]")
								  private static WebElement detailsHeaderPinOne;
						
								  @FindBy(xpath="//div[@id='tabId_70']//tr[9]//td[4]")
								  private static WebElement detailsHeaderSendEmailToCustomer;
						
								  @FindBy(xpath="//div[@id='tabId_70']//tr[10]//td[4]")
								  private static WebElement detailsHeaderAllowCustomerPortal;
								  
								  @FindBy(xpath="//div[@id='tabId_70']//tr[11]//td[4]")
								  private static WebElement detailsHeaderEmail;

								  @FindBy(xpath="//div[@id='tabId_70']//tr[12]//td[4]")
								  private static WebElement detailsHeaderPassword;

								  @FindBy(xpath="//div[@id='tabId_70']//tr[13]//td[4]")
								  private static WebElement detailsHeaderBankAccountName;

								  @FindBy(xpath="//div[@id='tabId_70']//tr[14]//td[4]")
								  private static WebElement detailsHeaderBankAccountNumber;

								  @FindBy(xpath="//div[@id='tabId_70']//tr[15]//td[4]")
								  private static WebElement detailsHeaderIFSCCode;

								  @FindBy(xpath="//div[@id='tabId_70']//tr[16]//td[4]")
								  private static WebElement detailsHeaderPaymentType;

								  @FindBy(xpath="//div[@id='tabId_70']//tr[17]//td[4]")
								  private static WebElement detailsHeaderFinanceEmail;							  
						
	                              @FindBy(xpath="//div[@id='tabId_70']//tr[18]//td[4]")
								  private static WebElement detailsHeaderPortalEmail;	

	                              @FindBy(xpath="//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Save')]")
								  private static WebElement detailsHeaderBankSaveBtn;

								  @FindBy(xpath="//div[contains(text(),'Move Field')]")
								  private static WebElement detailsHeaderMoveFieldBtn;

								  @FindBy(xpath="//div[contains(text(),'Add')]")
								  private static WebElement detailsHeaderAddBtn;

								  @FindBy(xpath="//div[contains(text(),'Preview')]")
								  private static WebElement detailsHeaderPreviewBtn;

								  @FindBy(xpath="//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
								  private static WebElement detailsHeaderCloseBtn;						  

						/*@FindBy(xpath="//div[@id='tabId_69']//span[@class='font-5'][contains(text(),'Body Details')]")
						private static WebElement detailsBodyDetailsTab;*/
						
						@FindBy(xpath="//div[@id='tabId_70']//span[@class='font-5'][contains(text(),'Body Details')]")
						private static WebElement detailsBodyDetailsTab;
						
						        @FindBy(xpath="//div[@id='BodyDetails_2']//div[@class='scrollable tbl-emptyfields-mastercust']")
						        private static WebElement detailsBodyDetailsEmptyTab;
			
			@FindBy(xpath="//*[@id='masterFields_list']/li[4]/span")
			private static WebElement accountMasterFieldPrintLayoutOption;
		
					 /* @FindBy(xpath="//div[@id='tabId_70']//span[@class='font-5'][contains(text(),'Header Details')]")
					  private static WebElement printlayoutHeaderDetailsTab;*/
					                  
					  @FindBy(xpath="//div[@id='tabId_71']//span[@class='font-5'][contains(text(),'Header Details')]")
					  private static WebElement printlayoutHeaderDetailsTab;
					  
					    @FindBy(xpath="//div[@id='HeaderDetails_3']//div[@class='scrollable tbl-emptyfields-mastercust']")
					    private static WebElement printLayoutHeaderDetailsEmpty;
			
					 /* @FindBy(xpath="//div[@id='tabId_70']//span[@class='font-5'][contains(text(),'Body Details')]")
					  private static WebElement printLayouBodyDetailsTab;*/
			
					  @FindBy(xpath="//div[@id='tabId_71']//span[@class='font-5'][contains(text(),'Body Details')]")
					  private static WebElement printLayouBodyDetailsTab;
					  
					      @FindBy(xpath="//div[@id='BodyDetails_3']/div/div/table/tbody/tr[1]/td[4]")
					      private static WebElement printlayoutBodyVoucherType;
			
			    	      @FindBy(xpath="//div[@id='tabId_70']//tr[2]//td[4]]")
					      private static WebElement printLayoutBodyPrintlayout;
			
					  @FindBy(xpath="//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Save')]")
					  private static WebElement printLayoutSaveBtn;
			
					  @FindBy(xpath="//div[@id='btnAdd_Field_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Add')]")
					  private static WebElement printLayoutAddBtn;
			
					  @FindBy(xpath="//div[contains(text(),'Preview')]")
					  private static WebElement printLayoutPreviewBtn;
					  
					  @FindBy(xpath="//div[@class='clsFieldButtons_MCust_div']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
					  private static WebElement printLayoutCloseBtn;
	
	    @FindBy(xpath="//span[contains(text(),'Unique Constraints')]")
		private static WebElement accountMasterFieldsUniqueConstraintsOption; 
		
			 @FindBy(xpath="//div[@id='tabContent_UConstraints_MasterCust']//div[@id='btnAddOrEdit_UniqueConstraints']")
			 private static WebElement UniqueConstraintsAddBtn;
					     
			 @FindBy(xpath="//div[@id='tabContent_UConstraints_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Delete')]")
			 private static WebElement UniqueConstraintsDeleteBtn;
					     
			 @FindBy(xpath="//div[@id='tabContent_UConstraints_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
			 private static WebElement UniqueConstraintsCloseBtn;
			 
				 @FindBy(xpath="//input[@id='txtConstraintName']")
				 private static WebElement ConstraintNameTxt;
			
			     @FindBy(xpath="//select[@id='ddlConstraintTabList']")
			     private static WebElement UniqueConstraints_TabsDropdown;

			     @FindBy(xpath="//*[@id='tab_Main_UcFields']/i")
			     private static WebElement UniqueConstraints_MainExpansionBtn;
			   
				     @FindBy(xpath="//*[@id='chkList_MainFields_1']/div[1]")
				     private static WebElement accountUniqueConstraints_Main_NameChekbox;
				     
				     @FindBy(xpath="//div[@id='chkList_MainFields_1']//div[2]")
				     private static WebElement accountUniqueConstraints_Main_CodeChekbox;
				     
				     @FindBy(xpath="//div[@id='chkList_MainFields_1']//div[3]")
				     private static WebElement accountUniqueConstraints_Main_AccountTypeCheckbox;
				     
				     @FindBy(xpath="//div[@id='chkList_MainFields_1']//div[4]")
				     private static WebElement accountUniqueConstraints_Main_CreditLimitCheckBox;
				     
				     @FindBy(xpath="//div[@id='chkList_MainFields_1']//div[5]")
				     private static WebElement accountUniqueConstraints_Main_CreditDays;
				     
				 @FindBy(xpath="//div[@id='tab_Header_UcFields']//i[@class='pull-right icon-expand theme_icon-color']")
				 private static WebElement UniqueConstraints_HeaderExpansionBtn;

				     @FindBy(xpath="//*[@id='chkList_HeaderFields_1']/div[1]")
				     private static WebElement UniqueConstraints_Header_ChequeDiscountLimitChekbox;
				     
				     @FindBy(xpath="//*[@id='chkList_HeaderFields_1']/div[2]")
				     private static WebElement UniqueConstraints_Header_RateOfInterestChekbox;
				     
				     @FindBy(xpath="//*[@id='chkList_HeaderFields_1']/div[3]")
				     private static WebElement UniqueConstraints_Header_BankAcChekbox;
				     
				     @FindBy(xpath="//*[@id='chkList_HeaderFields_1']/div[4]")
				     private static WebElement UniqueConstraints_Header_PDCDiscountedChekbox;
				  
				 @FindBy(xpath="//*[@id='tab_Body_UcFields']/i")
			     private static WebElement UniqueConstraints_BodyExpansionBtn;
				 
		@FindBy(xpath="//span[contains(text(),'Rules')]")
		private static WebElement RulesOption;
		
			 @FindBy(xpath="//div[@id='newRuleTab_div_MasterRules']//span[@class='icon-save icon-font7']")
			 private static WebElement accountRules_SaveBtn;
			               
			 @FindBy(xpath="//div[@class='pull-right']//span[@class='icon-new icon-font7']")
			 private static WebElement accountRules_NewBtn;
			
			 @FindBy(xpath="//div[@id='newRuleTab_div_MasterRules']//span[@class='icon-delete icon-font7']")
			 private static WebElement accountRules_DeleteBtn;
			 
			 @FindBy(xpath="//div[@id='newRuleTab_div_MasterRules']//div[@class='pull-right']//span[@class='icon-close icon-font7']")
			 private static WebElement accountRules_CloseBtn;   
			 
			 @FindBy(xpath="//*[@id='txtMasterRuleName_MasterRules']")
			 private static WebElement Rules_RuleNameText ;
				     

			 @FindBy(xpath="//*[@id='chkCreatingGroup_Rule_MasterRules']")
			 private static WebElement Rules_CreatingGroupChekbox ;

			 @FindBy(xpath="//*[@id='chkNewRecord_Rule_MasterRules']")
			 private static WebElement Rules_NewRecordChekbox ;

			 @FindBy(xpath="//*[@id='chkEdit_Rule_MasterRules']")
			 private static WebElement Rules_EditChekbox ;

			 @FindBy(xpath="//*[@id='chkLoad_Rule_MasterRules']")
			 private static WebElement Rules_LoadChekbox ;

			 @FindBy(xpath="//*[@id='chkBeforeSave_Rule_MasterRules']")
			 private static WebElement Rules_BeforeSaveChekbox ;

			 @FindBy(xpath="//*[@id='chkOnLeave_Rule_MasterRules']")
			 private static WebElement Rules_OnLeaveChekbox ;

			 @FindBy(xpath="//*[@id='chkBeforeDelete_Rule_MasterRules']")
			 private static WebElement Rules_BeforeDeleteChekbox ;

			 @FindBy(xpath="//*[@id='chkOnEnter_Rule_MasterRules']")
			 private static WebElement Rules_OnEnterChekbox ;
			 
			 @FindBy(xpath="//label[contains(text(),'Active')]")
			 private static WebElement Rules_ActiveChekbox ;

			 @FindBy(xpath="//*[@id='Conditions-Context_Menu']/li[3]")
			 private static WebElement Rules_NoConditionChekbox;	  

		@FindBy(xpath="//*[@id='allMasterCustTabs']/li[4]/span")
		private static WebElement rules_ExternalModules;

			@FindBy(xpath="//select[@id='ddlOnEvent_MasterExternalModule']")
			private static WebElement rules_ExternalModulesOnEventDropdown;	
			
			@FindBy(xpath="//input[@id='txtbuttonCaption_MasterExternalModule']")
			private static WebElement rules_ExternalModulesButtonCaption;

			@FindBy(xpath="//select[@id='ddlModuleType_MasterExternalModule']")
			private static WebElement rules_ExternalModules_ModuleTypeBtn;

			@FindBy(xpath="//select[@id='ddlAllDllFiles_MasterExternalModule']")
			private static WebElement rules_ExternalModules_ModuleNameDropdown;

			@FindBy(xpath="//select[@id='ddlClassNames_MasterExternalModule']")
			private static WebElement rules_ExternalModulesClassNameDropdown;
		
			@FindBy(xpath="//select[@id='ddlClassMethods_MasterExternalModule']")
			private static WebElement rules_ExternalModulesFuctionalNameDropdown;
	
			@FindBy(xpath="//div[contains(text(),'Update')]")
			private static WebElement rules_ExternalModulesUpdateBtn;

			@FindBy(xpath="//div[@id='ExternalModulesDiv_MasterExternalModule']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Save')]")
			private static WebElement rules_ExternalModulesSaveBtn;

			@FindBy(xpath="//div[contains(text(),'Edit')]")
			private static WebElement rules_ExternalModulesEditBtn;

			@FindBy(xpath="//div[@id='ExternalModulesDiv_MasterExternalModule']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Delete')]")
			private static WebElement rules_ExternalModulesDeleteBtn;

		    @FindBy(xpath="//div[@id='ExternalModulesDiv_MasterExternalModule']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
			private static WebElement rules_ExternalModulesCloseBtn;

	    @FindBy(xpath="//span[contains(text(),'Info Panel Customization')]")
		private static WebElement itemInfoPanelCustomizationTab;

		    @FindBy(xpath="//body[@class='custom_scrollbar']/section[@id='mainDiv']/div[@id='id_focus8_wrapper_default']/div[@class='content-wrapper']/section[@id='page_Content']/div[@class='container']/div[@id='mainLandingDiv']/div[@id='masterFirstDiv']/div[@class='tab-content']/div[@id='divLanding']/div[@id='divSecondsub']/div/div[@class='col-xs-12 col-sm-8 col-md-9 col-lg-9 padding-left-right5 mainDivCustomWidth']/div[@id='tabContent_IPCust_MasterCust']/div[@class='col-xs-12 col-sm-12 padding0']/div[@class='pull-right']/div[1]")
		    private static WebElement iteminfoPanelCust_AddBtn;

		    @FindBy(xpath="//div[@id='tabContent_IPCust_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Delete')]")
		    private static WebElement infoPanelCust_DeleteBtn;

		    @FindBy(xpath="//div[@id='tabContent_IPCust_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
		    private static WebElement infoPanelCust_CloseBtn;

		    @FindBy(xpath="//strong[contains(text(),'Credit History')]")
		    private static WebElement infoPanelCustCreditHistory;

		    @FindBy(xpath="//div[@id='divPanel_2']//strong[contains(text(),'Ageing Analysis')]")
		    private static WebElement infoPanelCustAgeingAnalysis;

		    @FindBy(xpath="//strong[contains(text(),'Ageing Analysis Base')]")
		    private static WebElement infoPanelCustAgeingAnalysisBase;

		    @FindBy(xpath="//strong[contains(text(),'Ageing Analysis Local')]")
		    private static WebElement infoPanelCustAgeingAnalysisLocal;

		    @FindBy(xpath="//strong[contains(text(),'Authorization')]")
		    private static WebElement infoPanelAuthorization;
		   
	    @FindBy(xpath="//span[@class='Flabel'][contains(text(),'Reports')]")
	    private static WebElement accountReportsTab;
			
			@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Save')]")
			private static WebElement  accountReports_SaveBtn;
			
			@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
			private static WebElement  accountReports_CloseBtn;
			
					@FindBy(xpath="//*[@id='AllReportsDiv_MasterCust']/div[1]/label/input")
					private static WebElement  reports_AdvancedbudgetReportChkBox;
					
					@FindBy(xpath="//div[@id='AllReportsDiv_MasterCust']//div[1]//label[1]")
					private static WebElement  reports_MasterInfoChkBox;
					
					
					@FindBy(xpath="//div[@class='col-xs-12 col-sm-8 col-md-8 col-lg-9 padding0']//div[2]//label[1]")
					private static WebElement  reports_ABCAnalysisAmountChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[3]//label[1]")
					private static WebElement  reports_BankBookChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[4]//label[1]")
					private static WebElement  reports_BankreconciliationstatementChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[5]//label[1]")
					private static WebElement  reports_CashBookChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[6]//label[1]")
					private static WebElement  reports_ComparativeAnalysisChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[7]//label[1]")
					private static WebElement  reports_CustomerAgeingDetailsAnalysisChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[8]//label[1]")
					private static WebElement  reports_CustomerageingdetailbyduedateChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[9]//label[1]")
					private static WebElement  reports_CustomerageingsummaryanalysisChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[10]//label[1]")
					private static WebElement  reports_CustomerbillwisesummaryChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[11]//label[1]")
					private static WebElement  reports_CustomerduedateanalysisChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[12]//label[1]")
					private static WebElement  reports_CustomerlistingofoutstandingbillsChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[13]//label[1]")
					private static WebElement  reports_CustomeroverdueanalysisChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[14]//label[1]")
					private static WebElement  reports_CustomeroverduesummaryChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[15]//label[1]")
					private static WebElement  reports_CustomerstatementsChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[16]//label[1]")
					private static WebElement  reports_CustomersummaryageingbyduedateChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[17]//label[1]")
					private static WebElement  reports_IncomeExpenseTrendChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[18]//label[1]")
					private static WebElement  reports_InterestcalculationChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[19]//label[1]")
					private static WebElement  reports_LedgerChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[20]//label[1]")
					private static WebElement  reports_LedgerdetailChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[21]//label[1]")
					private static WebElement  reports_MonthlySalesChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[22]//label[1]")
					private static WebElement  reports_PeakandlowbalanceamountChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[23]//label[1]")
					private static WebElement  reports_PettycashbookChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[24]//label[1]")
					private static WebElement  reports_PurchaseregisterChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[25]//label[1]")
					private static WebElement  reports_PurchasereturnregisterChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[26]//label[1]")
					private static WebElement  reports_PurchasesgroupedbydepartmentChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[27]//label[1]")
					private static WebElement  reports_PurchasesgroupedbyItemChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[28]//label[1]")
					private static WebElement  reports_PurchasesgroupedbyvendorChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[29]//label[1]")
					private static WebElement  reports_SalesdaybookChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[30]//label[1]")
					private static WebElement  reports_SalesgroupedbycustomerChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[31]//label[1]")
					private static WebElement  reports_salesgroupedbydepartmentrChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[32]//label[1]")
					private static WebElement  reports_SalesgroupedbyItemChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[33]//label[1]")
					private static WebElement  reports_SalesreturnregisterChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[34]//label[1]")
					private static WebElement  reports_SchedulesChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[35]//label[1]")
					private static WebElement  reports_SubledgerChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[36]//label[1]")
					private static WebElement  reports_SummarypurchaseregisterChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[37]//label[1]")
					private static WebElement  reports_SummarysalesregisterChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[38]//label[1]")
					private static WebElement  reports_TopCustomersChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[39]//label[1]")
					private static WebElement  reports_TradingaccountrChkBox;
					
					@FindBy(xpath="//div[40]//label[1]")
					private static WebElement  reports_TransactionstypeanalysisChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[41]//label[1]")
					private static WebElement  reports_VendorageingdetailanalysisChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[42]//label[1]")
					private static WebElement  reports_VendorageingdetailbyduedateChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[43]//label[1]")
					private static WebElement  reports_VendorageingsummaryanalysisChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[44]//label[1]")
					private static WebElement  reports_VendorbillwisesummaryChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[45]//label[1]")
					private static WebElement  reports_VendorduedateanalysisChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[46]//label[1]")
					private static WebElement  reports_VendorlistingofoutstandingbillsChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[47]//label[1]")
					private static WebElement  reports_VendoroverdueanalysisChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[48]//label[1]")
					private static WebElement  reports_VendoroverduesummaryChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[49]//label[1]")
					private static WebElement  reports_VendorstatementsChkBox;
					
					@FindBy(xpath="//div[@id='tabContent_Reports_MasterCust']//div[50]//label[1]")
					private static WebElement  reports_VendorsummaryageingbyduedateChkBox;
				
                 //Create Tabs
					
					@FindBy(xpath="//button[@id='btnCreateTab_MCust']")
					private static WebElement createTab;  
						
						@FindBy(xpath="//a[contains(text(),'Tree Tab')]")
						private static WebElement createTabTreeTab; 

							@FindBy(xpath="//input[@id='txtNewTreeTabCaption']")
							private static WebElement  treeTabCaptionTxt;
							
							@FindBy(xpath="//input[@id='txtNewTreeTabName']")
							private static WebElement  treeTabTabNameTxt;
							
							@FindBy(xpath="//input[@id='fOption_tab_masterName']")
							private static WebElement  treeTabNewMastertxt;
							
							@FindBy(xpath="//*[@id='fOption_tab_masterName_input_image']/span")
							private static WebElement  treeTabNewMasterdropdown;
							
							@FindBy(xpath="//*[@id='fOption_tab_masterName_input_settings']/span")
							private static WebElement  treeTab_NewMasterSettingsBtn;
	
							@FindBy(xpath="//input[@id='txtNewTreeTabFieldCaption']")
							private static WebElement  treeTab_FieldCaptiontxt;
							
							@FindBy(xpath="//input[@id='txtNewTreeTabFieldName']")
							private static WebElement  treeTab_FieldNameTxt;
							
							@FindBy(xpath="//input[@id='bIsTreeHiddenTab']")
							private static WebElement  treeTab_HiddenChkBox;
						
									@FindBy(xpath="//div[@id='fOption_tab_masterName_customize_popup_container']")
									private static WebElement  treeTab_NewMaster_ContainerTxt;
									
									@FindBy(xpath="//div[@id='fOption_tab_masterName_customize_popup_footer']//input[1]")
									private static WebElement  treeTab_NewMaster_StanadrdfieldsBtn;
									
											@FindBy(xpath="//select[@id='fOption_tab_masterName_customize_popup_standardfields_list']")
											private static WebElement treeTab_NewMaster_Stanadrdfields_Fielddropdown;
											
											@FindBy(xpath="//input[@id='fOption_tab_masterName_customize_popup_standardfields_header']")
											private static WebElement  treeTab_NewMaster_Stanadrdfields_headerTxt;
											
											@FindBy(xpath="//select[@id='fOption_tab_masterName_customize_popup_standardfields_alignment']")
											private static WebElement  treeTab_NewMaster_Stanadrdfields_Allignmentdropdown;
											
											@FindBy(xpath="//input[@id='fOption_tab_masterName_customize_popup_standardfields_width']")
											private static WebElement  treeTab_NewMaster_Stanadrdfields_WidthTxt;
											
											@FindBy(xpath="//div[@id='divSecondsub']//input[3]")
											private static WebElement  treeTab_NewMaster_Stanadrdfields_OkBtn;
											
											@FindBy(xpath="//div[@id='divSecondsub']//input[4]")
											private static WebElement  treeTab_NewMaster_Stanadrdfields_CancelBtn;
									
									@FindBy(xpath="//div[@id='fOption_tab_masterName_search_container']//input[2]")
									private static WebElement  treeTab_NewMaster_deleteColumnBtn;
									
									@FindBy(xpath="//div[@id='divSecondsub']//input[3]")
									private static WebElement  treeTab_NewMaster_OkBtn;
									
									@FindBy(xpath="//div[@id='divSecondsub']//input[4]")
									private static WebElement  treeTab_NewMaster_CancelBtn;
						
							@FindBy(xpath="//*[@id='MCustomize_CreateTreeTab_Modal']/div/div/div[3]/button[1]")
							private static WebElement  treeTab_OkBtn;

							@FindBy(xpath="//div[@id='MCustomize_EditTab_Modal']//button[@class='Fbutton font-5'][contains(text(),'OK')]")
							private static WebElement  treeTab_UpdateOkBtn;
							
							@FindBy(xpath="//*[@id='MCustomize_CreateTreeTab_Modal']/div/div/div[3]/button[2]")
							private static WebElement  treeTab_CancelBtn;
							
						@FindBy(xpath="//a[contains(text(),'Document Tab')]")
						private static WebElement createTab_DocumentTab;  

						            @FindBy(xpath="//input[@id='txtNewDocTabCaption']")
						            private static WebElement documentTab_TabCaption;  
						
									@FindBy(xpath="//input[@id='txtNewDocTabName']")
									private static WebElement documentTab_TabName;  
									
									@FindBy(xpath="//input[@id='bIsDocHiddenTab']")
									private static WebElement documentTab_HiddenChkbox;  
									
									@FindBy(xpath="//div[@id='MCustomize_CreateDocTab_Modal']//button[@class='Fbutton font-5'][contains(text(),'OK')]")
									private static WebElement documentTab_OkBtn;  
									
									@FindBy(xpath="//div[@id='MCustomize_CreateDocTab_Modal']//button[@class='Fbutton font-5'][contains(text(),'Cancel')]")
									private static WebElement documentTab_CancelBtn; 
									
									@FindBy(xpath="//label[@id='lblNameError_DocTab']")
									private static WebElement customizeMasterCreateTabDocumentTabMessage; 
								
								        @FindBy(xpath="//span[contains(text(),'Document Tab')]")
								        private static WebElement customizeMasterCreateDocumentTabNewlyCreate; 
									
										@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[6]/ul[1]/li[1]/a[1]/span[1]")
										private static WebElement customizeMasterCreateDocumentTabNewlyHeaderDetails; 	
													
										@FindBy(xpath="//div[@id='HeaderDetails_5']//div[@class='scrollable tbl-emptyfields-mastercust']")
										private static WebElement customizeMasterCreateDocumentTabNewlyBlankInHeaderDetails;
										
										@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[6]/ul[1]/li[2]/a[1]/span[1]")
										private static WebElement customizeMasterCreateDocumentTabNewlyBodyDetails; 			
										
										@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[6]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[4]")
										private static WebElement customizeMasterCreateDocumentTabFieldName; 
										
										@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[6]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[4]")
										private static WebElement customizeMasterCreateDocumentTabDocument; 
										
										@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[6]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[3]/td[4]")
										private static WebElement customizeMasterCreateDocumentTabCreatedDate; 
										
										@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[6]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[4]/td[4]")
										private static WebElement customizeMasterCreateDocumentTabLastModified; 
										
										@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[6]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[5]/td[4]")
										private static WebElement customizeMasterCreateDocumentFileSize; 
										
						@FindBy(xpath="//a[contains(text(),'Create Tab')]")
						private static WebElement createTab_CreateTab;  

							@FindBy(xpath="//div[@id='MCustomize_CreateDocTab_Modal']//h4[@class='modal-title'][contains(text(),'Tab Details')]")
							private static WebElement createTab_TabDetailsLabel;  
			                                                  
							@FindBy(xpath="//input[@id='txtNewTabCaption']")
							private static WebElement createTab_TabCaptionTxt;  
			                                                   
							@FindBy(xpath="//input[@id='txtNewTabName']")
							private static WebElement createTab_TabNameTxt;  
							
							@FindBy(xpath="//input[@id='bIsHiddenTab']")
							private static WebElement createTab_PopHiddenChkbox;  
							
							@FindBy(xpath="//*[@id='MCustomize_CreateTab_Modal']/div/div/div[3]/button[1]")
							private static WebElement createTab_PopOkBtn;  
							
							@FindBy(xpath="//*[@id='MCustomize_CreateTab_Modal']/div/div/div[3]/button[2]")
							private static WebElement createTab_PopCancelBtn; 

							@FindBy(xpath="//label[@id='lblNameError_NewTab']")
							private static WebElement customizeMasterCreateTabCreateTabMessage;  
							
							
                          //Newly Created Create Tab  

							@FindBy(xpath="//span[contains(text(),'Create Tab')]")
							private static WebElement customizeMasterCreateCreateTabNewlyCreateTabCaption; 	
							
							              
							@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/ul[1]/li[1]/a[1]/span[1]")
							private static WebElement customizeMasterCreateCreateTabHeaderDetails; 	
							
							@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/div[1]/div[1]/div[1]/div[1]")
							private static WebElement customizeMasterCreateCreateTabBlankHeaderDetails; 	
							
							@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/ul[1]/li[2]/a[1]/span[1]")
							private static WebElement customizeMasterCreateCreateTabBodyDetails; 	
							
							@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/div[1]/div[2]/div[1]/div[1]")
							private static WebElement customizeMasterCreateCreateTabBlankBodyDetails; 
							
							@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[8]/div[1]/div[2]")
							private static WebElement customizeMasterCreateCreateTabNewlyCreateTabSaveButton; 	
							
							@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[8]/div[3]/div[2]")
							private static WebElement customizeMasterCreateCreateTabAddButton; 	
							
							@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[8]/div[4]/div[2]")
							private static WebElement customizeMasterCreateCreateTabPreviewButton; 	
							
							@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[8]/div[5]/div[2]")
							private static WebElement customizeMasterCreateCreateTabCloseButton; 	

				@FindBy(xpath="//button[@id='btnDeleteTab_MCust']")
				private static WebElement tabs_DeleteBtn; 
		
		@FindBy(xpath="//label[@id='lblNameError_TreeTab']")
		private static WebElement customizeMasterCreateTabsMessage; 		
		
		@FindBy(xpath="//label[@id='lblMasterNameError_TreeTab']")
		private static WebElement customizeMasterCreateTabsMasterNameMessage; 	
	
		@FindBy(xpath="//span[contains(text(),'Tree Tab')]")
		private static WebElement customizeMasterCreateTreeTabNewlyCreate; 	
		
		@FindBy(xpath="//input[@id='txtEditTabCaption']")
		private static WebElement customizeMasterCreateTreeTabNewlyCreateTabCaption; 	
		
		@FindBy(xpath="//div[@id='MCustomize_EditTab_Modal']//label[@class='col-sm-12 col-xs-12 col-md-5 col-lg-5 form-control-static Flabel'][contains(text(),'Hidden')]")
		private static WebElement customizeMasterCreateTreeTabNewlyCreateHidden; 	
		
		@FindBy(xpath="//div[@id='MCustomize_EditTab_Modal']//button[@class='Fbutton font-5'][contains(text(),'OK')]")
		private static WebElement customizeMasterCreateTreeTabNewlyCreateOkButton; 	
		
		@FindBy(xpath="//div[@id='MCustomize_EditTab_Modal']//button[@class='Fbutton font-5'][contains(text(),'Cancel')]")
		private static WebElement customizeMasterCreateTreeTabNewlyCreateCancelButton; 	
		
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[5]/ul[1]/li[1]/a[1]/span[1]")
		private static WebElement customizeMasterCreateTreeTabNewlyHeaderDetails; 	
		
		//               
		@FindBy(xpath="//div[@id='HeaderDetails_4']//div[@class='scrollable tbl-emptyfields-mastercust']")
		private static WebElement customizeMasterCreateTreeTabNewlyBlankInHeaderDetails; 	
		
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[5]/ul[1]/li[2]/a[1]/span[1]")
		private static WebElement customizeMasterCreateTreeTabNewlyBodyDetails; 	
		
		
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[5]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[4]")
		private static WebElement customizeMasterCreateTreeTabNewlyFieldInBodyDetails; 	
		
		@FindBy(xpath="//span[contains(text(),'Tree Tab Delete')]")
		private static WebElement  CustomizeCreateTreeTab;
		
		@FindBy(xpath="//*[@id='lblWarnMessage_TabModal']")
		private static WebElement  CustomizeDeleteTabMessage;

		@FindBy(xpath="//div[@id='MCustomize_DeleteTab_Modal']//button[@class='Fbutton font-5'][contains(text(),'Yes')]")
		private static WebElement  CustomizeDeleteTabMessageYesBtn;

		@FindBy(xpath="//div[@id='MCustomize_DeleteTab_Modal']//button[@class='Fbutton font-5'][contains(text(),'No')]")
		private static WebElement  CustomizeDeleteTabMessageNoBtn;
		
		
		//Master Customization Extra Fields Creation On Click Add button
	
		
		     @FindBy(xpath="//input[@id='ExtraField_FieldCaption']")
		     private static WebElement extraFields_FieldDetailsTab;   
		     
		     @FindBy(xpath="//span[contains(text(),'Field Details')]")
		     private static WebElement extraFields_ClickOnFieldDetailsTab; 

			     @FindBy(xpath="//*[@id='ExtraField_FieldCaption']")
			     private static WebElement extraFields_FieldDetailsCaption;   
			     
			     @FindBy(xpath="//div[@id='tab_FieldDetails_ExtraField']//div[1]//div[2]//div[1]//input[1]")
			     private static WebElement extraFields_FieldDetailsName;  
			     
			     @FindBy(xpath="//*[@id='ddlDatatype_ExtraField']")
			     private static WebElement extraFields_FieldDetailsDatTypeDropdown;  
			     
			     @FindBy(xpath="//*[@id='ExtraField_DefaultValue']")
			     private static WebElement extraFields_FieldDetailsDefaultValue;  
			     
			     @FindBy(xpath="//*[@id='ExtraField_BannerText']")
			     private static WebElement extraFields_FieldDetailsBannerText;  
			     
			     @FindBy(xpath="//*[@id='ExtraField_MaxSize']")
			     private static WebElement extraFields_FieldDetails_MaxSize;  
			     
			     @FindBy(xpath="//*[@id='ExtraField_ControlType']")
			     private static WebElement extraFields_FieldDetailsControlTypeDropdown;  
			     
			     @FindBy(xpath="//*[@id='ExtraField_ToolTipText']")
			     private static WebElement extraFields_FieldDetailsToolTipText;  
			     
			     @FindBy(xpath="//*[@id='ExtraField_RegularExpression']")
			     private static WebElement extraFields_FieldDetailsRegularExpressionText;  
			     
			     @FindBy(xpath=" //*[@id='btnLoad_RExp']")
			     private static WebElement extraFields_RegularExperssionLoadExpansionBtn;
		
			     
		     @FindBy(xpath="//*[@id='extraFieldContainer_div']/ul/li[2]/a/span")
			 private static WebElement extraFields_PropertiesTab;		     
						     
				@FindBy(xpath="//*[@id='tab_Properties_ExtraField']/div[1]/div[1]")
				private static WebElement extraFields_PropertiesTabPartOfDeliveryAddressChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_PartOfBillingAddress']")
				private static WebElement extraFields_PropertiesPartOfBillingAddressChekbox;
				
				@FindBy(xpath="//input[@id='FieldsDetails_IsMandatory']")
				private static WebElement extraFields_PropertiesMandatoryChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_AllowAsParameter']")
				private static WebElement extraFields_PropertiesAllowAsParameterChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_IsHidden']")
				private static WebElement extraFields_PropertiesHiddenChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_HiddenInGroup']")
				private static WebElement extraFields_PropertiesHiddenInGroupChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_SpellCheck']")
				private static WebElement extraFields_PropertiesSpellCheckChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_MandatoryInGroup']")
				private static WebElement extraFields_PropertiesMandatoryInGroupChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_AuditTrial']")
				private static WebElement extraFields_PropertiesAuditTrialChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_NotAvailableForReports']")
				private static WebElement extraFields_PropertiesNotAvailableForReportsChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_MergeField']")
				private static WebElement extraFields_PropertiesMergeFieldChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_MassUpdate']")
				private static WebElement extraFields_PropertiesMassUpdateChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_IncludeInQuickCreate']")
				private static WebElement extraFields_PropertiesIncludeInQuickCreateChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_CannotBeExported']")
				private static WebElement extraFields_PropertiesCannotBeExportedChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_CannotBeImported']")
				private static WebElement extraFields_PropertiesCannotBeImportedChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_CopyFromParent']")
				private static WebElement extraFields_PropertiesCopyFromParentChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_ReadOnly']")
				private static WebElement extraFields_PropertiesReadOnlyChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_InformationField']")
				private static WebElement extraFields_PropertiesInformationFieldChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_IsMandatoryInRevision']")
				private static WebElement extraFields_PropertiesIsMandatoryInRevisionChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_AvailableinCustomerPortal']")
				private static WebElement extraFields_PropertiesAvailableinCustomerPortalChekbox;
				
				@FindBy(xpath="//*[@id='FieldsDetails_EditableinCustomerPortal']")
				private static WebElement extraFields_PropertiesEditableinCustomerPortalChekbox;
				
				@FindBy(xpath="//*[@id='Behaviour']")
				private static WebElement extraFields_PropertiesBehaviourDropdown;     

			@FindBy(xpath="//*[@id='extraFieldContainer_div']/ul/li[3]/a/span")
			private static WebElement  extraFields_FormattingTab;
		        
		        @FindBy(xpath="//*[@id='FieldsDetails_ColumnSpan']")
				private static WebElement extraFields_FormattingColumnSpan ;
		        
		        @FindBy(xpath="//*[@id='FieldsDetails_RowSpan']")
				private static WebElement extraFields_FormattingRowSpan ;
		        
		        @FindBy(xpath="//*[@id='CharacterCasing']")
				private static WebElement extraFields_FormattingCharacterCasingDropdown ;
		        
		        @FindBy(xpath="//*[@id='TextAlign']")
				private static WebElement extraFields_FormattingTextAlignDropdown ;
		        
		        @FindBy(xpath="//*[@id='FieldsDetails_Font']")
				private static WebElement extraFields_FormattingFieldsDetailsFont ;
		        
		        @FindBy(xpath="//*[@id='FieldsDetails_BackColor']")
		        private static WebElement extraFields_FormattingBackColor ;
		        
		        @FindBy(xpath="//*[@id='FieldsDetails_Font_Button']")
				private static WebElement extraFields_FormattingFieldsDetailsFont_Btn ;
			       
			        //Elements are used for font edit option pop in field details and rules 
			        
			        @FindBy(xpath="//*[@id='fontFamily_ExtraField_FontCtrl']")
					private static WebElement FontAreialDropdown ;
			        
			        @FindBy(xpath="//*[@id='fontStyle_ExtraField_FontCtrl']")
					private static WebElement fontStyleDropdown ;
			        
			        @FindBy(xpath="//*[@id='fontWeight_ExtraField_FontCtrl']")
					private static WebElement fontWeightDropdown ;

			        @FindBy(xpath="//*[@id='fontSizes_ExtraField_FontCtrl']")
					private static WebElement fontSizeDropdown ;
			        
			        @FindBy(xpath="//*[@id='fontForeColor_ExtraField_FontCtrl']")
					private static WebElement foreColourDropdown ;
			        
			        @FindBy(xpath="//*[@id='fontBackColor_ExtraField_FontCtrl']")
					private static WebElement backColorDropdown ;
			        
			        @FindBy(xpath="//*[@id='chkBaseline_ExtraField_FontCtrl']")
					private static WebElement baselineChekbox ;
			        
			        @FindBy(xpath="//*[@id='chkOverLine_ExtraField_FontCtrl']")
					private static WebElement overLineChekbox ;
			        
			        @FindBy(xpath="//*[@id='chkStrikeThrough_ExtraField_FontCtrl']")
					private static WebElement strikeThroughChekbox ;
			        
			        @FindBy(xpath="//*[@id='chkUnderline_ExtraField_FontCtrl']")
					private static WebElement underlineChekbox ;
			        
			        @FindBy(xpath="//*[@id='chkEnableFont_ExtraField_FontCtrl']")
					private static WebElement enableFontChekbox ;
			        
			        @FindBy(xpath="//*[@id='previewfont_ExtraField_FontCtrl']")
					private static WebElement frontPriveiw ;
			        
			        @FindBy(xpath="//*[@id='myFontControl_IF_Save']/i")
					private static WebElement frontSaveBtn ;
			        
			        @FindBy(xpath="//*[@id='myFontControl_IF_Close']")
					private static WebElement frontCloseBtn ;
			        
				   
			 @FindBy(xpath="//*[@id='extraFieldContainer_div']/ul/li[4]/a/span")
			 private static WebElement extraFields_RulesTab ;
					                   
			    @FindBy(xpath="//*[@id='newRuleTab_div_FieldRules']/div[1]/div/div[1]/div[2]")
				private static WebElement extraFields_RulesSaveBtn ;
		                  
		        @FindBy(xpath="//*[@id='newRuleTab_div_FieldRules']/div[1]/div/div[2]/div[2]")
				private static WebElement extraFields_RulesNewBtn ;
		        
		        @FindBy(xpath="//*[@id='newRuleTab_div_FieldRules']/div[1]/div/div[3]/div[2]")
				private static WebElement extraFields_RulesDeleteBtn ;
		        
		        @FindBy(xpath="//div[@id='newRuleTab_div_FieldRules']//div[@class='col-sm-12 btnheader_txt font-7'][contains(text(),'Close')]")
				private static WebElement extraFields_RulesCloseBtn ;
		        
	            @FindBy(xpath="//input[@class='btn btn-link']")
			    private static WebElement extraFields_RulesClickHereAddRuleButtonBtn ;
	        
		        @FindBy(xpath="//input[@id='txtMasterRuleName_FieldRules']")
				private static WebElement extraFields_RuleNameDropdown ;
		        
		        @FindBy(xpath="//*[@id='txtMasterRuleName_FieldRules_input_settings']/span")
				private static WebElement extraFields_RuleNameSetingBtn ;      
			        
			        
			        
	        @FindBy(xpath="//*[@id='extraFieldContainer_div']/ul/li[5]/a/span")
	        private static WebElement extraFields_ExternalModulesTab;			     
		         
		        @FindBy(xpath="//*[@id='ddlOnEvent_mExtraFieldExternalModule']")
		        private static WebElement extraFields_ExternalModulesOnEventDropdown;	

		        @FindBy(xpath="//*[@id='txtbuttonCaption_mExtraFieldExternalModule']")
		        private static WebElement extraFields_ExternalModulesButtonCaption;

		        @FindBy(xpath="//*[@id='ddlModuleType_mExtraFieldExternalModule']")
		        private static WebElement extraFields_ExternalModules_ModuleTypeBtn;

		        @FindBy(xpath="//*[@id='ddlAllDllFiles_mExtraFieldExternalModule']")
		        private static WebElement extraFields_ExternalModules_ModuleNameDropdown;

		        @FindBy(xpath="//*[@id='ddlClassNames_mExtraFieldExternalModule']")
		        private static WebElement extraFields_ExternalModulesClassNameDropdown;

		        @FindBy(xpath="//*[@id='ddlClassMethods_mExtraFieldExternalModule']")
		        private static WebElement extraFields_ExternalModulesFuctionalNameDropdown;

		        @FindBy(xpath="//*[@id='ddlAvailability_mExtraFieldExternalModule']")
		        private static WebElement extraFields_ExternalModulesAvabileInDropdown;        
		        
		        @FindBy(xpath="//*[@id='MCustomize_FontDialog_Modal']/div/div/div[3]/button[1]")
			    private static WebElement extraFields_ExternalModulesForMatFieldDeatailsOkBtn ;
			        
			    @FindBy(xpath="//*[@id='MCustomize_FontDialog_Modal']/div/div/div[3]/button[2]")
			    private static WebElement extraFields_externalModulesForMatFieldDeatailsCanceBtn ;
				        
			 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]")
			 private static WebElement extraFields_OkBtn;
			    
			 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[2]/div[1]/div[2]/div[1]/div[2]/div[2]")
			 private static WebElement extraFields_CloseBtn;
			 
	 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[4]")
	 private static WebElement extraFields_CreatedByUserInHeaderDetailsTab;
			 
	 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/i[1]")
	 private static WebElement extraFields_EditOptionInHeaderDetailsTab;
	 
	 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/i[2]")
	 private static WebElement extraFields_DeleteOptionInHeaderDetailsTab;
	 
	 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[7]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[2]/td[4]")
	 private static WebElement extraFields_SecondFieldInHeaderDetailsTab;
	 
	 
	 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[8]/div[2]/div[2]")
	 private static WebElement extraFields_MoveFieldOption;
	 
		 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[8]/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/select[1]")
		 private static WebElement extraFields_MoveFieldTabName;
		 
		 @FindBy(xpath="//div[@id='tableType_div']//div[1]//label[1]")
		 private static WebElement extraFields_MoveFieldHeader;
		 
		 @FindBy(xpath="//div[@id='rdbBodyDiv_MoveField_Modal']//label[1]")
		 private static WebElement extraFields_MoveFieldBody;
		 
		 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[8]/div[1]/div[1]/div[3]/button[1]")
		 private static WebElement extraFields_MoveFieldOkButton;
		 
		 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[8]/div[1]/div[1]/div[3]/button[2]")
		 private static WebElement extraFields_MoveFieldCancelButton;
		 
		 
	     //Settings Tab Move Field
		 
		 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[3]/td[1]/i[1]")
		 private static WebElement extraFields_SettingsTab_MoveField;
	 
		 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[4]")
		 private static WebElement extraFields_SettingsTab_First;
	 
     	 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/i[2]")
		 private static WebElement extraFields_SettingsTab_FirstDelete;
	 
     	 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[3]/td[1]/i[2]")
		 private static WebElement extraFields_SettingsTab_MoveFieldOptsDelete;
     	 
     	 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[9]/div[1]/div[1]/div[2]/div[1]/div[1]/label[1]")
		 private static WebElement extraFields_SettingsTab_DeleteMessage;
     	 
     	 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[9]/div[1]/div[1]/div[3]/button[2]")
		 private static WebElement extraFields_SettingsTab_DeleteMessageNoOption;
     	
     	 @FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[9]/div[1]/div[1]/div[3]/button[1]")
		 private static WebElement extraFields_SettingsTab_DeleteMessageYesOption;
     	
     	 

     	 @FindBy(xpath="//a[@class='icon-font7 theme_color-inverse']")
		 private static WebElement extraFields_GeneralTab;
     	 
     	 @FindBy(xpath="//a[contains(text(),'Settings')]")
		 private static WebElement extraFields_SettingsTab;
     	
     	 @FindBy(xpath="//a[contains(text(),'Details')]")
		 private static WebElement extraFields_DetailsTab;
     	
     	 @FindBy(xpath="//a[contains(text(),'Print Layout')]")
		 private static WebElement extraFields_PrintLayout;
     	
     	 @FindBy(xpath="//a[contains(text(),'Tree Tab Update')]")
		 private static WebElement extraFields_TreeTabUpdate;
     	
     	 @FindBy(xpath="//a[@class='icon-font7 theme_color'][contains(text(),'Document Tab')]")
		 private static WebElement extraFields_DocumentTab;
     	
     	 @FindBy(xpath="//a[@class='icon-font7 theme_color'][contains(text(),'Create Tab')]")
		 private static WebElement extraFields_CreateTab;
     	

     	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[2]/div[1]/div[1]/div[8]/div[4]/div[2]")
		private static WebElement extraFields_PreviewTab;
     	
     	
     	@FindBy(xpath="//input[@id='sFileName1']")
		private static WebElement extraFields_PreviewTab_DocumentTab_FileName;
     	
     	@FindBy(xpath="//input[@id='biDocument1']")
		private static WebElement extraFields_PreviewTab_DocumentTab_Document;
     	
     	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[6]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[5]/div[1]/div[1]/div[1]")
		private static WebElement extraFields_PreviewTab_DocumentTab_UploadDocument;
     	
     	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[6]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[5]/div[1]/div[1]/div[2]")
		private static WebElement extraFields_PreviewTab_DocumentTab_SaveDocument;
     	
     	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[6]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[5]/div[1]/div[1]/div[3]")
		private static WebElement extraFields_PreviewTab_DocumentTab_DeleteDocument;
     	
     	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[6]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[7]/table[1]/tbody[1]/tr[1]/td[2]/input[2]")
		private static WebElement extraFields_PreviewTab_DocumentTab_CreateDate;
     	
     	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[6]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[9]/div[1]/input[1]")
		private static WebElement extraFields_PreviewTab_DocumentTab_FileSize;
     	
     	
     	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[7]/div[1]/div[2]/div[2]/div[1]/div[1]/input[1]")
		private static WebElement extraFields_PreviewTab_CreateCreateTab_ExtraFieldUpdate;
     	
     	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[3]/div[2]/div[1]/ul[1]/li[2]/button[2]/i[1]")
		private static WebElement extraFields_PreviewTab_SaveBtn;
     	
     	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[23]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[3]/div[2]/div[1]/ul[1]/li[2]/button[5]")
		private static WebElement extraFields_PreviewTab_CloseBtn;
     	
     	
     	@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table/tbody/tr[1]/td/div/table/tbody/tr[2]/td[8]/div[1]/label/input")
		private static WebElement selectVendorB;
     	
     	
     	
     	@FindBy(xpath="//input[@id='iPDCDiscountedAC']")
		private static WebElement pdcDicountAccount;
     
    	
     	@FindBy(xpath="//input[@id='uploadBtnbiDocument1']")
		private static WebElement addGroupDocumentTab_UploadDocument;
     	
     	@FindBy(xpath="//tr[1]//td[5]//div[1]//div[1]//div[2]")
		private static WebElement addGroupDocumentTab_SaveDocument;
     	
     	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/form[1]/div[1]/div[6]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[5]/div[1]/div[1]/div[3]/i[1]")
		private static WebElement addGroupDocumentTab_DeleteDocument;
     	
     	
     	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/form[1]/div[1]/div[7]/div[1]/div[2]/div[2]/div[1]/div[1]/input[1]")
		private static WebElement addGroupCreateCreateTab_ExtraFieldUpdate;
     	
		@FindBy(xpath="//td[contains(@class,'tdName padding0')][contains(text(),'Vendor Group')]")
		public static WebElement accountGroupCreation;
		
		               
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[8]/div[1]/label[1]/input[1]")
		private static WebElement accSelect;
     	
		
	
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[11]")
		public static WebElement accSelectBySingleClick;
	
		
		
		

		@FindBy(xpath="//span[contains(text(),'Vendor Group')]")
		public static WebElement clickOnAccountGroupToDisplayTitle;
		
		
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/ol[1]/li[4]/span[1]")
		public static WebElement clickOnAccountGroup;
		
		
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/form[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/input[1]")
		public static WebElement getCpationOfGroup;
		
	
		@FindBy(xpath="//label[@class='msgInfo']")
		public static WebElement clickOnEmptyGroup;
		
		
		
		
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[3]/ul[1]/li[1]/ul[6]/li[1]/ul[3]/li[1]/ul[1]/li[1]/a[2]")
		public static WebElement selectVendorGroupCheckbox;
		
		@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table/tbody/tr[1]/td/div/table/tbody/tr[1]/td[8]/div[1]/label/input")
		public static WebElement selectVendorGroup;
			
		@FindBy(xpath="//*[@id='innerDivDeleteMasterModal']//div[1]/label")
		public static WebElement getMsgOnDelete;
		
		@FindBy(xpath="//button[@id='btnOkForDelete']")
		public static WebElement clickOnOkInDelete;
		
		@FindBy(xpath="//button[@id='btnCancelForDelete']")
		public static WebElement clickOnCancelInDelete;
		
		
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
		
		//@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[8]/div[1]/label[1]/input[1]")
		@FindBy(xpath="//*[@id='LandingGridBody']/tr[1]/td[8]")
		public static WebElement accountFirstCheckBox;
		                
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[2]/td[8]/div[1]/label[1]/input[1]")
		public static WebElement accountSecondCheckBox;
		
		@FindBy(xpath="//button[@class='btn-link theme_icon-color btnGenerateMasterclose closeMasterPartial clsMasterSaveClose']")
		public static WebElement closeAccMasterCreationScreen;
		
		
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/ul[1]/li[2]/button[5]/i[1]")
		public static WebElement closeMaster;
		
		
		//Dash board Page
		
		@FindBy(xpath="//*[@id='dashName']")
	    private static WebElement labelDashboard ;
		
		@FindBy(xpath="//*[@id='Select_dash']")
		private static WebElement selectDashboard ;
			
		@FindBy(xpath="//*[@id='Dashboard_AddDash']")
	    private static WebElement newAddDashBoard;
		 
	    @FindBy(xpath="//*[@id='Dashboard_Dash_Config']")
        private static WebElement dashboardCustomizationSettings;
	
		
	   
	    //Properties Page
	   
	    @FindBy(xpath="//span[contains(text(),'Account Properties')]")
	    public static WebElement accountPropertiesLabel;
	     
		@FindBy(xpath="//input[@id='rbnNormal']")
		public static WebElement normalRadioBtn;

		@FindBy(xpath="//input[@id='rbnDisplayDebitCreditTotalForEachAccountUnderGroup']")
		private static WebElement  displayDebitCreditTotalForEachAccountUnderGroup;
		
		@FindBy(xpath="//input[@id='rbnClubTranOfAllAcctsUnderGrpWhenDispInLedger']")
		private static WebElement  ClubTheTransactionsofAllTheAccountsUnderGroup;
							
		@FindBy(xpath="//input[@id='rbnShowSummaryOfAccount']")
		private static WebElement showSummaryOfTheAccountRadioBtn ;	
		
		@FindBy(xpath="//input[@id='chkSuspendCreditLimit']")
		private static WebElement suspendCreditLimitChkbox ;
								
		@FindBy(xpath="//input[@id='rbnDisplayDebitCreditTotalForEachMonth']")
		private static WebElement  displayDebitCreditTotalForEachMonthRadioBtn;
							
		@FindBy(xpath="//input[@id='rbnrbnDisplayDebitCreditTotalForEachDay']")
		private static WebElement  displayDebitCreditTotalForEachDayRadioBtn;
		
		@FindBy(xpath="//input[@id='rbnNone']")
		public static WebElement  noneRadioBtn;
		
		@FindBy(xpath="//input[@id='rbnConsolidateWhenBothAccountsAreSame']")
		public static WebElement consolidateWhenBothAccountsAreSameRadioBtn ;
		
		@FindBy(xpath="//input[@id='rbnConsolidateAlways']")
		public static WebElement consolidateAlwaysRadioBtn ;
		
		@FindBy(xpath="//select[@id='ddlStatus']")
		public static WebElement  statusDropdown;
		
		@FindBy(xpath="//input[@id='chkDoNotRestrictSelectionEvenIfRightsNotAllotted']")
		public static WebElement  doNotRestrictSelectionEvenIfRightsNotAllottedCheckbox;
		
		@FindBy(xpath="//input[@id='chkAllowOtherCompaniesToViewRecords']")
		public static WebElement allowOtherCompaniesToViewRecordsCheckbox;
		
		@FindBy(xpath="//input[@id='chkGenerateLedgerByCurrency']")
		public static WebElement generateLedgerByCurrencyCheckbox;
		
		@FindBy(xpath="//input[@id='chkSendSmsEmailWhereverAcDebitedCredited']")
		public static WebElement sendSmsEmailWhereverAcDebitedCreditedCheckbox;
		
		@FindBy(xpath="//span[@id='btnPropOk']")
		public static WebElement properties_okBtn;
		
		@FindBy(xpath="//*[@id='divMasterProperty']/ul/li/span[2]")
		public static WebElement  properties_cancelBtn;
	     
		
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/ul[1]/li[1]/ul[1]/li[5]/span[1]")
		public static WebElement  customizeTabWithTreeTabUpdate;
		
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/ul[1]/li[1]/ul[1]/li[6]/span[1]")
		public static WebElement  customizeTabWithDocumentTab;
		
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/ul[1]/li[1]/ul[1]/li[7]/span[1]")
		public static WebElement  customizeTabWithCreateTreeTab;
		
	    @FindBy(xpath="//div[@id='Layout_Popup']//input[2]")
		private static WebElement noIn;
	    
		@FindBy(xpath="//*[@id='dashName']")
		private static WebElement dashboard;
		
		
		@FindBy(xpath="//div[@id='chartdiv0']")
		public static WebElement dashboardGraph;
					
		@FindBy(xpath="//div[@id='DBdashlate1']")
		public static WebElement dashboardLedger;
		
		@FindBy(xpath="//div[@id='DBdashlate2']")
		public static WebElement dashboardInfoPanel;

		
		
		//Logout and Login Screen
		
		@FindBy(xpath="//*[@id='txtUsername']")
		private static WebElement username;

		@FindBy(id="txtPassword")
		private static WebElement password;

		@FindBy(id="btnSignin")
		private static WebElement signIn;

	    @FindBy(id="ddlCompany")
	    private static WebElement companyDropDownList;

		/*@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/header[1]/nav[1]/div[1]/ul[1]/li[5]/a[1]/span[1]")
		private static WebElement userNameDisplay;*/

		@FindBy(xpath="//span[@class='hidden-xs']")
		private static WebElement userNameDisplay;
		
		
		@FindBy(xpath="//*[@id='companyLogo']")
		private static WebElement companyLogo;

		@FindBy(xpath="//*[@id='ulCompanyDetails_HomePage']/li[1]")
		private static WebElement companyName;

	    @FindBy(xpath="//*[@id='userprofile']/li/span[2]")
        private static WebElement logoutOption;
		
	    
	    @FindBy(xpath="//input[@id='donotshow']")
		private static WebElement doNotShowCheckbox;
	    
	    @FindBy(xpath="//span[@class='pull-right']")
		private static WebElement closeBtnInDemoPopupScreen;
	    
	    
	    
	    @FindBy(xpath="//div[contains(text(),'Settings')]")
		private static WebElement  settingsmenuBtn;
		
		@FindBy(xpath="//span[contains(text(),'Configure Transactions')]")
		private static WebElement  configureTransactionBtn;
		
		@FindBy(xpath="//label[contains(text(),'Preferences')]")
		private static WebElement  preferencesBtn;
		
		@FindBy(xpath="//div[@id='docCustomization']")
		private static WebElement  documentCustomization45Btn;
		
		@FindBy(xpath="//span[@class='icon-general icon-font6 theme_button_color']")
		private static WebElement  createVoucherBtn;
		
		@FindBy(xpath="//span[@id='spanDeleteVoucher']")
		private static WebElement  deletevoucherBtn;
		
		@FindBy(xpath="//a[@id='4608']")
		private static WebElement  receiptsBtn;
		
		@FindBy(xpath="//a[@id='4864']")
		private static WebElement  paymentsBtn;
		
		@FindBy(xpath="//a[@id='5120']")
		private static WebElement  pettyCashBtn;
		
		@FindBy(xpath="//a[@id='5888']")
		private static WebElement  postDatedreceiptsBtn;
		
		@FindBy(xpath="//a[@id='7168']")
		private static WebElement  postDatedpaymentsBtn;
		
		@FindBy(xpath="//a[@id='768']")
		private static WebElement  purchaseVouchersBtn;
		
		@FindBy(xpath="//a[@id='1023']")
		private static WebElement  stockAdjustmentBtn;
		
		@FindBy(xpath="//a[@id='6400']")
		private static WebElement  purchaseReturnsBtn;
		
		@FindBy(xpath="//a[@id='2560']")
		private static WebElement  purcahseOrdersBtn;
		
		@FindBy(xpath="//a[@id='2561']")
		private static WebElement  jobOrdersBtn;
		
		@FindBy(xpath="//a[@id='1280']")
		private static WebElement  materialReceiptsNotesBtn;
		
		@FindBy(xpath="//a[@id='3328']//i[contains(@class,'icon-font7')]")
		private static WebElement  salesInvoiceBtn;
		
		@FindBy(xpath="//a[@id='3329']")
		private static WebElement  cashSalesBtn;
		
		@FindBy(xpath="//a[@id='3330']")
		private static WebElement  hirePurcahseSalesBtn;
		
		@FindBy(xpath="//a[@id='3331']")
		private static WebElement  posSalesBtn;
		
		@FindBy(xpath="//a[@id='1792']")
		private static WebElement  salesReturnsBtn;
		
		@FindBy(xpath="//a[@id='5632']")
		private static WebElement  salesOrdersBtn;
		
		@FindBy(xpath="//a[@id='5633']")
		private static WebElement  productionOrderBtn;
		
		@FindBy(xpath="//a[@id='6144']")
		private static WebElement  deliveryNotesBtn;
		
		@FindBy(xpath="//a[@id='3584']")
		private static WebElement  nonStandardjournalEntriesBtn;
		
		@FindBy(xpath="//a[@id='3585']")
		private static WebElement  interDepartmentalJVBtn;
		
		@FindBy(xpath="//a[@id='3586']")
		private static WebElement  debitNotesLineWiseBtn;
		
		@FindBy(xpath="//a[@id='3587']")
		private static WebElement  creditNotesLineWiseBtn;
		
		@FindBy(xpath="//a[@id='3588']")
		private static WebElement  fixedAssetDepreciationVoucherBtn;
		
		@FindBy(xpath="//a[@id='8704']")
		private static WebElement  journalEntriesBtn;
		
		@FindBy(xpath="//a[@id='3840']")
		private static WebElement  debitNotesBtn;
		
		@FindBy(xpath="//a[@id='4096']")
		private static WebElement  creditNotesBtn;
		
		@FindBy(xpath="//a[@id='4097']")
		private static WebElement  payRollPostingsBtn;
		
		@FindBy(xpath="//a[@id='256']")
		private static WebElement  openingBalancesBtn;
		
		@FindBy(xpath="//a[@id='5376']")
		private static WebElement  shortageInStocksBtn;
		
		@FindBy(xpath="//div[@class='col-xs-12']//a[@id='2048']")
		private static WebElement  excessInStocksBtn;
		
		@FindBy(xpath="//a[@id='3072']")
		private static WebElement  stockTransfersBtn;
		
		@FindBy(xpath="//div[@class='col-xs-12']//a[@id='512']")
		private static WebElement  openingStocksBtn;
		
		@FindBy(xpath="//a[@id='6656']")
		private static WebElement  issuesToProductionBtn;
		
		@FindBy(xpath="//a[@id='1024']")
		private static WebElement  receiptsFromProductionBtn;
		
		@FindBy(xpath="//a[@id='1056']")
		private static WebElement  returnsFromShopfloorBtn;
		
		@FindBy(xpath="//a[@id='6912']")
		private static WebElement  jobWorkIssuesBtn;
		
		@FindBy(xpath="//a[@id='1536']")
		private static WebElement  jobWorkReeceiptsBtn;
		
		@FindBy(xpath="//a[@id='2304']")
		private static WebElement  purchaseQuotationsBtn;
		
		@FindBy(xpath="//a[@id='7424']")
		private static WebElement  salesQuotationsBtn;
		
		@FindBy(xpath="//a[@id='7680']")
		private static WebElement  jobOrderBtn;
		
		@FindBy(xpath="//a[@id='2816']")
		private static WebElement  productionProcessBtn;
		
		@FindBy(xpath="//a[@id='7936']")
		private static WebElement  materialRequisitionBtn;
		
		@FindBy(xpath="//a[@id='8192']")
		private static WebElement  requestForQuoteBtn;
		
		@FindBy(xpath="//a[@id='8448']")
		private static WebElement  forecxJVBtn;
		
		@FindBy(xpath="//span[@id='updateButton']")
		private static WebElement  updateBtn;

		 
	    @FindBy(xpath="//input[@id='donotshow']")
		private static WebElement DontShowChkBox;
	   		    
	    @FindBy(xpath="//span[@class='pull-right']")
	    private static WebElement PopUpcloseBtn;	
		
	   
	    
	    @FindBy(xpath="//div[@class='font-5 theme_background-color-inverse theme_color-inverse']")
		private static WebElement  documentsTab;
		
		@FindBy(xpath="//li[@id='navigationtab1']")
		private static WebElement  documentsTabForMaterialReceiptNotes;
		

		@FindBy(xpath="//i[@class='icon-load-from icon-font7']")
		private static WebElement  loadMastersBtn;
		
					@FindBy(xpath="//select[@id='doc_voucherDropDown']")
					private static WebElement  loadmasterDropDown;
					
					@FindBy(xpath="//button[@id='btnSelectDependMaster']")
					private static WebElement  selectAllBtn;
					
					@FindBy(xpath="//button[contains(text(),'Reset')]")
					private static WebElement  resetBtn;
					
					@FindBy(xpath="//button[@class='Fbutton pull-right'][contains(text(),'Ok')]")
					private static WebElement  okBtn;
		
		@FindBy(xpath="//*[@id='editScreen_CustomizeButtons']/span[1]")
		private static WebElement  deleteTagBtn;
		
		@FindBy(xpath="//i[@class='icon-add-group icon-font7']")
		private static WebElement  addGroupBtn;
		
					@FindBy(xpath="//input[@id='doc_GroupName']")
					private static WebElement  groupNameTXt;
					
					@FindBy(xpath="//button[contains(text(),'Add')]")
					private static WebElement  addBtn;
					
					@FindBy(xpath="//div[@class='col-xs-12 form-group']//button[@id='btnCancel']")
					private static WebElement  cancelBtn;

		@FindBy(xpath="//div[@id='voucherTabContent']//span[2]")
		private static WebElement  deletegroupBtn;
		
		
		
		
		
		@FindBy(xpath="//input[@id='doc_baseDocument']")
		private static WebElement  baseDocumentTxt;
		
		@FindBy(xpath="//input[@id='txtbox_doc_Accountdepandency']")
		private static WebElement  accountDependencyTxt;
		
					@FindBy(xpath="//span[contains(text(),'iExchangeAdjustmentGainAC')]")
					private static WebElement  iExchangeAdjustmentGainACChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iExchangeAdjustmentLossAC')]")
					private static WebElement  iExchangeAdjustmentLossACChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iPrimaryAccount')]")
					private static WebElement  iPrimaryAccountChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iCity')]")
					private static WebElement  iCityChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iDeliveryCity')]")
					private static WebElement  iDeliverycityChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iBankAc')]")
					private static WebElement  iBankACChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iPDCDiscountedAC')]")
					private static WebElement  iPDCDiscountedACChkBox;
					
		@FindBy(xpath="//input[@id='txtbox_doc_Productdepandency']")
		private static WebElement  itemDependencyTxt;
		
					@FindBy(xpath="//span[contains(text(),'Outlet__')]")
					private static WebElement  outletChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iCostofShortageStockAC')]")
					private static WebElement  iCostofShortageStockACChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iCostofExcessStockAC')]")
					private static WebElement  iCostofExcessStockACChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iCostofSaleReturnAC')]")
					private static WebElement  iCostofSaleReturnACChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iPurchaseVarianceAC')]")
					private static WebElement  iPurchaseVarianceACChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iDefaultBaseUnit')]")
					private static WebElement  iDefaultBaseUnitChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iDefaultSalesUnit')]")
					private static WebElement  iDefaultSalesUnitChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iDefaultPurchaseUnit')]")
					private static WebElement  iDefaultPurchaseUnitChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iCostOfIssueAccount')]")
					private static WebElement  iCostOfIssueAccountChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iStocksAccount')]")
					private static WebElement  iStocksAccountChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iSalesAccount')]")
					private static WebElement  iSalesAccountChkBox;
					
					@FindBy(xpath="//li[12]//div[1]//label[1]//span[1]")
					private static WebElement  otherdetails1Box;
					
					@FindBy(xpath="//li[13]//div[1]//label[1]//span[1]")
					private static WebElement  replenishment1ChkBox;
					
					@FindBy(xpath="//span[contains(text(),'Reorder__')]")
					private static WebElement  reorderChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iCategory')]")
					private static WebElement  iCategoryChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iWIPAccount')]")
					private static WebElement  iWIPAccountChkBox;
					
					@FindBy(xpath="//span[contains(text(),'Classification__')]")
					private static WebElement  classificationChkBox;
					
					@FindBy(xpath="//li[18]//div[1]//label[1]//span[1]")
					private static WebElement  repelenishment2ChkBox;
					
					@FindBy(xpath="//li[19]//div[1]//label[1]//span[1]")
					private static WebElement  otherDetails2ChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iBin')]")
					private static WebElement  iBinChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iAlternateCategory')]")
					private static WebElement  iAlternateCategoryChkBox;
					
					@FindBy(xpath="//span[contains(text(),'iTaxCode')]")
					private static WebElement  iTaxCodeChkBox;
		
		@FindBy(xpath="//select[@id='doc_Accountdepandency_Filter']")
		private static WebElement  accountfilterDropdown;
		
		@FindBy(xpath="//select[@id='doc_Itemdepandency_Filter']")
		private static WebElement  itemfilterDropdown;
		
		
		
		@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/input[1]")
		private static WebElement  titleTxt;
		
		
		
		// Documents Grid Elements
		
			// first row 
		
		@FindBy(xpath="//td[@id='doc_TagsTable_col_1-0']")
		private static WebElement  firstrow;
		
		@FindBy(xpath="//div[contains(text(),'Delete Row')]")
		private static WebElement  deleterowBtn;
		
		@FindBy(xpath="//div[contains(text(),'Insert Row')]")
		private static WebElement  insertRowBtn;
		
		@FindBy(xpath="//td[@id='doc_TagsTable_col_1-1']")
		private static WebElement  masters1;
		
		@FindBy(xpath="//td[@id='doc_TagsTable_col_1-2']")
		private static WebElement  position1;
		
		@FindBy(xpath="//td[@id='doc_TagsTable_col_1-3']")
		private static WebElement  showDependency1;
		
		@FindBy(xpath="//td[@id='doc_TagsTable_col_1-4']")
		private static WebElement  group1;
		
		@FindBy(xpath="//td[@id='doc_TagsTable_col_1-5']")
		private static WebElement  filter1;
		
		@FindBy(xpath="//td[@id='doc_TagsTable_col_1-6']")
		private static WebElement  mandatory1;
		
		
		// Second Row... delete and insert row are same
		
		@FindBy(xpath="//td[@id='doc_TagsTable_col_2-0']")
		private static WebElement  secondRow;
		
		@FindBy(xpath="//td[@id='doc_TagsTable_col_2-1']")
		private static WebElement  masters2;
		
		@FindBy(xpath="//td[@id='doc_TagsTable_col_2-2']")
		private static WebElement  position2;
		
		@FindBy(xpath="//td[@id='doc_TagsTable_col_2-3']")
		private static WebElement  showDependency2;
		
		@FindBy(xpath="//td[@id='doc_TagsTable_col_2-4']")
		private static WebElement  group2;
		
		@FindBy(xpath="//td[@id='doc_TagsTable_col_2-5']")
		private static WebElement  filter2;
		
		@FindBy(xpath="//td[@id='doc_TagsTable_col_2-6']")
		private static WebElement  mandatory2;
		
		
		// Third Row... delete and insert row are same
		
			@FindBy(xpath="//td[@id='doc_TagsTable_col_3-0']")
			private static WebElement  thirdRow;
			
			@FindBy(xpath="//td[@id='doc_TagsTable_col_3-1']")
			private static WebElement  masters3;
			
			@FindBy(xpath="//td[@id='doc_TagsTable_col_3-2']")
			private static WebElement  position3;
			
			@FindBy(xpath="//td[@id='doc_TagsTable_col_3-3']")
			private static WebElement  showDependency3;
			
			@FindBy(xpath="//td[@id='doc_TagsTable_col_3-4']")
			private static WebElement  group3;
			
			@FindBy(xpath="//td[@id='doc_TagsTable_col_3-5']")
			private static WebElement  filter3;
			
			@FindBy(xpath="//td[@id='doc_TagsTable_col_3-6']")
			private static WebElement  mandatory3;
			
			
			// Forth Row... delete and insert row are same
			
				@FindBy(xpath="//td[@id='doc_TagsTable_col_4-0']")
				private static WebElement  forthRow;
				
				@FindBy(xpath="//td[@id='doc_TagsTable_col_4-1']")
				private static WebElement  masters4;
				
				@FindBy(xpath="//td[@id='doc_TagsTable_col_4-2']")
				private static WebElement  position4;
				
				@FindBy(xpath="//td[@id='doc_TagsTable_col_4-3']")
				private static WebElement  showDependency4;
				
				@FindBy(xpath="//td[@id='doc_TagsTable_col_4-4']")
				private static WebElement  group4;
				
				@FindBy(xpath="//td[@id='doc_TagsTable_col_4-5']")
				private static WebElement  filter4;
				
				@FindBy(xpath="//td[@id='doc_TagsTable_col_4-6']")
				private static WebElement  mandatory4;
		
		
		
		// After clicking *** Properties Changes which are same for all rows
		
		@FindBy(xpath="//select[@id='doc_TagsTableMasterDropDown']")
		private static WebElement  masterDropDown;
		
		@FindBy(xpath="//select[@id='doc_TagsPositionDropDown']")
		private static WebElement  positionDropdown;
		
		@FindBy(xpath="//input[@id='txtbox_doc_TagsDepedencyDropDown']")
		private static WebElement  showDependencyDropdown;
		
		@FindBy(xpath="//input[@id='doc_TagsGroupTxtbox']")
		private static WebElement  groupTxt;
		
		@FindBy(xpath="//select[@id='docFilter']")
		private static WebElement  filterDropdown;
		
		@FindBy(xpath="//select[@id='doc_MandatoryDropDown']")
		private static WebElement  mandatoryDropDown;
		
		// Grid Elements completed
		
		// Right Panel 
		
		// Home ***************************
		
		@FindBy(xpath="//*[@id='id_menu_tree_1']/a/i")
		private static WebElement  homeExpandBtn;
		
					@FindBy(xpath="//*[@id='id_menu_tree_2']/a/i")
					private static WebElement  companyExpandBtn;
					
								@FindBy(xpath="//*[@id='id_menu_tree_3']/a/i")
								private static WebElement  newComapanyBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_131']/a/i")
								private static WebElement  yearEndProcesBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_125']/a/i")
								private static WebElement  interComapnyTransactionBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_134']/a/i")
								private static WebElement  synchronizationBtn;
								
					@FindBy(xpath="//*[@id='id_menu_tree_6']/a/i")
					private static WebElement  securityBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_1000']/a/i")
					private static WebElement  mastersExpandBtn;
				
								@FindBy(xpath="//*[@id='id_menu_tree_220']/a/i")
								private static WebElement  currencyBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_221']/a/i")
								private static WebElement  itemBtn;
					
								@FindBy(xpath="//*[@id='id_menu_tree_8002']/a/i")
								private static WebElement  schemesExpandBtn;
								
											@FindBy(xpath="//*[@id='id_menu_tree_8015']/a/i")
											private static WebElement  discountDefinitionBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_224']/a/i")
								private static WebElement  mrpBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_225']/a/i")
								private static WebElement  qcBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_234']/a/i")
								private static WebElement  warehouseManagementSystemBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_251']/a/i")
								private static WebElement  maintainceBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_8001']/a/i")
								private static WebElement  pointofSaleExpandBtn;
								
											@FindBy(xpath="//*[@id='id_menu_tree_8005']/a/i")
											private static WebElement  memberBtn;
											
											@FindBy(xpath="//*[@id='id_menu_tree_8006']/a/i")
											private static WebElement  giftVoucherBtn;
											
											@FindBy(xpath="//*[@id='id_menu_tree_1025']/a/i")
											private static WebElement  pointOfSaleTableManagementBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_1103']/a/i")
								private static WebElement  payrollBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_1056']/a/i")
								private static WebElement  tdsAnnexureBtn;
								
					@FindBy(xpath="//*[@id='id_menu_tree_16']/a/i")
					private static WebElement  settingsExpandBtn;
					
								@FindBy(xpath="//*[@id='id_menu_tree_66']/a/i")
								private static WebElement  authorizationBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_25']/a/i")
					private static WebElement  datamanagementBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_20']/a/i")
					private static WebElement  utilitiesBtn;
		
					
		// Finance *************************			
					
		@FindBy(xpath="//*[@id='id_menu_tree_60']/a/i")
		private static WebElement  financeExpandBtn;
		
					@FindBy(xpath="//*[@id='id_menu_tree_61']/a/i")
					private static WebElement  financeTransactionExpandBtn;
					
								@FindBy(xpath="//*[@id='id_menu_tree_2001']/a/i")
								private static WebElement  cashBankBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_2007']/a/i")
								private static WebElement  financePurchasesBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_2014']/a/i")
								private static WebElement  financeSalesBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_2023']/a/i")
								private static WebElement  journalsBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_2039']/a/i")
								private static WebElement  manufacturingBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_219']/a/i")
								private static WebElement  autoPostingsBtn;
								
					@FindBy(xpath="//*[@id='id_menu_tree_62']/a/i")
					private static WebElement  creditManagementExpandBtn;
					
								@FindBy(xpath="//*[@id='id_menu_tree_217']/a/i")
								private static WebElement  delinquencyBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_81']/a/i")
					private static WebElement  financeReportsExpandBtn;
					
								@FindBy(xpath="//*[@id='id_menu_tree_82']/a/i")
								private static WebElement  cashAndBankBookBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_502']/a/i")
								private static WebElement  salesReportBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_506']/a/i")
								private static WebElement  purchaseReportBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_207']/a/i")
								private static WebElement  vatBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_514']/a/i")
								private static WebElement  registersBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_523']/a/i")
								private static WebElement  salesAndPurchaseAnalysisBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_568']/a/i")
								private static WebElement  miscellaneousAnalysisBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_536']/a/i")
								private static WebElement  masterInformationBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_520']/a/i")
								private static WebElement  financeReportsAuditTrialBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_50001']/a/i")
					private static WebElement  fixedAssetBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_530']/a/i")
					private static WebElement  finalAccountsBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_558']/a/i")
					private static WebElement  receiveableAndPayableAnalysisExpandBtn;
					
								@FindBy(xpath="//*[@id='id_menu_tree_605']/a/i")
								private static WebElement  customerDetailBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_607']/a/i")
								private static WebElement  customerSummaryBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_606']/a/i")
								private static WebElement  vendorDetailBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_608']/a/i")
								private static WebElement  vendorSummaryBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_3301']/a/i")
					private static WebElement  budgetExpandBtn;
					
								@FindBy(xpath="//*[@id='id_menu_tree_3304']/a/i")
								private static WebElement  reviseBudgetBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_3321']/a/i")
								private static WebElement  budgetReportsBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_3311']/a/i")
					private static WebElement  onlinepaymentsBtn;
					
	//Inventory   *************************
					
		@FindBy(xpath="//*[@id='id_menu_tree_135']/a/i")
		private static WebElement  inventoryExpandBtn;
		
					@FindBy(xpath="//*[@id='id_menu_tree_137']/a/i")
					private static WebElement  inventoryTransactionExpandBtn;
					
								@FindBy(xpath="//*[@id='id_menu_tree_139']/a/i")
								private static WebElement  inventoryPurchasesBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_140']/a/i")
								private static WebElement  inventorySalesBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_2033']/a/i")
								private static WebElement  stocksBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_200']/a/i")
					private static WebElement  inventoryReportsExpandBtn;
					
								@FindBy(xpath="//*[@id='id_menu_tree_540']/a/i")
								private static WebElement  binsReportBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_621']/a/i")
								private static WebElement  stockAgeingAnalysisBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_550']/a/i")
					private static WebElement  ordermanagementExpandBtn;
					
								@FindBy(xpath="//*[@id='id_menu_tree_600']/a/i")
								private static WebElement  analysisOfLinkedUnlinkedDocumentsBtn;
								
	//Fixed Asset************************							
		
		@FindBy(xpath="//*[@id='id_menu_tree_136']/a/i")
		private static WebElement  fixedAssetExpandBtn;
		
					@FindBy(xpath="//*[@id='id_menu_tree_138']/a/i")
					private static WebElement  fixedAssetTransactionExpandBtn;
					
								@FindBy(xpath="//*[@id='id_menu_tree_50010']/a/i")
								private static WebElement  fixedAssetReportsBtn;
								
	//Production ***********************							
		
		@FindBy(xpath="//*[@id='id_menu_tree_3000']/a/i")
		private static WebElement  productionExpandBtn;
		
					@FindBy(xpath="//*[@id='id_menu_tree_3001']/a/i")
					private static WebElement  productionTransactionBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_625']/a/i")
					private static WebElement  productionReportBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_3006']/a/i")
					private static WebElement  mrpSettingsBtn;
					
	//Point Of Sale ***********************************				
		
		@FindBy(xpath="//*[@id='id_menu_tree_8000']/a/i")
		private static WebElement  pointofSale_ExpandBtn;
		
					@FindBy(xpath="//*[@id='id_menu_tree_8004']/a/i")
					private static WebElement  mainScreenBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_8500']/a/i")
					private static WebElement  posReportsExpandBtn;
					
								@FindBy(xpath="//*[@id='id_menu_tree_8501']/a/i")
								private static WebElement  retailSaleReportBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_8507']/a/i")
								private static WebElement  summarysalesReportBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_8513']/a/i")
								private static WebElement  preOrderReportsBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_8523']/a/i")
								private static WebElement  memberReportsBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_8530']/a/i")
								private static WebElement  memberPointsReportsBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_8535']/a/i")
								private static WebElement  itemReportsBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_8540']/a/i")
								private static WebElement  cashReportsBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_8545']/a/i")
								private static WebElement  stockReportsBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_8550']/a/i")
								private static WebElement  periodicSalesReportsBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_8556']/a/i")
								private static WebElement  lastYearComaprisionReportsBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_8562']/a/i")
								private static WebElement  miscellaneousReportsBtn;
								
								@FindBy(xpath="//*[@id='id_menu_tree_8571']/a/i")
								private static WebElement  restaurantReportsBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_8900']/a/i")
					private static WebElement  reportsTableManagementBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_8003']/a/i")
					private static WebElement  utilitiesExpandBtn;
					
								@FindBy(xpath="//*[@id='id_menu_tree_8027']/a/i")
								private static WebElement  bbqConversionBtn;
		
		
	//Quality Control *****************  	
		
		
		
		@FindBy(xpath="//*[@id='id_menu_tree_150']/a/i")
		private static WebElement  qualityControlExpandBtn;
		
					@FindBy(xpath="//*[@id='id_menu_tree_151']/a/i")
					private static WebElement  qualityControlTransactionBtn;
					
					@FindBy(xpath="//*[@id='id_menu_tree_152']/a/i")
					private static WebElement  qualityControlReportsBtn;


	    
	 // Views Tab
		
	    
		@FindBy(xpath="//span[@class='icon-views'")
		private static WebElement  viewsTab;
		

		
	// Export Fields 
		
		@FindBy(xpath="//span[contains(text(),'Export Fields')]")
		private static WebElement exportFieldsTab;
		
	// Triggers
		
		@FindBy(xpath="//span[@class='icon-trigger icon-font7']")
		private static WebElement  triggersTab;	
		
	//  Document Numbering
		
		@FindBy(xpath="//span[contains(text(),'Document Numbering')]")
		private static WebElement  documentNumberingTab;
		
	// Hire Purchase	
		
		@FindBy(xpath="//span[contains(text(),'Hire Purchase')]")
		private static WebElement hirePurchasetab;
		

		
	// Reports Tab	
		
		@FindBy(xpath="//span[@class='icon-text70 icon-font7']")
		private static WebElement  reportsTab;	
		
	// Schemes

		@FindBy(xpath="//span[contains(text(),'Schemes')]")
		private static WebElement schemestab;	
			
	    
	    
		
		
	////////////////////////////////////////////////////////////////////////////////
		
		
		//-*-------------------------------------------------------------------------------------------------------------	
			
		// Miscellaneous Tab
						
			@FindBy(xpath="//span[@class='icon-miscellaneous']")
			private static WebElement  miscellaneousTab;
			
			@FindBy(xpath="//input[@id='misc_Abbrivation']")
			private static WebElement  abbreviationTxt;
			
			@FindBy(xpath="//input[@id='misc_defaultAcc']")
			private static WebElement  defaultAccount1Txt;
			
			@FindBy(xpath="//*[@id='misc_defaultAcc_input_image']/span")
			private static WebElement  defaultAccount1Dropdown;
			
			@FindBy(xpath="//*[@id='misc_defaultAcc_input_settings']/span")
			private static WebElement  defaultAccount1SettingBtn;
			
			@FindBy(xpath="//div[@id='misc_defaultAcc_customize_popup_container']")
			private static WebElement  defaultAccount1ContainerTxt;
			
			@FindBy(xpath="//div[@id='misc_defaultAcc_customize_popup_footer']//input[1]")
			private static WebElement  defaultAccount1StandardfieldsBtn;
			
						@FindBy(xpath="//select[@id='misc_defaultAcc_customize_popup_standardfields_list']")
						private static WebElement  defaultAccount1FieldsDropdown;
						
						@FindBy(xpath="//input[@id='misc_defaultAcc_customize_popup_standardfields_header']")
						private static WebElement  defaultAccount1HeaderTxt;
						
						@FindBy(xpath="//select[@id='misc_defaultAcc_customize_popup_standardfields_alignment']")
						private static WebElement  defaultAccount1AllignmentDropdown;
						
						@FindBy(xpath="//input[@id='misc_defaultAcc_customize_popup_standardfields_width']")
						private static WebElement  defaultAccount1WidthTxt;
						
						@FindBy(xpath="//section[@id='page_Content']//input[3]")
						private static WebElement  defaultAccount1StandardFieldsokBtn;
						
						@FindBy(xpath="//section[@id='page_Content']//input[4]")
						private static WebElement  defaultAccount1StanadardFieldscancelBtn;
			
			@FindBy(xpath="//div[@id='misc_defaultAcc_search_container']//input[2]")
			private static WebElement  defaultAccount1DeleteColumnBtn;
			
			@FindBy(xpath="//section[@id='page_Content']//input[3]")
			private static WebElement  defaultAccount1okBtn;
			
			@FindBy(xpath="//section[@id='page_Content']//input[4]")
			private static WebElement  defaultAccount1cancelBtn;
			
			@FindBy(xpath="//input[@id='misc_defaultAcc2']")
			private static WebElement  defaultAccount2Txt;

			@FindBy(xpath="//*[@id='misc_defaultAcc2_input_image']/span")
			private static WebElement  defaultAccount2Dropdown;
			
			@FindBy(xpath="//*[@id='misc_defaultAcc2_input_settings']/span")
			private static WebElement  defaultAccount2SettingBtn;
			
			@FindBy(xpath="//div[@id='misc_defaultAcc2_customize_popup_container']")
			private static WebElement  defaultAccount2ContainerTxt;
			
			@FindBy(xpath="//div[@id='misc_defaultAcc2_customize_popup_footer']//input[1]")
			private static WebElement  defaultAccount2StandardfieldsBtn;
			
						@FindBy(xpath="//select[@id='misc_defaultAcc2_customize_popup_standardfields_list']")
						private static WebElement  defaultAccount2FieldsDropdown;
						
						@FindBy(xpath="//input[@id='misc_defaultAcc2_customize_popup_standardfields_header']")
						private static WebElement  defaultAccount2HeaderTxt;
						
						@FindBy(xpath="//select[@id='misc_defaultAcc2_customize_popup_standardfields_alignment']")
						private static WebElement  defaultAccount2AllignmentDropdown;
						
						@FindBy(xpath="//input[@id='misc_defaultAcc2_customize_popup_standardfields_width']")
						private static WebElement  defaultAccount2WidthTxt;
						
						@FindBy(xpath="//section[@id='page_Content']//input[3]")
						private static WebElement  defaultAccount2StandardFieldsokBtn;
						
						@FindBy(xpath="//section[@id='page_Content']//input[4]")
						private static WebElement  defaultAccount2StanadardFieldscancelBtn;
			
			@FindBy(xpath="//div[@id='misc_defaultAcc2_search_container']//input[2]")
			private static WebElement  defaultAccount2DeleteColumnBtn;
			
			@FindBy(xpath="//section[@id='page_Content']//input[3]")
			private static WebElement  defaultAccount2okBtn;
			
			@FindBy(xpath="//section[@id='page_Content']//input[4]")
			private static WebElement  defaultAccount2cancelBtn;
			
			@FindBy(xpath="//input[@id='misc_chkallowDeptAppropriation']")
			private static WebElement  allowDeptAppropriationChkBox;
			
			@FindBy(xpath="//input[@id='misc_chkTransDateWhnAuthzed']")
			private static WebElement  changeTransactionDateWhenAuthorizedChkBox;
			
			@FindBy(xpath="//input[@id='misc_chkdateCntbePriorPrevDocDate']")
			private static WebElement  dateCannotBePriortoPrevdateChkBox;
			
			@FindBy(xpath="//input[@id='misc_chkdetailedDisinSummryBar']")
			private static WebElement  detailedDisplayInSummaryBarChkBox;
			
			@FindBy(xpath="//input[@id='misc_chkdontAllowEditThisDoc']")
			private static WebElement  dontAllowEditThisDocChkBox;
			
			@FindBy(xpath="//input[@id='misc_chkDontAllowInputDate']")
			private static WebElement  dontAllowInputDateChkBox;
			
			@FindBy(xpath="//input[@id='misc_chkdontRestrictAcc1Type']")
			private static WebElement  dontRestrictAccount1TypeChkBox;
			
			@FindBy(xpath="//input[@id='misc_chkdontRestrictAcc2Type']")
			private static WebElement  dontRestrictAccount2TypeChkBox;
			
			@FindBy(xpath="//input[@id='misc_chkdontRestrictonLoginAcc']")
			private static WebElement  dontRestrictOnLoginAccountChkBox;
			
			@FindBy(xpath="//input[@id='misc_chkdontRestrictLoginTag']")
			private static WebElement  dontRestrictOnLoginTagChkBox;
			
			@FindBy(xpath="//input[@id='misc_chkenableVersion']")
			private static WebElement  enableVersionChkBox;
			
			@FindBy(xpath="//input[@id='misc_chkPrintUsingDefaultPrintFormat']")
			private static WebElement  printUsingDefaultPrinterAndFormatChkBox;
			
			@FindBy(xpath="//select[@id='misc_chkCheckBudget']")
			private static WebElement  checkbudgetDropdown;
			
			@FindBy(xpath="//select[@id='misc_PickEmailFrom']")
			private static WebElement  pickEmailFromDropdown;
			
			// Restrictions Sections
			
			@FindBy(xpath="//input[@id='misc_restrictionRdStop']")
			private static WebElement  stopRadio;
			
			@FindBy(xpath="//input[@id='misc_restrictionRdwarnandAllow']")
			private static WebElement  warnAndAllowRadio;
			
			@FindBy(xpath="//input[@id='misc_restrictionentryCondition_textbox']")
			private static WebElement  entryRestrictConditionTxt;
			
						@FindBy(xpath="//input[@id='misc_restrictionentryCondition_formulaText']")
						private static WebElement  miscformulaTxt;
						
						@FindBy(xpath="//div[@id='misc_restrictionentryCondition_expandedFormula']")
						private static WebElement  miscexpandeFormula;
						
						 // Grid Elements 
						
						@FindBy(xpath="//span[@class='icon-collepse icon-font8']")
						private static WebElement  miscTransactionExpandBtn;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_2-2']")
						private static WebElement  miscQuantity;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_2-3']")
						private static WebElement  miscqty;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_3-2']")
						private static WebElement  miscRate;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_3-3']")
						private static WebElement  miscrt;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_4-2']")
						private static WebElement  miscAlternateQty;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_4-3']")
						private static WebElement  miscqtya;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_5-2']")
						private static WebElement  miscGrossAmount;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_5-3']")
						private static WebElement  miscgr;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_6-2']")
						private static WebElement  miscStockValue;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_6-3']")
						private static WebElement  miscstkv;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_7-2']")
						private static WebElement  miscAccount;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_7-3']")
						private static WebElement  miscacc;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_8-2']")
						private static WebElement  miscAccount2;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_8-3']")
						private static WebElement  miscacc2;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_9-2']")
						private static WebElement  miscExchangeRate;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_9-3']")
						private static WebElement  miscexrt;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_10-2']")
						private static WebElement  miscQtyinBaseUnit;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_10-3']")
						private static WebElement  miscqtyb;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_11-2']")
						private static WebElement  miscCurrency;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_11-3']")
						private static WebElement  misccrn;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_12-2']")
						private static WebElement  miscValueOfThetagMasterIDofTheTag;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_12-3']")
						private static WebElement  misctgval;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_13-2']")
						private static WebElement  miscNetAmount;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_13-3']")
						private static WebElement  miscnet;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_14-2']")
						private static WebElement  miscNetAmountInOriginalCurrency;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_14-3']")
						private static WebElement  miscneto;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_15-2']")
						private static WebElement  miscTotalQuantity;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_15-3']")
						private static WebElement  misctqty;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_16-2']")
						private static WebElement  miscTotalGross;
						
						@FindBy(xpath="//td[@id='formulaGridmisc_restrictionentryCondition_col_16-3']")
						private static WebElement  misctgr;
						
						// grid completed
						
						
						@FindBy(xpath="//input[@id='misc_restrictionentryCondition_availableVariables']")
						private static WebElement  searchTxt;
						
						@FindBy(xpath="//button[@id='misc_restrictionentryCondition_PlusOpt']")
						private static WebElement  plusBtn;
						
						@FindBy(xpath="//button[contains(text(),'-')]")
						private static WebElement  minusBtn;
						
						@FindBy(xpath="//button[contains(text(),'/')]")
						private static WebElement  divideBtn;
						
						@FindBy(xpath="//button[contains(text(),'*')]")
						private static WebElement  multiplyBtn;
				
						@FindBy(xpath="//button[5]")
						private static WebElement  brackets1Btn;
						
						@FindBy(xpath="//button[6]")
						private static WebElement  brackets2Btn;
						
						@FindBy(xpath="//button[contains(text(),'~')]")
						private static WebElement  tildeBtn;
						
						@FindBy(xpath="//button[contains(text(),'&')]")
						private static WebElement  andBtn;
						
						@FindBy(xpath="//button[contains(text(),'%')]")
						private static WebElement  percentBtn;
						
						@FindBy(xpath="//button[contains(text(),']')]")
						private static WebElement  squarebracketBtn;
						
						@FindBy(xpath="//button[contains(text(),'|')]")
						private static WebElement  verticalBarBtn;
						
						@FindBy(xpath="//button[contains(text(),'\')]")
						private static WebElement  backSlashBtn;
						
						@FindBy(xpath="//button[contains(text(),'<')]")
						private static WebElement  lessThanbtn;
						
						@FindBy(xpath="//button[contains(text(),'>')]")
						private static WebElement  greaterThanBtn;
						
						@FindBy(xpath="//button[contains(text(),'!')]")
						private static WebElement  exclamationBtn;
						
						@FindBy(xpath="//button[contains(text(),'=')]")
						private static WebElement  equalBtn;
						
						@FindBy(xpath="//button[contains(text(),'?:')]")
						private static WebElement  questionBtn;
					
						@FindBy(xpath="//button[contains(text(),'^')]")
						private static WebElement  caretBtn;
						
						@FindBy(xpath="//button[contains(text(),'DateDiff()')]")
						private static WebElement  dateDiffBtn;
						
						@FindBy(xpath="//button[contains(text(),'Cond()')]")
						private static WebElement  condBtn;
						
						@FindBy(xpath="//button[contains(text(),'Min()')]")
						private static WebElement  minBtn;
						
						@FindBy(xpath="//button[contains(text(),'Max()')]")
						private static WebElement  maxBtn;
						
						@FindBy(xpath="//button[@id='misc_restrictionentryCondition_Ok']")
						private static WebElement  entryRestrictConditionokBtn;
						
						@FindBy(xpath="//button[@id='misc_restrictionentryCondition_Cancel']")
						private static WebElement  entryRestrictConditioncancelBtn;
						
			@FindBy(xpath="//input[@id='misc_restrictionEntryMessage']")
			private static WebElement  entryRestrictMessageTxt;
			
		// Apply Rules On	
			
			@FindBy(xpath="//input[@id='misc_chkRulesExtMod']")
			private static WebElement applyRulesOnExternalModuleChkbox;
			
			@FindBy(xpath="//input[@id='misc_chkRulesLink']")
			private static WebElement applyRulesOnLinkChkBox;
			
			@FindBy(xpath="//input[@id='misc_chkRulesItems']")
			private static WebElement applyRulesOnSetItemsChkBox;
			
//			Apply Dependency On
			
			@FindBy(xpath="//input[@id='misc_chkdependencyExtMod']")
			private static WebElement applyDependencyOnExternalModuleChkbox;
			
			@FindBy(xpath="//input[@id='misc_chkdependencyLink']")
			private static WebElement applyDependencyOnLinkChkBox;
			
			@FindBy(xpath="//input[@id='misc_chkdependencyItems']")
			private static WebElement applyDependencyOnSetItemsChkBox;
			
		// Posting Details	
			
			@FindBy(xpath="//input[@id='misc_PostingUpdateFA']")
			private static WebElement updateFAChkBox;
			
			@FindBy(xpath="//input[@id='misc_PostingSelectAccInLine']")
			private static WebElement selectAmountInEveryLineChkBox;
			
			@FindBy(xpath="//input[@id='misc_chkDetailedPosting']")
			private static WebElement detailedPostingChkbox;
			
			@FindBy(xpath="//input[@id='misc_chkUpdateStock']")
			private static WebElement updateStockChkBox;
			
			@FindBy(xpath="//input[@id='misc_chkFixUpdateStock']")
			private static WebElement allowChangingUpdateStock;
			
			@FindBy(xpath="//input[@id='misc_PostingPostServiceTax']")
			private static WebElement postServicetaxChkBox;
			
			@FindBy(xpath="//input[@id='misc_PostingPostTDS']")
			private static WebElement postTDSchkBox;
			
			@FindBy(xpath="//input[@id='misc_PostingPostVAT']")
			private static WebElement postVATchkBox;
			
			@FindBy(xpath="//input[@id='misc_PostingLinkToLetterOfCredit']")
			private static WebElement linkToLetterOfCreditChkBox;
			
			@FindBy(xpath="//input[@id='misc_Dontcheckcreditlimit']")
			private static WebElement dontcheckcreditlimitChkBox;
			
		// Currency	
			
			@FindBy(xpath="//select[@id='misc_CurrencyAddCurrencyIn']")
			private static WebElement addCurrencyInDropdown;
			
			@FindBy(xpath="//input[@id='misc_currencyInputExchangeRate']")
			private static WebElement inputExchangeRateChkBox;
			
			@FindBy(xpath="//input[@id='misc_currencyPostExgRateDiff']")
			private static WebElement postExchangeRateChkDifferrenceBox;
			
			@FindBy(xpath="//input[@id='misc_currencyPostExchgRateBasedOnStd']")
			private static WebElement postExchangeRateDifferenceBasedOnStandardRuleChkBox;
			
		// AR/AP	
			
			
			@FindBy(xpath="//select[@id='misc_arapDueDate']")
			private static WebElement dueDateDropdown;
			
			@FindBy(xpath="//select[@id='misc_arapARAPOptions']")
			private static WebElement arapOptionsDropdown;
			
			@FindBy(xpath="//select[@id='misc_arapOptions']")
			private static WebElement optionsdDropdown;
					       
			@FindBy(xpath="//input[@id='misc_arapchkSelectbillsbeforeamt']")
			private static WebElement selectbillsbeforeAmountChkBox;
			
			@FindBy(xpath="//input[@id='misc_arapchkPostdiscountbasedoncrdays']")
			private static WebElement postDiscountBasedOnCrediDaysChkBox;
			
			@FindBy(xpath="//input[@id='misc_arapchkInputPaymentTerms']")
			private static WebElement inputPaymenttermsChkBox;
			
			@FindBy(xpath="//select[@id='misc_arapNarration']")
			private static WebElement narrationDropdown;
				

	//**----------------------------------------------------------------------------------------------------------    
	// Edit Screen Tab     
		    
		    @FindBy(xpath="//span[contains(text(),'Edit Screen')]")
		   	private static WebElement  editScreenTab; 
			
			@FindBy(xpath="//i[@class='icon-add icon-font7']")
			private static WebElement editScreenaddBtn;

			@FindBy(xpath="//a[contains(text(),'Field Details')]")
			private static WebElement editScreenFieldDetailsTab;
	   
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_Caption']")
			private static WebElement  editScreenCaptionTxt;

			@FindBy(xpath="//select[@id='editScreen_FieldsCustomization_DataType']")
			private static WebElement  editScreenCaptionDataTypeDropdown;
			
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_DefaultValue']")
			private static WebElement editScreenDefaultValueTxt;
		   
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_noofDecimals']")
			private static WebElement editScreenNoOfDecimalsTxt;
		   
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_toolTip']")
			private static WebElement editScreenToolTipTxt;
		   
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_bannerText']")
			private static WebElement editScreenBannerTextTxt;
		   
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_MinValue']")
			private static WebElement editScreenMinimunValueTxt;
		   
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_MaxValue']")
			private static WebElement editScreenMaximunValueTxt;
		   
			@FindBy(xpath="//select[@id='editScreen_FieldsCustomization_behaviour']")
			private static WebElement editScreenAddFieldsBehaviourDropdown;
			
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_Formula_textbox']")
			private static WebElement  editScreenAddBehaviourformulaTxt;
			
			@FindBy(xpath="//button[@id='editScreen_FieldsCustomization_Formula_Ok']")
			private static WebElement  editScreenAddBehaviourFormulaokBtn;
		   
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_PreLoad_textbox']")
			private static WebElement editScreenPreloadTxt;
		   
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_PreLoad_textbox']")
			private static WebElement editScreenRestrictformulaTxt;
		   
			@FindBy(xpath="//select[@id='editScreen_FieldsCustomization_RoundOff']")
			private static WebElement editScreenRoundOffsDropdown;
		   
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_RoundOffTo']")
			private static WebElement editScreenRoundOffToTxt;
		   
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_RestrictMsg']")
			private static WebElement editScreenRestrictMessageTxt;
		   
			@FindBy(xpath="//select[@id='editScreen_FieldsCustomization_AddToNet']")
			private static WebElement editScreenAddToNetDropdown;
		   
			@FindBy(xpath="//select[@id='editScreen_FieldsCustomization_AddToStock']")
			private static WebElement EditScreenAddToStockDropDown;
		
			@FindBy(xpath="//select[@id='editScreen_FieldsCustomization_Position']")
			private static WebElement editScreenpositionDropdown;  
		   
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_Columnwidth']")
			private static WebElement editScreenColumnWidthTxt;  
		   
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_chkValinBaseCurency']")
			private static WebElement editScreenValueInBaseCurrencyChkBox;  
		   
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_chkValinBaseCurency']")
			private static WebElement editScreenHideFromSummaryChkBox;  
		   
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_chkPostToAcc']")
			private static WebElement editScreenPostToAccountChkBox;  
		 
			@FindBy(xpath="//input[@id='acc1-1']")
			private static WebElement editScreenAccount1Radio;  
		   
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_defaultAcc1']")
			private static WebElement editScreenAccount1Txt;  
		   
			@FindBy(xpath="//*[@id='editScreen_FieldsCustomization_defaultAcc1_input_image']/span")
			private static WebElement editScreenAccount1ExpandBtn;  
		   
			@FindBy(xpath="//*[@id='editScreen_FieldsCustomization_defaultAcc1_input_settings']/span")
			private static WebElement editScreenAccount1SettingsBtn;  
		   
					    @FindBy(xpath="//div[@id='editScreen_FieldsCustomization_defaultAcc1_customize_popup_container']")
						private static WebElement editScreenAccount1Container;  
					    
					    @FindBy(xpath="//div[@id='editScreen_FieldsCustomization_defaultAcc1_customize_popup_footer']//input[1]")
						private static WebElement editScreenAccount1StandardFiledsBtn;  
					    
								    @FindBy(xpath="//select[@id='editScreen_FieldsCustomization_defaultAcc1_customize_popup_standardfields_list']")
									private static WebElement editScreenAccount1StandardFiledsDropdown;  
								    
								    @FindBy(xpath="//input[@id='editScreen_FieldsCustomization_defaultAcc1_customize_popup_standardfields_header']")
									private static WebElement editScreenAccount1StandardFiledsHeaderTxt;  
								    
								    @FindBy(xpath="//select[@id='editScreen_FieldsCustomization_defaultAcc1_customize_popup_standardfields_alignment']")
									private static WebElement editScreenAccount1StandardFiledsAllignmentDropdown; 
								    
								    @FindBy(xpath="//input[@id='editScreen_FieldsCustomization_defaultAcc1_customize_popup_standardfields_width']")
									private static WebElement editScreenAccount1StandardFiledsWidthTxt;  
								    
								    @FindBy(xpath="//section[@id='page_Content']//input[3]")
									private static WebElement editScreenAccount1StandardFiledsokBtn;  
								    
								    @FindBy(xpath="//section[@id='page_Content']//input[4]")
									private static WebElement editScreenAccount1StandardFiledscancelBtn;  
					    
					    @FindBy(xpath="//div[@id='editScreen_FieldsCustomization_defaultAcc1_search_container']//input[2]")
						private static WebElement editScreenAccount1deleteColumnBtn;
					    
					    @FindBy(xpath="//section[@id='page_Content']//input[3]")
					   	private static WebElement editScreenAccount1okBtn;  
					       
					       @FindBy(xpath="//section[@id='page_Content']//input[4]")
					   	private static WebElement editScreenAccount1cancelBtn;
		
		
	        @FindBy(xpath="//input[@id='chkPostToAccountAllowEntry']")
			private static WebElement editScreenAccount1AllowEntryChkBox;
	        
			@FindBy(xpath="//input[@id='acc2-2']")
			private static WebElement editScreenAccount2Radio;  
			    
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_defaultAcc2']")
			private static WebElement editScreenAccount2Txt;  
			    
			@FindBy(xpath="//*[@id='editScreen_FieldsCustomization_defaultAcc2_input_image']/span")
			private static WebElement editScreenAccount2ExpandBtn;
		  
		   @FindBy(xpath="//*[@id='editScreen_FieldsCustomization_defaultAcc2_input_settings']/span")
		  	private static WebElement editScreenAccount2SettingsBtn;  
		      
		  			    @FindBy(xpath="//div[@id='editScreen_FieldsCustomization_defaultAcc2_customize_popup_container']")
		  				private static WebElement editScreenAccount2Container;  
		  			    
		  			    @FindBy(xpath="//div[@id='editScreen_FieldsCustomization_defaultAcc2_customize_popup_footer']//input[1]")
		  				private static WebElement editScreenAccount2StandardFiledsBtn;  
		  			    
		  						    @FindBy(xpath="//select[@id='editScreen_FieldsCustomization_defaultAcc2_customize_popup_standardfields_list']")
		  							private static WebElement editScreenAccount2StandardFiledsDropdown;  
		  						    
		  						    @FindBy(xpath="//input[@id='editScreen_FieldsCustomization_defaultAcc2_customize_popup_standardfields_header']")
		  							private static WebElement editScreenAccount2StandardFiledsHeaderTxt;  
		  						    
		  						    @FindBy(xpath="//select[@id='editScreen_FieldsCustomization_defaultAcc2_customize_popup_standardfields_alignment']")
		  							private static WebElement editScreenAccount2StandardFiledsAllignmentDropdown; 
		  						    
		  						    @FindBy(xpath="//input[@id='editScreen_FieldsCustomization_defaultAcc2_customize_popup_standardfields_width']")
		  							private static WebElement editScreenAccount2StandardFiledsWidthTxt;  
		  						    
		  						    @FindBy(xpath="//section[@id='page_Content']//input[3]")
		  							private static WebElement editScreenAccount2StandardFiledsokBtn;  
		  						    
		  						    @FindBy(xpath="//section[@id='page_Content']//input[4]")
		  							private static WebElement editScreenAccount2StandardFiledscancelBtn;  
		  			    
		  			    @FindBy(xpath="//div[@id='editScreen_FieldsCustomization_defaultAcc2_search_container']//input[2]")
		  				private static WebElement editScreenAccount2deleteColumnBtn;
		  			    
		  			    @FindBy(xpath="//section[@id='page_Content']//input[3]")
		  			   	private static WebElement editScreenAccount2okBtn;  
		  			       
		  			       @FindBy(xpath="//section[@id='page_Content']//input[4]")
		  			   	private static WebElement editScreenAccount2cancelBtn;
		   
		   @FindBy(xpath="//input[@id='chkPostToAccount1AllowEntry']")
		   private static WebElement editScreenAccount2AllowEntryChkBox;
		  			       
		   @FindBy(xpath="//input[@id='acc1-3']")
		   private static WebElement editScreenAccount1VariableRadio; 
		  
		   @FindBy(xpath="//input[@id='Account_1Var']")
		   private static WebElement editScreenAccount1VariableTxt;
		   
		   @FindBy(xpath="//input[@id='acc2-4']")
		   private static WebElement editScreenAccount2VariableRadio; 
		  
		   @FindBy(xpath="//input[@id='Account_2Var']")
		   private static WebElement editScreenAccount2VariableTxt;
		   
		   @FindBy(xpath="//input[@id='editScreen_FieldsCustomization_RegularExpr']")
		   private static WebElement editScreenRegularExpTxt; 
		   
		   @FindBy(xpath="//input[@id='editScreen_FieldsCustomization_ErrorMsg']")
		   private static WebElement editScreenerrorMessageTxt; 
		
		 // properties  
		   
		  @FindBy(xpath="//a[contains(text(),'Properties')]")
		  private static WebElement editScreenPropertiesTab;  
		     
		  @FindBy(xpath="//input[@id='editScreen_FieldsCustomization_chkAuditTrail']")
		  private static WebElement editScreenauditTrailChkBox;  
		     
		  @FindBy(xpath="//input[@id='editScreen_FieldsCustomization_chkCantExprt']")
		  private static WebElement editScreencantExportChkBox;  
		  
		  @FindBy(xpath="//input[@id='editScreen_FieldsCustomization_chkCantImprt']")
		  private static WebElement editScreencantImportChkbox; 
		     
		  @FindBy(xpath="//input[@id='editScreen_FieldsCustomization_chkHidden']")
		  private static WebElement editScreenhiddenChkbox;  
		  
		  @FindBy(xpath="//input[@id='editScreen_FieldsCustomization_chkMandatory']")
		  private static WebElement editScreenmandatorychkBox; 
		     
		  @FindBy(xpath="//input[@id='editScreen_FieldsCustomization_chkNotAvailForReports']")
		  private static WebElement editScreennotAvailableForReportsChkBox;  
		     
		  @FindBy(xpath="//input[@id='editScreen_FieldsCustomization_chkReadOnly']")
		  private static WebElement editScreenreadOnlyChkBox;  
		 
		     
		  @FindBy(xpath="//*[@id='editScreen_FieldsCustomization_btns']/span[1]")
		  private static WebElement editScreenApplyBtn;  
		     
		  @FindBy(xpath="//span[@id='editScreen_FieldsCustomization_Close']")
		  private static WebElement editScreencloseBtn;  
		  
		
		 
		  @FindBy(xpath="//a[@href='#editScreen_FieldsCustomization_ExternalModules']")
		  private static WebElement editScreenExternalModulesTab; 
		  
		  @FindBy(xpath="//select[@id='cmbOnEvent_VT_769_7']")
		  private static WebElement editScreenOnEventDrpDwn; 
		  
		  @FindBy(xpath="//select[@id='cmbModuleType_VT_769_7']")
		  private static WebElement editScreenModuleTypeDrpDwn; 
		  
		  @FindBy(xpath="//select[@id='cmbClassname_VT_769_7']")
		  private static WebElement editScreenClassNameDrpDwn; 
		  
		  @FindBy(xpath="//input[@id='txtModuleName_VT_769_7']")
		  private static WebElement editScreenModuleNameTxt; 
		  
		  @FindBy(xpath="//select[@id='cmbFunctionName_VT_769_7']")
		  private static WebElement editScreenFunctionNameDrpDwn; 
		  
		  @FindBy(xpath="//input[@id='btnInsert_VT_769_7']")
		  private static WebElement editScreenInsertBtn; 
		  
		  @FindBy(xpath="//input[@id='btnRemove_VT_769_7']")
		  private static WebElement editScreenRemoveBtn; 
		  
		  
		// External Module Tab is named As EFES And ELV For Each Screen (ReceiptsEFES,ReceiptsELV )
		// External Functions(External Fields) Changes From Screen to Screen
		
		  
		  // Formatting Tab
		  
		  @FindBy(xpath="//*[@id='editScreen_FieldsCustomization_tabs']/li[3]/a")
		    private static WebElement editScreenformattingTab;

			@FindBy(xpath="//*[@id='fontFamily_editScreen_FieldsCustomization_FontCtrl']")
		    private static WebElement editScreenFontFamilyDropdown;
			
		    @FindBy(xpath="//*[@id='fontStyle_editScreen_FieldsCustomization_FontCtrl']")
		    private static WebElement editScreenFontStyleDropdown;

		    @FindBy(xpath="//*[@id='fontWeight_editScreen_FieldsCustomization_FontCtrl']")
		    private static WebElement editScreenFontWeightDropdown;

			@FindBy(xpath="//*[@id='fontSizes_editScreen_FieldsCustomization_FontCtrl']")
			private static WebElement editScreenFontSizeDropdown;
			
			@FindBy(xpath="//*[@id='fontForeColor_editScreen_FieldsCustomization_FontCtrl']")
			private static WebElement editScreenForeColourDropdown;
			
			@FindBy(xpath="//*[@id='fontBackColor_editScreen_FieldsCustomization_FontCtrl']")
			private static WebElement editScreenBackColourDropdown;
			
			@FindBy(xpath="//*[@id='chkBaseline_editScreen_FieldsCustomization_FontCtrl']")
			private static WebElement editScreenBaseLineCheckbox;
			
			@FindBy(xpath="//*[@id='chkOverLine_editScreen_FieldsCustomization_FontCtrl']")
			private static WebElement editScreenOverLineCheckbox;
			
			@FindBy(xpath="//*[@id='chkStrikeThrough_editScreen_FieldsCustomization_FontCtrl']")
			private static WebElement EditScreenStrikeThroughCheckbox;
			
			@FindBy(xpath="//*[@id='chkUnderline_editScreen_FieldsCustomization_FontCtrl']")
			private static WebElement editScreenUnderlineCheckbox;
			
			@FindBy(xpath="//*[@id='previewfont_editScreen_FieldsCustomization_FontCtrl']")
			private static WebElement editScreenFontPreviewTxt;
			
			@FindBy(xpath="//*[@id='editScreen_FieldsCustomization_Formatting_chkenableFont']")
			private static WebElement editScreenenableFontChkbox;
		
			
			@FindBy(xpath="//div[@id='EditScreen']//div[@id='editScreen_CustomizeButtons']//span[2]")
			private static WebElement editScreenLoadFieldsBtn;
			
			
			
		  /*@FindBy(xpath="//i[@class='icon-copyfields icon-font7']")
		  private static WebElement editScreenLoadFieldsBtn;*/
		  
		  @FindBy(xpath="//span[@class='dropdown-toggle theme_button_color']")
		  private static WebElement editScreenCustomizeGridBtn;
		     
		  @FindBy(xpath="//select[@id='editScreen_voucherDropDown']")
		  private static WebElement editScreenLoadFieldsDropdown;  
		     
		  @FindBy(xpath="//*[@id='editScreen_copyFieldsDiv']/div[3]/button[1]")
		  private static WebElement editScreenselectAllBtn;  
		  
		  @FindBy(xpath="//div[@id='editScreen_copyFieldsDiv']//button[@class='Fbutton pull-right'][contains(text(),'Ok')]")
		  private static WebElement editScreenokBtn;  
		     
		  @FindBy(xpath="//input[@id='searchlayout2']")
		  private static WebElement editScreenSearchTxt;  
		     
		  @FindBy(xpath="//button[@id='btnEditScreenPrev']")
		  private static WebElement editScreenPreviousBtn;  
		     
		  @FindBy(xpath="//button[@id='btnEditScreenNext']")
		  private static WebElement editScreenNextBtn;  
		  
		  @FindBy(xpath="//*[@id='editScreen_tabContent']/label")
			private static WebElement editScreenRatePreloadLabel;
		  
		  @FindBy(xpath="//select[@id='cmbBehaviour']")
		  private static WebElement editScreenBehaviourDropdown;
		  
		  @FindBy(xpath="//input[@id='editScreen_qtyPreLoaded_textbox']")
		  private static WebElement editScreenQuantityPreloadedTxt;
		  
		  
		  @FindBy(xpath="//input[@id='editScreen_qtyPreLoaded_formulaText']")
		  private static WebElement  formulaTxt;
			
		  @FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_expandedFormula']")
		  private static WebElement  expandeFormula;
			
		 // Grid Elements 
		
		@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_1-2']//span[@class='icon-collepse icon-font8']")
		private static WebElement  editScreenTransactionExpandBtn;
		
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_2-2']")
					private static WebElement  editScreenQuantity;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_2-3']")
					private static WebElement  editScreenqty;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_3-2']")
					private static WebElement  editScreenRate;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_3-3']")
					private static WebElement  editScreenrt;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_4-2']")
					private static WebElement  editScreenAlternateQty;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_4-3']")
					private static WebElement  editScreenqtya;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_5-2']")
					private static WebElement  editScreengrossAmount;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_5-3']")
					private static WebElement  editScreengr;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_6-2']")
					private static WebElement  editScreenStockValue;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_6-3']")
					private static WebElement  editScreenstkv;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_7-2']")
					private static WebElement  editScreenAccount;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_7-3']")
					private static WebElement  EditScreenacc;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_8-2']")
					private static WebElement  editScreenAccount2;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_8-3']")
					private static WebElement  editScreenacc2;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_9-2']")
					private static WebElement  editScreenExchangeRate;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_9-3']")
					private static WebElement  editScreenexrt;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_10-2']")
					private static WebElement  editScreenqtyinBaseUnit;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_10-3']")
					private static WebElement  editScreenqtyb;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_11-2']")
					private static WebElement  editScreenCurrency;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_11-3']")
					private static WebElement  editScreencrn;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_12-2']")
					private static WebElement  editScreenValueOfThetagMasterIDofTheTag;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_12-3']")
					private static WebElement  editScreentgval;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_13-2']")
					private static WebElement  editScreenNetAmount;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_13-3']")
					private static WebElement  EditScreennet;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_14-2']")
					private static WebElement  editScreenNetAmountInOriginalCurrency;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_14-3']")
					private static WebElement  editScreenneto;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_15-2']")
					private static WebElement  editScreenTotalQuantity;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_15-3']")
					private static WebElement  editScreentqty;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_16-2']")
					private static WebElement  editScreenTotalGross;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_16-3']")
					private static WebElement  editScreentgr;
					
		@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_17-2']//span[@class='icon-expand icon-font8']")
		private static WebElement  screenFiledExpandBtn;
		
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_18-2']")
					private static WebElement  mrnQtyInput;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_18-3']")
					private static WebElement  sb1;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_19-2']")
					private static WebElement  mrnQtyCalculatedValue;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_19-3']")
					private static WebElement  sbo1;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_20-2']")
					private static WebElement  qtyBalancedInput;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_20-3']")
					private static WebElement  sb2;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_21-2']")
					private static WebElement  qtyCalculatedValue;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_21-3']")
					private static WebElement  sbO2;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_22-2']")
					private static WebElement  averagerateInput;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_22-3']")
					private static WebElement  sb3;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_23-2']")
					private static WebElement  avreageRateCalculatedValue;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_23-3']")
					private static WebElement  sbO3;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_24-2']")
					private static WebElement  addValueInput;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_24-3']")
					private static WebElement  sb4;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_25-2']")
					private static WebElement  addValueInputCalculatedValue;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_25-3']")
					private static WebElement  sbO4;
		
		@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_26-2']//span[@class='icon-collepse icon-font8']")
		private static WebElement  layoutFieldExpandBtn;
		
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_27-2']")
					private static WebElement  appropriateBasedOnInput;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_27-3']")
					private static WebElement  lh0;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_28-2']")
					private static WebElement  additionalValueInput;
					
					@FindBy(xpath="//td[@id='formulaGrideditScreen_qtyPreLoaded_col_28-3']")
					private static WebElement  lh1;
					
		// grid completed
		
		
		@FindBy(xpath="//input[@id='editScreen_qtyPreLoaded_availableVariables']")
		private static WebElement  editScreenqtyFormulaSearchTxt;
		
		@FindBy(xpath="//button[@id='editScreen_qtyPreLoaded_PlusOpt']")
		private static WebElement  editScreenqtyPlusBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'-')]")
		private static WebElement  editScreenqtyMinusBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded']//button[3]")
		private static WebElement  editScreenqtyDivideBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'*')]")
		private static WebElement  editScreenqtyMultiplyBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded']//button[5]")
		private static WebElement  editScreenqtyBrackets1Btn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded']//button[6]")
		private static WebElement  editScreenqtyBrackets2Btn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'~')]")
		private static WebElement  editScreenqtyTildeBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'&')]")
		private static WebElement  editScreenqtyAndBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'%')]")
		private static WebElement  editScreenqtyPercentBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//div[@id='operationButtons']//button[@class='btn btn-link'][contains(text(),']')]")
		private static WebElement  editScreenqtySquarebracketBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'|')]")
		private static WebElement  editScreenqtyVerticalBarBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'\')]")
		private static WebElement  editScreenqtyBackSlashBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'<')]")
		private static WebElement  editScreenqtyLessThanbtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'>')]")
		private static WebElement  editScreenqtyGreaterThanBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'=')]")
		private static WebElement  editScreenqtyExclamationBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'=')]")
		private static WebElement  editScreenqtyEqualBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'?:')]")
		private static WebElement  editScreenqtyQuestionBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'^')]")
		private static WebElement  editScreenqtyCaretBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'DateDiff()')]")
		private static WebElement  editScreenqtyDateDiffBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'Cond()')]")
		private static WebElement  editScreenqtyCondBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'Min()')]")
		private static WebElement  editScreenqtyMinBtn;
		
		@FindBy(xpath="//div[@id='editScreen_qtyPreLoaded_formulaDiv']//button[@class='btn btn-link'][contains(text(),'Max()')]")
		private static WebElement  editScreenqtyMaxBtn;
		
		@FindBy(xpath="//button[@id='editScreen_qtyPreLoaded_Ok']")
		private static WebElement  editScreenqtyConditionokBtn;
		
		@FindBy(xpath="//button[@id='editScreen_qtyPreLoaded_Cancel']")
		private static WebElement  editScreenqtyConditioncancelBtn;
		
		
		///////////
		@FindBy(xpath="//*[@id='editScreen_qtyFormula_textbox']")
		private static WebElement  editScreenformulaTxt;
		
		
		// CustomizeGrid
		
		@FindBy(xpath="//i[@class='icon-custamize icon-font7']")
		private static WebElement editScreencustomizeGridBtn;  
		  
		@FindBy(xpath="//span[contains(text(),'Apportion')")
		private static WebElement editScreenApportionBtn;
		     
		@FindBy(xpath="//span[contains(text(),'Audit Trial')]")
		private static WebElement editScreenAuditTrialBtn;  
		    
		@FindBy(xpath="//span[contains(text(),'Back color')]")
		private static WebElement editScreenbackColorBtn;  
		     
		@FindBy(xpath="//span[contains(text(),'Banner text')]")
		private static WebElement editScreenbannerTextBtn;  
		 
		  @FindBy(xpath="//span[contains(text(),'Cannot be exported')]")
		  private static WebElement editScreencannotBeExportedBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Cannot be imported')]")
		  private static WebElement editScreencannotBeImportedBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Column Span')]")
		  private static WebElement editScreencolumnSpanBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Copy from Parent')]")
		  private static WebElement editScreencopyFromParentBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'DataType')]")
		  private static WebElement editScreendataTypeBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Default Value')]")
		  private static WebElement editScreenDefaultValueBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Error message')]")
		  private static WebElement editScreenerrorMessageBtn;  
		 
		  @FindBy(xpath="//span[contains(text(),'External module attached')]")
		  private static WebElement editScreenExternalModuleAttachedBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Field Order')]")
		  private static WebElement editScreenFilterOrderBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Font')]")
		  private static WebElement editScreenFontBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Group Name')]")
		  private static WebElement editScreengroupNameBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Hide left panel')]")
		  private static WebElement editScreenHideLeftPanelBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Information Field')]")
		  private static WebElement editScreenInformationFieldBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Mandatory')]")
		  private static WebElement editScreenMandatoryBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Mass Update')]")
		  private static WebElement editScreenMassupdateBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Maximum Value')]")
		  private static WebElement editScreenMaximumValueBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Merge Field')]")
		  private static WebElement editScreenMergerFieldBtn;  
		  
		  @FindBy(xpath="//span[contains(text(),'Minimum Value')]")
		  private static WebElement editScreenMinimumValueBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'No Of Decimals')]")
		  private static WebElement editScreenNoOfDecimalsBtn;  
		  
		  @FindBy(xpath="//span[contains(text(),'Not available for reports')]")
		  private static WebElement editScreenNotAvailableForReportsBtn; 
		     
		  @FindBy(xpath="//span[contains(text(),'Read Only')]")
		  private static WebElement editScreenReadOnlyBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Regular expression')]")
		  private static WebElement editScreenRegularExpressionBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Restrict formula')]")
		  private static WebElement editScreenRestrictFormulaBtn;  
		  
		  @FindBy(xpath="//span[contains(text(),'Restrict Message')]")
		  private static WebElement editScreenRestrictMessageBtn;  
		  
		  @FindBy(xpath="//span[contains(text(),'Round Offs')]")
		  private static WebElement editScreenRoundsOffsBtn;  
		
		  @FindBy(xpath="//span[contains(text(),'Row Span')]")
		  private static WebElement editScreenRowSpanBtn; 
		     
		  @FindBy(xpath="//span[contains(text(),'Rule Attached')]")
		  private static WebElement editScreenRuleAttachedBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Tooltip')]")
		  private static WebElement editScreenToolTipBtn;  
		     
		  @FindBy(xpath="//span[contains(text(),'Variable name')]")
		  private static WebElement editScreenVariableNameBtn;  
		 
		  
		  
		  
		
		 // Body Tab 
		  
		  @FindBy(xpath="//a[contains(text(),'Body')]")
		  private static WebElement  editScreenBodyTab;
		  
		  @FindBy(xpath="//*[@id='editScreen_bodyTable']/tbody/tr")
		  private static List<WebElement> editScreenBodyTabRowsCount;
		  
		  @FindBy(xpath="//*[@id='editScreen_bodyTable']/tbody")
		  private static WebElement editScreenBodyTabTable;
		
		  @FindBy(xpath="//td[@id='editScreen_bodyrow_1 _1']//span[@id='editIcon']")
		  private static WebElement  editScreenBody1stRowEditBtn; 
		  
		  @FindBy(xpath="//td[@id='editScreen_bodyrow_1 _1']//span[@id='deleteIcon']")
		  private static WebElement  editScreenBody1stRowDeleteBtn; 
		  
		  @FindBy(xpath="//td[@id='editScreen_bodyrow_1 _3']")
		  private static WebElement  editScreenBody1stRowCaption;
		  
		  @FindBy(xpath="//td[@id='editScreen_bodyrow_1 _8']")
		  private static WebElement  editScreenBody1stRowFormula;
		  
		  @FindBy(xpath="//td[@id='editScreen_bodyrow_2 _1']//span[@id='editIcon']")
		  private static WebElement  editScreenBody2ndRowEditBtn; 
		  
		  @FindBy(xpath="//td[@id='editScreen_bodyrow_2 _1']//span[@id='deleteIcon']")
		  private static WebElement  editScreenBody2ndRowDeleteBtn; 
		  
		  @FindBy(xpath="//td[@id='editScreen_bodyrow_2 _3']")
		  private static WebElement  editScreenBody2ndRowCaption;
		  
		  @FindBy(xpath="//td[@id='editScreen_bodyrow_2 _8']")
		  private static WebElement  editScreenBody2ndRowFormula;
		  
		  @FindBy(xpath="//div[@id='editScreen_body']//button[@class='icon-up-arrow Fbutton']")
		  private static WebElement  editScreenBodyUpArrowBtn; 
		
		  @FindBy(xpath="//div[@id='editScreen_body']//button[@class='icon-down-arrow Fbutton']")
		  private static WebElement  editScreenBodyDownArrowBtn;
						
		// Footer Tab
		  
		  @FindBy(xpath="//a[contains(text(),'Footer')]")
		  private static WebElement  editScreenFooterTab;  
		  
		  @FindBy(xpath="//*[@id='editScreen_footerTable']/tbody/tr")
		  private static List<WebElement> editScreenFooterTabRowsCount;
		  
		  @FindBy(xpath="//*[@id='editScreen_footerTable']/tbody")
		  private static WebElement editScreenFooterTabTable;
		  
		  @FindBy(xpath="//td[@id='editScreen_footerrow_1 _1']//span[@id='editIcon']")
		  private static WebElement  editScreenFooter1stRowEditBtn; 
		  
		  @FindBy(xpath="//td[@id='editScreen_footerrow_1 _1']//span[@id='deleteIcon']")
		  private static WebElement  editScreenFooter1stRowDeleteBtn; 
		  
		  @FindBy(xpath="//td[@id='editScreen_footerrow_1 _3']")
		  private static WebElement  editScreenFooter1stRowCaption; 
		  
		  @FindBy(xpath="//td[@id='editScreen_footerrow_1 _8']")
		  private static WebElement  editScreenFooter1stRowFormula; 
		  
		  @FindBy(xpath="//td[@id='editScreen_footerrow_2 _1']//span[@id='editIcon']")
		  private static WebElement  editScreenFooter2ndRowEditBtn; 
		  
		  @FindBy(xpath="//td[@id='editScreen_footerrow_2 _1']//span[@id='deleteIcon']")
		  private static WebElement  editScreenFooter2ndRowDeleteBtn; 
		  
		  
		  
		  @FindBy(xpath="//td[@id='editScreen_footerrow_4 _1']//span[@id='deleteIcon']")
		  private static WebElement  editScreenFooter4thRowDeleteBtn; 
		  
		  
		  @FindBy(xpath="//td[@id='editScreen_footerrow_5 _1']//span[@id='deleteIcon']")
		  private static WebElement  editScreenFooter5thRowDeleteBtn; 
		  
		  
		  @FindBy(xpath="//td[@id='editScreen_footerrow_6 _1']//span[@id='deleteIcon']")
		  private static WebElement  editScreenFooter6thRowDeleteBtn; 
		  
		  
		  
		  
		  
		  @FindBy(xpath="//div[@id='editScreen_footer']//button[@class='icon-up-arrow Fbutton']")
		  private static WebElement  editScreenFooterUpArrowBtn; 
		
		  @FindBy(xpath="//div[@id='editScreen_footer']//button[@class='icon-down-arrow Fbutton']")
		  private static WebElement  editScreenFooterDownArrowBtn; 
		  
		/////
		  	@FindBy(xpath="//label[contains(text(),'Quantity preloaded')]")
			private static WebElement qtyPreloadedLabel;
			
			@FindBy(xpath="//input[@id='editScreen_qtyPreLoaded_textbox']")
			private static WebElement qtyPreloadedTxt;
			
			@FindBy(xpath="//input[@id='editScreen_qtyPreLoaded_formulaText']")
			private static WebElement qtyPreloadedformulaTxt;
			
			@FindBy(xpath="//input[@id='editScreen_qtyPreLoaded_availableVariables']")
			private static WebElement qtyPreloadedFormulaSearchTxt;
			
			@FindBy(xpath="//button[@id='editScreen_qtyPreLoaded_Ok']")
			private static WebElement qtyPreloadedFormulaokBtn;
			
			@FindBy(xpath="//button[@id='editScreen_qtyPreLoaded_Cancel']")
			private static WebElement qtyPreloadedFormulacancelBtn;


			@FindBy(xpath="//*[@id='editScreen_copyFieldsOptFields']/div[1]/label/span")
			private static WebElement editScreenLoadFields1stChkBox;
			
			@FindBy(xpath="//*[@id='editScreen_copyFieldsOptFields']/div[2]/label/span")
			private static WebElement editScreenLoadFields2ndChkBox;
			
			@FindBy(xpath="//*[@id='editScreen_copyFieldsOptFields']/div[3]/label/span")
			private static WebElement editScreenLoadFields3rdChkBox;
			
			@FindBy(xpath="//*[@id='editScreen_copyFieldsOptFields']/div[4]/label/span")
			private static WebElement editScreenLoadFields4thChkBox;
			
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_PreLoad_formulaText']")
			private static WebElement  preloadformulaTxt;
			
			@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_PreLoad_availableVariables']")
			private static WebElement  preloadFormulaSearchTxt;
			
			@FindBy(xpath="//button[@id='editScreen_FieldsCustomization_PreLoad_Ok']")
			private static WebElement  preloadFormulaokBtn;
				
			@FindBy(xpath="//button[@id='editScreen_FieldsCustomization_PreLoad_Cancel']")
			private static WebElement  preloadFormulacancelBtn;
			
			
			///////////////////////////////////////////////////////////////////////////////////


	//**------------------------------------------------------------------------------------------------------------- 

	// Edit Layout Tab  
	   
	   @FindBy(xpath="//span[contains(text(),'Edit Layout')]")
		private static WebElement editLayoutTab;
	   
	   @FindBy(xpath="//*[@id='editLayout_tabContent']/div[1]/span[1]")
	  	private static WebElement editLayoutAddFieldsBtn;
	   
	   @FindBy(xpath="//a[contains(text(),'Field Details')]")
		private static WebElement editLayoutFieldDetailsTab;
	   
	   @FindBy(xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_Caption']")
		private static WebElement  editLayoutCaptionTxt;

		@FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_Caption_input_image']/span")
		private static WebElement  editLayOutCaptionExpandBtn;
		
		@FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_Caption_input_settings']/span")
		private static WebElement  editLayoutCaptionSettingsBtn;

					@FindBy(xpath="//div[@id='EditLayout_FieldsCustomization_FieldDetails_Caption_customize_popup_container']")
					private static WebElement  editLayoutCaptionSettingsContainer;
											
					@FindBy(xpath="//div[@id='EditLayout_FieldsCustomization_FieldDetails_Caption_customize_popup_footer']//input[1]")
					private static WebElement  editLayoutCaptionSettingsStandardFieldsBtn;
											
								@FindBy(xpath="//select[@id='EditLayout_FieldsCustomization_FieldDetails_Caption_customize_popup_standardfields_list']")
								private static WebElement  editLayoutCaptionSettingsStandardFieldFielddropdown;
																	
								@FindBy(xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_Caption_customize_popup_standardfields_header']")
								private static WebElement  editLayoutCaptionSettingsStandardFieldHeaderTxt;
																
								@FindBy(xpath="//select[@id='EditLayout_FieldsCustomization_FieldDetails_Caption_customize_popup_standardfields_alignment']")
								private static WebElement  editLayoutCaptionSettingsStandardFieldAlignmentdropdown;
																	
								@FindBy(xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_Caption_customize_popup_standardfields_width']")
								private static WebElement  editLayoutCaptionSettingStandardFieldsWidthTxt;
																	
								@FindBy(xpath="//section[@id='page_Content']//input[3]")
								private static WebElement  editLayoutCaptionSettingsStandardFieldokBtn;
																	
								@FindBy(xpath="//section[@id='page_Content']//input[4]")
								private static WebElement  editLayoutCaptionSettingsStandardFieldcancelBtn;
											
											
					@FindBy(xpath="//div[@id='EditLayout_FieldsCustomization_FieldDetails_Caption_search_container']//input[2]")
					private static WebElement  editLayoutCaptionSettingsDeleteColumnBtn;
											
					@FindBy(xpath="//section[@id='page_Content']//input[3]")
					private static WebElement  editLayoutCaptionSettingsokBtn;
								
					@FindBy(xpath="//section[@id='page_Content']//input[4]")
					private static WebElement  editLayoutCaptionSettingscancelBtn;
	   
	   @FindBy(xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_Name']")
		private static WebElement editLayoutNameTxt;
	   
	   @FindBy(xpath="//select[@id='EditLayout_FieldsCustomization_FieldDetails_DataType']")
		private static WebElement editLayoutDataTypeDropdown;
	   
	   @FindBy(xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_Maxsize']")
		private static WebElement editLayoutMaxSizeTxt;
	   
	   @FindBy(xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_defaultValue']")
		private static WebElement editLayoutDefaultValueTxt;
	   
	   @FindBy(xpath="//select[@id='EditLayout_FieldsCustomization_FieldDetails_controlType']")
		private static WebElement editLayoutControlTypeDropdown;
	   
	   @FindBy(xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_toolTip']")
		private static WebElement editLayoutToolTipTxt;
	   
	   @FindBy(xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_toolTip']")
		private static WebElement editlayoutUniqueDropdown;
	   
	   @FindBy(xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_groupName']")
		private static WebElement editLayoutgroupNameTXt;
	   
	   @FindBy(xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_entryValidation']")
		private static WebElement editLayoutEntryValidationTxt;
	   
	   @FindBy(xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_entryValidationMsg']")
		private static WebElement editLayoutValidationMessageTxt;
	   
	   @FindBy(xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_chkPostToAcc']")
		private static WebElement editLayoutPostToAccountChkBox;
	   
	   @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_FieldDetails_defaultAcc1']")
		private static WebElement editLayoutAccount1Txt;
	   
	   @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_FieldDetails_defaultAcc2']")
		private static WebElement editLayoutAccount2Txt;
	   
	   @FindBy(xpath="//input[@id='txtEditLayout_FieldsCustomization_FieldDetails_regularExp']")
		private static WebElement editLayoutRegularExpTxt;
	   
	   //  PopUp  RegularExpresiion Window
	   
	   @FindBy(xpath="//input[@id='txtREditLayout_FieldsCustomization_FieldDetails_regularExp']")
		private static WebElement regularExpTxt;  
	   
	   @FindBy(xpath="//input[@id='txtTestStringEditLayout_FieldsCustomization_FieldDetails_regularExp']")
		private static WebElement testStringTxt;  
	   
	   @FindBy(xpath="//button[contains(text(),'Test')]")
		private static WebElement testBtn;  
	   
	   @FindBy(xpath="//a[contains(text(),'Characters and Repetitions')]")
		private static WebElement characterAndRepetitionsTab;  
	   
	   @FindBy(xpath="//input[@id='txtCharAndRepititionEditLayout_FieldsCustomization_FieldDetails_regularExp']")
		private static WebElement characterAndRepetitionsRegularExpTxt;  
	 
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp0']/div[1]/div[2]/button")
		private static WebElement characterAndRepetitionsRegularExpInsertBtn;  
	   
	   @FindBy(xpath="//input[@id='chkMatchOnlyIfAbsentEditLayout_FieldsCustomization_FieldDetails_regularExp']")
		private static WebElement matchOnlyIfAbsentChkbox;  
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp0']/div[2]/fieldset/div/div[2]/div[1]/label/input")
		private static WebElement anyCharacterRadio;  
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp0']/div[2]/fieldset/div/div[2]/div[2]/label/input")
		private static WebElement alphanumericRadio;  
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp0']/div[2]/fieldset/div/div[2]/div[3]/label/input")
		private static WebElement digitRadio;  
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp0']/div[2]/fieldset/div/div[2]/div[4]/label/input")
		private static WebElement whiteSpaceRadio;  
	   
	   @FindBy(xpath="//div[@class='col-sm-10 Fradiobox padding']//input[1]")
		private static WebElement specificCharacterRadio;  
	   
	   @FindBy(xpath="//input[@id='txtSpecificCharEditLayout_FieldsCustomization_FieldDetails_regularExp']")
		private static WebElement specificCharacterBtn;  
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp0']/div[2]/fieldset/div/div[2]/div[7]/label/input")
		private static WebElement namedClassRadio; 
	   
	   @FindBy(xpath="//select[@id='ddlNamedClassEditLayout_FieldsCustomization_FieldDetails_regularExp']")
		private static WebElement namedClassdropdown;  
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp0']/div[2]/fieldset/div/div[2]/div[9]/label/input")
		private static WebElement specifiedSetRadio;  
	   
	   @FindBy(xpath="//select[@id='ddlSpecifiedSetEditLayout_FieldsCustomization_FieldDetails_regularExp']")
		private static WebElement specifiedsetDropdown;  
	   
	   @FindBy(xpath="//input[@id='chkAsFewAsPossibleEditLayout_FieldsCustomization_FieldDetails_regularExp']")
		private static WebElement asFewAsPossibleChkBox;  
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp0']/div[3]/fieldset/div/div[2]/div[1]/div[1]/label/input")
		private static WebElement justOnceRadio;  
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp0']/div[3]/fieldset/div/div[2]/div[1]/div[2]/label/input")
	  	private static WebElement anyNumberRadio; 
	  
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp0']/div[3]/fieldset/div/div[2]/div[1]/div[3]/label/input")
	  	private static WebElement oneOrMoreRadio; 
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp0']/div[3]/fieldset/div/div[2]/div[1]/div[4]/label/input")
	  	private static WebElement zeroOrOneRadio; 
	   
	   @FindBy(xpath="//div[@class='col-sm-12 padding border']//div[1]//label[1]//input[1]")
	  	private static WebElement exactlyRadio; 
	   
	   @FindBy(xpath="//input[@id='nExactlyNEditLayout_FieldsCustomization_FieldDetails_regularExp']")
		private static WebElement exactlyDropdown;  
	   
	   @FindBy(xpath="//div[@class='col-sm-12 padding border']//div[3]//label[1]//input[1]")
		private static WebElement atleastRadio;  
	   
	   @FindBy(xpath="//input[@id='mAtLeastNEditLayout_FieldsCustomization_FieldDetails_regularExp']")
		private static WebElement atleastDropdown;  
	   
	   @FindBy(xpath="//div[@class='Fradiobox col-sm-12']//input[1]")
		private static WebElement betweenNandMBtn;  
	   
	   @FindBy(xpath="//button[@class='Fbutton'][contains(text(),'Ok')]")
		private static WebElement regularExpokBtn;  
	   
	   @FindBy(xpath="//button[@class='Fbutton'][contains(text(),'Close')]")
		private static WebElement regularExpcancelBtn;  
	   
	   
	   //
	   
	   @FindBy(xpath="//a[contains(text(),'Groups and Options')]")
		private static WebElement groupAndOptionsTab;  
	   
	   @FindBy(xpath="//input[@id='txtGrpAndOptEditLayout_FieldsCustomization_FieldDetails_regularExp']")
		private static WebElement groupAndOptionRegularExpTxt;  
	   
	   @FindBy(xpath="//button[contains(text(),'Result')]")
		private static WebElement resultBtn;  
	   
	   @FindBy(xpath="//fieldset[@id='dvDefineAMatchGrpEditLayout_FieldsCustomization_FieldDetails_regularExp']//div[@class='col-sm-12']//div[1]//label[1]//input[1]")
		private static WebElement numbereedCapturedRadio;  
	   
	   @FindBy(xpath="//fieldset[@id='dvDefineAMatchGrpEditLayout_FieldsCustomization_FieldDetails_regularExp']//div[2]//label[1]//input[1]")
		private static WebElement nonCapturingGroupRadio;  
	   
	   @FindBy(xpath="//fieldset[@id='dvDefineAMatchGrpEditLayout_FieldsCustomization_FieldDetails_regularExp']//div[3]//label[1]//input[1]")
		private static WebElement namedCaptureRadio;  
	   
	   @FindBy(xpath="//input[@id='txtNamedCaptureEditLayout_FieldsCustomization_FieldDetails_regularExp']")
		private static WebElement namedCaptureTxt;  
	   
	   @FindBy(xpath="//fieldset[@id='dvDefineAMatchGrpEditLayout_FieldsCustomization_FieldDetails_regularExp']//div[5]//label[1]//input[1]")
		private static WebElement balancingGrpRadio;  
	   
	   @FindBy(xpath="//input[@id='txtBalancingGrpEditLayout_FieldsCustomization_FieldDetails_regularExp']")
		private static WebElement balancingGrpTxt;  
	   
	   @FindBy(xpath="//div[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp1']//div[7]//label[1]//input[1]")
		private static WebElement matchSuffixButEcludeItRadio;  
	   
	   @FindBy(xpath="//div[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp1']//div[8]//label[1]//input[1]")
	  	private static WebElement matchPrefixButEcludeItRadio;  
	      
	   @FindBy(xpath="//label[text()='Match if suffix is not present (?']")
	  	private static WebElement matchIfSuffixIsNotPresentRadio;  
	      
	   @FindBy(xpath="//div[@id='editLayout_FieldsCustomization_FieldDetails_validationDiv']//div[10]//label[1]")
	  	private static WebElement matchIfPrefixIsNotPresentRadio;  
	      
	    @FindBy(xpath="//div[@id='editLayout_FieldsCustomization_FieldDetails_validationDiv']//div[11]//label[1]")
	  	private static WebElement greedySubExpressionRadio;  
	      
	   @FindBy(xpath="//div[@id='editLayout_FieldsCustomization_FieldDetails_validationDiv']//div[12]//label[1]")
	  	private static WebElement commentRadio;  
	      
	   @FindBy(xpath="//input[@id='txtCommentEditLayout_FieldsCustomization_FieldDetails_regularExp']")
	  	private static WebElement commentTxt;  
	      
	   @FindBy(xpath="//input[@id='chkWithANewNonCaptureGrpEditLayout_FieldsCustomization_FieldDetails_regularExp']")
	  	private static WebElement withInANonCapturingGroupChkBox;  
	      
	   @FindBy(xpath="//input[@id='chkUntilEndOfEncloseGrpEditLayout_FieldsCustomization_FieldDetails_regularExp']")
	  	private static WebElement untilEndOfEnclosingGroupChkBox;  
	      
	   @FindBy(xpath="//fieldset[@id='dvOptionsEditLayout_FieldsCustomization_FieldDetails_regularExp']//div[@class='col-sm-12']//div[1]//label[1]//input[1]")
	  	private static WebElement ignoreCaseChkBox;  
	      
	   @FindBy(xpath="//fieldset[@id='dvOptionsEditLayout_FieldsCustomization_FieldDetails_regularExp']//div[2]//label[1]//input[1]")
	  	private static WebElement multiLineChkBox;  
	      
	   @FindBy(xpath="//fieldset[@id='dvOptionsEditLayout_FieldsCustomization_FieldDetails_regularExp']//div[3]//label[1]//input[1]")
	  	private static WebElement singleLineChkBox;  
	      
	   @FindBy(xpath="//div[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp1']//div[4]//label[1]//input[1]")
	  	private static WebElement explicitCaptureChkBox;  
	      
	   @FindBy(xpath="//fieldset[@id='dvOptionsEditLayout_FieldsCustomization_FieldDetails_regularExp']//div[5]//label[1]//input[1]")
	  	private static WebElement ignorePatternWhiteSpaceChkBox;  
	     
	  // 
	   
	   @FindBy(xpath="//a[contains(text(),'Special Characters')]")
	  	private static WebElement sepecialCharactersTab;  
	      
	   @FindBy(xpath="//input[@id='txtSpecialCharEditLayout_FieldsCustomization_FieldDetails_regularExp']")
	  	private static WebElement specialCharacterRegularExpTxt;  
	      
	   @FindBy(xpath="//div[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']//button[@class='Fbutton'][contains(text(),'Insert')]")
	  	private static WebElement specialCharacterRegularExpInsertBtn;  
	      
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[1]/div[1]/label/input")
	  	private static WebElement bellRadio;  
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[1]/div[2]/label/input")
	  	private static WebElement backSpaceRadio; 
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[1]/div[3]/label/input")
	  	private static WebElement tabRadio; 
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[1]/div[4]/label/input")
	  	private static WebElement carriageReturnRadio; 
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[1]/div[5]/label/input")
	  	private static WebElement verticalTabRadio; 
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[1]/div[6]/label/input")
	  	private static WebElement formfeedRadio; 
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[1]/div[7]/label/input")
	  	private static WebElement newLineRadio; 
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[1]/div[8]/label/input")
	  	private static WebElement escapeRadio; 
	      
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[2]/fieldset[1]/div/div[1]/label/input")
	  	private static WebElement beginningOfStringOrLineRadio; 
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[2]/fieldset[1]/div/div[2]/label/input")
	  	private static WebElement endOfStringOrLineRadio;
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[2]/fieldset[1]/div/div[3]/label/input")
	  	private static WebElement alterNationRadio;
	 
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[2]/fieldset[2]/div/div[1]/label/input")
	  	private static WebElement beginningOfStringRadio;
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[2]/fieldset[2]/div/div[2]/label/input")
	  	private static WebElement endOfStringOrBeforeNewLineAtEndRadio;
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[2]/fieldset[2]/div/div[3]/label/input")
	  	private static WebElement endOfStringRadio;
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[2]/fieldset[2]/div/div[4]/label/input")
	  	private static WebElement beginningOfCurrentSearchRadio;
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[2]/fieldset[2]/div/div[5]/label/input")
	  	private static WebElement firsOrLastCharacterInWordRadio;
	   
	   @FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp2']/div[2]/fieldset/div/div[2]/fieldset[2]/div/div[6]/label/input")
	  	private static WebElement noOfFirstOfLastCharacterInWordRadio;
	     
	   //
	   
	   @FindBy(xpath="//a[contains(text(),'Exiting')]")
	  	private static WebElement excitngTab;
	   
	   //
	      
	   @FindBy(xpath="//ul[@class='nav nav-tabs']//a[contains(text(),'Help')]")
	  	private static WebElement helpTab;  
	      
	   @FindBy(xpath="//div[@class='col-xs-12 col-sm-3']//div[1]//label[1]//input[1]")
	  	private static WebElement syntaxElementsRadio;  
	      
	   @FindBy(xpath="//div[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp4']//div[2]//label[1]//input[1]")
	  	private static WebElement charactersRadio;  
	      
	   @FindBy(xpath="//div[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp4']//div[3]//label[1]//input[1]")
	  	private static WebElement quantifierRadio;  
	      
	   @FindBy(xpath="//div[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp4']//div[4]//label[1]//input[1]")
	  	private static WebElement anchorsRadio;  
	      
	   @FindBy(xpath="//div[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp4']//div[5]//label[1]//input[1]")
	  	private static WebElement characterClassradio;  
	      
	   @FindBy(xpath="//div[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp4']//div[6]//label[1]//input[1]")
	  	private static WebElement extendedgroupsRadio;  
	      
	   @FindBy(xpath="//div[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp4']//div[7]//label[1]//input[1]")
	  	private static WebElement backReferenceRadio;  
	      
	   @FindBy(xpath="//div[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp4']//div[8]//label[1]//input[1]")
	  	private static WebElement subExpcallRadio;  
	      
	   @FindBy(xpath="//div[@id='EditLayout_FieldsCustomization_FieldDetails_regularExp4']//div[9]//label[1]//input[1]")
	  	private static WebElement capturedGroupRadio;  
	      
	   @FindBy(xpath="//pre[@id='preHelpEditLayout_FieldsCustomization_FieldDetails_regularExp']")
	  	private static WebElement helpField;  
	   
	    @FindBy(xpath="//span[contains(text(),'Apply')]")
	  	private static WebElement editLayoutApplyBtn;  
	      
	    @FindBy(xpath="//span[@id='editLayout_FieldsCustomization_Close']//i[@class='icon-cancel icon-font7']")
	  	private static WebElement editLayoutcloseBtn;
	   
	   
	   
	    @FindBy(xpath="//input[@id='EditLayout_FieldsCustomization_FieldDetails_errorMsg']")
	  	private static WebElement editLayouterrorMessageTxt;
	  
	    @FindBy(xpath="//*[@id='editLayout_copyFieldsOptFields']/div[1]/label/span")
		private static WebElement editLayoutLoadFields1stChkBox;
		
		@FindBy(xpath="//*[@id='editLayout_copyFieldsOptFields']/div[2]/label/span")
		private static WebElement editLayoutLoadFields2ndChkBox;
		
		@FindBy(xpath="//*[@id='editLayout_copyFieldsOptFields']/div[3]/label/span")
		private static WebElement editLayoutLoadFields3rdChkBox;
	   
	   
	   
	   
	   
	   
	   
	   
	 //Properties Tab 

	      
	    @FindBy(xpath="//a[contains(text(),'Properties')]")
	  	private static WebElement editLayoutPropertiesTab;  
	      
	    @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_Properties_chkauditTrail']")
	  	private static WebElement auditTrailChkBox;  
	      
	    @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_Properties_chkhidden']")
	  	private static WebElement hiddenChkbox;  
	      
	   @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_Properties_chkmergeFld']")
	  	private static WebElement mergeFieldChkBox;  
	      
	   @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_Properties_chknotAvailFrReports']")
	  	private static WebElement notAvailableForReportsChkBox;  
	      
	   @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_Properties_chkinfoFld']")
	  	private static WebElement informationFieldChkBox;  
	      
	   @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_Properties_chkreadOnly']")
	  	private static WebElement readOnlyChkBox;  
	      
	   @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_Properties_chkcantExport']")
	  	private static WebElement cantExportChkBox;  
	      
	   @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_Properties_chkmandatory']")
	  	private static WebElement mandatorychkBox;  
	      
	   @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_Properties_chkspellCheck']")
	  	private static WebElement spellCheckChkBox;  
	      
	   @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_Properties_chkcantImport']")
	  	private static WebElement cantImportChkbox;  
	      
	   @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_Properties_chkmassUpdate']")
	  	private static WebElement massUpdateChkBox;  
	      
	   @FindBy(xpath="//select[@id='editLayout_FieldsCustomization_Properties_behavoir']")
	  	private static WebElement behaviourDropdown;  
	      
	//   Formatting Tab 
	      
	   @FindBy(xpath="//a[contains(text(),'Formatting')]")
	  	private static WebElement formattingTab;  
	      
	   @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_Formatting_colSpan']")
	  	private static WebElement columnSpanDropdown;  
	      
	   @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_Formatting_rowSpan']")
	  	private static WebElement rowSpandropdown;  
	      
	   @FindBy(xpath="//select[@id='editLayout_FieldsCustomization_Formatting_charCasing']")
	  	private static WebElement characterCasingDropdown;  
	      
	   @FindBy(xpath="//select[@id='editLayout_FieldsCustomization_Formatting_textAlign']")
	  	private static WebElement textAlignDropdown;  
	      
	   @FindBy(xpath="//input[@id='editLayout_FieldsCustomization_Formatting_chkenableFont']")
	  	private static WebElement enableFontChkbox;  
	      
	   
	   // External Module is named As EFES And ELV For Each Screen (ReceiptsEFES,ReceiptsELV )
	   // External Functions(External Fields) is same but Changes From Screen to Screen
	   
	   
	   @FindBy(xpath="//ul[@id='editScreen_FieldsCustomization_tabs']//li[4]//a[1]")
		private static WebElement  editLayoutRulesTab;
			
			@FindBy(xpath="//td[@id='editLayout_headerrow_1 _4']")
			private static WebElement  editLayoutHeader1stRowCaption;
			
			@FindBy(xpath="//td[@id='editLayout_headerrow_2 _4']")
			private static WebElement  editLayoutHeader2ndRowCaption;
			
			@FindBy(xpath="//td[@id='editLayout_bodyrow_1 _4']")
			private static WebElement  editLayoutbody1stRowCaption;
			
			
			
			@FindBy(xpath="//td[@id='editLayout_bodyrow_2 _4']")
			private static WebElement  editLayoutbody2ndRowCaption;
		
		@FindBy(xpath="//ul[@id='editScreen_FieldsCustomization_tabs']//li[5]")
		private static WebElement  editlayoutExternalModulesTab;
	      
	    @FindBy(xpath="//i[@class='icon-copyfields icon-font7']")
	  	private static WebElement editLayoutLoadFieldsBtn;  
	      
	    @FindBy(xpath="//*[@id='editLayout_voucherDropDown']")
	  	private static WebElement editLayoutLoadFieldsDropdown;  
	      
	    @FindBy(xpath="//button[@class='Fbutton pull-left']")
	  	private static WebElement editLayoutselectAllBtn;  
	      
	   @FindBy(xpath="//div[@id='editLayout_copyFieldsDiv']//button[@class='Fbutton pull-right'][contains(text(),'Ok')]")
	   private static WebElement editLayoutokBtn;     
	   
	   @FindBy(xpath="//*[@id='editScreen_copyFieldsDiv']/div[3]/button[2]")
	   private static WebElement editScreenLayoutokBtn;
	   
	 
	   
	   
	    @FindBy(xpath="//input[@id='searchlayout']")
	  	private static WebElement editLayoutSearchTxt;  
	      
	    @FindBy(xpath="//button[@id='btnEditLayoutPrev']")
	  	private static WebElement editLayoutPreviousBtn;  
	      
	    @FindBy(xpath="//button[@id='btnEditLayoutNext']")
	  	private static WebElement editLayoutNextBtn;  
	      
	    @FindBy(xpath="//span[@class='dropdown-toggle theme_button_color']")
	  	private static WebElement customizeGridBtn;  
	      
	    @FindBy(xpath="//li[@id='editLayout_hiddenColumn_34']//span[contains(text(),'Audit Trail')]")
	  	private static WebElement editLayoutAuditTrialBtn;  
	      
	    @FindBy(xpath="//span[contains(text(),'Back color')]")
	  	private static WebElement backColorBtn;  
	      
	    @FindBy(xpath="//span[contains(text(),'Banner text')]")
	  	private static WebElement bannerTextBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Behaviour')]")
	  	private static WebElement behaviourBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Calender Type')]")
	  	private static WebElement calenderTypeBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Cannot be exported')]")
	  	private static WebElement cannotBeExportedBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Cannot be imported')]")
	  	private static WebElement cannotBeImportedBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Character casing')]")
	  	private static WebElement characterCasingBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Column Span')]")
	  	private static WebElement columnSpanBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Copy from Parent')]")
	  	private static WebElement copyFromParentBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'DataType')]")
	  	private static WebElement dataTypeBtn;  
	      
	   @FindBy(xpath="//li[@id='editLayout_hiddenColumn_11']//span[contains(text(),'Default')]")
	  	private static WebElement defaultBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Display Control Type')]")
	  	private static WebElement controlTypeBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'EntryValidation')]")
	  	private static WebElement entryValidationBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Error message')]")
	  	private static WebElement errorMessageBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'External Display Member')]")
	  	private static WebElement externalDisplaymemberBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'External module attached')]")
	  	private static WebElement externalModuleAttachedBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'External Table Name')]")
	  	private static WebElement externalTableNameBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'External Value Member')]")
	  	private static WebElement externalValuememberBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Filter condition')]")
	  	private static WebElement filterConditionBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Font')]")
	  	private static WebElement fontBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Group Name')]")
	  	private static WebElement groupNameBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Hidden in group')]")
	  	private static WebElement hiddenInGroupBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Hide left panel')]")
	  	private static WebElement hideLeftPanelBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Information Field')]")
	  	private static WebElement informationFieldBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Interval Value')]")
	  	private static WebElement intervalValueBtn;  
	      
	   @FindBy(xpath="//li[@id='editLayout_hiddenColumn_28']//span[contains(text(),'Mandatory')]")
	  	private static WebElement mandatoryBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Mandatory Fields')]")
	  	private static WebElement mandatoryFieldsBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Mandatory in group')]")
	  	private static WebElement mandatoryInGroupBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Mass Update')]")
	  	private static WebElement massupdateBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Maximum Value')]")
	  	private static WebElement maximumValueBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Merge Field')]")
	  	private static WebElement mergerFieldBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Minimum Value')]")
	  	private static WebElement minimumValueBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Not available for reports')]")
	  	private static WebElement notAvailableForReportsBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Read Only')]")
	  	private static WebElement readOnlyBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Regular expression')]")
	  	private static WebElement regularExpressionBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Row Span')]")
	  	private static WebElement rowSpanBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Rule Attached')]")
	  	private static WebElement ruleAttachedBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'ScrollBar')]")
	  	private static WebElement scrollBarBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Spellcheck')]")
	  	private static WebElement spellCheckBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Text Align')]")
	  	private static WebElement textAllignBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Tooltip')]")
	  	private static WebElement toolTipBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Validation Message')]")
	  	private static WebElement validationMessageBtn;  
	      
	   @FindBy(xpath="//span[contains(text(),'Word wrap')]")
	  	private static WebElement wordWrapBtn;  
	      
	 // Header Tab  
	   
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/ul/li[1]/a")
	  	private static WebElement  editLayoutHeaderTab;  
	   
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[1]/span[1]")
	  	private static WebElement  editLayoutHeader1stRowEditBtn; 
	   
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[1]/span[2]")
	  	private static WebElement  editLayoutHeader1stRowDeleteBtn; 
	   
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[1]/div[2]/table/tbody/tr[2]/td[1]/span[1]")
	  	private static WebElement  editLayoutHeader2ndRowEditBtn; 
	   
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[1]/div[2]/table/tbody/tr[2]/td[1]/span[2]")
	  	private static WebElement  editLayoutHeader2ndRowDeleteBtn; 
	    
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[1]/div[2]/table/tbody/tr[3]/td[1]/span[1]")
	 	private static WebElement  editLayoutHeader3rdRowEditBtn; 
	  
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[1]/div[2]/table/tbody/tr[3]/td[1]/span[2]")
	 	private static WebElement  editLayoutHeader3rdRowDeleteBtn; 
	  
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[1]/div[2]/table/tbody/tr[4]/td[1]/span[1]")
		private static WebElement  editLayoutHeader4thRowEditBtn; 
	 
	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[1]/div[2]/table/tbody/tr[4]/td[1]/span[2]")
		private static WebElement  editLayoutHeader4thRowDeleteBtn; 
	   
	   
	   
	   @FindBy(xpath="//div[@id='editLayout_header']//button[@class='icon-up-arrow Fbutton']")
	  	private static WebElement  editLayoutHeaderUpArrowBtn; 

	   @FindBy(xpath="//div[@id='editLayout_header']//button[@class='icon-down-arrow Fbutton']")
	  	private static WebElement  editLayoutHeaderDownArrowBtn; 
	   
	 
	   
	  // Body Tab 
	                   
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/ul/li[2]/a")
	  	private static WebElement  editLayoutBodyTab; 
	   
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr/td[1]/span[1]")
	  	private static WebElement  editLayoutBody1stRowEditBtn; 
	   
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr/td[1]/span[2]")
	  	private static WebElement  editLayoutBody1stRowDeleteBtn; 
	   
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr[2]/td[1]/span[1]")
	  	private static WebElement  editLayoutBody2ndRowEditBtn; 
	   
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr[2]/td[1]/span[2]")
	  	private static WebElement  editLayoutBody2ndRowDeleteBtn; 
	   
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr[3]/td[1]/span[1]")
	 	private static WebElement  editLayoutBody3rdRowEditBtn; 
	  
	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[2]/div[2]/table/tbody/tr[3]/td[1]/span[2]")
	 	private static WebElement  editLayoutBody3rdRowDeleteBtn; 
	 
	   
	   
	   @FindBy(xpath="//div[@id='editLayout_body']//button[@class='icon-up-arrow Fbutton']")
	  	private static WebElement  editLayoutBodyUpArrowBtn; 

	   @FindBy(xpath="//div[@id='editLayout_body']//button[@class='icon-down-arrow Fbutton']")
	  	private static WebElement  editLayoutBodyDownArrowBtn;
	   

	//**-------------------------------------------------------------------------------------------------------  

		// External Functions/ External fields Tab 
		// External Module in Edit Screen.Edit Layout,Views Are Also in this
		// External Functions/Edit Screen===EFES
		// Edit Layout/ Views====ELV
		
		
		@FindBy(xpath="//span[contains(text(),'External Functions')]")
	  	private static WebElement  externalFunctionsTab; 
		
		
		/////////////
	   
		// Receipts External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_4608_5']")
	  	private static WebElement  receiptsEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_4608_5']")
	  	private static WebElement  receiptsEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_4608_5']")
	  	private static WebElement  receiptsEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_4608_5']")
	  	private static WebElement  receiptsEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_4608_5']")
	  	private static WebElement  receiptsEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_4608_5']")
	  	private static WebElement  receiptsEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4608_5+_0']")
	  	private static WebElement  receiptsEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4608_5+_1']")
	  	private static WebElement  receiptsEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4608_5+_2']")
	  	private static WebElement  receiptsEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_4608_5']")
	  	private static WebElement  receiptsEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_4608_5']")
	  	private static WebElement  receiptsEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_4608_5']")
	  	private static WebElement  receiptsEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_4608_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  receiptsEFES_closeBtn; 
	   
	   
	   // Receipts Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_4608_7']")
	  	private static WebElement  receiptsELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_4608_7']")
	  	private static WebElement  receiptsELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_4608_7']")
	  	private static WebElement  receiptsELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_4608_7']")
	  	private static WebElement  receiptsELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_4608_7']")
	  	private static WebElement  receiptsELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_4608_7']")
	  	private static WebElement  receiptsELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4608_7+_0']")
	  	private static WebElement  receiptsELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4608_7+_1']")
	  	private static WebElement  receiptsELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4608_7+_2']")
	  	private static WebElement  receiptsELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_4608_7']")
	  	private static WebElement  receiptsELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_4608_7']")
	  	private static WebElement  receiptsELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_4608_7']")
	  	private static WebElement  receiptsELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_4608_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  receiptsELV_closeBtn;
		
	   
	  ////// 
	   
		// Payments External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_4864_5']")
	  	private static WebElement  paymentsEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_4864_5']")
	  	private static WebElement  paymentsEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_4864_5']")
	  	private static WebElement  paymentsEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_4864_5']")
	  	private static WebElement  paymentsEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_4864_5']")
	  	private static WebElement  paymentsEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_4864_5']")
	  	private static WebElement  paymentsEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4864_5+_0']")
	  	private static WebElement  paymentsEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4864_5+_1']")
	  	private static WebElement  paymentsEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4864_5+_2']")
	  	private static WebElement  paymentsEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_4864_5']")
	  	private static WebElement  paymentsEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_4864_5']")
	  	private static WebElement  paymentsEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_4864_5']")
	  	private static WebElement  paymentsEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_4864_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  paymentsEFES_closeBtn; 
	   
	   
	   // Payments Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_4864_7']")
	  	private static WebElement  paymentsELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_4864_7']")
	  	private static WebElement  paymentsELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_4864_7']")
	  	private static WebElement  paymentsELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_4864_7']")
	  	private static WebElement  paymentsELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_4864_7']")
	  	private static WebElement  paymentsELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_4864_7']")
	  	private static WebElement  paymentsELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4864_7+_0']")
	  	private static WebElement  paymentsELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4864_7+_1']")
	  	private static WebElement  paymentsELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4864_7+_2']")
	  	private static WebElement  paymentsELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_4864_7']")
	  	private static WebElement  paymentsELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_4864_7']")
	  	private static WebElement  paymentsELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_4864_7']")
	  	private static WebElement  paymentsELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_4864_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  paymentsELV_closeBtn;
	   
	   
	 ////////////////  
	   
		// PettyCash External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_5120_5']")
	  	private static WebElement  pettyCashEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_5120_5']")
	  	private static WebElement  pettyCashEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_5120_5']")
	  	private static WebElement  pettyCashEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_5120_5']")
	  	private static WebElement  pettyCashEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_5120_5']")
	  	private static WebElement  pettyCashEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_5120_5']")
	  	private static WebElement  pettyCashEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5120_5+_0']")
	  	private static WebElement  pettyCashEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5120_5+_1']")
	  	private static WebElement  pettyCashEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5120_5+_2']")
	  	private static WebElement  pettyCashEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_5120_5']")
	  	private static WebElement  pettyCashEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_5120_5']")
	  	private static WebElement  pettyCashEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_5120_5']")
	  	private static WebElement  pettyCashEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_5120_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  pettyCashEFES_closeBtn; 
	   
	   
	   // PettyCash Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_5120_7']")
	  	private static WebElement  pettyCashELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_5120_7']")
	  	private static WebElement  pettyCashELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_5120_7']")
	  	private static WebElement  pettyCashELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_5120_7']")
	  	private static WebElement  pettyCashELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_5120_7']")
	  	private static WebElement  pettyCashELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_5120_7']")
	  	private static WebElement  pettyCashELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5120_7+_0']")
	  	private static WebElement  pettyCashELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5120_7+_1']")
	  	private static WebElement  pettyCashELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5120_7+_2']")
	  	private static WebElement  pettyCashELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_5120_7']")
	  	private static WebElement  pettyCashELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_5120_7']")
	  	private static WebElement  pettyCashELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_5120_7']")
	  	private static WebElement  pettyCashELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_5120_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  pettyCashELV_closeBtn;
	   
	 /////////////////////////////////
	   
	   
		// PostDatedReceipts External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_5888_5']")
	  	private static WebElement  postDatedReceiptsEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_5888_5']")
	  	private static WebElement  postDatedReceiptsEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_5888_5']")
	  	private static WebElement  postDatedReceiptsEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_5888_5']")
	  	private static WebElement  postDatedReceiptsEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_5888_5']")
	  	private static WebElement  postDatedReceiptsEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_5888_5']")
	  	private static WebElement  postDatedReceiptsEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5888_5+_0']")
	  	private static WebElement  postDatedReceiptsEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5888_5+_1']")
	  	private static WebElement  postDatedReceiptsEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5888_5+_2']")
	  	private static WebElement  postDatedReceiptsEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_5888_5']")
	  	private static WebElement  postDatedReceiptsEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_5888_5']")
	  	private static WebElement  postDatedReceiptsEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_5888_5']")
	  	private static WebElement  postDatedReceiptsEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_5888_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  postDatedReceiptsEFES_closeBtn; 
	   
	   
	   // PostDatedReceipts Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_5888_7']")
	  	private static WebElement  postDatedReceiptsELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_5888_7']")
	  	private static WebElement  postDatedReceiptsELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_5888_7']")
	  	private static WebElement  postDatedReceiptsELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_5888_7']")
	  	private static WebElement  postDatedReceiptsELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_5888_7']")
	  	private static WebElement  postDatedReceiptsELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_5888_7']")
	  	private static WebElement  postDatedReceiptsELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5888_7+_0']")
	  	private static WebElement  postDatedReceiptsELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5888_7+_1']")
	  	private static WebElement  postDatedReceiptsELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5888_7+_2']")
	  	private static WebElement  postDatedReceiptsELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_5888_7']")
	  	private static WebElement  postDatedReceiptsELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_5888_7']")
	  	private static WebElement  postDatedReceiptsELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_5888_7']")
	  	private static WebElement  postDatedReceiptsELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_5888_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  postDatedReceiptsELV_closeBtn;
	   
	   
	   ///////////////////////////////
	   
		// PostDatedPayments External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_7168_5']")
	  	private static WebElement  postDatedPaymentsEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_7168_5']")
	  	private static WebElement  postDatedPaymentsEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_7168_5']")
	  	private static WebElement  postDatedPaymentsEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_7168_5']")
	  	private static WebElement  postDatedPaymentsEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_7168_5']")
	  	private static WebElement  postDatedPaymentsEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_7168_5']")
	  	private static WebElement  postDatedPaymentsEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7168_5+_0']")
	  	private static WebElement  postDatedPaymentsEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7168_5+_1']")
	  	private static WebElement  postDatedPaymentsEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7168_5+_2']")
	  	private static WebElement  postDatedPaymentsEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_7168_5']")
	  	private static WebElement  postDatedPaymentsEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_7168_5']")
	  	private static WebElement  postDatedPaymentsEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_7168_5']")
	  	private static WebElement  postDatedPaymentsEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_7168_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  postDatedPaymentsEFES_closeBtn; 
	   
	   
	   // PostDatedPayments Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_7168_7']")
	  	private static WebElement  postDatedPaymentsELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_7168_7']")
	  	private static WebElement  postDatedPaymentsELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_7168_7']")
	  	private static WebElement  postDatedPaymentsELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_7168_7']")
	  	private static WebElement  postDatedPaymentsELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_7168_7']")
	  	private static WebElement  postDatedPaymentsELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_7168_7']")
	  	private static WebElement  postDatedPaymentsELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7168_7+_0']")
	  	private static WebElement  postDatedPaymentsELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7168_7+_1']")
	  	private static WebElement  postDatedPaymentsELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7168_7+_2']")
	  	private static WebElement  postDatedPaymentsELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_7168_7']")
	  	private static WebElement  postDatedPaymentsELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_7168_7']")
	  	private static WebElement  postDatedPaymentsELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_7168_7']")
	  	private static WebElement  postDatedPaymentsELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_7168_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  postDatedPaymentsELV_closeBtn;
	   
	   
	   
	   /////////////////////////
	   
	   
		// purchaseVouchers External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_768_5']")
	  	private static WebElement  purchaseVouchersEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_768_5']")
	  	private static WebElement  purchaseVouchersEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_768_5']")
	  	private static WebElement  purchaseVouchersEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_768_5']")
	  	private static WebElement  purchaseVouchersEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_768_5']")
	  	private static WebElement  purchaseVouchersEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_768_5']")
	  	private static WebElement  purchaseVouchersEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_768_5+_0']")
	  	private static WebElement  purchaseVouchersEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_768_5+_1']")
	  	private static WebElement  purchaseVouchersEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_768_5+_2']")
	  	private static WebElement  purchaseVouchersEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_768_5']")
	  	private static WebElement  purchaseVouchersEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_768_5']")
	  	private static WebElement  purchaseVouchersEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_768_5']")
	  	private static WebElement  purchaseVouchersEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_768_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  purchaseVouchersEFES_closeBtn; 
	   
	   
	   // purchaseVouchers Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_768_7']")
	  	private static WebElement  purchaseVouchersELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_768_7']")
	  	private static WebElement  purchaseVouchersELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_768_7']")
	  	private static WebElement  purchaseVouchersELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_768_7']")
	  	private static WebElement  purchaseVouchersELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_768_7']")
	  	private static WebElement  purchaseVouchersELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_768_7']")
	  	private static WebElement  purchaseVouchersELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_768_7+_0']")
	  	private static WebElement  purchaseVouchersELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_768_7+_1']")
	  	private static WebElement  purchaseVouchersELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_768_7+_2']")
	  	private static WebElement  purchaseVouchersELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_768_7']")
	  	private static WebElement  purchaseVouchersELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_768_7']")
	  	private static WebElement  purchaseVouchersELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_768_7']")
	  	private static WebElement  purchaseVouchersELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_768_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  purchaseVouchersELV_closeBtn;
	   
	   /////////////////
	   
	   
		// PurchaseReturns External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_6400_5']")
	  	private static WebElement  purchaseReturnsEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_6400_5']")
	  	private static WebElement  purchaseReturnsEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_6400_5']")
	  	private static WebElement  purchaseReturnsEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_6400_5']")
	  	private static WebElement  purchaseReturnsEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_6400_5']")
	  	private static WebElement  purchaseReturnsEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_6400_5']")
	  	private static WebElement  purchaseReturnsEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_6400_5+_0']")
	  	private static WebElement  purchaseReturnsEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_6400_5+_1']")
	  	private static WebElement  purchaseReturnsEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_6400_5+_2']")
	  	private static WebElement  purchaseReturnsEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_6400_5']")
	  	private static WebElement  purchaseReturnsEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_6400_5']")
	  	private static WebElement  purchaseReturnsEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_6400_5']")
	  	private static WebElement  purchaseReturnsEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_6400_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  purchaseReturnsEFES_closeBtn; 
	   
	   
	   // PurchaseReturns Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_6400_7']")
	  	private static WebElement  purchaseReturnsELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_6400_7']")
	  	private static WebElement  purchaseReturnsELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_6400_7']")
	  	private static WebElement  purchaseReturnsELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_6400_7']")
	  	private static WebElement  purchaseReturnsELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_6400_7']")
	  	private static WebElement  purchaseReturnsELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_6400_7']")
	  	private static WebElement  purchaseReturnsELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_6400_7+_0']")
	  	private static WebElement  purchaseReturnsELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_6400_7+_1']")
	  	private static WebElement  purchaseReturnsELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_6400_7+_2']")
	  	private static WebElement  purchaseReturnsELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_6400_7']")
	  	private static WebElement  purchaseReturnsELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_6400_7']")
	  	private static WebElement  purchaseReturnsELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_6400_7']")
	  	private static WebElement  purchaseReturnsELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_6400_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  purchaseReturnsELV_closeBtn;
	   
	   
	   /////////////////////////
	   
		// SalesInvoices External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3328_5']")
	  	private static WebElement  salesInvoicesEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3328_5']")
	  	private static WebElement  salesInvoicesEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3328_5']")
	  	private static WebElement  salesInvoicesEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3328_5']")
	  	private static WebElement  salesInvoicesEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3328_5']")
	  	private static WebElement  salesInvoicesEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3328_5']")
	  	private static WebElement  salesInvoicesEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3328_5+_0']")
	  	private static WebElement  salesInvoicesEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3328_5+_1']")
	  	private static WebElement  salesInvoicesEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3328_5+_2']")
	  	private static WebElement  salesInvoicesEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3328_5']")
	  	private static WebElement  salesInvoicesEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3328_5']")
	  	private static WebElement  salesInvoicesEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3328_5']")
	  	private static WebElement  salesInvoicesEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3328_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  salesInvoicesEFES_closeBtn; 
	   
	   
	   // SalesInvoices Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3328_7']")
	  	private static WebElement  salesInvoicesELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3328_7']")
	  	private static WebElement  salesInvoicesELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3328_7']")
	  	private static WebElement  salesInvoicesELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3328_7']")
	  	private static WebElement  salesInvoicesELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3328_7']")
	  	private static WebElement  salesInvoicesELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3328_7']")
	  	private static WebElement  salesInvoicesELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3328_7+_0']")
	  	private static WebElement  salesInvoicesELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3328_7+_1']")
	  	private static WebElement  salesInvoicesELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3328_7+_2']")
	  	private static WebElement  salesInvoicesELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3328_7']")
	  	private static WebElement  salesInvoicesELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3328_7']")
	  	private static WebElement  salesInvoicesELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3328_7']")
	  	private static WebElement  salesInvoicesELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3328_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  salesInvoicesELV_closeBtn;
	   
	   
	   ////////////////
	   
		// SalesReturns External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_1792_5']")
	  	private static WebElement  salesReturnsEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_1792_5']")
	  	private static WebElement  salesReturnsEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_1792_5']")
	  	private static WebElement  salesReturnsEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_1792_5']")
	  	private static WebElement  salesReturnsEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_1792_5']")
	  	private static WebElement  salesReturnsEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_1792_5']")
	  	private static WebElement  salesReturnsEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1792_5+_0']")
	  	private static WebElement  salesReturnsEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1792_5+_1']")
	  	private static WebElement  salesReturnsEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1792_5+_2']")
	  	private static WebElement  salesReturnsEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_1792_5']")
	  	private static WebElement  salesReturnsEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_1792_5']")
	  	private static WebElement  salesReturnsEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_1792_5']")
	  	private static WebElement  salesReturnsEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_1792_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  salesReturnsEFES_closeBtn; 
	   
	   
	   // SalesReturns Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_1792_7']")
	  	private static WebElement  salesReturnsELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_1792_7']")
	  	private static WebElement  salesReturnsELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_1792_7']")
	  	private static WebElement  salesReturnsELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_1792_7']")
	  	private static WebElement  salesReturnsELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_1792_7']")
	  	private static WebElement  salesReturnsELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_1792_7']")
	  	private static WebElement  salesReturnsELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1792_7+_0']")
	  	private static WebElement  salesReturnsELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1792_7+_1']")
	  	private static WebElement  salesReturnsELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1792_7+_2']")
	  	private static WebElement  salesReturnsELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_1792_7']")
	  	private static WebElement  salesReturnsELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_1792_7']")
	  	private static WebElement  salesReturnsELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_1792_7']")
	  	private static WebElement  salesReturnsELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_1792_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  salesReturnsELV_closeBtn;
	   
	   ///////////
	   
		// CashSales External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3329_5']")
	  	private static WebElement  cashSalesEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3329_5']")
	  	private static WebElement  cashSalesEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3329_5']")
	  	private static WebElement  cashSalesEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3329_5']")
	  	private static WebElement  cashSalesEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3329_5']")
	  	private static WebElement  cashSalesEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3329_5']")
	  	private static WebElement  cashSalesEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3329_5+_0']")
	  	private static WebElement  cashSalesEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3329_5+_1']")
	  	private static WebElement  cashSalesEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3329_5+_2']")
	  	private static WebElement  cashSalesEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3329_5']")
	  	private static WebElement  cashSalesEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3329_5']")
	  	private static WebElement  cashSalesEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3329_5']")
	  	private static WebElement  cashSalesEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3329_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  cashSalesEFES_closeBtn; 
	   
	   
	   // CashSales Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3329_7']")
	  	private static WebElement  cashSalesELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3329_7']")
	  	private static WebElement  cashSalesELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3329_7']")
	  	private static WebElement  cashSalesELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3329_7']")
	  	private static WebElement  cashSalesELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3329_7']")
	  	private static WebElement  cashSalesELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3329_7']")
	  	private static WebElement  cashSalesELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3329_7+_0']")
	  	private static WebElement  cashSalesELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3329_7+_1']")
	  	private static WebElement  cashSalesELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3329_7+_2']")
	  	private static WebElement  cashSalesELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3329_7']")
	  	private static WebElement  cashSalesELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3329_7']")
	  	private static WebElement  cashSalesELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3329_7']")
	  	private static WebElement  cashSalesELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3329_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  cashSalesELV_closeBtn;
	   
	   ////////////
	   
		// HirePurchaseSales External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3330_5']")
	  	private static WebElement  hirePurchaseSalesEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3330_5']")
	  	private static WebElement  hirePurchaseSalesEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3330_5']")
	  	private static WebElement  hirePurchaseSalesEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3330_5']")
	  	private static WebElement  hirePurchaseSalesEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3330_5']")
	  	private static WebElement  hirePurchaseSalesEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3330_5']")
	  	private static WebElement  hirePurchaseSalesEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3330_5+_0']")
	  	private static WebElement  hirePurchaseSalesEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3330_5+_1']")
	  	private static WebElement  hirePurchaseSalesEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3330_5+_2']")
	  	private static WebElement  hirePurchaseSalesEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3330_5']")
	  	private static WebElement  hirePurchaseSalesEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3330_5']")
	  	private static WebElement  hirePurchaseSalesEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3330_5']")
	  	private static WebElement  hirePurchaseSalesEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3330_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  hirePurchaseSalesEFES_closeBtn; 
	   
	   
	   // HirePurchaseSales Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3330_7']")
	  	private static WebElement  hirePurchaseSalesELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3330_7']")
	  	private static WebElement  hirePurchaseSalesELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3330_7']")
	  	private static WebElement  hirePurchaseSalesELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3330_7']")
	  	private static WebElement  hirePurchaseSalesELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3330_7']")
	  	private static WebElement  hirePurchaseSalesELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3330_7']")
	  	private static WebElement  hirePurchaseSalesELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3330_7+_0']")
	  	private static WebElement  hirePurchaseSalesELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3330_7+_1']")
	  	private static WebElement  hirePurchaseSalesELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3330_7+_2']")
	  	private static WebElement  hirePurchaseSalesELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3330_7']")
	  	private static WebElement  hirePurchaseSalesELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3330_7']")
	  	private static WebElement  hirePurchaseSalesELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3330_7']")
	  	private static WebElement  hirePurchaseSalesELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3330_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  hirePurchaseSalesELV_closeBtn;
	   
	   
	   /////////////////
	   
	   
		// JournalEntries External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_8704_5']")
	  	private static WebElement  journalEntriesEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_8704_5']")
	  	private static WebElement  journalEntriesEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_8704_5']")
	  	private static WebElement  journalEntriesEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_8704_5']")
	  	private static WebElement  journalEntriesEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_8704_5']")
	  	private static WebElement  journalEntriesEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_8704_5']")
	  	private static WebElement  journalEntriesEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8704_5+_0']")
	  	private static WebElement  journalEntriesEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8704_5+_1']")
	  	private static WebElement  journalEntriesEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8704_5+_2']")
	  	private static WebElement  journalEntriesEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_8704_5']")
	  	private static WebElement  journalEntriesEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_8704_5']")
	  	private static WebElement  journalEntriesEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_8704_5']")
	  	private static WebElement  journalEntriesEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_8704_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  journalEntriesEFES_closeBtn; 
	   
	   
	   // JournalEntries Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_8704_7']")
	  	private static WebElement  JournalEntriesELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_8704_7']")
	  	private static WebElement  JournalEntriesELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_8704_7']")
	  	private static WebElement  JournalEntriesELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_8704_7']")
	  	private static WebElement  JournalEntriesELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_8704_7']")
	  	private static WebElement  JournalEntriesELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_8704_7']")
	  	private static WebElement  JournalEntriesELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8704_7+_0']")
	  	private static WebElement  JournalEntriesELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8704_7+_1']")
	  	private static WebElement  JournalEntriesELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8704_7+_2']")
	  	private static WebElement  JournalEntriesELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_8704_7']")
	  	private static WebElement  JournalEntriesELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_8704_7']")
	  	private static WebElement  JournalEntriesELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_8704_7']")
	  	private static WebElement  JournalEntriesELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_8704_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  JournalEntriesELV_closeBtn;
	   
	   //////////////
	   
		// ForexJV External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_8448_5']")
	  	private static WebElement  ForexJVEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_8448_5']")
	  	private static WebElement  ForexJVEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_8448_5']")
	  	private static WebElement  ForexJVEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_8448_5']")
	  	private static WebElement  ForexJVEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_8448_5']")
	  	private static WebElement  ForexJVEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_8448_5']")
	  	private static WebElement  ForexJVEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8448_5+_0']")
	  	private static WebElement  ForexJVEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8448_5+_1']")
	  	private static WebElement  ForexJVEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8448_5+_2']")
	  	private static WebElement  ForexJVEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_8448_5']")
	  	private static WebElement  ForexJVEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_8448_5']")
	  	private static WebElement  ForexJVEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_8448_5']")
	  	private static WebElement  ForexJVEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_8448_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  ForexJVEFES_closeBtn; 
	   
	   
	   // ForexJV Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_8448_7']")
	  	private static WebElement  ForexJVELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_8448_7']")
	  	private static WebElement  ForexJVELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_8448_7']")
	  	private static WebElement  ForexJVELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_8448_7']")
	  	private static WebElement  ForexJVELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_8448_7']")
	  	private static WebElement  ForexJVELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_8448_7']")
	  	private static WebElement  ForexJVELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8448_7+_0']")
	  	private static WebElement  ForexJVELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8448_7+_1']")
	  	private static WebElement  ForexJVELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8448_7+_2']")
	  	private static WebElement  ForexJVELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_8448_7']")
	  	private static WebElement  ForexJVELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_8448_7']")
	  	private static WebElement  ForexJVELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_8448_7']")
	  	private static WebElement  ForexJVELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_8448_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  ForexJVELV_closeBtn;
	   
	   ///////////
	   
		// InterDepartmentalJV External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3585_5']")
	  	private static WebElement  InterDepartmentalJVEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3585_5']")
	  	private static WebElement  InterDepartmentalJVEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3585_5']")
	  	private static WebElement  InterDepartmentalJVEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3585_5']")
	  	private static WebElement  InterDepartmentalJVEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3585_5']")
	  	private static WebElement  InterDepartmentalJVEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3585_5']")
	  	private static WebElement  InterDepartmentalJVEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3585_5+_0']")
	  	private static WebElement  InterDepartmentalJVEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3585_5+_1']")
	  	private static WebElement  InterDepartmentalJVEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3585_5+_2']")
	  	private static WebElement  InterDepartmentalJVEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3585_5']")
	  	private static WebElement  InterDepartmentalJVEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3585_5']")
	  	private static WebElement  InterDepartmentalJVEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3585_5']")
	  	private static WebElement  InterDepartmentalJVEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3585_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  InterDepartmentalJVEFES_closeBtn; 
	   
	   
	   // InterDepartmentalJV Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3585_7']")
	  	private static WebElement  InterDepartmentalJVELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3585_7']")
	  	private static WebElement  InterDepartmentalJVELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3585_7']")
	  	private static WebElement  InterDepartmentalJVELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3585_7']")
	  	private static WebElement  InterDepartmentalJVELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3585_7']")
	  	private static WebElement  InterDepartmentalJVELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3585_7']")
	  	private static WebElement  InterDepartmentalJVELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3585_7+_0']")
	  	private static WebElement  InterDepartmentalJVELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3585_7+_1']")
	  	private static WebElement  InterDepartmentalJVELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3585_7+_2']")
	  	private static WebElement  InterDepartmentalJVELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3585_7']")
	  	private static WebElement  InterDepartmentalJVELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3585_7']")
	  	private static WebElement  InterDepartmentalJVELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3585_7']")
	  	private static WebElement  InterDepartmentalJVELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3585_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  InterDepartmentalJVELV_closeBtn;
	   
	   //////////////
	   
		// NonStandardJournalEntries External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3584_5']")
	  	private static WebElement  NonStandardjournalEntriesEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3584_5']")
	  	private static WebElement  NonStandardjournalEntriesEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3584_5']")
	  	private static WebElement  NonStandardjournalEntriesEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3584_5']")
	  	private static WebElement  NonStandardjournalEntriesEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3584_5']")
	  	private static WebElement  NonStandardjournalEntriesEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3584_5']")
	  	private static WebElement  NonStandardjournalEntriesEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3584_5+_0']")
	  	private static WebElement  NonStandardjournalEntriesEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3584_5+_1']")
	  	private static WebElement  NonStandardjournalEntriesEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3584_5+_2']")
	  	private static WebElement  NonStandardjournalEntriesEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3584_5']")
	  	private static WebElement  NonStandardjournalEntriesEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3584_5']")
	  	private static WebElement  NonStandardjournalEntriesEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3584_5']")
	  	private static WebElement  NonStandardjournalEntriesEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3584_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  NonStandardjournalEntriesEFES_closeBtn; 
	   
	   
	   // NonStandardJournalEntries Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3584_7']")
	  	private static WebElement  nonStandardJournalEntriesELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3584_7']")
	  	private static WebElement  nonStandardJournalEntriesELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3584_7']")
	  	private static WebElement  nonStandardJournalEntriesELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3584_7']")
	  	private static WebElement  nonStandardJournalEntriesELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3584_7']")
	  	private static WebElement  nonStandardJournalEntriesELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3584_7']")
	  	private static WebElement  nonStandardJournalEntriesELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3584_7+_0']")
	  	private static WebElement  nonStandardJournalEntriesELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3584_7+_1']")
	  	private static WebElement  nonStandardJournalEntriesELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3584_7+_2']")
	  	private static WebElement  nonStandardJournalEntriesELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3584_7']")
	  	private static WebElement  nonStandardJournalEntriesELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3584_7']")
	  	private static WebElement  nonStandardJournalEntriesELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3584_7']")
	  	private static WebElement  nonStandardJournalEntriesELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3584_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  nonStandardJournalEntriesELV_closeBtn;
	   
	   ///////////////
	   
		// DebitNotes External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3840_5']")
	  	private static WebElement  debitNotesEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3840_5']")
	  	private static WebElement  debitNotesEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3840_5']")
	  	private static WebElement  debitNotesEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3840_5']")
	  	private static WebElement  debitNotesEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3840_5']")
	  	private static WebElement  debitNotesEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3840_5']")
	  	private static WebElement  debitNotesEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3840_5+_0']")
	  	private static WebElement  debitNotesEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3840_5+_1']")
	  	private static WebElement  debitNotesEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3840_5+_2']")
	  	private static WebElement  debitNotesEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3840_5']")
	  	private static WebElement  debitNotesEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3840_5']")
	  	private static WebElement  debitNotesEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3840_5']")
	  	private static WebElement  debitNotesEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3840_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  debitNotesEFES_closeBtn; 
	   
	   
	   // DebitNotes Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3840_7']")
	  	private static WebElement  debitNotesELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3840_7']")
	  	private static WebElement  debitNotesELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3840_7']")
	  	private static WebElement  debitNotesELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3840_7']")
	  	private static WebElement  debitNotesELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3840_7']")
	  	private static WebElement  debitNotesELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3840_7']")
	  	private static WebElement  debitNotesELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3840_7+_0']")
	  	private static WebElement  debitNotesELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3840_7+_1']")
	  	private static WebElement  debitNotesELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3840_7+_2']")
	  	private static WebElement  debitNotesELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3840_7']")
	  	private static WebElement  debitNotesELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3840_7']")
	  	private static WebElement  debitNotesELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3840_7']")
	  	private static WebElement  debitNotesELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3840_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  debitNotesELV_closeBtn;
	   
	   /////////////
	   
		// DebitNotesLineWise External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3586_5']")
	  	private static WebElement  debitNotesLineWiseEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3586_5']")
	  	private static WebElement  debitNotesLineWiseEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3586_5']")
	  	private static WebElement  debitNotesLineWiseEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3586_5']")
	  	private static WebElement  debitNotesLineWiseEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3586_5']")
	  	private static WebElement  debitNotesLineWiseEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3586_5']")
	  	private static WebElement  debitNotesLineWiseEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3586_5+_0']")
	  	private static WebElement  debitNotesLineWiseEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3586_5+_1']")
	  	private static WebElement  debitNotesLineWiseEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3586_5+_2']")
	  	private static WebElement  debitNotesLineWiseEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3586_5']")
	  	private static WebElement  debitNotesLineWiseEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3586_5']")
	  	private static WebElement  debitNotesLineWiseEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3586_5']")
	  	private static WebElement  debitNotesLineWiseEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3586_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  debitNotesLineWiseEFES_closeBtn; 
	   
	   
	   // DebitNotesLineWise Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3586_7']")
	  	private static WebElement  debitNotesLineWiseELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3586_7']")
	  	private static WebElement  debitNotesLineWiseELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3586_7']")
	  	private static WebElement  debitNotesLineWiseELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3586_7']")
	  	private static WebElement  debitNotesLineWiseELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3586_7']")
	  	private static WebElement  debitNotesLineWiseELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3586_7']")
	  	private static WebElement  debitNotesLineWiseELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3586_7+_0']")
	  	private static WebElement  debitNotesLineWiseELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3586_7+_1']")
	  	private static WebElement  debitNotesLineWiseELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3586_7+_2']")
	  	private static WebElement  debitNotesLineWiseELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3586_7']")
	  	private static WebElement  debitNotesLineWiseELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3586_7']")
	  	private static WebElement  debitNotesLineWiseELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3586_7']")
	  	private static WebElement  debitNotesLineWiseELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3586_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  debitNotesLineWiseELV_closeBtn;
	   
	   /////////////////////
	   
		// CreditNotes External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_4096_5']")
	  	private static WebElement  creditNotesEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_4096_5']")
	  	private static WebElement  creditNotesEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_4096_5']")
	  	private static WebElement  creditNotesEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_4096_5']")
	  	private static WebElement  creditNotesEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_4096_5']")
	  	private static WebElement  creditNotesEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_4096_5']")
	  	private static WebElement  creditNotesEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4096_5+_0']")
	  	private static WebElement  creditNotesEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4096_5+_1']")
	  	private static WebElement  creditNotesEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4096_5+_2']")
	  	private static WebElement  creditNotesEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_4096_5']")
	  	private static WebElement  creditNotesEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_4096_5']")
	  	private static WebElement  creditNotesEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_4096_5']")
	  	private static WebElement  creditNotesEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_4096_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  creditNotesEFES_closeBtn; 
	   
	   
	   // CreditNotes Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_4096_7']")
	  	private static WebElement  creditNotesELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_4096_7']")
	  	private static WebElement  creditNotesELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_4096_7']")
	  	private static WebElement  creditNotesELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_4096_7']")
	  	private static WebElement  creditNotesELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_4096_7']")
	  	private static WebElement  creditNotesELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_4096_7']")
	  	private static WebElement  creditNotesELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4096_7+_0']")
	  	private static WebElement  creditNotesELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4096_7+_1']")
	  	private static WebElement  creditNotesELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_4096_7+_2']")
	  	private static WebElement  creditNotesELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_4096_7']")
	  	private static WebElement  creditNotesELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_4096_7']")
	  	private static WebElement  creditNotesELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_4096_7']")
	  	private static WebElement  creditNotesELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_4096_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  creditNotesELV_closeBtn;
	   
	   ///////////////////
	   
		// CreditNotesLineWise External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3587_5']")
	  	private static WebElement  creditNotesLineWiseEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3587_5']")
	  	private static WebElement  creditNotesLineWiseEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3587_5']")
	  	private static WebElement  creditNotesLineWiseEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3587_5']")
	  	private static WebElement  creditNotesLineWiseEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3587_5']")
	  	private static WebElement  creditNotesLineWiseEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3587_5']")
	  	private static WebElement  creditNotesLineWiseEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3587_5+_0']")
	  	private static WebElement  creditNotesLineWiseEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3587_5+_1']")
	  	private static WebElement  creditNotesLineWiseEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3587_5+_2']")
	  	private static WebElement  creditNotesLineWiseEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3587_5']")
	  	private static WebElement  creditNotesLineWiseEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3587_5']")
	  	private static WebElement  creditNotesLineWiseEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3587_5']")
	  	private static WebElement  creditNotesLineWiseEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3587_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  creditNotesLineWiseEFES_closeBtn; 
	   
	   
	   // CreditNotesLineWise Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3587_7']")
	  	private static WebElement  creditNotesLineWiseELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3587_7']")
	  	private static WebElement  creditNotesLineWiseELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3587_7']")
	  	private static WebElement  creditNotesLineWiseELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3587_7']")
	  	private static WebElement  creditNotesLineWiseELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3587_7']")
	  	private static WebElement  creditNotesLineWiseELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3587_7']")
	  	private static WebElement  creditNotesLineWiseELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3587_7+_0']")
	  	private static WebElement  creditNotesLineWiseELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3587_7+_1']")
	  	private static WebElement  creditNotesLineWiseELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3587_7+_2']")
	  	private static WebElement  creditNotesLineWiseELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3587_7']")
	  	private static WebElement  creditNotesLineWiseELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3587_7']")
	  	private static WebElement  creditNotesLineWiseELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3587_7']")
	  	private static WebElement  creditNotesLineWiseELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3587_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  creditNotesLineWiseELV_closeBtn;
	   
	   
	   ///////////////////////////
	   
		// OpeningBalances External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_256_5']")
	  	private static WebElement  openingBalancesEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_256_5']")
	  	private static WebElement  openingBalancesEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_256_5']")
	  	private static WebElement  openingBalancesEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_256_5']")
	  	private static WebElement  openingBalancesEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_256_5']")
	  	private static WebElement  openingBalancesEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_256_5']")
	  	private static WebElement  openingBalancesEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_256_5+_0']")
	  	private static WebElement  openingBalancesEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_256_5+_1']")
	  	private static WebElement  openingBalancesEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_256_5+_2']")
	  	private static WebElement  openingBalancesEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_256_5']")
	  	private static WebElement  openingBalancesEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_256_5']")
	  	private static WebElement  openingBalancesEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_256_5']")
	  	private static WebElement  openingBalancesEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_256_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  openingBalancesEFES_closeBtn; 
	   
	   
	   // OpeningBalances Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_256_7']")
	  	private static WebElement  openingBalancesELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_256_7']")
	  	private static WebElement  openingBalancesELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_256_7']")
	  	private static WebElement  openingBalancesELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_256_7']")
	  	private static WebElement  openingBalancesELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_256_7']")
	  	private static WebElement  openingBalancesELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_256_7']")
	  	private static WebElement  openingBalancesELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_256_7+_0']")
	  	private static WebElement  openingBalancesELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_256_7+_1']")
	  	private static WebElement  openingBalancesELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_256_7+_2']")
	  	private static WebElement  openingBalancesELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_256_7']")
	  	private static WebElement  openingBalancesELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_256_7']")
	  	private static WebElement  openingBalancesELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_256_7']")
	  	private static WebElement  openingBalancesELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_256_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  openingBalancesELV_closeBtn;
	   
	   
	   //////////////////////////////////
	   
		// FixedAssetDepreciationVoucher External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3588_5']")
	  	private static WebElement  fixedAssetDepreciationVoucherEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3588_5']")
	  	private static WebElement  fixedAssetDepreciationVoucherEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3588_5']")
	  	private static WebElement  fixedAssetDepreciationVoucherEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3588_5']")
	  	private static WebElement  fixedAssetDepreciationVoucherEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3588_5']")
	  	private static WebElement  fixedAssetDepreciationVoucherEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3588_5']")
	  	private static WebElement  fixedAssetDepreciationVoucherEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3588_5+_0']")
	  	private static WebElement  fixedAssetDepreciationVoucherEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3588_5+_1']")
	  	private static WebElement  fixedAssetDepreciationVoucherEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3588_5+_2']")
	  	private static WebElement  fixedAssetDepreciationVoucherEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3588_5']")
	  	private static WebElement  fixedAssetDepreciationVoucherEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3588_5']")
	  	private static WebElement  fixedAssetDepreciationVoucherEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3588_5']")
	  	private static WebElement  fixedAssetDepreciationVoucherEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3588_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  fixedAssetDepreciationVoucherEFES_closeBtn; 
	   
	   
	   // FixedAssetDepreciationVoucher Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3588_7']")
	  	private static WebElement  fixedAssetDepreciationVoucherELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3588_7']")
	  	private static WebElement  fixedAssetDepreciationVoucherELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3588_7']")
	  	private static WebElement  fixedAssetDepreciationVoucherELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3588_7']")
	  	private static WebElement  fixedAssetDepreciationVoucherELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3588_7']")
	  	private static WebElement  fixedAssetDepreciationVoucherELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3588_7']")
	  	private static WebElement  fixedAssetDepreciationVoucherELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3588_7+_0']")
	  	private static WebElement  fixedAssetDepreciationVoucherELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3588_7+_1']")
	  	private static WebElement  fixedAssetDepreciationVoucherELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3588_7+_2']")
	  	private static WebElement  fixedAssetDepreciationVoucherELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3588_7']")
	  	private static WebElement  fixedAssetDepreciationVoucherELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3588_7']")
	  	private static WebElement  fixedAssetDepreciationVoucherELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3588_7']")
	  	private static WebElement  fixedAssetDepreciationVoucherELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3588_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  fixedAssetDepreciationVoucherELV_closeBtn;
	   
	   /////////////////////
	   
		// RequestForQuote External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_8192_5']")
	  	private static WebElement  requestForQuoteEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_8192_5']")
	  	private static WebElement  requestForQuoteEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_8192_5']")
	  	private static WebElement  requestForQuoteEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_8192_5']")
	  	private static WebElement  requestForQuoteEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_8192_5']")
	  	private static WebElement  requestForQuoteEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_8192_5']")
	  	private static WebElement  requestForQuoteEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8192_5+_0']")
	  	private static WebElement  requestForQuoteEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8192_5+_1']")
	  	private static WebElement  requestForQuoteEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8192_5+_2']")
	  	private static WebElement  requestForQuoteEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_8192_5']")
	  	private static WebElement  requestForQuoteEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_8192_5']")
	  	private static WebElement  requestForQuoteEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_8192_5']")
	  	private static WebElement  requestForQuoteEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_8192_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  requestForQuoteEFES_closeBtn; 
	   
	   
	   // RequestForQuote Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_8192_7']")
	  	private static WebElement  requestForQuoteELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_8192_7']")
	  	private static WebElement  requestForQuoteELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_8192_7']")
	  	private static WebElement  requestForQuoteELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_8192_7']")
	  	private static WebElement  requestForQuoteELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_8192_7']")
	  	private static WebElement  requestForQuoteELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_8192_7']")
	  	private static WebElement  requestForQuoteELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8192_7+_0']")
	  	private static WebElement  requestForQuoteELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8192_7+_1']")
	  	private static WebElement  requestForQuoteELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_8192_7+_2']")
	  	private static WebElement  requestForQuoteELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_8192_7']")
	  	private static WebElement  requestForQuoteELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_8192_7']")
	  	private static WebElement  requestForQuoteELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_8192_7']")
	  	private static WebElement  requestForQuoteELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_8192_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  requestForQuoteELV_closeBtn;
	   
	   //////////////////////
	   
		// PurchasesQuotations External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_2304_5']")
	  	private static WebElement  purchasesQuotationsEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_2304_5']")
	  	private static WebElement  purchasesQuotationsEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_2304_5']")
	  	private static WebElement  purchasesQuotationsEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_2304_5']")
	  	private static WebElement  purchasesQuotationsEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_2304_5']")
	  	private static WebElement  purchasesQuotationsEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_2304_5']")
	  	private static WebElement  purchasesQuotationsEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2304_5+_0']")
	  	private static WebElement  purchasesQuotationsEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2304_5+_1']")
	  	private static WebElement  purchasesQuotationsEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2304_5+_2']")
	  	private static WebElement  purchasesQuotationsEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_2304_5']")
	  	private static WebElement  purchasesQuotationsEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_2304_5']")
	  	private static WebElement  purchasesQuotationsEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_2304_5']")
	  	private static WebElement  purchasesQuotationsEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_2304_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  purchasesQuotationsEFES_closeBtn; 
	   
	   
	   // PurchasesQuotations Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_2304_7']")
	  	private static WebElement  purchasesQuotationsELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_2304_7']")
	  	private static WebElement  purchasesQuotationsELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_2304_7']")
	  	private static WebElement  purchasesQuotationsELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_2304_7']")
	  	private static WebElement  purchasesQuotationsELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_2304_7']")
	  	private static WebElement  purchasesQuotationsELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_2304_7']")
	  	private static WebElement  purchasesQuotationsELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2304_7+_0']")
	  	private static WebElement  purchasesQuotationsELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2304_7+_1']")
	  	private static WebElement  purchasesQuotationsELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2304_7+_2']")
	  	private static WebElement  purchasesQuotationsELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_2304_7']")
	  	private static WebElement  purchasesQuotationsELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_2304_7']")
	  	private static WebElement  purchasesQuotationsELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_2304_7']")
	  	private static WebElement  purchasesQuotationsELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_2304_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  purchasesQuotationsELV_closeBtn;
	   
	   
	   ///////////////////////////
	   
		// PurchasesOrders External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_2560_5']")
	  	private static WebElement  purchasesOrdersEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_2560_5']")
	  	private static WebElement  purchasesOrdersEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_2560_5']")
	  	private static WebElement  purchasesOrdersEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_2560_5']")
	  	private static WebElement  purchasesOrdersEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_2560_5']")
	  	private static WebElement  purchasesOrdersEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_2560_5']")
	  	private static WebElement  purchasesOrdersEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2560_5+_0']")
	  	private static WebElement  purchasesOrdersEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2560_5+_1']")
	  	private static WebElement  purchasesOrdersEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2560_5+_2']")
	  	private static WebElement  purchasesOrdersEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_2560_5']")
	  	private static WebElement  purchasesOrdersEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_2560_5']")
	  	private static WebElement  purchasesOrdersEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_2560_5']")
	  	private static WebElement  purchasesOrdersEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_2560_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  purchasesOrdersEFES_closeBtn; 
	   
	   
	   // PurchasesOrders Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_2560_7']")
	  	private static WebElement  purchasesOrdersELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_2560_7']")
	  	private static WebElement  purchasesOrdersELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_2560_7']")
	  	private static WebElement  purchasesOrdersELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_2560_7']")
	  	private static WebElement  purchasesOrdersELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_2560_7']")
	  	private static WebElement  purchasesOrdersELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_2560_7']")
	  	private static WebElement  purchasesOrdersELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2560_7+_0']")
	  	private static WebElement  purchasesOrdersELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2560_7+_1']")
	  	private static WebElement  purchasesOrdersELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2560_7+_2']")
	  	private static WebElement  purchasesOrdersELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_2560_7']")
	  	private static WebElement  purchasesOrdersELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_2560_7']")
	  	private static WebElement  purchasesOrdersELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_2560_7']")
	  	private static WebElement  purchasesOrdersELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_2560_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  purchasesOrdersELV_closeBtn;
	   
	   /////////////////////////
	   

		// MaterialReceiptNotes External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_1280_5']")
	  	private static WebElement  materialReceiptNotesEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_1280_5']")
	  	private static WebElement  materialReceiptNotesEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_1280_5']")
	  	private static WebElement  materialReceiptNotesEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_1280_5']")
	  	private static WebElement  materialReceiptNotesEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_1280_5']")
	  	private static WebElement  materialReceiptNotesEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_1280_5']")
	  	private static WebElement  materialReceiptNotesEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1280_5+_0']")
	  	private static WebElement  materialReceiptNotesEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1280_5+_1']")
	  	private static WebElement  materialReceiptNotesEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1280_5+_2']")
	  	private static WebElement  materialReceiptNotesEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_1280_5']")
	  	private static WebElement  materialReceiptNotesEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_1280_5']")
	  	private static WebElement  materialReceiptNotesEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_1280_5']")
	  	private static WebElement  materialReceiptNotesEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_1280_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  materialReceiptNotesEFES_closeBtn; 
	   
	   
	   // MaterialReceiptNotes Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[2]/div/div[2]/div[5]/div/div[1]/fieldset/div[1]/div[1]/div[1]/div/div[1]/select")
	  	private static WebElement  materialReceiptNotesELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_1280_7']")
	  	private static WebElement  materialReceiptNotesELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[2]/div/div[2]/div[5]/div/div[1]/fieldset/div[1]/div[1]/div[2]/div/div[1]/select")
	  	private static WebElement  materialReceiptNotesELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[2]/div/div[2]/div[5]/div/div[1]/fieldset/div[1]/div[2]/div[2]/div/input")
	  	private static WebElement  materialReceiptNotesELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[2]/div/div[2]/div[5]/div/div[1]/fieldset/div[1]/div[1]/div[3]/div/div/select")
	  	private static WebElement  materialReceiptNotesELV_ClassNameDropdown;
			
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[2]/div/div[2]/div[5]/div/div[1]/fieldset/div[1]/div[2]/div[3]/div/select")
	  	private static WebElement  materialReceiptNotesELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1280_7+_0']")
	  	private static WebElement  materialReceiptNotesELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1280_7+_1']")
	  	private static WebElement  materialReceiptNotesELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1280_7+_2']")
	  	private static WebElement  materialReceiptNotesELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_1280_7']")
	  	private static WebElement  materialReceiptNotesELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[2]/div/div[1]/span[1]/i")
	  	private static WebElement  materialReceiptNotesELV_InsertBtn;
			
	   @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[2]/div[2]/div[2]/div/div/div[2]/div[3]/div[2]/div/div[1]/span[2]/i")
	  	private static WebElement  materialReceiptNotesELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_1280_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  materialReceiptNotesELV_closeBtn;
	   
	   
	   /////////////////////
	   
	//SalesQuotations External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_7424_5']")
	  	private static WebElement  salesQuotationsEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_7424_5']")
	  	private static WebElement  salesQuotationsEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_7424_5']")
	  	private static WebElement  salesQuotationsEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_7424_5']")
	  	private static WebElement  salesQuotationsEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_7424_5']")
	  	private static WebElement  salesQuotationsEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_7424_5']")
	  	private static WebElement  salesQuotationsEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7424_5+_0']")
	  	private static WebElement  salesQuotationsEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7424_5+_1']")
	  	private static WebElement  salesQuotationsEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7424_5+_2']")
	  	private static WebElement  salesQuotationsEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_7424_5']")
	  	private static WebElement  salesQuotationsEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_7424_5']")
	  	private static WebElement  salesQuotationsEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_7424_5']")
	  	private static WebElement  salesQuotationsEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_7424_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  salesQuotationsEFES_closeBtn; 
	   
	   
	   // SalesQuotations Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_7424_7']")
	  	private static WebElement  salesQuotationsELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_7424_7']")
	  	private static WebElement  salesQuotationsELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_7424_7']")
	  	private static WebElement  salesQuotationsELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_7424_7']")
	  	private static WebElement  salesQuotationsELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_7424_7']")
	  	private static WebElement  salesQuotationsELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_7424_7']")
	  	private static WebElement  salesQuotationsELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7424_7+_0']")
	  	private static WebElement  salesQuotationsELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7424_7+_1']")
	  	private static WebElement  salesQuotationsELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7424_7+_2']")
	  	private static WebElement  salesQuotationsELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_7424_7']")
	  	private static WebElement  salesQuotationsELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_7424_7']")
	  	private static WebElement  salesQuotationsELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_7424_7']")
	  	private static WebElement  salesQuotationsELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_7424_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  salesQuotationsELV_closeBtn;
	   
	   
	   ////////////////////////////
	   
	//DeliveryNotes External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_6144_5']")
	  	private static WebElement  deliveryNotesEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_6144_5']")
	  	private static WebElement  deliveryNotesEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_6144_5']")
	  	private static WebElement  deliveryNotesEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_6144_5']")
	  	private static WebElement  deliveryNotesEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_6144_5']")
	  	private static WebElement  deliveryNotesEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_6144_5']")
	  	private static WebElement  deliveryNotesEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_6144_5+_0']")
	  	private static WebElement  deliveryNotesEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_6144_5+_1']")
	  	private static WebElement  deliveryNotesEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_6144_5+_2']")
	  	private static WebElement  deliveryNotesEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_6144_5']")
	  	private static WebElement  deliveryNotesEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_6144_5']")
	  	private static WebElement  deliveryNotesEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_6144_5']")
	  	private static WebElement  deliveryNotesEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_6144_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  deliveryNotesEFES_closeBtn; 
	   
	   
	   // DeliveryNotes Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_6144_7']")
	  	private static WebElement  deliveryNotesELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_6144_7']")
	  	private static WebElement  deliveryNotesELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_6144_7']")
	  	private static WebElement  deliveryNotesELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_6144_7']")
	  	private static WebElement  deliveryNotesELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_6144_7']")
	  	private static WebElement  deliveryNotesELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_6144_7']")
	  	private static WebElement  deliveryNotesELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_6144_7+_0']")
	  	private static WebElement  deliveryNotesELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_6144_7+_1']")
	  	private static WebElement  deliveryNotesELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_6144_7+_2']")
	  	private static WebElement  deliveryNotesELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_6144_7']")
	  	private static WebElement  deliveryNotesELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_6144_7']")
	  	private static WebElement  deliveryNotesELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_6144_7']")
	  	private static WebElement  deliveryNotesELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_6144_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  deliveryNotesELV_closeBtn;
	   
	   
	   /////////////////
	   

		// PosSales External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3331_5']")
	  	private static WebElement  posSalesEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3331_5']")
	  	private static WebElement  posSalesEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3331_5']")
	  	private static WebElement  posSalesEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3331_5']")
	  	private static WebElement  posSalesEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3331_5']")
	  	private static WebElement  posSalesEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3331_5']")
	  	private static WebElement  posSalesEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3331_5+_0']")
	  	private static WebElement  posSalesEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3331_5+_1']")
	  	private static WebElement  posSalesEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3331_5+_2']")
	  	private static WebElement  posSalesEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3331_5']")
	  	private static WebElement  posSalesEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3331_5']")
	  	private static WebElement  posSalesEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3331_5']")
	  	private static WebElement  posSalesEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3331_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  posSalesEFES_closeBtn; 
	   
	   
	   // PosSales Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3331_7']")
	  	private static WebElement  posSalesELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3331_7']")
	  	private static WebElement  posSalesELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3331_7']")
	  	private static WebElement  posSalesELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3331_7']")
	  	private static WebElement  posSalesELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3331_7']")
	  	private static WebElement  posSalesELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3331_7']")
	  	private static WebElement  posSalesELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3331_7+_0']")
	  	private static WebElement  posSalesELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3331_7+_1']")
	  	private static WebElement  posSalesELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3331_7+_2']")
	  	private static WebElement  posSalesELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3331_7']")
	  	private static WebElement  posSalesELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3331_7']")
	  	private static WebElement  posSalesELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3331_7']")
	  	private static WebElement  posSalesELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3331_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  posSalesELV_closeBtn;
	   
	   /////////////////
	   
		// ProductionOrders External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_5633_5']")
	  	private static WebElement  productionOrdersEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_5633_5']")
	  	private static WebElement  productionOrdersEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_5633_5']")
	  	private static WebElement  productionOrdersEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_5633_5']")
	  	private static WebElement  productionOrdersEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_5633_5']")
	  	private static WebElement  productionOrdersEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_5633_5']")
	  	private static WebElement  productionOrdersEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5633_5+_0']")
	  	private static WebElement  productionOrdersEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5633_5+_1']")
	  	private static WebElement  productionOrdersEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5633_5+_2']")
	  	private static WebElement  productionOrdersEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_5633_5']")
	  	private static WebElement  productionOrdersEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_5633_5']")
	  	private static WebElement  productionOrdersEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_5633_5']")
	  	private static WebElement  productionOrdersEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_5633_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  productionOrdersEFES_closeBtn; 
	   
	   
	   // ProductionOrders Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_5633_7']")
	  	private static WebElement  productionOrdersELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_5633_7']")
	  	private static WebElement  productionOrdersELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_5633_7']")
	  	private static WebElement  productionOrdersELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_5633_7']")
	  	private static WebElement  productionOrdersELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_5633_7']")
	  	private static WebElement  productionOrdersELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_5633_7']")
	  	private static WebElement  productionOrdersELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5633_7+_0']")
	  	private static WebElement  productionOrdersELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5633_7+_1']")
	  	private static WebElement  productionOrdersELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5633_7+_2']")
	  	private static WebElement  productionOrdersELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_5633_7']")
	  	private static WebElement  productionOrdersELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_5633_7']")
	  	private static WebElement  productionOrdersELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_5633_7']")
	  	private static WebElement  productionOrdersELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_5633_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  productionOrdersELV_closeBtn;
	   
	   
	   /////////////////////////
	   
		// SalesOrders External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_5632_5']")
	  	private static WebElement  salesOrdersEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_5632_5']")
	  	private static WebElement  salesOrdersEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_5632_5']")
	  	private static WebElement  salesOrdersEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_5632_5']")
	  	private static WebElement  salesOrdersEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_5632_5']")
	  	private static WebElement  salesOrdersEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_5632_5']")
	  	private static WebElement  salesOrdersEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5632_5+_0']")
	  	private static WebElement  salesOrdersEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5632_5+_1']")
	  	private static WebElement  salesOrdersEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5632_5+_2']")
	  	private static WebElement  salesOrdersEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_5632_5']")
	  	private static WebElement  salesOrdersEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_5632_5']")
	  	private static WebElement  salesOrdersEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_5632_5']")
	  	private static WebElement  salesOrdersEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_5632_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  salesOrdersEFES_closeBtn; 
	   
	   
	   // SalesOrders Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_5632_7']")
	  	private static WebElement  salesOrdersELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_5632_7']")
	  	private static WebElement  salesOrdersELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_5632_7']")
	  	private static WebElement  salesOrdersELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_5632_7']")
	  	private static WebElement  salesOrdersELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_5632_7']")
	  	private static WebElement  salesOrdersELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_5632_7']")
	  	private static WebElement  salesOrdersELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5632_7+_0']")
	  	private static WebElement  salesOrdersELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5632_7+_1']")
	  	private static WebElement  salesOrdersELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5632_7+_2']")
	  	private static WebElement  salesOrdersELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_5632_7']")
	  	private static WebElement  salesOrdersELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_5632_7']")
	  	private static WebElement  salesOrdersELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_5632_7']")
	  	private static WebElement  salesOrdersELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_5632_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  salesOrdersELV_closeBtn;
	   
	   
	   ////////////////////////////
	   
		// ShortageInStocks External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_5376_5']")
	  	private static WebElement  shortageInStocksEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_5376_5']")
	  	private static WebElement  shortageInStocksEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_5376_5']")
	  	private static WebElement  shortageInStocksEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_5376_5']")
	  	private static WebElement  shortageInStocksEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_5376_5']")
	  	private static WebElement  shortageInStocksEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_5376_5']")
	  	private static WebElement  shortageInStocksEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5376_5+_0']")
	  	private static WebElement  shortageInStocksEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5376_5+_1']")
	  	private static WebElement  shortageInStocksEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5376_5+_2']")
	  	private static WebElement  shortageInStocksEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_5376_5']")
	  	private static WebElement  shortageInStocksEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_5376_5']")
	  	private static WebElement  shortageInStocksEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_5376_5']")
	  	private static WebElement  shortageInStocksEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_5376_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  shortageInStocksEFES_closeBtn; 
	   
	   
	   // ShortageInStocks Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_5376_7']")
	  	private static WebElement  shortageInStocksELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_5376_7']")
	  	private static WebElement  shortageInStocksELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_5376_7']")
	  	private static WebElement  shortageInStocksELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_5376_7']")
	  	private static WebElement  shortageInStocksELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_5376_7']")
	  	private static WebElement  shortageInStocksELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_5376_7']")
	  	private static WebElement  shortageInStocksELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5376_7+_0']")
	  	private static WebElement  shortageInStocksELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5376_7+_1']")
	  	private static WebElement  shortageInStocksELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_5376_7+_2']")
	  	private static WebElement  shortageInStocksELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_5376_7']")
	  	private static WebElement  shortageInStocksELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_5376_7']")
	  	private static WebElement  shortageInStocksELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_5376_7']")
	  	private static WebElement  shortageInStocksELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_5376_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  shortageInStocksELV_closeBtn;
	   
	   /////////////////////////////
	   
		// ExcessInStocks External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_2048_5']")
	  	private static WebElement  excessInStocksEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_2048_5']")
	  	private static WebElement  excessInStocksEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_2048_5']")
	  	private static WebElement  excessInStocksEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_2048_5']")
	  	private static WebElement  excessInStocksEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_2048_5']")
	  	private static WebElement  excessInStocksEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_2048_5']")
	  	private static WebElement  excessInStocksEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2048_5+_0']")
	  	private static WebElement  excessInStocksEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2048_5+_1']")
	  	private static WebElement  excessInStocksEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2048_5+_2']")
	  	private static WebElement  excessInStocksEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_2048_5']")
	  	private static WebElement  excessInStocksEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_2048_5']")
	  	private static WebElement  excessInStocksEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_2048_5']")
	  	private static WebElement  excessInStocksEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_2048_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  excessInStocksEFES_closeBtn; 
	   
	   
	   // ExcessInStocks Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_2048_7']")
	  	private static WebElement  excessInStocksELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_2048_7']")
	  	private static WebElement  excessInStocksELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_2048_7']")
	  	private static WebElement  excessInStocksELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_2048_7']")
	  	private static WebElement  excessInStocksELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_2048_7']")
	  	private static WebElement  excessInStocksELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_2048_7']")
	  	private static WebElement  excessInStocksELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2048_7+_0']")
	  	private static WebElement  excessInStocksELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2048_7+_1']")
	  	private static WebElement  excessInStocksELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_2048_7+_2']")
	  	private static WebElement  excessInStocksELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_2048_7']")
	  	private static WebElement  excessInStocksELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_2048_7']")
	  	private static WebElement  excessInStocksELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_2048_7']")
	  	private static WebElement  excessInStocksELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_2048_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  excessInStocksELV_closeBtn;
	   
	   ///////////////////////
	   
		// StockTransfers External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3072_5']")
	  	private static WebElement  stockTransfersEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3072_5']")
	  	private static WebElement  stockTransfersEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3072_5']")
	  	private static WebElement  stockTransfersEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3072_5']")
	  	private static WebElement  stockTransfersEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3072_5']")
	  	private static WebElement  stockTransfersEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3072_5']")
	  	private static WebElement  stockTransfersEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3072_5+_0']")
	  	private static WebElement  stockTransfersEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3072_5+_1']")
	  	private static WebElement  stockTransfersEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3072_5+_2']")
	  	private static WebElement  stockTransfersEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3072_5']")
	  	private static WebElement  stockTransfersEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3072_5']")
	  	private static WebElement  stockTransfersEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3072_5']")
	  	private static WebElement  stockTransfersEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3072_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  stockTransfersEFES_closeBtn; 
	   
	   
	   // StockTransfers Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_3072_7']")
	  	private static WebElement  stockTransfersELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_3072_7']")
	  	private static WebElement  stockTransfersELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_3072_7']")
	  	private static WebElement  stockTransfersELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_3072_7']")
	  	private static WebElement  stockTransfersELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_3072_7']")
	  	private static WebElement  stockTransfersELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_3072_7']")
	  	private static WebElement  stockTransfersELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3072_7+_0']")
	  	private static WebElement  stockTransfersELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3072_7+_1']")
	  	private static WebElement  stockTransfersELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_3072_7+_2']")
	  	private static WebElement  stockTransfersELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_3072_7']")
	  	private static WebElement  stockTransfersELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_3072_7']")
	  	private static WebElement  stockTransfersELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_3072_7']")
	  	private static WebElement  stockTransfersELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_3072_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  stockTransfersELV_closeBtn;
	   
	   
	   /////////////////////////
	   
		// OpeningStocks External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_512_5']")
	  	private static WebElement  openingStocksEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_512_5']")
	  	private static WebElement  openingStocksEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_512_5']")
	  	private static WebElement  openingStocksEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_512_5']")
	  	private static WebElement  openingStocksEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_512_5']")
	  	private static WebElement  openingStocksEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_512_5']")
	  	private static WebElement  openingStocksEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_512_5+_0']")
	  	private static WebElement  openingStocksEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_512_5+_1']")
	  	private static WebElement  openingStocksEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_512_5+_2']")
	  	private static WebElement  openingStocksEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_512_5']")
	  	private static WebElement  openingStocksEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_512_5']")
	  	private static WebElement  openingStocksEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_512_5']")
	  	private static WebElement  openingStocksEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_512_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  openingStocksEFES_closeBtn; 
	   
	   
	   // OpeningStocks Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_512_7']")
	  	private static WebElement  openingStocksELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_512_7']")
	  	private static WebElement  openingStocksELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_512_7']")
	  	private static WebElement  openingStocksELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_512_7']")
	  	private static WebElement  openingStocksELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_512_7']")
	  	private static WebElement  openingStocksELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_512_7']")
	  	private static WebElement  openingStocksELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_512_7+_0']")
	  	private static WebElement  openingStocksELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_512_7+_1']")
	  	private static WebElement  openingStocksELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_512_7+_2']")
	  	private static WebElement  openingStocksELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_512_7']")
	  	private static WebElement  openingStocksELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_512_7']")
	  	private static WebElement  openingStocksELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_512_7']")
	  	private static WebElement  openingStocksELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_512_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  openingStocksELV_closeBtn;
	   
	   /////////////////////////////
	   
		// MaterialRequisition External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_7936_5']")
	  	private static WebElement  materialRequisitionEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_7936_5']")
	  	private static WebElement  materialRequisitionEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_7936_5']")
	  	private static WebElement  materialRequisitionEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_7936_5']")
	  	private static WebElement  materialRequisitionEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_7936_5']")
	  	private static WebElement  materialRequisitionEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_7936_5']")
	  	private static WebElement  materialRequisitionEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7936_5+_0']")
	  	private static WebElement  materialRequisitionEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7936_5+_1']")
	  	private static WebElement  materialRequisitionEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7936_5+_2']")
	  	private static WebElement  materialRequisitionEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_7936_5']")
	  	private static WebElement  materialRequisitionEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_7936_5']")
	  	private static WebElement  materialRequisitionEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_7936_5']")
	  	private static WebElement  materialRequisitionEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_7936_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  materialRequisitionEFES_closeBtn; 
	   
	   
	   // MaterialRequisition Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_7936_7']")
	  	private static WebElement  materialRequisitionELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_7936_7']")
	  	private static WebElement  materialRequisitionELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_7936_7']")
	  	private static WebElement  materialRequisitionELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_7936_7']")
	  	private static WebElement  materialRequisitionELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_7936_7']")
	  	private static WebElement  materialRequisitionELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_7936_7']")
	  	private static WebElement  materialRequisitionELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7936_7+_0']")
	  	private static WebElement  materialRequisitionELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7936_7+_1']")
	  	private static WebElement  materialRequisitionELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_7936_7+_2']")
	  	private static WebElement  materialRequisitionELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_7936_7']")
	  	private static WebElement  materialRequisitionELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_7936_7']")
	  	private static WebElement  materialRequisitionELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_7936_7']")
	  	private static WebElement  materialRequisitionELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_7936_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  materialRequisitionELV_closeBtn;
	   
	   
	   //////////////////////////////////////
	   
	   // stockAdjustment External Functions/ Edit Screen---EFES
		
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_1023_5']")
	  	private static WebElement  stockAdjustmentEFES_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_1023_5']")
	  	private static WebElement  stockAdjustmentEFES_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_1023_5']")
	  	private static WebElement  stockAdjustmentEFES_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_1023_5']")
	  	private static WebElement  stockAdjustmentEFES_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_1023_5']")
	  	private static WebElement  stockAdjustmentEFES_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_1023_5']")
	  	private static WebElement  stockAdjustmentEFES_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1023_5+_0']")
	  	private static WebElement  stockAdjustmentEFES_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1023_5+_1']")
	  	private static WebElement  stockAdjustmentEFES_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1023_5+_2']")
	  	private static WebElement  stockAdjustmentEFES_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_1023_5']")
	  	private static WebElement  stockAdjustmentEFES_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_1023_5']")
	  	private static WebElement  stockAdjustmentEFES_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_1023_5']")
	  	private static WebElement  stockAdjustmentEFES_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_1023_5']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  stockAdjustmentEFES_closeBtn; 
	   
	   
	   // StockAdjustment Edit Layout/ Views-----ELV
	   
	   @FindBy(xpath="//select[@id='cmbOnEvent_VT_1023_7']")
	  	private static WebElement  stockAdjustmentELV_onEventDropdown; 

	   @FindBy(xpath="//input[@id='txtButtonCaption_VT_1023_7']")
	  	private static WebElement  stockAdjustmentELV_ButtonCaptionTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbModuleType_VT_1023_7']")
	  	private static WebElement  stockAdjustmentELV_ModuleTypeDropdown;
			
	   @FindBy(xpath="//input[@id='txtModuleName_VT_1023_7']")
	  	private static WebElement  stockAdjustmentELV_ModuleNameOrURlTxt; 
	   
	   @FindBy(xpath="//select[@id='cmbClassname_VT_1023_7']")
	  	private static WebElement  stockAdjustmentELV_ClassNameDropdown;
			
	   @FindBy(xpath="//select[@id='cmbFunctionName_VT_1023_7']")
	  	private static WebElement  stockAdjustmentELV_FunctionNameDropdown; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1023_7+_0']")
	  	private static WebElement  stockAdjustmentELV_AvailableInToolBarRadio;
			
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1023_7+_1']")
	  	private static WebElement  stockAdjustmentELV_AvailableOnScreenRadio; 
	   
	   @FindBy(xpath="//input[@id='dvAvailable_VT_1023_7+_2']")
	  	private static WebElement  stockAdjustmentELV_BothRadio;
			
	   @FindBy(xpath="//input[@id='chkExecuteOnServer_VT_1023_7']")
	  	private static WebElement  stockAdjustmentELV_ExecuteOnServerChkBox; 
	   
	   @FindBy(xpath="//input[@id='btnInsert_VT_1023_7']")
	  	private static WebElement  stockAdjustmentELV_InsertBtn;
			
	   @FindBy(xpath="//input[@id='btnRemove_VT_1023_7']")
	  	private static WebElement  stockAdjustmentELV_RemoveBtn; 
	   
	   @FindBy(xpath="//div[@id='dvDLLS_VT_1023_7']//button[@class='btn btn-default'][contains(text(),'Close')]")
	  	private static WebElement  stockAdjustmentELV_closeBtn;
	   
	   
	   
	   //////////////////////////////////////////////////////////////////////
	   
	   
	   
	   //Grid Elements
	   
	   @FindBy(xpath="//section[@id='page_Content']//tr[1]//td[1]//input[1]")
	  	private static WebElement  selectRow1ChkBox; 
	   
	   @FindBy(xpath="//section[@id='page_Content']//tr[2]//td[1]//input[1]")
	  	private static WebElement  selectRow2ChkBox; 
	   
	   ////////////////////////////////
	   
	   @FindBy(xpath="//a[contains(text(),'antlr.runtime.dll')]")
	  	private static WebElement  antlrruntimedllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Antlr3.Runtime.dll')]")
	  	private static WebElement  antlr3RuntimedllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Aspose.Pdf.Builder.dll')]")
	  	private static WebElement  asposePdfBuilderdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Aspose.Pdf.dll')]")
	  	private static WebElement  asposePdfdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Common.Logging.dll')]")
	  	private static WebElement  commonLoggingdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Common.Logging.Log4Net.dll')]")
	  	private static WebElement  commonLoggingLog4NetdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'DocumentFormat.OpenXml.dll')]")
	  	private static WebElement  documentFormatOpenXmldllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'EPPlus.dll')]")
	  	private static WebElement  ePPlusdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Facebook.dll')]")
	  	private static WebElement  facebookdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'FControls.dll')]")
	  	private static WebElement  fControlsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.AccountMgmt.BL.dll')]")
	  	private static WebElement  focusAccountMgmtBLdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.AccountMgmt.DataStructs.dll')]")
	  	private static WebElement  focusAccountMgmtDataStructsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Common.BL.dll')]")
	  	private static WebElement  focusCommonBLdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Common.DataStructs.dll')]")
	  	private static WebElement  focusCommonDataStructsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Company.BL.dll')]")
	  	private static WebElement  focusCompanyBLdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Company.DataStructs.dll')]")
	  	private static WebElement  focusCompanyDataStructsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.CoreMasters.BL.dll')]")
	  	private static WebElement  focusCoreMastersBLdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.CoreMasters.DataStructs.dll')]")
	  	private static WebElement  focusCoreMastersDataStructsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.CoreTransaction.BL.dll')]")
	  	private static WebElement  focusCoreTransactionBLdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.CoreTransaction.DataStructs.dll')]")
	  	private static WebElement  focusCoreTransactionDataStructsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.CreditManagement.BL.dll')]")
	  	private static WebElement  focusCreditManagementBLdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.CreditManagement.DataStructs.dll')]")
	  	private static WebElement  focusCreditManagementDataStructsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.DashBoard.BL.dll')]")
	  	private static WebElement  focusDashBoardBLdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.DashBoard.DataStructs.dll')]")
	  	private static WebElement  focusDashBoardDataStructsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.DatabaseFactory.dll')]")
	  	private static WebElement  focusDatabaseFactorydllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.DataLayer.dll')]")
	  	private static WebElement  focusDataLayerdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.DataManagement.BL.dll')]")
	  	private static WebElement  focusDataManagementBLdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.DataManagement.DataStructs.dll')]")
	  	private static WebElement  focusDataManagementDataStructsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.DBMnt.BL.dll')]")
	  	private static WebElement  focusDBMntBLdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Export.dll')]")
	  	private static WebElement  focusExportdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.ExternalCall.BL.dll')]")
	  	private static WebElement  focusExternalCallBLdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.ExternalPosting.BL.dll')]")
	  	private static WebElement  focusExternalPostingBLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.ExternalPosting.DataStructs.dll')]")
	  	private static WebElement  focusExternalPostingDataStructsdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.FixedAssets.BL.dll')]")
	  	private static WebElement  focusFixedAssetsBLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.FixedAssets.DataStructs.dll')]")
	  	private static WebElement  focusFixedAssetsDataStructsdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.GST.BL.DLL')]")
	  	private static WebElement  focusGSTBLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.GST.DataStructs.DLL')]")
	  	private static WebElement  focusGSTDataStructsdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Home.BL.dll')]")
	  	private static WebElement  focusHomeBLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Home.DataStructs.dll')]")
	  	private static WebElement  focusHomeDataStructsdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.HRMS.DS.DLL')]")
	  	private static WebElement  focusHRMSDSdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.HRMS.Infrastructure.dll')]")
	  	private static WebElement  focusHRMSInfrastructuredllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Maintenance.BL.dll')]")
	  	private static WebElement  focusMaintenanceBLdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Maintenance.DataStructs.dll')]")
	  	private static WebElement  focusMaintenanceDataStructsdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Masters.BL.dll')]")
	  	private static WebElement  focusMastersBLdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Masters.DataStructs.dll')]")
	  	private static WebElement  focusMastersDataStructsdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Migration.BL.dll')]")
	  	private static WebElement  focusMigrationBLdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.MRP.DataStructs.dll')]")
	  	private static WebElement  focusMRPDataStructsdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.MRP8.BL.dll')]")
	  	private static WebElement  focusMRP8BLdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.MRP8.DataStructs.dll')]")
	  	private static WebElement  focusMRP8DataStructsdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Organization.BL.dll')]")
	  	private static WebElement  focusOrganizationBLdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Organization.DataStructs.dll')]")
	  	private static WebElement  focusOrganizationDataStructsdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Pay.Datastructs.dll')]")
	  	private static WebElement  focusPayDatastructsdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Pos.BusinessLayer.dll')]")
	  	private static WebElement  focusPosBusinessLayerdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Pos.BusinessObjects.dll')]")
	  	private static WebElement  focusPosBusinessObjectsdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Pos.Common.dll')]")
	  	private static WebElement  focusPosCommondllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Pos.DataAccessLayer.dll')]")
	  	private static WebElement  focusPosDataAccessLayerdllBtn;
			
	   @FindBy(xpath="//a[contains(text(),'Focus.Printing.dll')]")
	  	private static WebElement  focusPrintingdllBtn;
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Pronghorn.BL.dll')]")
	  	private static WebElement  focusPronghornBLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.QC.BL.dll')]")
	  	private static WebElement  focusQCBLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.QC.DataStructs.dll')]")
	  	private static WebElement  focusQCDataStructsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.RD.BL.dll')]")
	  	private static WebElement  focusRDBLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.RD.DataStructs.dll')]")
	  	private static WebElement  focusRDDataStructsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Security.BL.dll')]")
	  	private static WebElement  focusSecurityBLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Security.DataStructs.dll')]")
	  	private static WebElement  focusSecurityDataStructsdllBtn; 
	      
	   @FindBy(xpath="//a[contains(text(),'Focus.Security.Web.dll')]")
	  	private static WebElement  focusSecurityWebdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Session.BL.dll')]")
	  	private static WebElement  focusSessionBLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Synchronization.BL.dll')]")
	  	private static WebElement  focusSynchronizationBLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Synchronization.DataStructs.dll')]")
	  	private static WebElement  focusSynchronizationDataStructsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Tds.BL.dll')]")
	  	private static WebElement  focusTdsBLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Transactions.BL.dll')]")
	  	private static WebElement  focusTransactionsBLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Transactions.DataStructs.dll')]")
	  	private static WebElement  focusTransactionsDataStructsdllBtn; 
	      
	   @FindBy(xpath="//a[contains(text(),'Focus.TranSettings.BL.dll')]")
	  	private static WebElement  focusTranSettingsBLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.TranSettings.DataStructs.dll')]")
	  	private static WebElement  focusTranSettingsDataStructsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Web.Mvc.Controls.dll')]")
	  	private static WebElement  focusWebMvcControlsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Web.Mvc.dll')]")
	  	private static WebElement  focusWebMvcdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Web.Mvc.Resources.dll')]")
	  	private static WebElement  focusWebMvcResourcesdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.Web.Mvc.Session.dll')]")
	  	private static WebElement  focusWebMvcSessiondllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus.WMS.BL.dll')]")
	  	private static WebElement  focusWMSBLdllBtn; 
	      
	   @FindBy(xpath="//a[contains(text(),'Focus.WMS.DataStructs.dll')]")
	  	private static WebElement  focusWMSDataStructsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Focus8W.dll')]")
	  	private static WebElement  focus8WdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'FocusWpfControls.dll')]")
	  	private static WebElement  focusWpfControlsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'FusionCharts.dll')]")
	  	private static WebElement  fusionChartsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'FWebControls.dll')]")
	  	private static WebElement  fWebControlsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Google.Apis.Auth.dll')]")
	  	private static WebElement  googleApisAuthdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Google.Apis.Auth.PlatformServices.dll')]")
	  	private static WebElement  googleApisAuthPlatformServicesdllBtn; 
	      
	   @FindBy(xpath="//a[contains(text(),'Google.Apis.Core.dll')]")
	  	private static WebElement  googleApisCoredllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Google.Apis.dll')]")
	  	private static WebElement  googleApisdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Google.Apis.PlatformServices.dll')]")
	  	private static WebElement  googleApisPlatformServicesdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Hammock.dll')]")
	  	private static WebElement  hammockdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'ICSharpCode.SharpZipLib.dll')]")
	  	private static WebElement  iCSharpCodeSharpZipLibdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Ionic.Zip.dll')]")
	  	private static WebElement  ionicZipdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'itextsharp.dll')]")
	  	private static WebElement  itextsharpdllBtn; 
	      
	   @FindBy(xpath="//a[contains(text(),'log4net.dll')]")
	  	private static WebElement  log4netdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Microsoft.Activities.Build.dll')]")
	  	private static WebElement  microsoftActivitiesBuilddllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Microsoft.AspNet.SignalR.Core.dll')]")
	  	private static WebElement  microsoftAspNetSignalRCoredllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Microsoft.AspNet.SignalR.SystemWeb.dll')]")
	  	private static WebElement  microsoftAspNetSignalRSystemWebdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Microsoft.Office.Interop.Outlook.DLL')]")
	  	private static WebElement  microsoftOfficeInteropOutlookdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Microsoft.Owin.dll')]")
	  	private static WebElement  microsoftOwindllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Microsoft.Owin.Host.SystemWeb.dll')]")
	  	private static WebElement  microsoftOwinHostSystemWebdllBtn; 
	      
	   @FindBy(xpath="//a[contains(text(),'Microsoft.Owin.Security.Cookies.dll')]")
	  	private static WebElement  microsoftOwinSecurityCookiesdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Microsoft.Owin.Security.dll')]")
	  	private static WebElement  microsoftOwinSecuritydllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Microsoft.Practices.EnterpriseLibrary.Common.dll')]")
	  	private static WebElement  microsoftPracticesEnterpriseLibraryCommondllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Microsoft.Practices.EnterpriseLibrary.Data.dll')]")
	  	private static WebElement  microsoftPracticesEnterpriseLibraryDatadllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Microsoft.Practices.ObjectBuilder2.dll')]")
	  	private static WebElement  microsoftPracticesObjectBuilder2dllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Microsoft.Practices.Unity.dll')]")
	  	private static WebElement  microsoftPracticesUnitydllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Microsoft.VisualStudio.QualityTools.UnitTestFramew')]")
	  	private static WebElement  microsoftVisualStudioQualityToolsUnitTestFramew; 
	    
	   @FindBy(xpath="//a[contains(text(),'Microsoft.Web.Infrastructure.dll')]")
	  	private static WebElement  microsoftWebInfrastructuredllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'MVC4ControlsToolkit.dll')]")
	  	private static WebElement  mVvc4ControlsToolkitdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'MVCControlsToolkit.Business.dll')]")
	  	private static WebElement  mvcControlsToolkitBusinessdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'MySqlSL.dll')]")
	  	private static WebElement  mySqlSLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Newtonsoft.Json.dll')]")
	  	private static WebElement  newtonsoftJsondllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'NPOI.dll')]")
	  	private static WebElement  npoidllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'NPOI.Extension.dll')]")
	  	private static WebElement  npoiExtensiondllBtn; 
	      
	   @FindBy(xpath="//a[contains(text(),'NPOI.OOXML.dll')]")
	  	private static WebElement  npoiOOXMLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'NPOI.OpenXml4Net.dll')]")
	  	private static WebElement  npoiOpenXml4NetdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'NPOI.OpenXmlFormats.dll')]")
	  	private static WebElement  npoiOpenXmlFormatsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'office.DLL')]")
	  	private static WebElement  officedllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'OracleSL.dll')]")
	  	private static WebElement  oracleSLdllBtn; 
	   
	   @FindBy(xpath="//a[119]")
	  	private static WebElement  owindllBtn; 
	   
	   @FindBy(xpath="//a[120]")
	  	private static WebElement  resourcesdll1Btn; 
	      
	   @FindBy(xpath="//a[contains(text(),'ResPOS.dll')]")
	  	private static WebElement  resPOSdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Spring.Core.dll')]")
	  	private static WebElement  springCoredllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Spring.Web.dll')]")
	  	private static WebElement  springWebdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Spring.Web.Mvc3.dll')]")
	  	private static WebElement  springWebMvc3dllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'SqlServerSL.dll')]")
	  	private static WebElement  sqlServerSLdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'SRVTextToImage.dll')]")
	  	private static WebElement  srvTextToImagedllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'StackExchange.Redis.dll')]")
	  	private static WebElement  stackExchangeRedisdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'stdole.dll')]")
	  	private static WebElement  stdoledllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'SyntaxLayer.dll')]")
	  	private static WebElement  syntaxLayerdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'System.ComponentModel.Composition.Registration.dll')]")
	  	private static WebElement  systemComponentModelCompositionRegistrationdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'System.IO.Compression.dll')]")
	  	private static WebElement  systemIOCompressiondllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'System.IO.Compression.FileSystem.dll')]")
	  	private static WebElement  systemIOCompressionFileSystemdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'System.Net.Http.dll')]")
	  	private static WebElement  systemNetHttpdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'System.Net.Http.WebRequest.dll')]")
	  	private static WebElement  systemNetHttpWebRequestdllBtn; 
	      
	   @FindBy(xpath="//a[contains(text(),'System.Reflection.Context.dll')]")
	  	private static WebElement  systemReflectionContextdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'System.ServiceModel.Internals.dll')]")
	  	private static WebElement  systemServiceModelInternalsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'System.Web.Helpers.dll')]")
	  	private static WebElement  systemWebHelpersdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'System.Web.Mvc.dll')]")
	  	private static WebElement  systemWebMvcdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'System.Web.Optimization.dll')]")
	  	private static WebElement  systemWebOptimizationdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'System.Web.Razor.dll')]")
	  	private static WebElement  systemWebRazordllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'System.Web.WebPages.Deployment.dll')]")
	  	private static WebElement  systemWebWebPagesDeploymentdllBtn; 
	      
	   @FindBy(xpath="//a[contains(text(),'System.Web.WebPages.dll')]")
	  	private static WebElement  systemWebWebPagesdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'System.Web.WebPages.Razor.dll')]")
	  	private static WebElement  systemWebWebPagesRazordllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'System.Windows.dll')]")
	  	private static WebElement  systemWindowsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'System.Xml.Serialization.dll')]")
	  	private static WebElement  systemXmlSerializationdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'Transitionals.dll')]")
	  	private static WebElement  transitionalsdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'TweetSharp.dll')]")
	  	private static WebElement  tweetSharpdllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'WebActivator.dll')]")
	  	private static WebElement  webActivatordllBtn; 
	      
	   @FindBy(xpath="//a[contains(text(),'WebGrease.dll')]")
	  	private static WebElement  webGreasedllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'WebMatrix.Data.dll')]")
	  	private static WebElement  webMatrixDatadllBtn; 
	   
	   @FindBy(xpath="//a[contains(text(),'WebMatrix.WebData.dll')]")
	  	private static WebElement  webMatrixWebDatadllBtn; 
	   
	   @FindBy(xpath="//a[152]")
	  	private static WebElement  resourcesdll2Btn; 
	   
	   @FindBy(xpath="//a[153]")
	  	private static WebElement  resourcesdll3Btn; 
	   
	   
		
	 //**---------------------------------------------------------------------------------------------	
	 	
	 // Rules Tab	
	 	
	 	
	 	@FindBy(xpath="//span[contains(text(),'Rules')]")
	 	private static WebElement  rulesTab;

	 	@FindBy(xpath="//input[@class='btn btn-link']")
	 	private static WebElement clickhereToAddRuleBtn;
	 	
	 	@FindBy(xpath="//input[@id='txtMasterRuleName_DocCustRules']")
	 	private static WebElement  ruleNameTxt;
	 	
	 	@FindBy(xpath="//span[@class='optioncontrol_arrow_margin icon-down-arrow']")
	 	private static WebElement  ruleNameDropdown;

	     @FindBy(xpath="//*[@id='txtMasterRuleName_DocCustRules_input_settings']/span")
	 	private static WebElement  ruleNameSetingsBtn ;
	         
	         ////////////*****************pop
	         
	         @FindBy(xpath="//div[@id='txtMasterRuleName_DocCustRules_customize_popup_container']")
	     	private static WebElement ruleNameContainer;
	      
	     	   @FindBy(xpath="//div[@id='txtMasterRuleName_DocCustRules_customize_popup_footer']//input[1]")
	     	   private static WebElement ruleNameStandardFieldsBtn;
	     	   
	     				     @FindBy(xpath="//select[@id='txtMasterRuleName_DocCustRules_customize_popup_standardfields_list']")
	     				     private static WebElement ruleNameStandardField_Fielddropdown;
	     				
	     				     @FindBy(xpath="//input[@id='txtMasterRuleName_DocCustRules_customize_popup_standardfields_header']")
	     				     private static WebElement ruleNameStandardField_HeaderText;
	     				
	     				     @FindBy(xpath="//select[@id='txtMasterRuleName_DocCustRules_customize_popup_standardfields_alignment']")
	     				     private static WebElement ruleNameStandardField_AlignmentDropdown;
	     				
	     				     @FindBy(xpath="//input[@id='txtMasterRuleName_DocCustRules_customize_popup_standardfields_width']")
	     				     private static WebElement ruleNameStandardField_WidthText;
	     				
	     				     @FindBy(xpath="//section[@id='page_Content']//input[3]")
	     				     private static WebElement ruleNameStandardField_okBtn;
	     				
	     				     @FindBy(xpath="//section[@id='page_Content']//input[4]")
	     				     private static WebElement ruleNameStandardField_CanceBtn;
	     	
	 	    				     
	 	    	  @FindBy(xpath="//div[@id='txtMasterRuleName_DocCustRules_search_container']//input[2]")
	 	    	  private static WebElement ruleNameDeleteColoumnBtn;
	 	    	 
	 	    	  @FindBy(xpath="//section[@id='page_Content']//input[3]")
	 	    	  private static WebElement ruleNameokBtn;
	 	    	 
	 	    	  @FindBy(xpath="//section[@id='page_Content']//input[4]")
	 	    	  private static WebElement ruleNamecancelBtn;
	 	
	 	@FindBy(xpath="//input[@id='chkNewRecord_Rule_DocCustRules']")
	 	private static WebElement rulesNewRecordChekbox ;
	     
	     @FindBy(xpath="//input[@id='chkEdit_Rule_DocCustRules']")
	 	private static WebElement rulesEditChekbox ;
	     
	     @FindBy(xpath="//input[@id='chkLoad_Rule_DocCustRules']")
	 	private static WebElement rulesLoadChekbox ;
	   
	     @FindBy(xpath="//input[@id='chkOnLeave_Rule_DocCustRules']")
	 	private static WebElement rulesOnLeaveChekbox ;
	     
	     @FindBy(xpath="//input[@id='chkBeforeDelete_Rule_DocCustRules']")
	 	private static WebElement rulesBeforeDeleteChekbox ;
	     
	     @FindBy(xpath="//input[@id='chkOnEnter_Rule_DocCustRules']")
	 	private static WebElement rulesOnEnterChekbox ;
	     
	     @FindBy(xpath="//input[@id='chkIsRuleActive_Rule_DocCustRules']")
	 	private static WebElement rulesActiveChekbox ;
	     
	     @FindBy(xpath="//input[@id='chkNoRuleCondition_DocCustRules']")
	 	private static WebElement rulesNoConditionChekbox;
	      
	      
	     //
	     @FindBy(xpath="//li[@class='IF-Tab']//a")
	 	private static WebElement rulesIfTab ;
	     
	     @FindBy(xpath="//li[@class='ELSE-Tab']//a")
	 	private static WebElement rulesElseTab ;
	     
	     // if 
	 	
	     @FindBy(xpath="//input[@id='txtField_RC_1_1_DocCustRules']")
	 	private static WebElement rulesIf_SelectFieldDropdown ;
	     
	     @FindBy(xpath="//select[@id='ddlOperators_RC_1_1_DocCustRules']")
	 	private static WebElement rulesIf_OPerationTypeDropdown ;
	     
	     @FindBy(xpath="//select[@id='ddlCompareWith_RC_1_1_DocCustRules']")
	 	private static WebElement rulesIf_CompareTypeDropdown;          //Value,Field,Formula---Related
	     
	     @FindBy(xpath="//input[@id='txtCompareType_RC_1_1_DocCustRules']")
	 	private static WebElement rulesIf_ValueTxt;
	     
	     @FindBy(xpath="//select[@id='ddlConjunction_RC_1_1_DocCustRules']")
	 	private static WebElement rulesIf_ConditionDropdown;          //Only,And,Or--RelatedDD
	     
	     
	   //IF **** Customize Formatting Grid

	 	@FindBy(xpath="//*[@id='Rule_IF_DocCustRules']/div[2]/div[1]/ul/li[1]/a/span")
	 	private static WebElement rulesIFformattingTabBtn ;
	     
	 	
	 	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-0']")
	 	private static WebElement rulesIFFormattingRow1Btn;
	 	
	 	@FindBy(xpath="//div[contains(text(),'Select Row')]")
	 	private static WebElement rulesIFFormattingSelectRowBtn;
	 	
	 	@FindBy(xpath="//div[contains(text(),'Delete Row')]")
	 	private static WebElement rulesIFFormattingDeleteRowBtn;
	 	
	 	@FindBy(xpath="//div[contains(text(),'Insert Row')]")
	 	private static WebElement rulesIFFormattingInsertRowBtn;
	 	
	 	
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-1']")
	 	private static WebElement rulesIFFormattingGridFieldName1 ;
	     
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-2']")
	 	private static WebElement rulesIFFormattingGridChangeValue1;
	     
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-3']")
	 	private static WebElement rulesIFFormattingGridField1;
	     
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-4']")
	 	private static WebElement rulesIFFormattingGridValue1;
	     
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-5']")
	 	private static WebElement rulesIFFormattingGridCaption1;
	     
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-6']")
	 	private static WebElement rulesIFFormattingGridRegularExpression1;
	     
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-7']")
	 	private static WebElement rulesIFFormattingGridErrorMesage1;
	     
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-8']")
	 	private static WebElement rulesIFFormattingGridBannerText1;
	     
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-9']")
	 	private static WebElement rulesIFFormattingGridToolTip1;
	     
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-10']")
	 	private static WebElement rulesIFFormattingGridmandatory1;
	     
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-11']")
	 	private static WebElement rulesIFFormattingGridHidden1;
	     
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-12']")
	 	private static WebElement rulesIFFormattingGridDisable1;
	     
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-13']")
	 	private static WebElement rulesIFFormattingGridSpellCheck1;
	     
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-14']")
	 	private static WebElement rulesIFFormattingGridAligement1;
	     
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-15']")
	 	private static WebElement rulesIFFormattingGridCharacterCasting1;
	     
	     @FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-16']")
	 	private static WebElement rulesIFFormattingGridSelectFont1;
	     
	     
	    //
	     @FindBy(xpath="//select[@id='ddlFieldsDiv_Rule_IF_DocCustRules']")
	 	private static WebElement rulesIFFormattingGridFieldNameDropdown ;
	     
	     @FindBy(xpath="//td[@id='ddlChangeValueFormat_IF_DocCustRules']")
	 	private static WebElement rulesIFFormattingGridChangeValueDropdown;
	     
	     @FindBy(xpath="//input[@id='txtCaptionFormat_IF_DocCustRules']")
	 	private static WebElement rulesIFFormattingGridCaptionTxt;
	     
	     @FindBy(xpath="//input[@id='txtRegExFormat_IF_DocCustRules']")
	 	private static WebElement rulesIFFormattingGridRegularExpressionTxt;
	     
	     @FindBy(xpath="//input[@id='txtErrorMsgForamt_IF_DocCustRules']")
	 	private static WebElement rulesIFFormattingGridErrorMesageTxt;
	     
	     @FindBy(xpath="//input[@id='txtBannertextFormat_IF_DocCustRules']")
	 	private static WebElement rulesIFFormattingGridBannerTextTxt;
	     
	     @FindBy(xpath="//input[@id='txtTooltipFormat_IF_DocCustRules']")
	 	private static WebElement rulesIFFormattingGridToolTipTxt;
	     
	     @FindBy(xpath="//input[@id='chkMandatoryFormat_IF_DocCustRules']")
	 	private static WebElement rulesIFFormattingGridmandatorychkBox;
	     
	     @FindBy(xpath="//input[@id='chkHiddenFormat_IF_DocCustRules']")
	 	private static WebElement rulesIFFormattingGridhiddenChkbox;
	     
	     @FindBy(xpath="//input[@id='chkDisableFormat_IF_DocCustRules']")
	 	private static WebElement rulesIFFormattingGridDisableChkBox;
	     
	     @FindBy(xpath="//input[@id='chkSpecllCheckFormat_IF_DocCustRules']")
	 	private static WebElement rulesIFFormattingGridspellCheckChkBox;
	     
	     @FindBy(xpath="//select[@id='ddlAlignmentFormat_IF_DocCustRules']")
	 	private static WebElement rulesIFFormattingGridAligementDropdown;
	     
	     @FindBy(xpath="//select[@id='ddlCharCasingFormat_IF_DocCustRules']")
	 	private static WebElement rulesIFFormattingGridCharacterCastingDropdown;
	     
	     @FindBy(xpath="//button[@id='btnFontFormat_IF_DocCustRules']")
	 	private static WebElement rulesIFFormattingGridSelectFontBtn;
	     
	                          
	     
	     			@FindBy(xpath="//*[@id='myFontControl_IF_Save']")
	     			private static WebElement fontIfsaveBtn;

	     			@FindBy(xpath="//*[@id='myFontControl_IF_Close']")
	 				private static WebElement fontIfcloseBtn;
	 			    
	     			@FindBy(xpath="//*[@id='fontFamily_ctrlFontControl_IF']")
	 			    private static WebElement fontIfFontFamilyDropdown;
	     
	 			    @FindBy(xpath="//*[@id='fontStyle_ctrlFontControl_IF']")
	 			    private static WebElement fontIfFontStyleDropdown;
	     
	 			    @FindBy(xpath="//*[@id='fontWeight_ctrlFontControl_IF']")
	 			    private static WebElement fontIfFontWeightDropdown;
	  
	 				@FindBy(xpath="//*[@id='fontSizes_ctrlFontControl_IF']")
	 				private static WebElement fontIfFontSizeDropdown;
	 				
	 				@FindBy(xpath="//*[@id='fontForeColor_ctrlFontControl_IF']")
	 				private static WebElement fontIfForeColourDropdown;
	 				
	 				@FindBy(xpath="//*[@id='fontBackColor_ctrlFontControl']")
	 				private static WebElement fontIfBackColourDropdown;
	 				
	 				@FindBy(xpath="//*[@id='chkBaseline_ctrlFontControl']")
	 				private static WebElement fontIfBaseLineCheckbox;
	 				
	 				@FindBy(xpath="//*[@id='chkOverLine_ctrlFontControl_IF']")
	 				private static WebElement fontIfOverLineCheckbox;
	 				
	 				@FindBy(xpath="//*[@id='chkStrikeThrough_ctrlFontControl_IF']")
	 				private static WebElement fontIfStrikeThroughCheckbox;
	 				
	 				@FindBy(xpath="//*[@id='chkUnderline_ctrlFontControl_IF']")
	 				private static WebElement fontIfUnderlineCheckbox;
	 				
	 				@FindBy(xpath="//*[@id='previewfont_ctrlFontControl_IF']")
	 				private static WebElement fontIfFontPreviewTxt;
	 				
	 				@FindBy(xpath="//*[@id='chkEnableFont_ctrlFontControl_IF']")
	 				private static WebElement fontIfEnableFontCheckbox;
	     
	 	@FindBy(xpath="//button[@id='btnAddRule']")
	 	private static WebElement saveRuleBtn;
	 	
	 	@FindBy(xpath="//button[contains(text(),'Delete Rule')]")
	 	private static WebElement deleteRuleBtn;
	 	

	 	/// Rules--------IF Tab-------Message

	    @FindBy(xpath="//*[@id='Rule_IF_DocCustRules']/div[2]/div[1]/ul/li[2]/a/span")
	 	private static WebElement rulesIFMessageTab ;
	     
	    @FindBy(xpath="//input[@id='txtGeneralMsg_RuleIfMsg_DocCustRules']")
	 	private static WebElement rulesIFMessagesGenralMessageTxt;
	     
	    @FindBy(xpath="//*[@id='lstRdoMessageType_RuleIF_DocCustRules']/div[1]/div[2]/label/input")
	 	private static WebElement rulesIFMessagesInformationRadio;
	     
	    @FindBy(xpath="//*[@id='lstRdoMessageType_RuleIF_DocCustRules']/div[2]/div[2]/label/input")
	 	private static WebElement rulesIFMessagesWarnAndProceddRadio;
	     
	    @FindBy(xpath="//*[@id='lstRdoMessageType_RuleIF_DocCustRules']/div[3]/div[2]/label/input")
	 	private static WebElement rulesIFMessagesWarnAndStopRadio;
	    
	 	
	 ///// IF Tab ****Alerts-----Grid Elements

	    @FindBy(xpath="//*[@id='Rule_IF_DocCustRules']/div[2]/div[1]/ul/li[3]/a/span")
	 	private static WebElement rulesIFAlertTab ;

	    @FindBy(xpath="//input[@id='txtAlertMsg_RuleIfAlert_DocCustRules']")
	 	private static WebElement rulesIfAlertMessageTxt;
	     
	    @FindBy(xpath="//button[@id='custmizeAlertTemplate_RuleIF_DocCustRules']")
	    private static WebElement rulesIfCustmizeTemplateDropdown;
	     
	     @FindBy(xpath="//td[@id='tblAlert_IF_Rule_DocCustRules_col_1-0']")
	 	private static WebElement rulesIfAlertRow1;
	 	
	 	@FindBy(xpath="//td[@id='tblAlert_IF_Rule_DocCustRules_col_1-1']")
	 	private static WebElement rulesIfAlertType1;
	     
	 	@FindBy(xpath="//td[@id='tblAlert_IF_Rule_DocCustRules_col_1-2']")
	 	private static WebElement rulesIfAlertToWhom1;
	     
	 	@FindBy(xpath="//td[@id='tblAlert_IF_Rule_DocCustRules_col_1-3']")
	 	private static WebElement rulesIfAlertSms1;
	     
	 	@FindBy(xpath="//td[@id='tblAlert_IF_Rule_DocCustRules_col_1-4']")
	 	private static WebElement rulesIfAlertEmail1;
	     
	 	@FindBy(xpath="//td[@id='tblAlert_IF_Rule_DocCustRules_col_1-5']")
	 	private static WebElement rulesIfAlertDashBoard1;
	     
	 	@FindBy(xpath="//td[@id='tblAlert_IF_Rule_DocCustRules_col_1-6']")
	 	private static WebElement rulesIfAlertPopUp1;
	     
	 	@FindBy(xpath="//td[@id='tblAlert_IF_Rule_DocCustRules_col_1-7']")
	 	private static WebElement rulesIfAlertGTalk1;
	     
	 	@FindBy(xpath="//td[@id='tblAlert_IF_Rule_DocCustRules_col_1-8']")
	 	private static WebElement rulesIfAlertSendSepcificMessage1;
	   
	 	@FindBy(xpath="//td[@id='tblAlert_IF_Rule_DocCustRules_col_1-9']")
	 	private static WebElement rulesIfAlertSpecificMesage1;
	 		
	    @FindBy(xpath="//select[@id='ddlTypeAlert_IF_DocCustRules']")
	 	private static WebElement rulesIFAlertsGridTypeDropdown;
	     
	    @FindBy(xpath="//Select[@id='ddlRolesAlert_IF_DocCustRules']")
	 	private static WebElement rulesIFAlertsGridToWhomDropdown;
	     
	    @FindBy(xpath="//select[@id='ddlUsersAlert_IF_DocCustRules']")
	 	private static WebElement rulesIFAlertsGridSmsDropdown;
	     
	    @FindBy(xpath="//select[@id='ddlEmailAlert_IF_DocCustRules']")
	 	private static WebElement rulesIFAlertsGridEmailDropdown;
	     
	    @FindBy(xpath="//select[@id='ddlDashboardAlert_IF_DocCustRules']")
	 	private static WebElement rulesIFAlertsGridDashBoardDropdown;
	     
	    @FindBy(xpath="//select[@id='ddlPopupAlert_IF_DocCustRules']")
	 	private static WebElement rulesIFAlertsGridPopUpDropdown;
	     
	    @FindBy(xpath="//select[@id='ddlGTalkAlert_IF_DocCustRules']")
	 	private static WebElement rulesIFAlertsGridGTalkDropdown;
	     
	    @FindBy(xpath="//input[@id='chkSendMsgAlert_IF_DocCustRules']")
	 	private static WebElement rulesIFAlertsGridSendSpecificMessageChkBox;
	     
	    @FindBy(xpath="//input[@id='txtSpecificMsgAlert_IF_DocCustRules']")
	 	private static WebElement rulesIFAlertsGridSpecificMessageTxt;
	     


	     // Right Click On Page
	     
	    @FindBy(xpath="//a[@id='btnAddNewCondition']")
	  	private static WebElement rulesAddNewConditionBtn;
	      
	    @FindBy(xpath="//a[@id='btnRemoveCondition']")
	  	private static WebElement rulesRemoveConditionBtn;
	      
	     
	    //ELSE **** Customize Formatting Grid
	     
	 	@FindBy(xpath="//div[@id='Rule_Else_DocCustRules']//ul[@class='nav nav-tabs tab-container-div theme_sub_Tab-Dark']//li[1]//a[1]")
	 	private static WebElement rulesELSEformattingTabBtn ;
	     
	 	
	 	@FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-0']")
	 	private static WebElement rulesElseFormattingRow1Btn;
	 	
	 	@FindBy(xpath="//div[contains(text(),'Select Row')]")
	 	private static WebElement rulesElseFormattingSelectRowBtn;
	 	
	 	@FindBy(xpath="//div[contains(text(),'Delete Row')]")
	 	private static WebElement rulesElseFormattingDeleteRowBtn;
	 	
	 	@FindBy(xpath="//div[contains(text(),'Insert Row')]")
	 	private static WebElement rulesElseFormattingInsertRowBtn;
	 	
	 	
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-1']")
	 	private static WebElement rulesElseFormattingGridFieldName1 ;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-2']")
	 	private static WebElement rulesElseFormattingGridChangeValue1;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-3']")
	 	private static WebElement rulesElseFormattingGridField1;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-4']")
	 	private static WebElement rulesElseFormattingGridValue1;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-5']")
	 	private static WebElement rulesElseFormattingGridCaption1;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-6']")
	 	private static WebElement rulesElseFormattingGridRegularExpression1;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-7']")
	 	private static WebElement rulesElseFormattingGridErrorMesage1;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-8']")
	 	private static WebElement rulesElseFormattingGridBannerText1;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-9']")
	 	private static WebElement rulesElseFormattingGridToolTip1;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-10']")
	 	private static WebElement rulesElseFormattingGridmandatory1;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-11']")
	 	private static WebElement rulesElseFormattingGridHidden1;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-12']")
	 	private static WebElement rulesElseFormattingGridDisable1;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-13']")
	 	private static WebElement rulesElseFormattingGridSpellCheck1;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-14']")
	 	private static WebElement rulesElseFormattingGridAligement1;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-15']")
	 	private static WebElement rulesElseFormattingGridCharacterCasting1;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-16']")
	 	private static WebElement rulesElseFormattingGridSelectFont1;
	     
	     
	  
	    @FindBy(xpath="//select[@id='ddlFieldsDiv_Rule_ELSE_DocCustRules']")
	 	private static WebElement rulesElseFormattingGridFieldNameDropdown ;
	     
	    @FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-2']//select[@id='ddlChangeValueFormat_IF_DocCustRules']")
	 	private static WebElement rulesElseFormattingGridChangeValueDropdown;
	     
	    @FindBy(xpath="//input[@id='txtCaptionFormat_Else_DocCustRules']")
	 	private static WebElement rulesElseFormattingGridCaptionTxt;
	     
	    @FindBy(xpath="//input[@id='txtRegExFormat_Else_DocCustRules']")
	 	private static WebElement rulesElseFormattingGridRegularExpressionTxt;
	     
	    @FindBy(xpath="//input[@id='txtErrorMsgForamt_Else_DocCustRules']")
	 	private static WebElement rulesElseFormattingGridErrorMesageTxt;
	     
	    @FindBy(xpath="//input[@id='txtBannertextFormat_Else_DocCustRules']")
	 	private static WebElement rulesElseFormattingGridBannerTextTxt;
	     
	    @FindBy(xpath="//input[@id='txtTooltipFormat_Else_DocCustRules']")
	 	private static WebElement rulesElseFormattingGridToolTipTxt;
	     
	    @FindBy(xpath="//input[@id='chkMandatoryFormat_Else_DocCustRules']")
	 	private static WebElement rulesElseFormattingGridmandatorychkBox;
	     
	    @FindBy(xpath="//input[@id='chkHiddenFormat_Else_DocCustRules']")
	 	private static WebElement rulesElseFormattingGridhiddenChkbox;
	     
	    @FindBy(xpath="//input[@id='chkDisableFormat_Else_DocCustRules']")
	 	private static WebElement rulesElseFormattingGridDisableChkBox;
	     
	    @FindBy(xpath="//input[@id='chkSpecllCheckFormat_Else_DocCustRules']")
	 	private static WebElement rulesElseFormattingGridspellCheckChkBox;
	     
	    @FindBy(xpath="//select[@id='ddlAlignmentFormat_Else_DocCustRules']")
	 	private static WebElement rulesElseFormattingGridAligementDropdown;
	     
	    @FindBy(xpath="//select[@id='ddlCharCasingFormat_Else_DocCustRules']")
	 	private static WebElement rulesElseFormattingGridCharacterCastingDropdown;
	     
	    @FindBy(xpath="//button[@id='btnFontFormat_Else_DocCustRules']")
	 	private static WebElement rulesElseFormattingGridSelectFontBtn;
	     
	                          
	     
	     			@FindBy(xpath="//*[@id='myFontControl_ELSE_Save']")
	     			private static WebElement fontElsesaveBtn;

	     			@FindBy(xpath="//*[@id='myFontControl_ELSE_Close']")
	 				private static WebElement fontElsecloseBtn;
	 			    
	     			@FindBy(xpath="//*[@id='fontFamily_ctrlFontControl_ELSE']")
	 			    private static WebElement fontElseFontFamilyDropdown;
	     
	 			    @FindBy(xpath="//*[@id='fontStyle_ctrlFontControl_ELSE']")
	 			    private static WebElement fontElseFontStyleDropdown;
	     
	 			    @FindBy(xpath="//*[@id='fontWeight_ctrlFontControl_ELSE']")
	 			    private static WebElement fontElseFontWeightDropdown;
	  
	 				@FindBy(xpath="//*[@id='fontSizes_ctrlFontControl_ELSE']")
	 				private static WebElement fontElseFontSizeDropdown;
	 				
	 				@FindBy(xpath="//*[@id='fontForeColor_ctrlFontControl_ELSE']")
	 				private static WebElement fontElseForeColourDropdown;
	 				
	 				@FindBy(xpath="//*[@id='fontBackColor_ctrlFontControl']")
	 				private static WebElement fontElseBackColourDropdown;
	 				
	 				@FindBy(xpath="//*[@id='chkBaseline_ctrlFontControl']")
	 				private static WebElement fontElseBaseLineCheckbox;
	 				
	 				@FindBy(xpath="//*[@id='chkOverLine_ctrlFontControl_ELSE']")
	 				private static WebElement fontElseOverLineCheckbox;
	 				
	 				@FindBy(xpath="//*[@id='chkStrikeThrough_ctrlFontControl_ELSE']")
	 				private static WebElement fontElseStrikeThroughCheckbox;
	 				
	 				@FindBy(xpath="//*[@id='chkUnderline_ctrlFontControl_ELSE']")
	 				private static WebElement fontElseUnderlineCheckbox;
	 				
	 				@FindBy(xpath="//*[@id='previewfont_ctrlFontControl_ELSE']")
	 				private static WebElement fontElseFontPreviewTxt;
	 				
	 				@FindBy(xpath="//*[@id='chkEnableFont_ctrlFontControl_ELSE']")
	 				private static WebElement fontElseEnableFontCheckbox;
	     
	 	

	 	/// Rules--------ELSE Tab-------Message
	 	
	    @FindBy(xpath="//div[@id='Rule_Else_DocCustRules']//ul[@class='nav nav-tabs tab-container-div theme_sub_Tab-Dark']//li[2]//a[1]")
	 	private static WebElement rulesElseEMessageTab ;
	     
	    @FindBy(xpath="//input[@id='txtGeneralMsg_RuleElseMsg_DocCustRules']")
	 	private static WebElement rulesElseMessagesGenralMessageTxt;
	     
	    @FindBy(xpath="//*[@id='lstRdoMessageType_RuleElse_DocCustRules']/div[1]/div[2]/label/input")
	 	private static WebElement rulesElseMessagesInformationRadio;
	     
	    @FindBy(xpath="//*[@id='lstRdoMessageType_RuleElse_DocCustRules']/div[2]/div[2]/label/input")
	 	private static WebElement rulesElseMessagesWarnAndProceddRadio;
	     
	    @FindBy(xpath="//*[@id='lstRdoMessageType_RuleElse_DocCustRules']/div[3]/div[2]/label/input")
	 	private static WebElement rulesElseMessagesWarnAndStopRadio;
	    
	 	
	     ///// ELSE Tab ****Alerts-----
	     

	    @FindBy(xpath="//div[@id='Rule_Else_DocCustRules']//li[3]//a[1]")
	 	private static WebElement rulesElseAlertTab ;

	    @FindBy(xpath="//input[@id='txtAlertMsg_RuleElseAlert_DocCustRules']")
	 	private static WebElement rulesElseAlertMessageTxt;
	     
	     @FindBy(xpath="//button[@id='custmizeAlertTemplate_RuleELse_DocCustRules']")
	     private static WebElement rulesElseCustmizeTemplateDropdown;
	     
	     @FindBy(xpath="//div[@class='dropup pull-right open']//a[contains(text(),'E-mail Template')]")
	     private static WebElement rulesElseCustmizeTempEmailTemplate;
	     
	     // Pop Up  Window Of EmailTemplate 
	     @FindBy(xpath="//i[@class='icon-font6 icon-new']")
	 	 private static WebElement  emailTempNewdropdown;
	 	
	 				@FindBy(xpath="//*[@id='forHf']/span[1]/ul/li/a[1]")
	 				private static WebElement  emailTempNew_LayoutBtn;
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[1]/ul/li/a[2]")
	 				private static WebElement  emailTempNew_PageBtn;
	 				
	 	@FindBy(xpath="//i[@class='icon-font6 icon-open']")
	 	private static WebElement  emailTempOpendropdown;
	 	
	 				@FindBy(xpath="//*[@id='forHf']/span[2]/ul/li/a[1]")
	 				private static WebElement  emailTempOpenBtn;
	 				
	 							@FindBy(xpath="//*[@id='3']")
	 							private static WebElement  emailTempOpenBtn_First;
	 							
	 							@FindBy(xpath="//*[@id='7']")
	 							private static WebElement  emailTempOpenBtn_Second;
	 							
	 							@FindBy(xpath="//button[@id='btn_PopUpOk']")
	 							private static WebElement  emailTempOpenBtn_okBtn;
	 							
	 							@FindBy(xpath="//*[@id='id_PopUp']/div/div/div[3]/button[2]")
	 							private static WebElement  emailTempOpenBtn_cancelBtn;
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[2]/ul/li/a[2]")
	 				private static WebElement emailTempFromXMLBtn;
	 	
	 	@FindBy(xpath="//i[@class='icon-font6 icon-save']")
	 	private static WebElement  emailTempSavedropdown;
	 	
	 				@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[1]")
	 				private static WebElement  emailTempSaveBtn;
	 				
	 							@FindBy(xpath="//input[@id='id_PopuplayoutName']")
	 							private static WebElement  emailTempSaveTxtField;
	 							
	 							@FindBy(xpath="//input[@id='id_PopuplayoutName']")
	 							private static WebElement  emailTempSave_okBtn;
	 							
	 							@FindBy(xpath="//*[@id='id_PopUp']/div/div/div[3]/button[2]")
	 							private static WebElement  emailTempSave_CancekBtn;
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[2]")
	 				private static WebElement  emailTempSaveAsBtn; //Clicking popup is same as savebtn popup
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[3]")
	 				private static WebElement  emailTempSaveAsImageBtn; //No Response
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[4]")
	 				private static WebElement  emailTempSaveAsXMLBtn; // Downloads xml file
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[23]/ul/li/a[5]")
	 				private static WebElement  emailTempSaveAsHTMLBtn; // Showing Error
	 				
	 	@FindBy(xpath="//*[@id='forHf']/span[4]")
	 	private static WebElement  emailTempRemovedropdown;
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[4]/ul/li/a[1]")
	 				private static WebElement  emailTempRemoveBtn; //Raises popup
	 							
	 				@FindBy(xpath="//*[@id='forHf']/span[4]/ul/li/a[2]")
	 				private static WebElement  emailTempRemoveCurrentPageBtn; //Raises popup
	 							
	 				@FindBy(xpath="//*[@id='forHf']/span[4]/ul/li/a[3]")
	 				private static WebElement  emailTempRemoveControlBtn;				
	 				
	 	@FindBy(xpath="//*[@id='forHf']/span[5]")
	 	private static WebElement  emailTempPageNo;
	 	
	 	@FindBy(xpath="//*[@id='forHf']/span[6]")
	 	private static WebElement  emailTempControldrpdown;
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[1]")
	 				private static WebElement  emailTempStaticTextBtn;
	 							
	 				@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[2]")
	 				private static WebElement  emailTempBitmapBtn;
	 							
	 				@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[3]")
	 				private static WebElement  emailTempBodyGridBtn;	
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[4]")
	 				private static WebElement  emailTempRectangleBtn;	
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[5]")
	 				private static WebElement  emailTempEcllipseBtn;	
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[6]")
	 				private static WebElement  emailTempLineBtn;	
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[7]")
	 				private static WebElement  emailTempAreaBtn;
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[8]")
	 				private static WebElement  emailTempTableBtn;	
	 				
	 	@FindBy(xpath="//*[@id='forHf']/span[7]")
	 	private static WebElement  emailTempAlignmentdropdown;
	 							
	 				@FindBy(xpath="//*[@id='forHf']/span[7]/ul/li/a[1]")
	 				private static WebElement  emailTempLeftAlignmentBtn;
	 										
	 				@FindBy(xpath="//*[@id='forHf']/span[7]/ul/li/a[2]")
	 				private static WebElement  emailTempRightAlignmentBtn;
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[7]/ul/li/a[3]")
	 				private static WebElement  emailTempTopAlignmentBtn;	
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[7]/ul/li/a[4]")
	 				private static WebElement  emailTempBottomAlignmentBtn;
	 				
	 	@FindBy(xpath="//*[@id='forHf']/span[8]")
	 	private static WebElement  emailTempSizedropdown;
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[8]/ul/li/a[1]")
	 				private static WebElement  emailTempSameSizeBtn;
	 										
	 				@FindBy(xpath="//*[@id='forHf']/span[8]/ul/li/a[2]")
	 				private static WebElement  emailTempSameWidthBtn;
	 				
	 				@FindBy(xpath="//*[@id='forHf']/span[8]/ul/li/a[3]")
	 				private static WebElement  emailTempSameHeightBtn;
	 				
	 	@FindBy(xpath="//*[@id='forHf']/span[9]")
	 	private static WebElement  emailTempPrintdropdown;
	 							
	 				@FindBy(xpath="//*[@id='forHf']/span[9]/ul/li/a[1]")
	 				private static WebElement  emailTempPrint_PrintBtn;
	 													
	 				@FindBy(xpath="//*[@id='forHf']/span[9]/ul/li/a[2]")
	 				private static WebElement  emailTempPrint_PreviewBtn;
	 				
	 							@FindBy(xpath="//*[@id='fit-button']")
	 							private static WebElement emailTempPreview_FitToWidthBtn;
	 							
	 							@FindBy(xpath="//*[@id='zoom-in-button']")
	 							private static WebElement emailTempPreview_ZoomInBtn;
	 							
	 							@FindBy(xpath="//*[@id='zoom-out-button']")
	 							private static WebElement emailTempPreview_ZoomOutBtn;
	 							
	 				@FindBy(xpath="//*[@id='forHf']/span[9]/ul/li/a[3]")
	 				private static WebElement  emailTempPageSetUpBtn;
	 				
	 							@FindBy(xpath="//*[@id='id_PageSetUpPreview']")
	 							private static WebElement  emailTempPageSetupPagePreviewField;
	 							
	 							@FindBy(xpath="//*[@id='id_PageType']")
	 							private static WebElement  emailTempPageSetupPaperSizedropdown;
	 							
	 							@FindBy(xpath="//*[@id='id_PaperOrientation'][@value='0']")
	 							private static WebElement  emailTempPageSetupPotraitRadio;
	 							
	 							@FindBy(xpath="//*[@id='id_PaperOrientation'][@value='1']")
	 							private static WebElement  emailTempPageSetupLandScapeRadio;
	 							
	 							@FindBy(xpath="//*[@id='id_PaperUnit'][@value='0']")
	 							private static WebElement  emailTempPageSetupCentimetersRadio;
	 							
	 							@FindBy(xpath="//*[@id='id_PaperUnit'][@value='1']")
	 							private static WebElement  emailTempPageSetupInchesRadio;
	 							
	 							@FindBy(xpath="//*[@id='id_PageSizeLeftMarginDisplay']")
	 							private static WebElement  emailTempPageSetupLeftTxt;
	 							
	 							@FindBy(xpath="//*[@id='id_PageSizeRightMarginDisplay']")
	 							private static WebElement  emailTempPageSetupRightTxt;
	 							
	 							@FindBy(xpath="//*[@id='id_PageSizeTopMarginDisplay']")
	 							private static WebElement  emailTempPageSetupTopTxt;
	 							
	 							@FindBy(xpath="//*[@id='id_PageSizeTopMarginDisplay']")
	 							private static WebElement  emailTempPageSetupBottomTxt;
	 							
	 							@FindBy(xpath="//*[@id='id_PageSetupokBtn']")
	 							private static WebElement  emailTempPageSetupokBtn;
	 							
	 							@FindBy(xpath="//*[@id='id_PageSetupcancelBtn']")
	 							private static WebElement  emailTempPageSetUpcancelBtn;

	 	@FindBy(xpath="//*[@id='id_InvoiceDesingCancel'")
	 	private static WebElement  emailTempExitBtn;	
	 	
	 	// Fields Tab
	 	
	 	@FindBy(xpath="//*[@id='LiDesignField'")
	 	private static WebElement  emailTempFieldsTab;
	 	
	 	@FindBy(xpath="//span[contains(text(),'Master Fields')]")
	 	private static WebElement  masterFieldsExpandBtn;
	 	
	 				@FindBy(xpath="//span[contains(text(),'DocNo')]")
	 				private static WebElement  masterFieldsDocNodraganddrop;
	 	
	 				@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Date')]")
	 				private static WebElement  masterFieldsDatedraganddrop;
	 	
	 				@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Time')]")
	 				private static WebElement  masterFieldsTimedraganddrop;
	 	
	 				@FindBy(xpath="//span[contains(text(),'PurchaseAC')]")
	 				private static WebElement  masterFieldsPurchaseACdraganddrop;
	 	
	 				@FindBy(xpath="//span[contains(text(),'VendorAC')]")
	 				private static WebElement  masterFieldsVendorACdraganddrop;
	 	
	 				@FindBy(xpath="//span[contains(text(),'UpdateStock')]")
	 				private static WebElement  masterFieldsUpdateStockdraganddrop;
	 	
	 				@FindBy(xpath="//span[contains(text(),'RaiseReceipt')]")
	 				private static WebElement  masterFieldsRaiseReceiptdraganddrop;
	 	
	 				@FindBy(xpath="//span[contains(text(),'AppropriateOn')]")
	 				private static WebElement  masterFieldsAppropriateOndraganddrop;
	 	
	 				@FindBy(xpath="//span[contains(text(),'AdditionalValue')]")
	 				private static WebElement  masterFieldsAdditionalValuedraganddrop;
	 	
	 				@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Item')]")
	 				private static WebElement  masterFieldsItemdraganddrop;
	 	
	 				@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Quantity')]")
	 				private static WebElement  masterFieldsQuantitydraganddrop;
	 	
	 				@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Rate')]")
	 				private static WebElement  masterFieldsRatedraganddrop;
	 	
	 				@FindBy(xpath="//span[contains(text(),'MRN Qty')]")
	 				private static WebElement  masterFieldsMRNqtydraganddrop;
	 	
	 				@FindBy(xpath="//span[contains(text(),'Gross')]")
	 				private static WebElement  masterFieldsGrossdraganddrop;
	 	
	 				@FindBy(xpath="//span[contains(text(),'Qty balance')]")
	 				private static WebElement  masterFieldsQtyBalancedraganddrop;
	 	
	 				@FindBy(xpath="//span[contains(text(),'Average rate')]")
	 				private static WebElement  masterFieldsAverageRatedraganddrop;
	 				
	 				@FindBy(xpath="//span[contains(text(),'Add Value')]")
	 				private static WebElement  masterFieldsAddValuedraganddrop;
	 	
	 	// Properties Tab
	 				
	 	@FindBy(xpath="//*[@id='LiProp'")
	 	private static WebElement   emailTempPropertiesTab;				

	 	// RightPanel expand Btn
	 	
	 	@FindBy(xpath="//*[@id='LiFieldDivList'")
	 	private static WebElement emailTempRightPanelMaxMinIcon;
	     
	     
	    ////////////////// 
	     
	     
	     @FindBy(xpath="//div[@class='dropup pull-right open']//a[contains(text(),'SMS Template')]")
	    	private static WebElement rulesElseCustmizeTempSmsTemplate;

	     
	     // Same as Trigger PopUp Customize Template Window
	     // Named As Template
	     
	     
	     // Grid Elements
	     
	    @FindBy(xpath="//td[@id='tblAlert_Else_Rule_DocCustRules_col_1-0']")
	 	private static WebElement rulesElseAlertRow1;
	 	
	 	@FindBy(xpath="//td[@id='tblAlert_Else_Rule_DocCustRules_col_1-1']")
	 	private static WebElement rulesElseAlertType1;
	     
	 	@FindBy(xpath="//td[@id='tblAlert_Else_Rule_DocCustRules_col_1-2']")
	 	private static WebElement rulesElseAlertToWhom1;
	     
	 	@FindBy(xpath="//td[@id='tblAlert_Else_Rule_DocCustRules_col_1-3']")
	 	private static WebElement rulesElseAlertSms1;
	     
	 	@FindBy(xpath="//td[@id='tblAlert_Else_Rule_DocCustRules_col_1-4']")
	 	private static WebElement rulesElseAlertEmail1;
	     
	 	@FindBy(xpath="//td[@id='tblAlert_Else_Rule_DocCustRules_col_1-5']")
	 	private static WebElement rulesElseAlertDashBoard1;
	     
	 	@FindBy(xpath="//td[@id='tblAlert_Else_Rule_DocCustRules_col_1-6']")
	 	private static WebElement rulesElseAlertPopUp1;
	     
	 	@FindBy(xpath="//td[@id='tblAlert_Else_Rule_DocCustRules_col_1-7']")
	 	private static WebElement rulesElseAlertGTalk1;
	     
	 	@FindBy(xpath="//td[@id='tblAlert_Else_Rule_DocCustRules_col_1-8']")
	 	private static WebElement rulesElseAlertSendSepcificMessage1;
	   
	 	@FindBy(xpath="//td[@id='tblAlert_Else_Rule_DocCustRules_col_1-9']")
	 	private static WebElement rulesElseAlertSpecificMesage1;
	 		
	 	
	    @FindBy(xpath="//select[@id='ddlTypeAlert_Else_DocCustRules']")
	 	private static WebElement rulesELSEAlertsGridTypeDropdown;
	     
	    @FindBy(xpath="//td[@id='tblAlert_Else_Rule_DocCustRules_col_1-2']//select[@id='ddlRolesAlert_IF_DocCustRules']")
	 	private static WebElement rulesELSEAlertsGridToWhomDropdown;
	     
	    @FindBy(xpath="//td[@id='tblAlert_Else_Rule_DocCustRules_col_1-3']//select[@id='ddlUsersAlert_IF_DocCustRules']")
	 	private static WebElement rulesELSEAlertsGridSmsDropdown;
	     
	    @FindBy(xpath="//select[@id='ddlEmailAlert_Else_DocCustRules']")
	 	private static WebElement rulesElseAlertsGridEmailDropdown;
	     
	    @FindBy(xpath="//select[@id='ddlDashboardAlert_Else_DocCustRules']")
	 	private static WebElement rulesElseAlertsGridDashBoardDropdown;
	     
	    @FindBy(xpath="//select[@id='ddlPopupAlert_Else_DocCustRules']")
	 	private static WebElement rulesElseAlertsGridPopUpDropdown;
	     
	    @FindBy(xpath="//select[@id='ddlGTalkAlert_Else_DocCustRules']")
	 	private static WebElement rulesElseAlertsGridGTalkDropdown;
	     
	    @FindBy(xpath="//input[@id='chkSendMsgAlert_Else_DocCustRules']")
	 	private static WebElement rulesElseAlertsGridSendSpecificMessageChkBox;
	     
	    @FindBy(xpath="//input[@id='txtSpecElseicMsgAlert_Else_DocCustRules']")
	 	private static WebElement rulesElseAlertsGridSpecificMessageTxt;
	   
	 	
	   //*-------------------------------------------------------------------------------------------------------------------	
	   	
//	   	Inventory Options
	   	
	   	@FindBy(xpath="//span[@class='icon-inventory']")
	   	private static WebElement  inventoryoptionsTab;
	   	
	   	@FindBy(xpath="//input[@id='InventoryOpt_costOfIssueAcc']")
	   	private static WebElement  costOfIssueAccountTxt;
	   	
	   	@FindBy(xpath="//*[@id='InventoryOpt_costOfIssueAcc_input_image']/span")
	   	private static WebElement  costOfIssueAccountDropdown;
	   	
	   	@FindBy(xpath="//*[@id='InventoryOpt_costOfIssueAcc_input_settings']/span")
	   	private static WebElement  costOfIssueAccountSettingBtn;
	   	
	   	@FindBy(xpath="//div[@id='InventoryOpt_costOfIssueAcc_customize_popup_container']")
	   	private static WebElement  costOfIssueAccountContainerTxt;
	   	
	   	@FindBy(xpath="//div[@id='InventoryOpt_costOfIssueAcc_customize_popup_footer']//input[1]")
	   	private static WebElement  costOfIssueAccountStandardfieldsBtn;
	   	
	   				@FindBy(xpath="//select[@id='InventoryOpt_costOfIssueAcc_customize_popup_standardfields_list']")
	   				private static WebElement  costOfIssueAccountFieldsDropdown;
	   				
	   				@FindBy(xpath="//input[@id='InventoryOpt_costOfIssueAcc_customize_popup_standardfields_header']")
	   				private static WebElement  costOfIssueAccountHeaderTxt;
	   				
	   				@FindBy(xpath="//select[@id='InventoryOpt_costOfIssueAcc_customize_popup_standardfields_alignment']")
	   				private static WebElement  costOfIssueAccountAllignmentDropdown;
	   				
	   				@FindBy(xpath="//input[@id='InventoryOpt_costOfIssueAcc_customize_popup_standardfields_width']")
	   				private static WebElement  costOfIssueAccountWidthTxt;
	   				
	   				@FindBy(xpath="//section[@id='page_Content']//input[3]")
	   				private static WebElement  costOfIssueAccountStandardFieldsokBtn;
	   				
	   				@FindBy(xpath="//section[@id='page_Content']//input[4]")
	   				private static WebElement  costOfIssueAccountStandardFieldscancelBtn;
	   	
	   	@FindBy(xpath="//div[@id='InventoryOpt_costOfIssueAcc_search_container']//input[2]")
	   	private static WebElement  costOfIssueAccountDeleteColumnBtn;
	   	
	   	@FindBy(xpath="//section[@id='page_Content']//input[3]")
	   	private static WebElement  costOfIssueAccountokBtn;
	   	
	   	@FindBy(xpath="//section[@id='page_Content']//input[4]")
	   	private static WebElement  costOfIssueAccountcancelBtn;
	   	
	   	@FindBy(xpath="//input[@id='InventoryOpt_salesPurchaseACC']")
	   	private static WebElement  salesPurchaseACTxt;
	   	
	   	@FindBy(xpath="//*[@id='InventoryOpt_salesPurchaseACC_input_image']/span")
	   	private static WebElement  salesPurchaseACDropdown;

	   	@FindBy(xpath="//*[@id='InventoryOpt_salesPurchaseACC_input_settings']/span")
	   	private static WebElement  salesPurchaseACSettingBtn;
	   	
	   	@FindBy(xpath="//div[@id='InventoryOpt_salesPurchaseACC_customize_popup_container']")
	   	private static WebElement  salesPurchaseACContainerTxt;
	   	
	   	@FindBy(xpath="//div[@id='InventoryOpt_salesPurchaseACC_customize_popup_container']")
	   	private static WebElement  salesPurchaseACStandardfieldsBtn;
	   	
	   				@FindBy(xpath="//select[@id='InventoryOpt_salesPurchaseACC_customize_popup_standardfields_list']")
	   				private static WebElement  salesPurchaseACFieldsDropdown;
	   				
	   				@FindBy(xpath="//input[@id='InventoryOpt_salesPurchaseACC_customize_popup_standardfields_header']")
	   				private static WebElement  salesPurchaseACHeaderTxt;
	   				
	   				@FindBy(xpath="//select[@id='InventoryOpt_salesPurchaseACC_customize_popup_standardfields_alignment']")
	   				private static WebElement  salesPurchaseACAllignmentDropdown;
	   				
	   				@FindBy(xpath="//input[@id='InventoryOpt_salesPurchaseACC_customize_popup_standardfields_width']")
	   				private static WebElement  salesPurchaseACWidthTxt;
	   				
	   				@FindBy(xpath="//section[@id='page_Content']//input[3]")
	   				private static WebElement  salesPurchaseACStandardFieldsokBtn;
	   				
	   				@FindBy(xpath="//section[@id='page_Content']//input[4]")
	   				private static WebElement  salesPurchaseACStandardFieldscancelBtn;
	   	
	   	@FindBy(xpath="//div[@id='InventoryOpt_salesPurchaseACC_search_container']//input[2]")
	   	private static WebElement  salesPurchaseACDeleteColumnBtn;
	   	
	   	@FindBy(xpath="//section[@id='page_Content']//input[3]")
	   	private static WebElement  salesPurchaseACokBtn;
	   	
	   	@FindBy(xpath="//section[@id='page_Content']//input[4]")
	   	private static WebElement  salesPurchaseACcancelBtn;
	   	
	   	@FindBy(xpath="//select[@id='inventoryOpt_ProductRate']")
	   	private static WebElement itemRateDropdown;
	   	
	   	@FindBy(xpath="//select[@id='inventoryOpt_ReservationType']")
	   	private static WebElement reservationTypeDropdown;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkApplyRateSchemes']")
	   	private static WebElement applyrateSchemesChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkPickFreeSchemesQntyAfterLn']")
	   	private static WebElement pickFreeSchemesQtyAfterEveryLineChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkDisplayUnitNameBforeQnty']")
	   	private static WebElement displayUnitNamebeforeEveryQtyChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkDontInputProducts']")
	   	private static WebElement dontInputItemsChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkDontInputQntyRate']")
	   	private static WebElement dontInputQtyAndRateChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkDontPopupBinSelAfterQnty']")
	   	private static WebElement dontPopupBinSelectionAfterQuantityChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkHideRateAndGross']")
	   	private static WebElement hideRateAndGrossChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkHideRateAndGrossForNormalUsers']")
	   	private static WebElement hideRateAndGrossfromNormalUsersChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkHirePurchaseBehavior']")
	   	private static WebElement hirePurchaseBehaviourChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkInputCounterBill']")
	   	private static WebElement inputAsCounterBillChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkInputBinLocation']")
	   	private static WebElement inputBinLocationChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkInputQntyBreakup']")
	   	private static WebElement inputItemByAttributeChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkQCRequiredForThisDoc']")
	   	private static WebElement qcrequiredforThisDocChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkRecalculateRateInReceiptFromProd']")
	   	private static WebElement recalculateTheRateInReceiptFromProductionOnSavingChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkRestrictItmFromRepeatedEntry']")
	   	private static WebElement restricttheitemFromBeingRepeatedInEntryChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkShowvendorassignproducts']")
	   	private static WebElement showVendorAssignedItemsChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkInputBarcode']")
	   	private static WebElement inputBarcodeChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkShowHeldStock']")
	   	private static WebElement showHeldStockChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkCrossDoc']")
	   	private static WebElement crossDocChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkConsumeOrAllocateStock']")
	   	private static WebElement allocateStockChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkEnableRFID']")
	   	private static WebElement enableRFIDChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkEnableQC']")
	   	private static WebElement enableQCChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkRateZero']")
	   	private static WebElement dontmakeThertaeZeroWhenItemisChangedChkBox;
	   	
	   // Batch	
	   	
	   	@FindBy(xpath="//select[@id='inventoryOpt_BatchPick']")
	   	private static WebElement batchPickDropdown;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkAllowExpiredBatchesToBeSelected']")
	   	private static WebElement allowExpiredbatchesToBeSelectedChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkLoadBatchesBforeQnty']")
	   	private static WebElement loadbatchesbeforeQuantityChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkSelectBatchesEvenIfStksNotUpdated']")
	   	private static WebElement reserveStockByBatchesChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_chkInputBatchEvenIfStockIsNotUpdated']")
	   	private static WebElement inputBatchEvenIfStockIsnotUpdatedChkBox;
	   	
	   // Barcode	
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_rdbNotApplicable']")
	   	private static WebElement notApplicbleChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_rdbFromTemplate']")
	   	private static WebElement fromTemplateChkBox;
	   	
	   	@FindBy(xpath="//input[@id='inventoryOpt_rdbFromField']")
	   	private static WebElement fromFieldChkBox;
	  
					
		///////////////toggle
			
			@FindBy(xpath="//span[@class='dropdown icon-menuicon1 icon-font6 theme_color-inverse pull-right']")
			private static WebElement  toogleExpandBtn;
			
			@FindBy(xpath="//*[@id='navHidetab4']")
			private static WebElement  viewsBtn;
			
			@FindBy(xpath="//*[@id='navHidetab15']")
			private static WebElement  schemesBtn;
			
			@FindBy(xpath="//*[@id='navHidetab6']")
			private static WebElement  exportFieldsBtn;
			
			@FindBy(xpath="//*[@id='navHidetab7']")
			private static WebElement  triggersBtn;
			
			@FindBy(xpath="//*[@id='navHidetab8']")
			private static WebElement  documentNumberingBtn;
			
			@FindBy(xpath="//*[@id='navHidetab10']")
			private static WebElement  hirePurchaseBtn;
						
			@FindBy(xpath="//*[@id='navHidetab11']")
			private static WebElement  rulesBtn;
			
			@FindBy(xpath="//*[@id='navHidetab12']")
			private static WebElement  reportsBtn;
						
			@FindBy(xpath="//span[@class='icon-document icon-font7']")
			private static WebElement  documentsBtn;
			
			@FindBy(xpath="//span[@class='icon-miscellaneous']")
			private static WebElement  miscellaneousBtn;
						
			@FindBy(xpath="//span[@class='icon-edit icon-font8 ImagesinArabic']")
			private static WebElement  editLayoutBtn;
			
			/*@FindBy(xpath="//span[@class='icon-export']")
			private static WebElement  externalFunctionsBtn;*/
						
			@FindBy(xpath="//li[@id='navHidetab5']")
			private static WebElement  externalFunctionsBtn;
						
			@FindBy(xpath="//span[@class='icon-edit icon-font8 ImagesinArabic']")
			private static WebElement  editScreenBtn;
						
			@FindBy(xpath="//span[@class='icon-inventory']")
			private static WebElement  inventoryOptionsBtn;
			
			
			

			@FindBy(xpath="//input[@id='misc_chkdntPostCogs']")
			private static WebElement  dontpostCOGS;

			@FindBy(xpath="//input[@id='misc_PostingSelectAccInLine']")
			private static WebElement selectAccountInEveryLineChkBox;

			                //input[@id='misc_currencyPostExgRateDiff']
			@FindBy(xpath="//input[@id='misc_currencyInputLocalExchangeRate']")
			private static WebElement inputLocalExchangeRateChkBox;

			@FindBy(xpath="//input[@id='misc_currencyPostExgRateDiff']")
			private static WebElement postExchangeRateDifferrenceChkBox;
			
						
		    //Document Customization for State and City
		    	   
		    	   @FindBy (xpath="//input[@id='iCountry']")
		    	   private static WebElement countryComboBox;
		    	   
		    	   @FindBy (xpath="//tbody[@id='iCountry_table_data_body']/tr/td")
		    	   private static List<WebElement> countryComboBoxList;
		    	   
		    	   @FindBy (xpath="//input[@id='iState']")
		    	   private static WebElement stateComboBox;
		    	   
		    	   @FindBy (xpath="//tbody[@id='iState_table_data_body']/tr/td")
		    	   private static List<WebElement> stateComboBoxList;
		    	   
		    	 
		    	   @FindBy (xpath="//div[@id='dvSearchBx']/div[2]/span[1]")
		    	   private static WebElement docCustCreateVoucherBtn;
		
		    	   
		    	   
		
		
		 
	public static void checkPopUpWindow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 
	        
	        try 
	        {
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(doNotShowCheckbox));
				doNotShowCheckbox.click();
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtnInDemoPopupScreen));
				closeBtnInDemoPopupScreen.click(); 
				
				System.err.println("POP UP DISPLAYED AND CLOSED SUCCESSFULLY");
				
			} 
	        catch (Exception e)
	        {
	        	System.err.println("NO POP UP DISPLAYED");
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

	
	
}
	    
	    
	     
/*	 private static String xlfile;
	 private static String resPass="Pass";
	 private static String resFail="Fail";
	 private static ExcelReader excelReader;
*/


public static boolean checkLoginAndOpenAccountsMenu() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
{
	LoginPage lp=new LoginPage(getDriver()); 
		
    String unamelt="su";
  
    String pawslt="su";
  
    lp.enterUserName(unamelt);

    lp.enterPassword(pawslt);

   /* Select s = new Select(companyDropDownList);
     int size=s.getOptions().size();
     
     System.err.println("SIZE : "+size);
     
     for(int i=0;i<size;i++)
     {
    	 String data;
    	 
    	 data=s.getOptions().get(i).getText();
    	 
    	 System.err.println("DATA : "+data);
    	 
    	 if(data.contains("ss [070]"))
    	 {
    		 
    		 s.selectByVisibleText(data);
    		 
    		 break;
    	 }
     }*/
    
    lp.clickOnSignInBtn();
    
    //checkRefershPopOnlogin();
    
    //checkPopUpWindow();
    
    Thread.sleep(5000);
          
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
   	userNameDisplay.click();
           	
	String userInfo=userNameDisplay.getText();
	
	System.out.println("User Info : "+userInfo);
	
	System.out.println("User Info Capture Text :"+userNameDisplay.getText());
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	
	companyLogo.click();
	
	/*String getCompanyTxt=companyName.getText();
	String getLoginCompanyName=getCompanyTxt.substring(0, );
	System.out.println("company name :"+ getLoginCompanyName);
	companyLogo.click();
	*/
	
	System.out.println("***********************************checkOpenAccountsMenu*********************************");
    
	if(companyLogo.isDisplayed()==true)
	{	
		return true;
	}	 
	else
	{
		return false;
	}
 }



@FindBy (xpath="//*[@id='220']/span")
private static WebElement currency;

@FindBy (xpath="//*[@id='221']/span")
private static WebElement products;


@FindBy (xpath="//*[@id='224']/span")
private static WebElement mrp;

@FindBy (xpath="//a[@id='225']//span[contains(text(),'QC')]")
private static WebElement qc;

@FindBy (xpath="//*[@id='251']/span")
private static WebElement maintainance;

	@FindBy (xpath="//*[@id='1164']/span")
	private static WebElement breakDownStandardReason;

@FindBy (xpath="//*[@id='8001']/span")
private static WebElement pointOfSale;

@FindBy (xpath="//*[@id='8006']/span")
private static WebElement giftVoucher;


@FindBy (xpath="//*[@id='8035']/span")
private static WebElement discountVoucher;

@FindBy (xpath="//*[@id='1025']/span")
private static WebElement tableManagement;

@FindBy (xpath="//*[@id='1103']/span")
private static WebElement payroll;

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



//DataManagement menu
@FindBy (xpath="//*[@id='25']/span")
private static WebElement dataMangementMenu;

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




public static boolean checkLoginAndRestoreOptionsCompanyAndLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
{
	LoginPage lp=new LoginPage(getDriver()); 
	
    String unamelt="su";
  
    String pawslt="su";
      
    lp.enterUserName(unamelt);
    
    lp.enterPassword(pawslt);
    
    lp.clickOnSignInBtn();
    
    //checkRefershPopOnlogin();
    
    //checkPopUpWindow();
  
	Thread.sleep(8000);
	
	String actUserInfo=userNameDisplay.getText();
	
	System.out.println("User Info  : "+actUserInfo);
	
	System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	companyLogo.click();
	
	String getCompanyTxt=companyName.getText();
	String getLoginCompanyName=getCompanyTxt.substring(0, 19);
	System.out.println("company name  :  "+ getLoginCompanyName);
	companyLogo.click();
	
	
	String expUserInfo           ="SU";
	String expLoginCompanyName   ="Automation";
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataMangementMenu));
	dataMangementMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restore));
	restore.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(folderpathExpandBtn));
	folderpathExpandBtn.click();
	
	Thread.sleep(3000);
     
     Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\optionsCompanyRestore.exe");
     
     Thread.sleep(3000);
     
     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreCompanyBtn));
     restoreCompanyBtn.click();
     
     try
     {
    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(overRideYesBtn));
 		 overRideYesBtn.click();
 		 Thread.sleep(20000);
     }
     catch(Exception e)
     {
    	 System.err.println("NO OLDER COMPANY EXISTS");
     }
     
     Thread.sleep(3000);
     
     if(getIsAlertPresent())
     {
    	 getWaitForAlert();
    	 System.out.println(getAlert().getText());
    	 getAlert().accept();
     }
     
     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
     userNameDisplay.click();
	
	
     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
     logoutOption.click();

     Thread.sleep(3000);
     
     lp.enterUserName(unamelt);
     
     lp.enterPassword(pawslt);
     
     Thread.sleep(2000);
     
     Select s = new Select(companyDropDownList);
     int size=s.getOptions().size();
     
     System.err.println("SIZE : "+size);
     
     for(int i=0;i<size;i++)
     {
    	 String data;
    	 
    	 data=s.getOptions().get(i).getText();
    	 
    	 System.err.println("DATA : "+data);
    	 
    	 if(data.contains("Automation Options"))
    	 {
    		 
    		 s.selectByVisibleText(data);
    		 
    		 break;
    	 }
     }
     
     lp.clickOnSignInBtn();
     
     Thread.sleep(8000);
     
     String actUserInfo1=userNameDisplay.getText();
		
	System.out.println("User Info  : "+actUserInfo1);
	
	System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	companyLogo.click();
	
	String getCompanyTxt1=companyName.getText();
	String getLoginCompanyName1=getCompanyTxt1.substring(0, 19);
	System.out.println("company name  :  "+ getLoginCompanyName1);
	companyLogo.click();
	
	String expUserInfo1           ="SU";
	String expLoginCompanyName1   ="Automation Options ";
     
     
	System.out.println("UserInfo             : "+actUserInfo            +" Value Expected : "+expUserInfo);
	System.out.println("LoginCompanyName     : "+getLoginCompanyName    +" Value Expected : "+expLoginCompanyName);
	
	System.out.println("UserInfo1             : "+actUserInfo1            +" Value Expected : "+expUserInfo1);
	System.out.println("LoginCompanyName1     : "+getLoginCompanyName1    +" Value Expected : "+expLoginCompanyName1);
	
	if(actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName) 
			
			&& actUserInfo1.equalsIgnoreCase(expUserInfo1) && getLoginCompanyName1.contains(expLoginCompanyName1))
	{

		return true;
	}
	else
	{
		return false;
	}
}




@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[11]")
private static List<WebElement> masterAccountsList;


@FindBy(xpath="//*[@id='GenerateNewMasterModel']/div/div[1]/div[2]/div/ul/li[2]/button[5]")
public static WebElement masterCloseBtn;


@FindBy(xpath="//*[@id='221']/span")
private static WebElement  homeMasterItemMenu;

@FindBy(xpath="//*[@id='1105']/span")
private static WebElement  homeMasterItem_ItemMenu;

 @FindBy(xpath="//span[@id='spnHeaderText']")
 private static WebElement item_ItemLabel;

 
 @FindBy(xpath="//a[@id='72']/span")
 private static WebElement  homeMasterItem_UnitConversionMenu;
 

@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[12]")
private static List<WebElement> masterItemsList;


	//Currency Menu
		@FindBy(xpath="//a[@id='220']//span[contains(text(),'Currency')]")
		private static WebElement currencyMenu;
		
		@FindBy(xpath="//span[contains(text(),'Currency Master')]")
		private static WebElement currencyMasterMenu;
		
		@FindBy(xpath="//a[@id='71']//span[contains(text(),'Exchange Rate')]")
		private static WebElement exchangeRateMenu;
		
		@FindBy(xpath="//span[contains(text(),'Exchange Rate History')]")
		private static WebElement exchangeRateHistoryMenu;
		
		@FindBy(xpath="//span[@class='navText']//span[contains(text(),'Currency Master')]")
		private static WebElement currencyMasterlabel;
		
		@FindBy(xpath="//*[@id='ISOCurrencyCode']")
		private static WebElement ISOCurrencyCodeTxt;
		
		
		// ISO CurrencyPopUp
		//---------------------------------------------------------------------------------------------------------------------
		
		@FindBy(xpath="//td[@id='ISOCurrencyCode_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement  ISOCurrencyCodeSettingsBtn;
		
			@FindBy(xpath="//h4[@id='ISOCurrencyCode_customize_popup_heading']")
			private static WebElement  ISOCurrencyCode_CustomizeDisplayColumnsLabel;
		
				@FindBy(xpath="//div[@id='ISOCurrencyCode_customize_popup_container']")
				private static WebElement  ISOCurrencysCodeContainerTxt;
				
				@FindBy(xpath="//div[@id='ISOCurrencyCode_customize_popup_footer']//input[1]")
				private static WebElement  ISOCurrencyCodeStanadrdfieldsBtn;
				
					@FindBy(xpath="//select[@id='ISOCurrencyCode_customize_popup_standardfields_list']")
					private static WebElement  ISOCurrencyCodeStanadrdfields_Fielddropdown;
					
					@FindBy(xpath="//input[@id='ISOCurrencyCode_customize_popup_standardfields_header']")
					private static WebElement  ISOCurrencyCodeStanadrdfields_headerTxt;
					
					@FindBy(xpath="//select[@id='ISOCurrencyCode_customize_popup_standardfields_alignment']")
					private static WebElement  ISOCurrencyCodeStanadrdfields_Allignmentdropdown;
					
					@FindBy(xpath="//input[@id='ISOCurrencyCode_customize_popup_standardfields_width']")
					private static WebElement  ISOCurrencyCodeStanadrdfields_WidthTxt;
					
					@FindBy(xpath="//div[@id='currencyMasterWidth']//input[3]")
					private static WebElement  ISOCurrencyCodeStanadrdfields_OkBtn;
					
					@FindBy(xpath="//section[@id='page_Content']//input[4]")
					private static WebElement  ISOCurrencyCodeStanadrdfields_CancelBtn;
				
				@FindBy(xpath="//div[@id='ISOCurrencyCode_search_container']//input[2]")
				private static WebElement  ISOCurrencyCodedeleteColumnBtn;
				
				@FindBy(xpath="//div[@id='currencyMasterWidth']//input[3]")
				private static WebElement  ISOCurrencyCodeOkBtn;
				
				@FindBy(xpath="//section[@id='page_Content']//input[4]")
				private static WebElement  ISOCurrencyCodeCancelBtn;
		

				@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[2]/fieldset[1]/div[1]/div[2]/div[2]/input[1]")
				private static WebElement coinsNameTxt;
				
				@FindBy(xpath="//span[@id='btnSave']")
				private static WebElement saveIcon;
				
				@FindBy(xpath="//div[@id='myNavbar']/ul/li[4]/div")
				private static WebElement currencyMasterCancelBtn;
				

				@FindBy(xpath="//span[contains(text(),'Exchange Rate Definition')]")
				private static WebElement exchangeRateDefinitionLabel;
				
				@FindBy(xpath="//span[@id='DeleteData']")
				private static WebElement exchnageRateDeleteOption;
				
				@FindBy(xpath="//div[@class='navText']//span[contains(text(),'Exchange Rate History')]")
				private static WebElement exchangeRateHistoryLabel;
				
				@FindBy(xpath="//input[@id='id_option']")
				private static WebElement baseCurrencyInExchangeRateHistory;
				
				@FindBy(xpath="//*[@id='DateOption']")
				private static WebElement dateOptionDropdown;
				
				@FindBy(xpath="//*[@id='FromDate']")
				private static WebElement fromDateTxt;
				
				@FindBy(xpath="//*[@id='ToDate']")
				private static WebElement toDateTxt ;
				
				@FindBy(xpath="//i[@class='icon-font7 icon-refresh']")
				private static WebElement loadIcon;
				
				@FindBy(xpath="//i[@class='icon-font7 icon-clear']")
				private static WebElement exchangeRateHistoryclearIcon;
				

				@FindBy(xpath="//span[@id='exchange_currency']")
				private static WebElement exchangeRateIcon;
				
				@FindBy(xpath="//*[@id='id_option_table_data_body']/tr/td")
			    private static List<WebElement> baseCurrencyListCount;
				
				
				@FindBy(xpath="//*[@id='ExchangeTable_CurrencyName_table_data_body']/tr/td")
			    private static List<WebElement> currencyNameListCount;
				
				@FindBy(xpath="//*[@id='ExchangeTable_DefinedAs_table_data_body']/tr/td")
			    private static List<WebElement> definedAsListCount;
				
				@FindBy(xpath="//input[@id='id_option']")
				private static WebElement baseCurrencyTxt;
				
				@FindBy(xpath="//input[@id='EffectiveDate']")
				private static WebElement withEffectiveDateTxt;
				
				@FindBy(xpath="//td[@id='EffectiveDate_input_image']//span[@class='icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow']")
				private static WebElement withEffectiveDateCalendarBtn;
				
				@FindBy(xpath="//select[@id='EffectiveDate_year']")
				private static WebElement withEffectiveDateYearDropDown;
				
				@FindBy(xpath="//select[@id='EffectiveDate_month']")
				private static WebElement withEffectiveDateMonthDropDown;
						
				@FindBy(xpath="//*[@id='EffectiveDate_input_image']/span")
				private static WebElement dOne;
				
				@FindBy(xpath="//*[@id='EffectiveDate_day_today']/td/span[1]")
				private static WebElement dTwo;
				
				@FindBy(xpath="//input[@id='ExchangeTable_DefinedAs']")
				private static WebElement gridDefineAsTxt;
				
				@FindBy(xpath="//input[@id='ExchangeTable_Rate']")
				private static WebElement gridRateTxt;
				
				@FindBy(xpath="//input[@id='ExchangeTable_CurrencyName']")
				private static WebElement gridcurrencyName1Txt;
				
				@FindBy(xpath="//td[@id='ExchangeTable_CurrencyName_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
				private static WebElement  currencyNameSettingsBtn;
				
				
				@FindBy(xpath="//i[@class='icon-font7 icon-save']")
				private static WebElement exchangeRateDefinitionSaveIcon;
				                    
				@FindBy(xpath="//span[@id='btnCurrencyMasterCancel']//i[@class='icon-font7 icon-close']")
				private static WebElement exchangeRateDefinitionPopCloseIcon;
				
				@FindBy(xpath="//span[@id='closeExchangeRate']")
				private static WebElement exchangeRateDefinitionPageCloseIcon;
				
				
				
				@FindBy(xpath="//i[@class='icon-font6 icon-close']")
				private static WebElement exchangeRateHistoryCloseIcon;
				
				@FindBy(xpath="//tbody[@id='HistoryTableBody']/tr[2]/td[3]")
				private static WebElement valueUSD;
				
				@FindBy(xpath="//tbody[@id='HistoryTableBody']/tr[2]/td[4]")
				private static WebElement valueAED;
				
				@FindBy(xpath="//td[@id='ExchangeTable_col_1-2']")
				private static WebElement selCurrencyOne;
				
				@FindBy(xpath="//td[@id='ExchangeTable_col_2-2']")
				private static WebElement selCurrecnyTwo;
				
				@FindBy(xpath="//td[@id='ExchangeTable_col_1-3']")
				private static WebElement defCurrecnyOne;
				
				@FindBy(xpath="//td[@id='ExchangeTable_col_2-3']")
				private static WebElement defCurrecnyTwo;
				
				@FindBy(xpath="//*[@id='80']/span")
				private static WebElement  homeMasterItem_buyerPriceBookMenu;
				
				@FindBy(xpath="//*[@id='76']/span")
				private static WebElement  homeMasterItem_sellerPriceBookMenu;
				
				
				@FindBy(xpath="//*[@id='ctrlOptionProPriceBookH']")
				private static WebElement buyerPriceBook_PriceBook;
				
				
				@FindBy(xpath="//input[@id='iNoOfDecimals']")
				private static WebElement noOfDecimalsTxt;
				
				@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[10]")
				private static List<WebElement> masterItemUnitsList;
				
				@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[8]/div[1]/label/input")
				private static List<WebElement> masterItemUnitsChkboxList;
				
				
				@FindBy(xpath="//tbody[@id='LandingGridBody']/tr[1]/td[12]")
				public static WebElement itemNewCreationName;
				
				@FindBy(xpath="//tbody[@id='LandingGridBody']/tr[1]/td[10]")
				public static WebElement unitNewCreationName;
	
				
				//Department Menu	
				 @FindBy (xpath="//a[@id='1106']//span[contains(text(),'Department')]")
				 public static WebElement departmentMenu;
				 
				 
				//Warehouse Menu	
				 @FindBy (xpath="//a[@id='1107']")
				 public static WebElement warehouseMenu;
				
				
				 

	 public  boolean checkWarehouseMasterInMAsterMenu() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		 System.out.println("***********checkWarehouseMasterInMAsterMenu Method Excutes **********");
			
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	      	homeMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	      	mastersMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseMenu));
		    warehouseMenu.click();
		   
		    Thread.sleep(2000);
		    
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_ItemLabel));
		    System.out.println("item_ItemLabel  "+item_ItemLabel.getText());
		   
	      	if(item_ItemLabel.getText().equalsIgnoreCase("Warehouse"))
			{	
				return true;
			}	 
			else
			{
				return false;
			}
	     }
		
		
		
		public static boolean checkSavingWarehouseInWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 System.out.println("********************************checkSavingWarehouseInWarehouseMaster Method Excutes **********");
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
			 
			 masterNewBtn.click();
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			 nameTxt.click();
			 nameTxt.clear();
			 nameTxt.sendKeys("Trial Warehouse");
			 nameTxt.sendKeys(Keys.TAB);
			 
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
			 codeTxt.click();
			 codeTxt.clear();
			 codeTxt.sendKeys("TrialWarehouse");
			 codeTxt.sendKeys(Keys.TAB);
			 

			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			 saveBtn.click();
			 
			 String expMessage="Saved Successfully";
				
			 String actMessage=checkValidationMessage(expMessage);
			 
			 Thread.sleep(1000);
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
			 masterCloseBtn.click();
			 
			 System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
				
			 if(actMessage.equalsIgnoreCase(expMessage))
			 {
				 return true;
			 }
			 
			 else
			 {
				 return false;
			 }
		}
		
		
		
		
		
		public static boolean checkEditingInSavedWarehouseINWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitNewCreationName));
			 
			 String actaccountNewCreationName  =unitNewCreationName.getText();
			 String expaccountNewCreationName  ="Trial Warehouse";
				  		 
			 System.out.println("********************************checkEditingInSavedWarehouseINWarehouseMaster Method Executes ********************");
			 
			 System.out.println("Saved Warehouse : "+actaccountNewCreationName+"  Value Expected  "+expaccountNewCreationName);
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountFirstCheckBox));
			 accountFirstCheckBox.click();
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterEditBtn));
			 masterEditBtn.click();
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			 nameTxt.click();
			 nameTxt.clear();
			 nameTxt.sendKeys("Trial Warehouse Update");
			 nameTxt.sendKeys(Keys.TAB);
			 
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
			 codeTxt.click();
			 codeTxt.clear();
			 codeTxt.sendKeys("TrialWarehouseUpdate");
			 codeTxt.sendKeys(Keys.TAB);
			 
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			 saveBtn.click();
			 
			 String expMessage="Updated Successfully";
				
			 String actMessage=checkValidationMessage(expMessage);
			 
			 System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
				
			 if(actMessage.equalsIgnoreCase(expMessage))
			 {
				 return true;
			 }
			 
			 else
			 {
				 return false;
			 }
			 }
		
	
		
		
		public static boolean checkUpdatedWarehouseMasterDisplayInWarehouse() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitNewCreationName));
			 
			 String actaccountNewCreationName  =unitNewCreationName.getText();
			 String expaccountNewCreationName  ="Trial Warehouse Update";
			 
			 System.out.println("************************checkUpdatedWarehouseMasterDisplayInWarehouse***********************");
			 
			 System.err.println("accountNewCreationName  : "+actaccountNewCreationName+" Value Expected : "+expaccountNewCreationName);
			 
			 if(actaccountNewCreationName.equalsIgnoreCase(expaccountNewCreationName))
			 { 
				 return true;
			 }
			 else
			 {
				 return false;
			 }
		}
		
		
		
	
		public static boolean checkDeleteInWarehouseInWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 System.out.println("************************checkDeleteInWarehouseInWarehouseMaster  Method EXCUTES ***************");
			 
			/* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountFirstCheckBoxToSelection));
			 
			 if (accountFirstCheckBoxToSelection.isSelected()==false) 
			 {
				 accountFirstCheckBoxToSelection.click();		
			 }
			 */
	         getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDeleteBtn));
	     
	         masterDeleteBtn.click();
	         
	         getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getMsgOnDelete));
	         System.out.println(getMsgOnDelete.getText());
	         
	         getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelete));
	         clickOnOkInDelete.click();
	         
	         String expMessage="Record Deleted Successfully.";
	 		
	         String actMessage=checkValidationMessage(expMessage);
	         
	         System.out.println("********************************checkDeleteItemInItemMaster***************************");
				
			 System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
				
			 if(actMessage.equalsIgnoreCase(expMessage))
			 {
				 return true;
			 }
			 else
			 {
				 return false;
			 }	 
		}
				
		
		@FindBy(xpath="//*[@id='drpdownSlider']")
		public static WebElement toggleRibbon;

		@FindBy(xpath="//*[@id='btnAdvImportForMaster']")
		public static WebElement advanceMasterImportExportBtn;;

		@FindBy(xpath="//*[@id='viewheading']")
		public static WebElement advanceMasterLabel;
						 
		@FindBy(xpath="//button[@id='btnFile']")
		public static WebElement FileUploadBtn;
			
		@FindBy(xpath="//*[@class='icon-font7 icon-close']")
		public static WebElement advMasterCloseBtn;



		@FindBy(xpath="//*[@id='btnImportDataForAdvanceMasterPopUp']")
		public static WebElement advMasterImportDataBtn;;


		@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[10]")
		private static List<WebElement> accNameList;

		@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[8]//input")
		private static List<WebElement> masterGridBodyChkbox;

		  @FindBy(xpath="//*[@id='LandingGridBody']/tr/td[10]")
		  private static List<WebElement> masterGridBodyName;
		
	
		
		
		public boolean checkWarehouseListBeforeImportingDataFromExcel() throws InterruptedException
		{
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			   	homeMenu.click();
			   	
			   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
			   	mastersMenu.click();
			   	
			   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseMenu));
				     
			   	warehouseMenu.click();
				   
			   	Thread.sleep(2000);
			   	
			
			 
			 int accCount = masterGridBodyName.size();
			

			 ArrayList<String> acclist = new ArrayList<String>();

			 		for (int i = 0; i < accCount; i++)
			 		{
			 			 String data = masterGridBodyName.get(i).getText();
			 			acclist.add(data);
			 			
			 		}
			 		
			 		String actDeptList=acclist.toString();
			 		String  expDeptList="[STWH, SECUNDERABAD, HYDERABAD, MUMBAI, WHTOINACTIVE, WHG]";
			 		
			 		System.out.println("Actual Warehouse List:  "+actDeptList);
			 		System.out.println("Expected Warehouse List:  "+expDeptList);
			 		
			 		if(actDeptList.equalsIgnoreCase(expDeptList))
			 		{
			 			return true;
			 		}
			 		else
			 		{
			 			return false;
			 		}
			 		
		}

		public boolean checkAdvanceMasterImportExportWarehouse() throws InterruptedException
		{
			Thread.sleep(2000);
			
		   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleRibbon));
		   	toggleRibbon.click();
		   	Thread.sleep(2000);
		   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterImportExportBtn));
		   	advanceMasterImportExportBtn.click();
		   	
		   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterLabel));
		    System.out.println("Advance Master import/Export Label  "+advanceMasterLabel.getText());
		      	
		  	if(advanceMasterLabel.getText().equalsIgnoreCase("Advance Master Import/Export----> Warehouse"))
			{	
				return true;
			}	 
			else
			{
				return false;
			}
		   	
		   	
		}
		public boolean checkImportDatainAdvanceMasterImportExportWarehouse() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
		{
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterImportDataBtn));
			 advMasterImportDataBtn.click();
			 Thread.sleep(2000);
			 String expMessage="Please load excel sheet";
			String actMessage=checkValidationMessage(expMessage);
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FileUploadBtn));
			FileUploadBtn.click();
			Thread.sleep(2000);
			
			 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\WarehouseImportFromExcel.exe");
			 Thread.sleep(5000);
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterImportDataBtn));
			 advMasterImportDataBtn.click();
			 Thread.sleep(2000);
			 
		/*	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterCloseBtn));
			 advMasterCloseBtn.click();
			 Thread.sleep(2000);*/
			 
			 String expMessage1="8 Records Imported Successfully";
			 String actMessage1=checkValidationMessage(expMessage);
			 
			
			 
			 if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1))
			 {
				 return true;
			 }
			 else
			 {
				 return false;
			 }
			
			}

		@FindBy(xpath="//*[@id='LandingGridBody']/tr[7]/td")
		public static List<WebElement> Warehouse7thRowList;

		public boolean checkWarehouseListAfterImportingDataFromExcel() throws InterruptedException
		{
			getDriver().navigate().refresh();
			Thread.sleep(2000);

			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			   	homeMenu.click();
			   	
			   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
			   	mastersMenu.click();
			   	
			   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseMenu));
				     
			   	warehouseMenu.click();
				   
			   	Thread.sleep(2000);
			   	
			
			Thread.sleep(2000);
			 int accCount = masterGridBodyName.size();
				

			 ArrayList<String> acclist = new ArrayList<String>();

			 		for (int i = 0; i < accCount; i++)
			 		{
			 			 String data = masterGridBodyName.get(i).getText();
			 			acclist.add(data);
			 			
			 		}
			 		
			 		String actDeptList=acclist.toString();
			 		String  expDeptList="[STWH, SECUNDERABAD, HYDERABAD, MUMBAI, WHTOINACTIVE, WHG, WarehouseImport]";
			 		
			 		System.out.println("Actual Warehouse List:  "+actDeptList);
			 		System.out.println("Expected Warehouse List:  "+expDeptList);
			 		
			 		if(actDeptList.equalsIgnoreCase(expDeptList))
			 		{
			 			return true;
			 		}
			 		else
			 		{
			 			return false;
			 		}
			 		
		}
		
		
	public boolean checkImportedWarehouseDelete() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{		

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		   	homeMenu.click();
		   	
		   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		   	mastersMenu.click();
		   	
		   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseMenu));
			     
		   	warehouseMenu.click();
			   
		   	Thread.sleep(2000);
		   	
		 int count = masterGridBodyName.size();
			
			for (int i = 0; i < count; i++) 
			{
				String data = masterGridBodyName.get(i).getText();
				
				if (data.equalsIgnoreCase("WarehouseImport")) 
				{
					if (masterGridBodyChkbox.get(i).isSelected()==false) 
					{
						masterGridBodyChkbox.get(i).click();
						break;
					}
				}
			}
		    
		    Thread.sleep(2000);	
		 
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDeleteBtn));

	  masterDeleteBtn.click();
	  
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getMsgOnDelete));
	  System.out.println(getMsgOnDelete.getText());
	  
	  //Thread.sleep(2000);
	  
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelete));
	  clickOnOkInDelete.click();
	  
	  String expMessage="Record Deleted Successfully.";
		
	  String actMessage=checkValidationMessage(expMessage);
	  
	  System.out.println("********************************checkDeleteWarehouseBeforeCustomization***************************");
			
		 System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
			
		 if(actMessage.equalsIgnoreCase(expMessage))
		 {
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
			 closeBtn.click();
			 return true;
		 }
		 
		 else
		 {
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
			 closeBtn.click();
			 return false;
		 }
	}

		
	
	
/////////////////////Adding Groups and subgroups//////////////////////////

public boolean checkSavingWarehouseGroupinWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseMenu));
warehouseMenu.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAddGroupBtn));
masterAddGroupBtn.click();

System.out.println("********************************checkSavingWarehouseGroupWarehouseMaster*****************************");

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
nameTxt.click();
nameTxt.clear();
nameTxt.sendKeys("Group Warehouse");
nameTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
codeTxt.click();
codeTxt.clear();
codeTxt.sendKeys("GroupWarehouse");
codeTxt.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
saveBtn.click();

String expMessage="Saved Successfully";

String actMessage=checkValidationMessage(expMessage);

System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);

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


@FindBy(xpath="//*[@id='0']/a")
public static WebElement accountList;



@FindBy(xpath="(//*[@id='0']/li)[1]/a[1]")
public static WebElement accountGrp;




@FindBy(xpath="//*[@id='ol_treeNavigation']/li[2]/span")
public static WebElement GroupAccountLabel;

public static boolean checkSavingWarehouseSubGroupinWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
int accNameListCount = accNameList.size();

for(int i=0;i<accNameListCount;i++)
{
String data = accNameList.get(i).getText();
Thread.sleep(2000);
if(data.equals("Group Warehouse"))
{
Thread.sleep(2000);
getAction().doubleClick(accChkBoxList.get(i)).build().perform();
Thread.sleep(2000);
break;
}
}


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAddGroupBtn));
masterAddGroupBtn.click();

System.out.println("********************************checkSavingWarehouseSubGroupWarehouseMaster*****************************");

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
nameTxt.click();
nameTxt.clear();
nameTxt.sendKeys("SubGroup1 Warehouse");
nameTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
codeTxt.click();
codeTxt.clear();
codeTxt.sendKeys("SubGroup1Warehouse");
codeTxt.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
saveBtn.click();

Thread.sleep(2000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
nameTxt.click();
nameTxt.clear();
nameTxt.sendKeys("SubGroup2 Warehouse");
nameTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
codeTxt.click();
codeTxt.clear();
codeTxt.sendKeys("SubGroup2Warehouse");
codeTxt.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
saveBtn.click();


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

public boolean checkCreatedGroupDisplayedonLeftPanel()
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountGrp));

boolean s=accountGrp.isDisplayed();
if(s)
{
return true;
}
else
{
return false;
}
}


@FindBy(xpath="//*[@id='lblnchecked']/input")
public static List<WebElement> accChkBoxList;


public boolean checkSavingWarehousesinSubGroupsinWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

int accNameListCount = accNameList.size();

for(int i=0;i<accNameListCount;i++)
{
String data = accNameList.get(i).getText();
Thread.sleep(2000);
if(data.equals("SubGroup1 Warehouse"))
{
Thread.sleep(2000);
getAction().doubleClick(accChkBoxList.get(i)).build().perform();
Thread.sleep(2000);
break;
}
}


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
masterNewBtn.click();

System.out.println("********************************checkSavingWarehouseSubGroupWarehouseMaster*****************************");

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
nameTxt.click();
nameTxt.clear();
nameTxt.sendKeys("Warehouse1");
nameTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
codeTxt.click();
codeTxt.clear();
codeTxt.sendKeys("Warehouse1");
codeTxt.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
saveBtn.click();

Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
masterCloseBtn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GroupAccountLabel));
GroupAccountLabel.click();
Thread.sleep(2000);

int accNameListCount1 = accNameList.size();

for(int i=0;i<accNameListCount1;i++)
{
String data = accNameList.get(i).getText();
Thread.sleep(2000);
if(data.equals("SubGroup2 Warehouse"))
{
Thread.sleep(2000);
getAction().doubleClick(accChkBoxList.get(i)).build().perform();
Thread.sleep(2000);
break;
}
}


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
masterNewBtn.click();




getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
nameTxt.click();
nameTxt.clear();
nameTxt.sendKeys("Warehouse2");
nameTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
codeTxt.click();
codeTxt.clear();
codeTxt.sendKeys("Warehouse2");
codeTxt.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
saveBtn.click();


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


public static boolean checkWarehousesSavedinGroupsandSubGroupsinWarehouseMaster() throws InterruptedException
{
getDriver().navigate().refresh();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseMenu));
warehouseMenu.click();
Thread.sleep(2000);


int accCount = accNameList.size();

ArrayList<String> acclist = new ArrayList<String>();

for (int i = 0; i < accCount; i++)
{
String data = accNameList.get(i).getText();
acclist.add(data);

}

String actAccList=acclist.toString();
String  expAccList="[Group Warehouse, STWH, SECUNDERABAD, HYDERABAD, MUMBAI, WHTOINACTIVE, WHG]";


int accNameListCount = accNameList.size();

for(int i=0;i<accNameListCount;i++)
{
String data = accNameList.get(i).getText();
Thread.sleep(2000);
if(data.equals("Group Warehouse"))
{
Thread.sleep(2000);
getAction().doubleClick(accChkBoxList.get(i)).build().perform();
Thread.sleep(2000);
break;
}
}


int accCount1= accNameList.size();
ArrayList<String> acclist1 = new ArrayList<String>();

for (int i = 0; i < accCount1; i++)
{
String data = accNameList.get(i).getText();
acclist1.add(data);
}

String actGrpAccList=acclist1.toString();
String  expGrpAccList="[SubGroup2 Warehouse, SubGroup1 Warehouse]";


int accNameListCount1 = accNameList.size();
for(int i=0;i<accNameListCount1;i++)
{
String data = accNameList.get(i).getText();
Thread.sleep(2000);
if(data.equals("SubGroup1 Warehouse"))
{
Thread.sleep(2000);
getAction().doubleClick(accChkBoxList.get(i)).build().perform();
Thread.sleep(2000);
break;
}
}


int accCount2= accNameList.size();
ArrayList<String> acclist2 = new ArrayList<String>();
for (int i = 0; i < accCount2; i++)
{
String data = accNameList.get(i).getText();
acclist2.add(data);
}

String actSubGrpAccList=acclist2.toString();
String  expSubGrpAccList="[Warehouse1]";



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GroupAccountLabel));
GroupAccountLabel.click();
Thread.sleep(2000);

int accNameListCount3 = accNameList.size();
for(int i=0;i<accNameListCount3;i++)
{
String data = accNameList.get(i).getText();
Thread.sleep(2000);
if(data.equals("SubGroup2 Warehouse"))
{
Thread.sleep(2000);
getAction().doubleClick(accChkBoxList.get(i)).build().perform();
Thread.sleep(2000);
break;
}
}


int accCount3= accNameList.size();
ArrayList<String> acclist3 = new ArrayList<String>();
for (int i = 0; i < accCount3; i++)
{
String data = accNameList.get(i).getText();
acclist3.add(data);
}

String actSubGrpAcc2List=acclist3.toString();
String  expSubGrpAcc2List="[Warehouse2]";

System.out.println("********************************checkSavingGroupsandSubGroupWarehouseMaster*****************************");


System.out.println("Actual Warehouses List:  "+actAccList);
System.out.println("Expected Warehouses List:  "+expAccList);

System.out.println("Actual Sub Group Warehouses List:  "+actGrpAccList);
System.out.println("Expected Sub Group Warehouses List:  "+expGrpAccList);

System.out.println("Actual  Warehouses in Sub Group1 List:  "+actSubGrpAccList);
System.out.println("Expected Warehouses in Sub Group1 List:  "+expSubGrpAccList);

System.out.println("Actual  Warehouses in Sub Group2 List:  "+actSubGrpAcc2List);
System.out.println("Expected Warehouses in Sub Group2 List:  "+expSubGrpAcc2List);


if(actAccList.equalsIgnoreCase(expAccList) && actGrpAccList.equalsIgnoreCase(expGrpAccList)
&& actSubGrpAccList.equalsIgnoreCase(expSubGrpAccList) && actSubGrpAcc2List.equalsIgnoreCase(expSubGrpAcc2List))
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
closeBtn.click();

return true;
}
else
{

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
closeBtn.click();
return false;
}


}


public boolean checkDeleteGroupWarehouseinWarehouseMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseMenu));
warehouseMenu.click();
Thread.sleep(2000);


int count1 = masterGridBodyName.size();

for (int i = 0; i < count1; i++) 
{
String data = masterGridBodyName.get(i).getText();

if (data.equalsIgnoreCase("Group Warehouse")) 
{
if (masterGridBodyChkbox.get(i).isSelected()==false) 
{
	masterGridBodyChkbox.get(i).click();
	break;
}
}
}




Thread.sleep(2000);	

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDeleteBtn));

masterDeleteBtn.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getMsgOnDelete));
System.out.println(getMsgOnDelete.getText());

//Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelete));
clickOnOkInDelete.click();

String expMessage="Record Deleted Successfully.";

String actMessage=checkValidationMessage(expMessage);

System.out.println("********************************checkDeleteAccountBeforeCustomization***************************");

System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);

if(actMessage.equalsIgnoreCase(expMessage))
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
closeBtn.click();
return true;
}

else
{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
closeBtn.click();
return false;
}


}


	
	
	
	
	public WarehousePage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}
	
	
	

}
