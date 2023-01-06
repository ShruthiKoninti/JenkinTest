package com.focus.Pages;


import java.awt.RenderingHints.Key;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import javax.sound.midi.SysexMessage;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ddf.EscherColorRef.SysIndexProcedure;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;
import org.testng.SkipException;

import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;
import com.sun.tools.xjc.model.SymbolSpace;


public class SmokeBudgetVouchersMRPOPage extends BaseEngine
{
	
	
	//Vouchers Home Page Ribbon Control
		@FindBy(xpath="//div[@id='myNavbar']//span[@class='icon-new icon-font6']")
		private static WebElement  newBtn;
		
		@FindBy(xpath="//div[@id='id_transaction_homescreen_edit']//span[@class='icon-edit icon-font6 ImagesinArabic']")
		private static WebElement  editBtn;

		@FindBy(xpath="//span[@class='icon-print icon-font6']")
		private static WebElement  printBtn;
		
		@FindBy(xpath="//div[@id='id_transaction_homescreen_Delete']//span[@class='icon-delete icon-font6']")
		private static WebElement  deleteBtn;
		
		@FindBy(xpath="//span[@class='icon-authorize icon-font6']")
		private static WebElement  authorizeBtn;
		
		@FindBy(xpath="//span[@class='icon-suspend icon-font6']")
		private static WebElement  suspendBtn;
		
		@FindBy(xpath="//span[@class='icon-reject2 icon-font6']")
		private static WebElement  rejectBtn;
		
		@FindBy(xpath="//span[@class='icon-cheque icon-font6']")
		private static WebElement  chequeReturnBtn;
		
		@FindBy(xpath="//span[@class='icon-printbarcode icon-font6']")
		private static WebElement  printBarCodeBtn;
		
		@FindBy(xpath="//span[@class='icon-convert icon-font6']")
		private static WebElement  convertBtn;
		
		@FindBy(xpath="//span[@class='icon-export icon-font6']")
		private static WebElement  exportToXMLBtn;  
			
		@FindBy(xpath="//div[@id='myNavbar']//span[@class='icon-settings icon-font6']")
		private static WebElement  settingsBtn;
		
		@FindBy(xpath="//div[@id='dvHomeTransClose']//span[@class='icon-close icon-font6']")  
		private static WebElement  homeCloseBtn;
				
		//Vouchers Home Page Options
		@FindBy(xpath="//a[@class='transaction_viewname_anchor theme_color-inverse font-5']")
		private static WebElement  allVouchersOption;
		
		@FindBy(xpath="//a[@class='transaction_viewname_anchor font-5']")
		private static WebElement  pendingBillsOption;
		
		@FindBy(xpath="//input[@id='SelectDefaultView']")
		private static WebElement  selectDefaultViewChkBox;
			
		@FindBy(xpath="//span[@class='icon-sorting icon-font6 toolbar_button_image theme_button_color dropdown-toggle']")
		private static WebElement  sortingBtn;
		
	    @FindBy(xpath="//div[@id='btnCreateView']//span[@class='icon-new icon-font6']")
	    private static WebElement  createViewBtn;

	    @FindBy(xpath="//div[@id='btnCustomize']//span[@class='icon-settings icon-font6']")
	    private static WebElement  customizeBtn;
	   
	    @FindBy(xpath="//span[@class='icon-filter icon-font6']")
	    private static WebElement  filterBtn;
	   
	    @FindBy(xpath="//*[@id='transhomeRefresh']")
	    private static WebElement  refreshBtn;
	   
	    @FindBy(xpath="//*[@id='transhomeRefresh']")
	    private static WebElement  refreshBtnInPVVAT;
	    
	    @FindBy(xpath="//i[@class='icon-font6 icon-export']")
	    private static WebElement  exportBtn;
	   
		// Options Button  (Display Option)	
	    @FindBy(xpath="//span[@id='transOptions']")
	    private static WebElement  optionsBtn;
	   
	    @FindBy(xpath="//a[contains(text(),'Auto Adjust Width')]")
	    private static WebElement  autoAdjustWidthBtn;
	   
	    @FindBy(xpath="//a[contains(text(),'Fit To Screen')]")
	    private static WebElement  fitToScreenBtn;
	   
	    @FindBy(xpath="//a[contains(text(),'Normal')]")
	    private static WebElement  normalBtn;
	   
					
		//Voucher Home Page Grid Header Columns
		@FindBy(xpath="//input[@id='HeaderChkBox']")
		private static WebElement  grid_HeaderChkBox;	
		
		@FindBy(xpath="//*[@id='tblHeadTransRenderRow']/th[3]")
		private static WebElement  grid_Header_Date;
		
		@FindBy(xpath="//*[@id='tblHeadTransRenderRow']/th[4]")
		private static WebElement  grid_Header_VoucherNumber;
		
		@FindBy(xpath="//*[@id='tblHeadTransRenderRow']/th[5]")
		private static WebElement  grid_Header_CreatedBy;
		
		@FindBy(xpath="//*[@id='tblHeadTransRenderRow']/th[6]")
		private static WebElement  grid_Header_ModifiedBy;
		
		@FindBy(xpath="//*[@id='tblHeadTransRenderRow']/th[7]")
		private static WebElement  grid_Header_CreatedDate;
		
		@FindBy(xpath="//*[@id='tblHeadTransRenderRow']/th[8]")
		private static WebElement  grid_Header_ModifiedDate;
		
		@FindBy(xpath="//*[@id='tblHeadTransRenderRow']/th[9]")
		private static WebElement  grid_Header_CreatedTime;
		
		@FindBy(xpath="//*[@id='tblHeadTransRenderRow']/th[10]")
		private static WebElement  grid_Header_ModifiedTime;
		
		@FindBy(xpath="//*[@id='tblHeadTransRenderRow']/th[11]")
		private static WebElement  grid_Header_Suspended;
		
		@FindBy(xpath="//*[@id='tblHeadTransRenderRow']/th[12]")
		private static WebElement  grid_Header_Cancelled;
		
		@FindBy(xpath="//*[@id='tblHeadTransRenderRow']/th[13]")
		private static WebElement  grid_Header_AuthorizationStatus;
		
		
		//Grid Voucher DetailsBtn
		@FindBy(xpath="//tr[@id='trRender_1']//i[@class='icon-info icon-font7']")
		private static WebElement  grid_VoucherDetailsBtn1;
		
		@FindBy(xpath="//tr[@id='trRender_2']//i[@class='icon-info icon-font7']")
		private static WebElement  grid_VoucherDetailsBtn2;
		
		
		//Grid CheckBox 1 And 2
		
		@FindBy(xpath="//tr[@id='trRender_1']//td//input")
		private static WebElement  grid_ChkBox1;
		
		@FindBy(xpath="//tr[@id='trRender_2']//td//input")
		private static WebElement  grid_ChkBox2;
		
		@FindBy(xpath="//tr[@id='trRender_3']//td//input")
		private static WebElement  grid_ChkBox3;
		
		@FindBy(xpath="//tr[@id='trRender_4']//td//input")
		private static WebElement  grid_ChkBox4;
		
		
		// Vouchers Home Page Footer
		@FindBy(xpath="//div[@id='tblFooterReportRender']//input[@id='txtSearch']")
		private static WebElement  searchField;
								
		@FindBy(xpath="//button[@id='frstPage']")
		private static WebElement  firstPageBtn;

		@FindBy(xpath="//button[@id='id_previousPage']")
		private static WebElement  previousPageBtn;

		@FindBy(xpath="//button[@id='id_pagenumber']")
		private static WebElement  currentPageNo;

		@FindBy(xpath="//button[@id='id_nextPage']")
		private static WebElement  nextPageBtn;

		@FindBy(xpath="//button[@id='lstPage']")
		private static WebElement  lastPageBtn;
		
		
		
		//Vouchers Entry Page Ribbon Control Options
		@FindBy(xpath="//span[@class='icon-scroll icon-font6']")
		private static WebElement  freeFlowBtn;
		
		@FindBy(xpath="//span[@class='icon-header icon-font6']")
		private static WebElement  jumpToHeaderSectionBtn;
		
		@FindBy(xpath="//span[@class='icon-panel icon-font6']")
		private static WebElement  jumpToBodySectionBtn;
		
		@FindBy(xpath="//span[@class='icon-footer icon-font6']")
		private static WebElement  jumpToFooterBtn;
		
		@FindBy(xpath="//*[@id='id_transactionentry_new']")
		private static WebElement  new_newBtn;
		
		@FindBy(xpath="//*[@id='id_transactionentry_save']/div[2]")
		private static WebElement  saveBtn;
		
		@FindBy(xpath="//*[@id='id_transactionentry_previous']")
		private static WebElement  previousBtn;
		
		@FindBy(xpath="//*[@id='id_transactionentry_next']")
		private static WebElement  nextBtn;
		
		@FindBy(xpath="//*[@id='id_transactionentry_print']")
		private static WebElement  new_PrintBtn;
		
		@FindBy(xpath="//*[@id='id_transactionentry_suspend']")
		private static WebElement  new_SuspendBtn;
		
		@FindBy(xpath="//*[@id='id_transactionentry_close']")
		private static WebElement  new_CloseBtn;
		
		@FindBy(xpath="//div[@id='id_transactionentry_delete']//span[contains(@class,'icon-delete icon-font6')]")
		private static WebElement  new_DeleteBtn;
		
		@FindBy(xpath="//span[@class='icon-menu icon-font4']")
		private static WebElement  toggleBtn;
		
		@FindBy(xpath="//*[@id='id_transactionentry_copydocument']")
		private static WebElement  copyDocumentOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_copytoclipboard']")
		private static WebElement  copyToClipBoardOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_pastefromclipboard']")
		private static WebElement  pasteFromClipBoardOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_raisechequereturn']")
		private static WebElement  raiseCheckReturnOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_addtostock']")
		private static WebElement  addToStockOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_markconvert']")
		private static WebElement  markConvertedOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_bominput']")
		private static WebElement  bOMinputOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_autoallocate']")
		private static WebElement  autoAllocateOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_reverseentry']")
		private static WebElement  reverseEntryOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_printbarcode']")
		private static WebElement  printBarCodeOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_postingdetails']")
		private static WebElement  postingDetailsOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_exporttoxml']")
		private static WebElement  exportToXMLOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_settings']")
		private static WebElement  settingsOption;
		
		
		
		//Voucher Entry Page Header Fields
		@FindBy(xpath="//input[@id='id_header_1']")
		private static WebElement  documentNumberTxt;
		
		@FindBy(xpath="//*[@id='id_header_1_input_image']/span")
		private static WebElement  documentNumberdropdown_ExpansionBtn;
		
		@FindBy(xpath="//input[@id='id_header_2']")
		private static WebElement  dateTxt;
		
		@FindBy(xpath="//td[@id='id_header_2_input_image']//span[@class='icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow']")
		private static WebElement  dateTxt_CalenderBtn;
		
		@FindBy(xpath="//tr[@id='id_header_2_day_today']//span[@class='theme_color-inverse'][contains(text(),'Today')]")
		private static WebElement  calender_TodayBtn;
		
		@FindBy(xpath="//input[@id='id_header_4']")
		private static WebElement  caskBankAccountTxt;
		
		@FindBy(xpath="//*[@id='id_header_4_input_image']/span")
		private static WebElement  caskBankAccount_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='id_header_4_input_settings']/span")
		private static WebElement  cashBankAccount_SettingsBtn;

			@FindBy(xpath="//div[@id='id_header_4_customize_popup_container']")
			private static WebElement  cashBankAccount_Settings_Container;
									
			@FindBy(xpath="//div[@id='id_header_4_customize_popup_footer']//input[1]")
			private static WebElement  cashBankAccount_Settings_StandardFieldsBtn;
									
						@FindBy(xpath="//select[@id='id_header_4_customize_popup_standardfields_list']")
						private static WebElement  cashBankAccount_Settings_StandardFieldsBtn_Fielddropdown;
															
						@FindBy(xpath="//input[@id='id_header_4_customize_popup_standardfields_header']")
						private static WebElement  cashBankAccount_Settings_StandardFieldsBtn_HeaderTxt;
														
						@FindBy(xpath="//select[@id='id_header_4_customize_popup_standardfields_alignment']")
						private static WebElement  cashBankAccount_Settings_StandardFieldsBtn_Alignmentdropdown;
															
						@FindBy(xpath="//input[@id='id_header_4_customize_popup_standardfields_width']")
						private static WebElement  cashBankAccount_Settings_StandardFieldsBtn_WidthTxt;
															
						@FindBy(xpath="//div[@class='panel']//input[3]")
						private static WebElement  cashBankAccount_Settings_StandardFieldsBtn_OkBtn;
															
						@FindBy(xpath="//div[@class='panel']//input[4]")
						private static WebElement  cashBankAccount_Settings_StandardFieldsBtn_CancelBtn;
									
									
			@FindBy(xpath="//div[@id='id_header_4_search_container']//input[2]")
			private static WebElement  cashBankAccount_Settings_DeleteColumnBtn;
									
			@FindBy(xpath="//div[@class='panel']//input[3]")
			private static WebElement  cashBankAccount_Settings_OkBtn;
							
			@FindBy(xpath="//div[@class='panel']//input[4]")
			private static WebElement  cashBankAccount_Settings_CancelBtn;
				
		@FindBy(xpath="//div[4]/table/tbody/tr[2]/td/input[2]")
		private static WebElement  departmentTxt;
		
		@FindBy(xpath="//*[@id='id_header_268435459']")
		private static WebElement  departmentValuetxt;
		
		@FindBy(xpath="//*[@id='id_header_268435459_input_image']/span")
		private static WebElement  department_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='id_header_268435459_input_settings']/span")
		private static WebElement  department_SettingsBtn;

			@FindBy(xpath="//div[@id='id_header_268435459_customize_popup_container']")
			private static WebElement  department_Settings_Container;
									
			@FindBy(xpath="//div[@id='id_header_268435459_customize_popup_footer']//input[1]")
			private static WebElement  department_Settings_StandardFieldsBtn;
							
				@FindBy(xpath="//select[@id='id_header_268435459_customize_popup_standardfields_list']")
				private static WebElement  department_Settings_StandardFieldsBtn_Fielddropdown;
													
				@FindBy(xpath="//input[@id='id_header_268435459_customize_popup_standardfields_header']")
				private static WebElement  department_Settings_StandardFieldsBtn_HeaderTxt;
												
				@FindBy(xpath="//select[@id='id_header_268435459_customize_popup_standardfields_alignment']")
				private static WebElement  department_Settings_StandardFieldsBtn_Alignmentdropdown;
													
				@FindBy(xpath="//input[@id='id_header_268435459_customize_popup_standardfields_width']")
				private static WebElement  department_Settings_StandardFieldsBtn_WidthTxt;
													
				@FindBy(xpath="//div[@class='panel']//input[3]")
				private static WebElement  department_Settings_StandardFieldsBtn_OkBtn;
													
				@FindBy(xpath="//div[@class='panel']//input[4]")
				private static WebElement  department_Settings_StandardFieldsBtn_CancelBtn;
							
									
			@FindBy(xpath="//div[@id='id_header_268435459_search_container']//input[2]")
			private static WebElement  department_Settings_DeleteColumnBtn;
									
			@FindBy(xpath="//div[@class='panel']//input[3]")
			private static WebElement  department_Settings_OkBtn;
							
			@FindBy(xpath="//div[@class='panel']//input[4]")
			private static WebElement  department_Settings_CancelBtn;

		@FindBy(xpath="//*[@id='id_header_6']")
		private static WebElement  maturityDateTxt;
		
		@FindBy(xpath="//*[@id='id_header_6_input_image']/span']")
		private static WebElement  maturityDateTxt_CalenderBtn;
		
		@FindBy(xpath="//input[@id='id_header_3']")
		private static WebElement  purchaseAccountTxt;
		
		@FindBy(xpath="//*[@id='id_header_3_input_image']/span")
		private static WebElement  purchaseAccount_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='id_header_3_input_settings']/span")
		private static WebElement  purchaseAccount_SettingsBtn;

			@FindBy(xpath="//div[@id='id_header_3_customize_popup_container']")
			private static WebElement  purchaseAccount_Settings_Container;
									
			@FindBy(xpath="//div[@id='id_header_3_customize_popup_footer']//input[1]")
			private static WebElement  purchaseAccount_Settings_StandardFieldsBtn;
							
				@FindBy(xpath="//select[@id='id_header_3_customize_popup_standardfields_list']")
				private static WebElement  purchaseAccount_Settings_StandardFieldsBtn_Fielddropdown;
													
				@FindBy(xpath="//input[@id='id_header_3_customize_popup_standardfields_header']")
				private static WebElement  purchaseAccount_Settings_StandardFieldsBtn_HeaderTxt;
												
				@FindBy(xpath="//select[@id='id_header_3_customize_popup_standardfields_alignment']")
				private static WebElement  purchaseAccount_Settings_StandardFieldsBtn_Alignmentdropdown;
													
				@FindBy(xpath="//input[@id='id_header_3_customize_popup_standardfields_width']")
				private static WebElement  purchaseAccount_Settings_StandardFieldsBtn_WidthTxt;
													
				@FindBy(xpath="//div[@class='panel']//input[3]")
				private static WebElement  purchaseAccount_Settings_StandardFieldsBtn_OkBtn;
													
				@FindBy(xpath="//div[@class='panel']//input[4]")
				private static WebElement  purchaseAccount_Settings_StandardFieldsBtn_CancelBtn;
									
									
			@FindBy(xpath="//div[@id='id_header_3_search_container']//input[2]")
			private static WebElement  purchaseAccount_Settings_DeleteColumnBtn;
									
			@FindBy(xpath="//div[@class='panel']//input[3]")
			private static WebElement  purchaseAccount_Settings_OkBtn;
							
			@FindBy(xpath="//div[@class='panel']//input[4]")
			private static WebElement  purchaseAccount_Settings_CancelBtn;


	@FindBy(xpath="//input[@id='id_header_4']")
	private static WebElement  vendorAccountTxt;

	@FindBy(xpath="//*[@id='id_header_4_input_image']/span")
	private static WebElement  vendorAccount_ExpansionBtn;

	@FindBy(xpath="//*[@id='id_header_4_input_settings']/span")
	private static WebElement  vendorAccount_SettingsBtn;

			@FindBy(xpath="//div[@id='id_header_4_customize_popup_container']")
			private static WebElement  vendorAccount_Settings_Container;
									
			@FindBy(xpath="//div[@id='id_header_4_customize_popup_footer']//input[1]")
			private static WebElement  vendorAccount_Settings_StandardFieldsBtn;
									
						@FindBy(xpath="//select[@id='id_header_4_customize_popup_standardfields_list']")
						private static WebElement  vendorAccount_Settings_StandardFieldsBtn_Fielddropdown;
															
						@FindBy(xpath="//input[@id='id_header_4_customize_popup_standardfields_header']")
						private static WebElement  vendorAccount_Settings_StandardFieldsBtn_HeaderTxt;
														
						@FindBy(xpath="//select[@id='id_header_4_customize_popup_standardfields_alignment']")
						private static WebElement  vendorAccount_Settings_StandardFieldsBtn_Alignmentdropdown;
															
						@FindBy(xpath="//input[@id='id_header_4_customize_popup_standardfields_width']")
						private static WebElement  vendorAccount_Settings_StandardFieldsBtn_WidthTxt;
															
						@FindBy(xpath="//div[@class='panel']//input[3]")
						private static WebElement  vendorAccount_Settings_StandardFieldsBtn_OkBtn;
															
						@FindBy(xpath="//div[@class='panel']//input[4]")
						private static WebElement  vendorAccount_Settings_StandardFieldsBtn_CancelBtn;
									
									
			@FindBy(xpath="//div[@id='id_header_4_search_container']//input[2]")
			private static WebElement  vendorAccount_Settings_DeleteColumnBtn;
									
			@FindBy(xpath="//div[@class='panel']//input[3]")
			private static WebElement  vendorAccount_Settings_OkBtn;
							
			@FindBy(xpath="//div[@class='panel']//input[4]")
			private static WebElement  vendorAccount_Settings_CancelBtn;

	@FindBy(xpath="//input[@id='id_header_3']")
	private static WebElement  salesAccountTxt;

	@FindBy(xpath="//*[@id='id_header_3_input_image']/span")
	private static WebElement  salesAccount_ExpansionBtn;

	@FindBy(xpath="//*[@id='id_header_3_input_settings']/span")
	private static WebElement  salesAccount_SettingsBtn;

			@FindBy(xpath="//div[@id='id_header_3_customize_popup_container']")
			private static WebElement  salesAccount_Settings_Container;
									
			@FindBy(xpath="//div[@id='id_header_3_customize_popup_footer']//input[1]")
			private static WebElement  salesAccount_Settings_StandardFieldsBtn;
									
						@FindBy(xpath="//select[@id='id_header_3_customize_popup_standardfields_list']")
						private static WebElement  salesAccount_Settings_StandardFieldsBtn_Fielddropdown;
															
						@FindBy(xpath="//input[@id='id_header_3_customize_popup_standardfields_header']")
						private static WebElement  salesAccount_Settings_StandardFieldsBtn_HeaderTxt;
														
						@FindBy(xpath="//select[@id='id_header_3_customize_popup_standardfields_alignment']")
						private static WebElement  salesAccount_Settings_StandardFieldsBtn_Alignmentdropdown;
															
						@FindBy(xpath="//input[@id='id_header_3_customize_popup_standardfields_width']")
						private static WebElement  salesAccount_Settings_StandardFieldsBtn_WidthTxt;
															
						@FindBy(xpath="//div[@class='panel']//input[3]")
						private static WebElement  salesAccount_Settings_StandardFieldsBtn_OkBtn;
															
						@FindBy(xpath="//div[@class='panel']//input[4]")
						private static WebElement  salesAccount_Settings_StandardFieldsBtn_CancelBtn;
									
									
			@FindBy(xpath="//div[@id='id_header_3_search_container']//input[2]")
			private static WebElement  salesAccount_Settings_DeleteColumnBtn;
									
			@FindBy(xpath="//div[@class='panel']//input[3]")
			private static WebElement  salesAccount_Settings_OkBtn;
							
			@FindBy(xpath="//div[@class='panel']//input[4]")
			private static WebElement  salesAccount_Settings_CancelBtn;

		@FindBy(xpath="//input[@id='id_header_4']")
		private static WebElement  customerAccountTxt;
		
		@FindBy(xpath="//*[@id='id_header_4_input_image']/span")
		private static WebElement  customerAccount_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='id_header_4_input_settings']/span")
		private static WebElement  customerAccount_SettingsBtn;

			@FindBy(xpath="//div[@id='id_header_4_customize_popup_container']")
			private static WebElement  customerAccount_Settings_Container;
									
			@FindBy(xpath="//div[@id='id_header_4_customize_popup_footer']//input[1]")
			private static WebElement  customerAccount_Settings_StandardFieldsBtn;
									
						@FindBy(xpath="//select[@id='id_header_4_customize_popup_standardfields_list']")
						private static WebElement  customerAccount_Settings_StandardFieldsBtn_Fielddropdown;
															
						@FindBy(xpath="//input[@id='id_header_4_customize_popup_standardfields_header']")
						private static WebElement  customerAccount_Settings_StandardFieldsBtn_HeaderTxt;
														
						@FindBy(xpath="//select[@id='id_header_4_customize_popup_standardfields_alignment']")
						private static WebElement  customerAccount_Settings_StandardFieldsBtn_Alignmentdropdown;
															
						@FindBy(xpath="//input[@id='id_header_4_customize_popup_standardfields_width']")
						private static WebElement  customerAccount_Settings_StandardFieldsBtn_WidthTxt;
															
						@FindBy(xpath="//div[@class='panel']//input[3]")
						private static WebElement  customerAccount_Settings_StandardFieldsBtn_OkBtn;
															
						@FindBy(xpath="//div[@class='panel']//input[4]")
						private static WebElement  customerAccount_Settings_StandardFieldsBtn_CancelBtn;
									
									
			@FindBy(xpath="//div[@id='id_header_4_search_container']//input[2]")
			private static WebElement  customerAccount_Settings_DeleteColumnBtn;
									
			@FindBy(xpath="//div[@class='panel']//input[3]")
			private static WebElement  customerAccount_Settings_OkBtn;
							
			@FindBy(xpath="//div[@class='panel']//input[4]")
			private static WebElement  customerAccount_Settings_CancelBtn;
		
		@FindBy(xpath="//input[@id='id_header_12']")
		private static WebElement  accountTxt;
		
		@FindBy(xpath="//*[@id='id_header_12_input_image']/span")
		private static WebElement  account_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='id_header_4_input_settings']/span")
		private static WebElement  account_SettingsBtn;

			@FindBy(xpath="//div[@id='id_header_4_customize_popup_container']")
			private static WebElement  account_Settings_Container;
									
			@FindBy(xpath="//div[@id='id_header_4_customize_popup_footer']//input[1]")
			private static WebElement  account_Settings_StandardFieldsBtn;
									
				@FindBy(xpath="//select[@id='id_header_4_customize_popup_standardfields_list']")
				private static WebElement  account_Settings_StandardFieldsBtn_Fielddropdown;
													
				@FindBy(xpath="//input[@id='id_header_4_customize_popup_standardfields_header']")
				private static WebElement  account_Settings_StandardFieldsBtn_HeaderTxt;
												
				@FindBy(xpath="//select[@id='id_header_4_customize_popup_standardfields_alignment']")
				private static WebElement  account_Settings_StandardFieldsBtn_Alignmentdropdown;
													
				@FindBy(xpath="//input[@id='id_header_4_customize_popup_standardfields_width']")
				private static WebElement  account_Settings_StandardFieldsBtn_WidthTxt;
													
				@FindBy(xpath="//div[@class='panel']//input[3]")
				private static WebElement  account_Settings_StandardFieldsBtn_OkBtn;
													
				@FindBy(xpath="//div[@class='panel']//input[4]")
				private static WebElement  account_Settings_StandardFieldsBtn_CancelBtn;
									
									
			@FindBy(xpath="//div[@id='id_header_4_search_container']//input[2]")
			private static WebElement  account_Settings_DeleteColumnBtn;
									
			@FindBy(xpath="//div[@class='panel']//input[3]")
			private static WebElement  account_Settings_OkBtn;
							
			@FindBy(xpath="//div[@class='panel']//input[4]")
			private static WebElement  account_Settings_CancelBtn;
		
		@FindBy(xpath="//input[@id='id_header_31']")
		private static WebElement  issuesReceiptsTxt;
		
		@FindBy(xpath="//input[@id='id_header_268435460']")
		private static WebElement  wareHouseTxt;
		
		@FindBy(xpath="//*[@id='id_header_268435460_input_image']/span")
		private static WebElement  wareHouse_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='id_header_268435460_input_settings']/span")
		private static WebElement  wareHouse_SettingsBtn;

			@FindBy(xpath="//div[@id='id_header_268435460_customize_popup_container']")
			private static WebElement  wareHouse_Settings_Container;
									
			@FindBy(xpath="//div[@id='id_header_268435460_customize_popup_footer']//input[1]")
			private static WebElement  wareHouse_Settings_StandardFieldsBtn;
									
				@FindBy(xpath="//select[@id='id_header_268435460_customize_popup_standardfields_list']")
				private static WebElement  wareHouse_Settings_StandardFieldsBtn_Fielddropdown;
													
				@FindBy(xpath="//input[@id='id_header_268435460_customize_popup_standardfields_header']")
				private static WebElement  wareHouse_Settings_StandardFieldsBtn_HeaderTxt;
												
				@FindBy(xpath="//select[@id='id_header_268435460_customize_popup_standardfields_alignment']")
				private static WebElement  wareHouse_Settings_StandardFieldsBtn_Alignmentdropdown;
													
				@FindBy(xpath="//input[@id='id_header_268435460_customize_popup_standardfields_width']")
				private static WebElement  wareHouse_Settings_StandardFieldsBtn_WidthTxt;
													
				@FindBy(xpath="//div[@class='panel']//input[3]")
				private static WebElement  wareHouse_Settings_StandardFieldsBtn_OkBtn;
													
				@FindBy(xpath="//div[@class='panel']//input[4]")
				private static WebElement  wareHouse_Settings_StandardFieldsBtn_CancelBtn;
									
			@FindBy(xpath="//div[@id='id_header_268435460_search_container']//input[2]")
			private static WebElement  wareHouse_Settings_DeleteColumnBtn;
									
			@FindBy(xpath="//div[@class='panel']//input[3]")
			private static WebElement  wareHouse_Settings_OkBtn;
							
			@FindBy(xpath="//div[@class='panel']//input[4]")
			private static WebElement  wareHouse_Settings_CancelBtn;
		
		@FindBy(xpath="//span[contains(text(),'sName')]")
		private static WebElement  customize_sName;
		
		@FindBy(xpath="//span[@class='vcenter'][contains(text(),'Code')]")
		private static WebElement  customize_sCode;
		
		@FindBy(xpath="//span[contains(text(),'Alias')]")
		private static WebElement  customize_Alias;
		
		@FindBy(xpath="//span[@class='vcenter'][contains(text(),'Group')]")
		private static WebElement  customize_bGroup;
		
		@FindBy(xpath="//span[contains(@class,'vcenter')][contains(text(),'Balance')]")
		private static WebElement  customize_Balance;
		
		@FindBy(xpath="//input[@id='id_header_67108920']")
		private static WebElement  appropriateBasedOndropdown;
		
		@FindBy(xpath="//input[@id='id_header_67108921']")
		private static WebElement  additionalValueTxt;
		
		@FindBy(xpath="//input[@id='id_header_21']")
		private static WebElement  raiseReceiptsChkBox;
		
		@FindBy(xpath="//input[@id='id_header_67108865']")
		private static WebElement  receipts_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108868']")
		private static WebElement  payments_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108871']")
		private static WebElement  pettyCash_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108874']")
		private static WebElement  postDatedReceipts_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108865']")
		private static WebElement  postDatedPayments_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108880']")
		private static WebElement  purchaseVouchers_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108881']")
		private static WebElement  purchaseReturns_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108884']")
		private static WebElement  salesInvoice_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108886']")
		private static WebElement  salesOrder_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108928']")
		private static WebElement  salesInvoiceNewNarrationTxt;
		
		
		@FindBy(xpath="//input[@id='id_header_67108885']")
		private static WebElement  salesReturns_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108913']")
		private static WebElement  cashSales_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108914']")
		private static WebElement  hirePurchaseSales_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108890']")
		private static WebElement  journalEntries_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108916']")
		private static WebElement  forexJV_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108910']")
		private static WebElement  interDepartmentalJV_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108888']")
		private static WebElement  nonStandardJournalEntries_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108892']")
		private static WebElement  debitNotes_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108911']")
		private static WebElement  debitNotesLineWise_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108893']")
		private static WebElement  creditNotes_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108912']")
		private static WebElement  creditNotesLineWisePayroll_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108894']")
		private static WebElement  openingBalance_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108891']")
		private static WebElement  fixedAssetDepreciationVoucher_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108909']")
		private static WebElement  requestForQuote_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108903']")
		private static WebElement  purchasesQuotations_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108882']")
		private static WebElement  purchasesOrders_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108882']")
		private static WebElement  materialReceiptsNotes_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108904']")
		private static WebElement  salesQuotationss_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108887']")
		private static WebElement  deliveryNotes_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108917']")
		private static WebElement  posSales_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108918']")
		private static WebElement  productionOrder_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108886']")
		private static WebElement  salesOrders_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108895']")
		private static WebElement  shortagesinStock_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108896']")
		private static WebElement  excessesinStocks_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108897']")
		private static WebElement  stockTransfers_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108898']")
		private static WebElement  openingStocks_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108941']")
		private static WebElement  openingStocksNEW_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108908']")
		private static WebElement  materialRequisition_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108866']")
		private static WebElement  receipts_ChequeNoTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108869']")
		private static WebElement  payments_ChequeNoTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108872']")
		private static WebElement  pettyCash_ChequeNoTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108875']")
		private static WebElement  postDatedReceipts_ChequeNoTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108878']")
		private static WebElement  postDatedPayments_ChequeNoTxt;



		//Voucher Entry Page Body Fields
		
		@FindBy(xpath="//input[@id='id_body_536870916']")
		private static WebElement  pvWareHouseTxt;
		
		@FindBy(xpath="//td[@id='id_body_536870916_input_image']//span[@class='icon-down-arrow optioncontrol_arrow_margin']")
		private static WebElement  pvWareHouse_ExpansionBtn;
		
		@FindBy(xpath="//td[@id='id_body_536870916_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement  pvWareHouse_SettingsBtn;

			@FindBy(xpath="//div[@id='id_body_536870916_customize_popup_container']")
			private static WebElement  pvWareHouse_Settings_Container;
									
			@FindBy(xpath="//div[@id='id_body_536870916_customize_popup_footer']//input[1]")
			private static WebElement  pvWareHouse_Settings_StandardFieldsBtn;
									
				@FindBy(xpath="//select[@id='id_body_536870916_customize_popup_standardfields_list']")
				private static WebElement  pvWareHouse_Settings_StandardFieldsBtn_Fielddropdown;
													
				@FindBy(xpath="//input[@id='id_body_536870916_customize_popup_standardfields_header']")
				private static WebElement  pvWareHouse_Settings_StandardFieldsBtn_HeaderTxt;
												
				@FindBy(xpath="//select[@id='id_body_536870916_customize_popup_standardfields_alignment']")
				private static WebElement  pvWareHouse_Settings_StandardFieldsBtn_Alignmentdropdown;
													
				@FindBy(xpath="//input[@id='id_body_536870916_customize_popup_standardfields_width']")
				private static WebElement  pvWareHouse_Settings_StandardFieldsBtn_WidthTxt;
													
				@FindBy(xpath="//div[@class='panel']//input[3]")
				private static WebElement  pvWareHouse_Settings_StandardFieldsBtn_OkBtn;
													
				@FindBy(xpath="//div[@class='panel']//input[4]")
				private static WebElement  pvWareHouse_Settings_StandardFieldsBtn_CancelBtn;
									
			@FindBy(xpath="//div[@id='id_body_536870916_search_container']//input[2]")
			private static WebElement  pvWareHouse_Settings_DeleteColumnBtn;
									
			@FindBy(xpath="//div[@class='panel']//input[3]")
			private static WebElement  pvWareHouse_Settings_OkBtn;
							
			@FindBy(xpath="//div[@class='panel']//input[4]")
			private static WebElement  pvWareHouse_Settings_CancelBtn;
		
		// Grid Elements
		@FindBy(xpath="//th[@id='id_transaction_entry_detail_table_control_heading_1']")
		private static WebElement  grid_Header_Account;
		
		@FindBy(xpath="//th[@id='id_transaction_entry_detail_table_control_heading_1']")
		private static WebElement  grid_Header_Item;
		
		@FindBy(xpath="//th[@id='id_transaction_entry_detail_table_control_heading_1']")
		private static WebElement  grid_Header_DebitAC;
		
		@FindBy(xpath="//th[@id='id_transaction_entry_detail_table_control_heading_2']")
		private static WebElement  grid_Header_CreditAC;



		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[1]")
		private static WebElement  firstRowIndex;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[1]")
		private static WebElement  secondRowIndex;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[1]")
		private static WebElement  thirdRowIndex;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[1]")
		private static WebElement  fourthRowIndex;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[1]")
		private static WebElement  fifthRowIndex;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[1]")
		private static WebElement  sixthRowIndex;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[1]")
		private static WebElement  seventhRowIndex;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[1]")
		private static WebElement  eigthRowIndex;
		
		@FindBy(xpath="//span[@class='icon-delete icon-font8']")
		private static WebElement  deleteRowBtn;
		
		@FindBy(xpath="//span[@class='icon-insertrow icon-font8']")
		private static WebElement  insertRowBtn;
		
		@FindBy(xpath="//span[@class='icon-selectall icon-font8']")
		private static WebElement  selectRowBtn;
		
		@FindBy(xpath="//span[@class='icon-unselectall1 icon-font8']")
		private static WebElement  unSelectRowBtn;
		
		@FindBy(xpath="//span[@class='icon-clone icon-font8']")
		private static WebElement  duplicateRowBtn;

		
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[2]")
		private static WebElement  select1stRow_1stColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[3]")
		private static WebElement  select1stRow_2ndColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[4]")
		private static WebElement  select1stRow_3rdColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[5]")
		private static WebElement  select1stRow_4thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[6]")
		private static WebElement  select1stRow_5thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[7]")
		private static WebElement  select1stRow_6thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[8]")
		private static WebElement  select1stRow_7thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[9]")
		private static WebElement  select1stRow_8thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[10]")
		private static WebElement  select1stRow_9thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[11]")
		private static WebElement  select1stRow_10thColumn;

		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[12]")
		private static WebElement  select1stRow_11thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[13]")
		private static WebElement  select1stRow_12thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[14]")
		private static WebElement  select1stRow_13thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[15]")
		private static WebElement  select1stRow_14thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[16]")
		private static WebElement  select1stRow_15thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[17]")
		private static WebElement  select1stRow_16thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[18]")
		private static WebElement  select1stRow_17thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[19]")
		private static WebElement  select1stRow_18thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[2]")
		private static WebElement  select2ndRow_1stColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[3]")
		private static WebElement  select2ndRow_2ndColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[4]")
		private static WebElement  select2ndRow_3rdColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[5]")
		private static WebElement  select2ndRow_4thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[6]")
		private static WebElement  select2ndRow_5thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[7]")
		private static WebElement  select2ndRow_6thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[8]")
		private static WebElement  select2ndRow_7thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[9]")
		private static WebElement  select2ndRow_8thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[10]")
		private static WebElement  select2ndRow_9thColumn;

		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[12]")
		private static WebElement  select2ndRow_11thColumn;

		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[13]")
		private static WebElement  select2ndRow_12thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[11]")
		private static WebElement  select2ndRow_10thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[15]")
		private static WebElement  select2ndRow_14thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[16]")
		private static WebElement  select2ndRow_15thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[17]")
		private static WebElement  select2ndRow_16thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[18]")
		private static WebElement  select2ndRow_17thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[19]")
		private static WebElement  select2ndRow_18thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[20]")
		private static WebElement  select2ndRow_19thColumn;

		
		
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[2]")
		private static WebElement  select3rdRow_1stColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[3]")
		private static WebElement  select3rdRow_2ndColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[4]")
		private static WebElement  select3rdRow_3rdColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[5]")
		private static WebElement  select3rdRow_4thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[6]")
		private static WebElement  select3rdRow_5thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[7]")
		private static WebElement  select3rdRow_6thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[8]")
		private static WebElement  select3rdRow_7thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[9]")
		private static WebElement  select3rdRow_8thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[10]")
		private static WebElement  select3rdRow_9thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[11]")
		private static WebElement  select3rdRow_10thColumn;

	    @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[12]")
		private static WebElement  select3rdRow_11thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[13]")
		private static WebElement  select3rdRow_12thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[14]")
		private static WebElement  select3rdRow_13thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[15]")
		private static WebElement  select3rdRow_14thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[16]")
		private static WebElement  select3rdRow_15thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[17]")
		private static WebElement  select3rdRow_16thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[18]")
		private static WebElement  select3rdRow_17thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[3]/td[19]")
		private static WebElement  select3rdRow_18thColumn;




		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[2]")
		private static WebElement  select4thRow_1stColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[3]")
		private static WebElement  select4thRow_2ndColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[4]")
		private static WebElement  select4thRow_3rdColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[5]")
		private static WebElement  select4thRow_4thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[6]")
		private static WebElement  select4thRow_5thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[7]")
		private static WebElement  select4thRow_6thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[8]")
		private static WebElement  select4thRow_7thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[9]")
		private static WebElement  select4thRow_8thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[10]")
		private static WebElement  select4thRow_9thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[11]")
		private static WebElement  select4thRow_10thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[12]")
		private static WebElement  select4thRow_11thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[13]")
		private static WebElement  select4thRow_12thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[14]")
		private static WebElement  select4thRow_13thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[15]")
		private static WebElement  select4thRow_14thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[16]")
		private static WebElement  select4thRow_15thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[17]")
		private static WebElement  select4thRow_16thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[18]")
		private static WebElement  select4thRow_17thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[19]")
		private static WebElement  select4thRow_18thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[20]")
		private static WebElement  select4thRow_19thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[4]/td[21]")
		private static WebElement  select4thRow_20thColumn;



		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[2]")
		private static WebElement  select5thRow_1stColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[3]")
		private static WebElement  select5thRow_2ndColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[4]")
		private static WebElement  select5thRow_3rdColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[5]")
		private static WebElement  select5thRow_4thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[6]")
		private static WebElement  select5thRow_5thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[7]")
		private static WebElement  select5thRow_6thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[8]")
		private static WebElement  select5thRow_7thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[9]")
		private static WebElement  select5thRow_8thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[10]")
		private static WebElement  select5thRow_9thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[11]")
		private static WebElement  select5thRow_10thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[12]")
		private static WebElement  select5thRow_11thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[13]")
		private static WebElement  select5thRow_12thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[14]")
		private static WebElement  select5thRow_13thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[15]")
		private static WebElement  select5thRow_14thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[16]")
		private static WebElement  select5thRow_15thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[17]")
		private static WebElement  select5thRow_16thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[18]")
		private static WebElement  select5thRow_17thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[19]")
		private static WebElement  select5thRow_18thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[20]")
		private static WebElement  select5thRow_19thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[5]/td[21]")
		private static WebElement  select5thRow_20thColumn;




		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[2]")
		private static WebElement  select6thRow_1stColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[3]")
		private static WebElement  select6thRow_2ndColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[4]")
		private static WebElement  select6thRow_3rdColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[5]")
		private static WebElement  select6thRow_4thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[6]")
		private static WebElement  select6thRow_5thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[7]")
		private static WebElement  select6thRow_6thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[8]")
		private static WebElement  select6thRow_7thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[9]")
		private static WebElement  select6thRow_8thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[10]")
		private static WebElement  select6thRow_9thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[11]")
		private static WebElement  select6thRow_10thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[12]")
		private static WebElement  select6thRow_11thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[13]")
		private static WebElement  select6thRow_12thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[14]")
		private static WebElement  select6thRow_13thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[15]")
		private static WebElement  select6thRow_14thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[16]")
		private static WebElement  select6thRow_15thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[17]")
		private static WebElement  select6thRow_16thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[18]")
		private static WebElement  select6thRow_17thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[19]")
		private static WebElement  select6thRow_18thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[20]")
		private static WebElement  select6thRow_19thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[6]/td[21]")
		private static WebElement  select6thRow_20thColumn;


		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[2]")
		private static WebElement  select7thRow_1stColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[3]")
		private static WebElement  select7thRow_2ndColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[4]")
		private static WebElement  select7thRow_3rdColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[5]")
		private static WebElement  select7thRow_4thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[6]")
		private static WebElement  select7thRow_5thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[7]")
		private static WebElement  select7thRow_6thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[8]")
		private static WebElement  select7thRow_7thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[9]")
		private static WebElement  select7thRow_8thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[7]/td[10]")
		private static WebElement  select7thRow_9thColumn;



		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[2]")
		private static WebElement  select8thRow_1stColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[3]")
		private static WebElement  select8thRow_2ndColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[4]")
		private static WebElement  select8thRow_3rdColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[5]")
		private static WebElement  select8thRow_4thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[6]")
		private static WebElement  select8thRow_5thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[7]")
		private static WebElement  select8thRow_6thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[8]")
		private static WebElement  select8thRow_7thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[9]")
		private static WebElement  select8thRow_8thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[8]/td[10]")
		private static WebElement  select8thRow_9thColumn;


		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[2]")
		private static WebElement  select9thRow_1stColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[3]")
		private static WebElement  select9thRow_2ndColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[4]")
		private static WebElement  select9thRow_3rdColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[5]")
		private static WebElement  select9thRow_4thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[6]")
		private static WebElement  select9thRow_5thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[7]")
		private static WebElement  select9thRow_6thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[8]")
		private static WebElement  select9thRow_7thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[9]")
		private static WebElement  select9thRow_8thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[9]/td[10]")
		private static WebElement  select9thRow_9thColumn;



		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[2]")
		private static WebElement  select10thRow_1stColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[3]")
		private static WebElement  select10thRow_2ndColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[4]")
		private static WebElement  select10thRow_3rdColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[5]")
		private static WebElement  select10thRow_4thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[6]")
		private static WebElement  select10thRow_5thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[7]")
		private static WebElement  select10thRow_6thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[8]")
		private static WebElement  select10thRow_7thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[9]")
		private static WebElement  select10thRow_8thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[10]/td[10]")
		private static WebElement  select10thRow_9thColumn;


		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[2]")
		private static WebElement  select11thRow_1stColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[3]")
		private static WebElement  select11thRow_2ndColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[4]")
		private static WebElement  select11thRow_3rdColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[5]")
		private static WebElement  select11thRow_4thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[6]")
		private static WebElement  select11thRow_5thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[7]")
		private static WebElement  select11thRow_6thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[8]")
		private static WebElement  select11thRow_7thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[9]")
		private static WebElement  select11thRow_8thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[11]/td[10]")
		private static WebElement  select11thRow_9thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[2]")
		private static WebElement  select12thRow_1stColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[3]")
		private static WebElement  select12thRow_2ndColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[4]")
		private static WebElement  select12thRow_3rdColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[5]")
		private static WebElement  select12thRow_4thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[6]")
		private static WebElement  select12thRow_5thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[7]")
		private static WebElement  select12thRow_6thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[8]")
		private static WebElement  select12thRow_7thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[9]")
		private static WebElement  select12thRow_8thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[12]/td[10]")
		private static WebElement  select12thRow_9thColumn;

		@FindBy(xpath="//input[@id='id_body_12']")
		private static WebElement  enter_AccountTxt;
		
		@FindBy(xpath="//*[@id='id_body_12_input_image']/span")
		private static WebElement  enter_Account_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='id_body_12_input_settings']/span")
		private static WebElement  enter_Account_SettingsBtn;

			@FindBy(xpath="//div[@id='id_header_12_customize_popup_container']")
			private static WebElement  enter_Account_Settings_Container;
									
			@FindBy(xpath="//div[@id='id_header_12_customize_popup_footer']//input[1]")
			private static WebElement  enter_Account_Settings_StandardFieldsBtn;
									
						@FindBy(xpath="//select[@id='id_header_12_customize_popup_standardfields_list']")
						private static WebElement  enter_Account_Settings_StandardFieldsBtn_Fielddropdown;
															
						@FindBy(xpath="//input[@id='id_header_12_customize_popup_standardfields_header']")
						private static WebElement  enter_Account_Settings_StandardFieldsBtn_HeaderTxt;
														
						@FindBy(xpath="//select[@id='id_header_12_customize_popup_standardfields_alignment']")
						private static WebElement  enter_Account_Settings_StandardFieldsBtn_Alignmentdropdown;
															
						@FindBy(xpath="//input[@id='id_header_12_customize_popup_standardfields_width']")
						private static WebElement  enter_Account_Settings_StandardFieldsBtn_WidthTxt;
															
						@FindBy(xpath="//div[contains(@class,'panel')]//input[3]")
						private static WebElement  enter_Account_Settings_StandardFieldsBtn_OkBtn;
															
						@FindBy(xpath="//div[contains(@class,'panel')]//input[4]")
						private static WebElement  enter_Account_Settings_StandardFieldsBtn_CancelBtn;
									
			@FindBy(xpath="//div[@id='id_header_12_search_container']//input[2]")
			private static WebElement  enter_Account_Settings_DeleteColumnBtn;
									
			@FindBy(xpath="//div[contains(@class,'panel')]//input[3]")
			private static WebElement  enter_Account_Settings_OkBtn;
							
			@FindBy(xpath="//div[contains(@class,'panel')]//input[4]")
			private static WebElement  enter_Account_Settings_CancelBtn;

		@FindBy(xpath="//input[@id='id_body_16']")
		private static WebElement  enter_Amount;
		
		@FindBy(xpath="//textarea[@id='id_body_16777219']")
		private static WebElement  enter_Receipts_Remarks;
		
		@FindBy(xpath="//textarea[@id='id_body_16777222']")
		private static WebElement  enter_Payments_Remarks;
		
		@FindBy(xpath="//textarea[@id='id_body_16777225']")
		private static WebElement  enter_PettyCash_Remarks;
		
		@FindBy(xpath="//textarea[@id='id_body_16777228']")
		private static WebElement  enter_PostDatedReceipts_Remarks;
		
		@FindBy(xpath="//textarea[@id='id_body_16777231']")
		private static WebElement  enter_PostDatedPaymentss_Remarks;
		
		@FindBy(xpath="//input[@id='id_body_23']")
		private static WebElement  enter_ItemTxt;
		
		@FindBy(xpath="//*[@id='id_body_23_input_image']/span")
		private static WebElement  item_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='id_body_23_input_settings']/span")
		private static WebElement  item_SettingsBtn;

			@FindBy(xpath="//div[@id='id_body_23_customize_popup_container']")
			private static WebElement  item_Settings_Container;
									
			@FindBy(xpath="//div[@id='id_body_23_customize_popup_footer']//input[1]")
			private static WebElement  item_Settings_StandardFieldsBtn;
									
				@FindBy(xpath="//select[@id='id_body_23_customize_popup_standardfields_list']")
				private static WebElement  item_Settings_StandardFieldsBtn_Fielddropdown;
													
				@FindBy(xpath="//input[@id='id_body_23_customize_popup_standardfields_header']")
				private static WebElement  item_Settings_StandardFieldsBtn_HeaderTxt;
												
				@FindBy(xpath="//select[@id='id_body_23_customize_popup_standardfields_alignment']")
				private static WebElement  item_Settings_StandardFieldsBtn_Alignmentdropdown;
													
				@FindBy(xpath="//input[@id='id_body_23_customize_popup_standardfields_width']")
				private static WebElement  item_Settings_StandardFieldsBtn_WidthTxt;
													
				@FindBy(xpath="//div[contains(@class,'panel')]//input[3]")
				private static WebElement  item_Settings_StandardFieldsBtn_OkBtn;
													
				@FindBy(xpath="//div[contains(@class,'panel')]//input[4]")
				private static WebElement  item_Settings_StandardFieldsBtn_CancelBtn;
									
			@FindBy(xpath="//div[@id='id_body_23_search_container']//input[2]")
			private static WebElement  item_Settings_DeleteColumnBtn;
									
			@FindBy(xpath="//div[contains(@class,'panel')]//input[3]")
			private static WebElement  item_Settings_OkBtn;
							
			@FindBy(xpath="//div[contains(@class,'panel')]//input[4]")
			private static WebElement  item_Settings_CancelBtn;

		@FindBy(xpath="//input[@id='id_body_24']")
		private static WebElement  enter_UnitTxt;
		
		@FindBy(xpath="//*[@id='id_body_24_input_image']/span")
		private static WebElement  unit_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='id_body_24_input_settings']/span")
		private static WebElement  unit_SettingsBtn;

			@FindBy(xpath="//div[@id='id_body_24_customize_popup_container']")
			private static WebElement  unit_Settings_Container;
									
			@FindBy(xpath="//div[@id='id_body_24_customize_popup_footer']//input[1]")
			private static WebElement  unit_Settings_StandardFieldsBtn;
									
				@FindBy(xpath="//select[@id='id_body_24_customize_popup_standardfields_list']")
				private static WebElement  unit_Settings_StandardFieldsBtn_Fielddropdown;
													
				@FindBy(xpath="//input[@id='id_body_24_customize_popup_standardfields_header']")
				private static WebElement  unit_Settings_StandardFieldsBtn_HeaderTxt;
												
				@FindBy(xpath="//select[@id='id_body_24_customize_popup_standardfields_alignment']")
				private static WebElement  unit_Settings_StandardFieldsBtn_Alignmentdropdown;
													
				@FindBy(xpath="//input[@id='id_body_24_customize_popup_standardfields_width']")
				private static WebElement  unit_Settings_StandardFieldsBtn_WidthTxt;
													
				@FindBy(xpath="//div[contains(@class,'panel')]//input[3]")
				private static WebElement  unit_Settings_StandardFieldsBtn_OkBtn;
													
				@FindBy(xpath="//div[contains(@class,'panel')]//input[4]")
				private static WebElement  unit_Settings_StandardFieldsBtn_CancelBtn;
									
			@FindBy(xpath="//div[@id='id_body_24_search_container']//input[2]")
			private static WebElement  unit_Settings_DeleteColumnBtn;
									
			@FindBy(xpath="//div[contains(@class,'panel')]//input[3]")
			private static WebElement  unit_Settings_OkBtn;
							
			@FindBy(xpath="//div[contains(@class,'panel')]//input[4]")
			private static WebElement  unit_Settings_CancelBtn;

				
		@FindBy(xpath="//input[@id='id_body_26']")
		private static WebElement  enter_Quantity;
		
		@FindBy(xpath="//input[@id='id_body_27']")
		private static WebElement  enter_Rate;
		
		@FindBy(xpath="//*[@id='id_body_28']")
		private static WebElement  enter_Gross;
		
		@FindBy(xpath="//input[@id='id_body_13']")
		private static WebElement  enter_Batch;
		
		@FindBy(xpath="//input[@id='id_body_12']")
		private static WebElement  enter_DebitACTxt;
		
		@FindBy(xpath="//*[@id='id_body_12_input_image']/span")
		private static WebElement  enter_DebitAC_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='id_body_12_input_settings']/span")
		private static WebElement  enter_DebitAC_SettingsBtn;

			@FindBy(xpath="//div[@id='id_header_12_customize_popup_container']")
			private static WebElement  enter_DebitAC_Settings_Container;
									
			@FindBy(xpath="//div[@id='id_header_12_customize_popup_footer']//input[1]")
			private static WebElement  enter_DebitAC_Settings_StandardFieldsBtn;
									
				@FindBy(xpath="//select[@id='id_header_12_customize_popup_standardfields_list']")
				private static WebElement  enter_DebitAC_Settings_StandardFieldsBtn_Fielddropdown;
													
				@FindBy(xpath="//input[@id='id_header_12_customize_popup_standardfields_header']")
				private static WebElement  enter_DebitAC_Settings_StandardFieldsBtn_HeaderTxt;
												
				@FindBy(xpath="//select[@id='id_header_12_customize_popup_standardfields_alignment']")
				private static WebElement  enter_DebitAC_Settings_StandardFieldsBtn_Alignmentdropdown;
													
				@FindBy(xpath="//input[@id='id_header_12_customize_popup_standardfields_width']")
				private static WebElement  enter_DebitAC_Settings_StandardFieldsBtn_WidthTxt;
													
				@FindBy(xpath="//div[contains(@class,'panel')]//input[3]")
				private static WebElement  enter_DebitAC_Settings_StandardFieldsBtn_OkBtn;
													
				@FindBy(xpath="//div[contains(@class,'panel')]//input[4]")
				private static WebElement  enter_DebitAC_Settings_StandardFieldsBtn_CancelBtn;
									
			@FindBy(xpath="//div[@id='id_header_12_search_container']//input[2]")
			private static WebElement  enter_DebitAC_Settings_DeleteColumnBtn;
									
			@FindBy(xpath="//div[contains(@class,'panel')]//input[3]")
			private static WebElement  enter_DebitAC_Settings_OkBtn;
							
			@FindBy(xpath="//div[contains(@class,'panel')]//input[4]")
			private static WebElement  enter_DebitAC_Settings_CancelBtn;

		@FindBy(xpath="//input[@id='id_body_39']")
		private static WebElement  enter_CreditACTxt;
		
		@FindBy(xpath="//*[@id='id_body_39_input_image']/span")
		private static WebElement  enter_CreditAC_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='id_body_39_input_settings']/span")
		private static WebElement  enter_CreditAC_SettingsBtn;

			@FindBy(xpath="//div[@id='id_header_39_customize_popup_container']")
			private static WebElement  enter_CreditAC_Settings_Container;
									
			@FindBy(xpath="//div[@id='id_header_39_customize_popup_footer']//input[1]")
			private static WebElement  enter_CreditAC_Settings_StandardFieldsBtn;
									
				@FindBy(xpath="//select[@id='id_header_39_customize_popup_standardfields_list']")
				private static WebElement  enter_CreditAC_Settings_StandardFieldsBtn_Fielddropdown;
													
				@FindBy(xpath="//input[@id='id_header_39_customize_popup_standardfields_header']")
				private static WebElement  enter_CreditAC_Settings_StandardFieldsBtn_HeaderTxt;
												
				@FindBy(xpath="//select[@id='id_header_39_customize_popup_standardfields_alignment']")
				private static WebElement  enter_CreditAC_Settings_StandardFieldsBtn_Alignmentdropdown;
													
				@FindBy(xpath="//input[@id='id_header_39_customize_popup_standardfields_width']")
				private static WebElement  enter_CreditAC_Settings_StandardFieldsBtn_WidthTxt;
													
				@FindBy(xpath="//div[contains(@class,'panel')]//input[3]")
				private static WebElement  enter_CreditAC_Settings_StandardFieldsBtn_OkBtn;
													
				@FindBy(xpath="//div[contains(@class,'panel')]//input[4]")
				private static WebElement  enter_CreditAC_Settings_StandardFieldsBtn_CancelBtn;
									
			@FindBy(xpath="//div[@id='id_header_39_search_container']//input[2]")
			private static WebElement  enter_CreditAC_Settings_DeleteColumnBtn;
									
			@FindBy(xpath="//div[contains(@class,'panel')]//input[3]")
			private static WebElement  enter_CreditAC_Settings_OkBtn;
							
			@FindBy(xpath="//div[contains(@class,'panel')]//input[4]")
			private static WebElement  enter_CreditAC_Settings_CancelBtn;

		@FindBy(xpath="//input[@id='id_body_18']")
		private static WebElement  enter_DebitTxt;
		
		@FindBy(xpath="//input[@id='id_body_19']")
		private static WebElement  enter_CreditTxt;
		
		@FindBy(xpath="//textarea[@id='id_body_16777241']")
		private static WebElement  enter_NonStandardJournalEntries_RemarksTxt;
		
		@FindBy(xpath="//input[@id='id_body_87']")
		private static WebElement  enter_Warehouse2Txt;
		
		@FindBy(xpath="//*[@id='id_body_87_input_image']/span")
		private static WebElement  enter_Warehouse2_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='id_body_87_input_settings']/span")
		private static WebElement  enter_Warehouse2_SettingsBtn;

			@FindBy(xpath="//div[@id='id_header_87_customize_popup_container']")
			private static WebElement  enter_WareHouse2_Settings_Container;
									
			@FindBy(xpath="//div[@id='id_header_87_customize_popup_footer']//input[1]")
			private static WebElement  enter_WareHouse2_Settings_StandardFieldsBtn;
									
				@FindBy(xpath="//select[@id='id_header_87_customize_popup_standardfields_list']")
				private static WebElement  enter_WareHouse2_Settings_StandardFieldsBtn_Fielddropdown;
													
				@FindBy(xpath="//input[@id='id_header_87_customize_popup_standardfields_header']")
				private static WebElement  enter_WareHouse2_Settings_StandardFieldsBtn_HeaderTxt;
												
				@FindBy(xpath="//select[@id='id_header_87_customize_popup_standardfields_alignment']")
				private static WebElement  enter_WareHouse2_Settings_StandardFieldsBtn_Alignmentdropdown;
													
				@FindBy(xpath="//input[@id='id_header_87_customize_popup_standardfields_width']")
				private static WebElement  enter_WareHouse2_Settings_StandardFieldsBtn_WidthTxt;
													
				@FindBy(xpath="//div[contains(@class,'panel')]//input[3]")
				private static WebElement  enter_WareHouse2_Settings_StandardFieldsBtn_OkBtn;
													
				@FindBy(xpath="//div[contains(@class,'panel')]//input[4]")
				private static WebElement  enter_WareHouse2_Settings_StandardFieldsBtn_CancelBtn;
									
			@FindBy(xpath="//div[@id='id_header_87_search_container']//input[2]")
			private static WebElement  enter_WareHouse2_Settings_DeleteColumnBtn;
									
			@FindBy(xpath="//div[contains(@class,'panel')]//input[3]")
			private static WebElement  enter_WareHouse2_Settings_OkBtn;
							
			@FindBy(xpath="//div[contains(@class,'panel')]//input[4]")
			private static WebElement  enter_WareHouse2_Settings_CancelBtn;


		// Right Panel InfoSide Bar	
		@FindBy(xpath="//*[@id='id_transactionentry_infopanel_container']/div[1]/div[2]/span")
		private static WebElement  infoSideBarCustomizeBtn;

		@FindBy(xpath="//*[@id='id_transactionentry_infopanel_container']/div[1]/div[3]/span")
		private static WebElement  infoSideBarMinimizeExpandBtn;
			
			@FindBy(xpath="//li[@id='Trans_Dash_Save']//span[text()='Save']")
			private static WebElement  customizeSaveBtn;
			
			@FindBy(xpath="//span[@class='noWrap clsBlueColor'][contains(text(),'Cancel')]")
			private static WebElement  customizeCancelBtn;
			
			@FindBy(xpath="//input[@id='searchBoxTrans']")
			private static WebElement  customizeSearchTxt;
						
			@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//i[@class='icon-search searchicon']")
			private static WebElement  customizeSearchBtn;
						
			@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//i[@id='g']")
			private static WebElement  graphBtn;
						
				@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//li[@id='Newgraph']")
				private static WebElement  newGraphOption;
					
			@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//i[@id='r']")
			private static WebElement  reportBtn;
				
				@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//li[@id='NewReport']")
				private static WebElement  newReportOption;
						
			@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//i[@id='i']")
			private static WebElement  infoPanelBtn;
						
				@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//li[@id='NewInfopanel']")
				private static WebElement  newInfoPanelOption;
				
			@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//i[@id='w']")
			private static WebElement  workFlowBtn;
							
					@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//li[@id='NewWorkflow']")
					private static WebElement  newWorkFlowOption;
							
			@FindBy(xpath="//a[@id='DocumentInfoOption']")
			private static WebElement  otherDashlets;
							
				@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//label[contains(@class,'theme_icon-color')][contains(text(),'Document Info')]")
				private static WebElement  documentInfo;
		
		@FindBy(xpath="//span[@class='icon-left-and-right-panel-icon icon-font6 no_padding_left_right']")
		private static WebElement  infoSideBarExpandBtn;


		// Footer Section
		@FindBy(xpath="//label[@id='id_transactionentry_footer_panel_summary_value_net']")
		private static WebElement  netLabel;
		
		@FindBy(xpath="//*[@id='id_transactionentry_footer_panel_summary_value_net']/span[2]")
		private static WebElement  netAmount;

		
		//Validation and Confirmation messages
		@FindBy(xpath="//div[@class='theme_color font-6']")
		public static WebElement errorMessage;
		 
		@FindBy(xpath="//span[@class='icon-reject2 theme_color']")
		public static WebElement errorMessageCloseBtn;
		
		//Dash board Page
		@FindBy(xpath="//*[@id='dashName']")
	    private static WebElement labelDashboard ;
		
		@FindBy(xpath="//*[@id='Select_dash']")
		private static WebElement selectDashboard ;
			
		@FindBy(xpath="//*[@id='Dashboard_AddDash']")
	    private static WebElement newAddDashBoard;
		 
	    @FindBy(xpath="//*[@id='Dashboard_Dash_Config']")
	    private static WebElement dashboardCustomizationSettings;
	    
	    
	 // Login Page
		@FindBy(xpath = "//*[@id='txtUsername']")
		private static WebElement username;
		

			
		private static String xlfile;
		private static String resPass="Pass";
		private static String resFail="Fail";
		private static ExcelReader excelReader;
		
		// Home Menu
		@FindBy(xpath = "//*[@id='1']/div/span")
		public static WebElement homeMenu;

		//Financial Menu
		@FindBy (xpath="//span[@class='icon-financial icon-font1']")
		private static WebElement financialMenu;
		
		//Fixed Assets Menu
		@FindBy (xpath="//span[@class='icon-assets icon-font1']")
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
		
		private static String xlSheetName = "SmokeVouchers";
		
		private static boolean Status;

	    
		//Inventory Menu and Sub Menus		
		@FindBy(xpath="//div[contains(text(),'Inventory')]")
		private static WebElement  inventoryMenu; 
		
			@FindBy(xpath="//a[@id='137']//span[contains(text(),'Transactions')]")
			private static WebElement  inventoryTransactionsMenu; 
				
				@FindBy(xpath="//*[@id='139']/span")
				private static WebElement  inventoryTransactionsPurchasesMenu; 
			
					@FindBy(xpath="//*[@id='2010']/span")
					private static WebElement  requestForQuoteVoucher;
					
					@FindBy(xpath="//*[@id='2011']/span")
					private static WebElement  purchasesQuotationsVoucher;
						
					@FindBy(xpath="//a[@id='2012']//span[contains(text(),'Purchases Orders')]")
					private static WebElement  purchasesOrdersVoucher;
				
					@FindBy(xpath="//*[@id='2013']/span")
					private static WebElement  materialReceiptNotesVoucher;
						
				@FindBy(xpath="//*[@id='140']/span")
				private static WebElement  inventoyTransactionsSalesMenu; 
					
					@FindBy(xpath="//*[@id='2018']/span")
					private static WebElement  salesQuotationsVoucher;
					
					@FindBy(xpath="//*[@id='2019']/span")
					private static WebElement  deliveryNotesVoucher;
					
					@FindBy(xpath="//*[@id='2022']/span")
					private static WebElement  posSalesVoucher;
					
					@FindBy(xpath="//*[@id='2045']/span")
					private static WebElement  productionOrdersVoucher;
					
					@FindBy(xpath="//*[@id='2017']/span")
					private static WebElement  salesOrdersVoucher;
								
				@FindBy(xpath="//a[@id='2033']//span[contains(text(),'Stocks')]")
				private static WebElement  inventoryTransactionsStocksMenu; 
					
					@FindBy(xpath="//*[@id='2034']/span")
					private static WebElement  shortagesInStockVoucher;
					
					@FindBy(xpath="//*[@id='2035']/span")
					private static WebElement  excessesInStocksVoucher;
						
					@FindBy(xpath="//*[@id='2036']/span")
					private static WebElement  stockTransfersVoucher;
					
					@FindBy(xpath="//a[@id='2037']//span[contains(text(),'Opening Stocks')]")
					private static WebElement  openingStocksVoucher;
					
					@FindBy(xpath="//a[@id='2061']//span[contains(text(),'Opening Stock New')]")
					private static WebElement  openingStocksNewVoucher;
					
					@FindBy(xpath="//*[@id='2038']/span")
					private static WebElement  materialRequisitionVoucher;
							
					@FindBy(xpath="//*[@id='2049']/span")
					private static WebElement  stockAdjustmentVoucher;
						
				@FindBy(xpath="//*[@id='164']/span")
				private static WebElement  holdAndUnholdStockVoucher;
				
				@FindBy(xpath="//*[@id='93']/span")
				private static WebElement  stockReconciliation;
				
				@FindBy(xpath="//*[@id='99']/span")
				private static WebElement  stockAllocation;
	    
	    
				@FindBy(xpath = "//ul[@id='navigation_menu']/li/a/div/div")
				private static List<WebElement> menusList;
				
				@FindBy(xpath = "//ul[@id='navigation_menu']/li[3]/ul/li/a/span")
				private static List<WebElement> inventoryMenusList;
				
				@FindBy(xpath = "//ul[@id='navigation_menu']/li[3]/ul/li[1]/ul/li/a/span")
				private static List<WebElement> inventoryTransactionsMenusList;
				
				@FindBy(xpath = "//ul[@id='navigation_menu']/li[3]/ul/li[1]/ul/li[3]/ul/li/a/span")
				private static List<WebElement> inventoryTransactionsStocksMenusList;
				
				@FindBy(xpath = "//*[@id='dv_TranHomeHeading']/nav/div[2]/div/div/div[2]")
				private static List<WebElement> openingStocksNewToolBarButtons;
				
				
				@FindBy(xpath="//*[@id='ulCompanyDetails_HomePage']/li[1]")
				private static WebElement companyName;

				@FindBy(xpath="//*[@id='companyLogo']")
				private static WebElement companyLogo;

				
				//User Name Display
				@FindBy(xpath="/html/body/section/div[2]/header/nav/div/ul/li[6]/a/span")
				private static WebElement userNameDisplay;
				
	 public boolean checkSignInToVerifyBudgetsMRPOVouchers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	 {
	  	excelReader = new ExcelReader(POJOUtility.getExcelPath());
	  	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	  	   
	  	try
	  	{  
	 		 getDriver().navigate().refresh();
	 		 Thread.sleep(1999);
	 		 
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
	      	 
	  		LoginPage lp                          = new LoginPage(getDriver());
	        
	          String unamelt                        = excelReader.getCellData(xlSheetName, 7, 6);
	          String pawslt                         = excelReader.getCellData(xlSheetName, 8, 6);
	  		
	  		LoginPage.enterUserName(unamelt);
	  		LoginPage.enterPassword(pawslt);
	  		
	  		LoginPage.clickOnSignInBtn();
	  		
	  		
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	  		
	  		boolean actHomeMenuIsDisplay           = homeMenu.isDisplayed();
	  		boolean actFinancialMenuIsDisplay      = financialMenu.isDisplayed();
	  		boolean actInventoryMenuIsDisplay      = inventoryMenu.isDisplayed();
	  		boolean actFixedAssetsMenuIsDisplay    = fixedAssetsMenu.isDisplayed();
	  		boolean actProductionMenuIsDisplay     = productionMenu.isDisplayed();
	  		boolean actPointOfSaleMenuIsDisplay    = pointOfSaleMenu.isDisplayed();
	  		boolean actQualityControlMenuIsDisplay = qualityControlMenu.isDisplayed();
	  		boolean actSettingsMenuIsDisplay       = SettingsMenu.isDisplayed();
	  		
	  		boolean expHomeMenuIsDisplay           = true;
	  		boolean expFinancialMenuIsDisplay      = true;
	  		boolean expInventoryMenuIsDisplay      = true;
	  		boolean expFixedAssetsMenuIsDisplay    = true;
	  		boolean expProductionMenuIsDisplay     = true;
	  		boolean expPointOfSaleMenuIsDisplay    = true;
	  		boolean expQualityControlMenuIsDisplay = true;
	  		boolean expSettingsMenuIsDisplay       = true;
	  		
	  		System.out.println("Home Menu Display Value Actual            : " + actHomeMenuIsDisplay +           " Value Expected : " + expHomeMenuIsDisplay);
	  		System.out.println("Financial Menu Display Value Actual       : " + actFinancialMenuIsDisplay +      " Value Expected : " + expFinancialMenuIsDisplay);
	  		System.out.println("Inventory Menu Display Value Actual       : " + actInventoryMenuIsDisplay +      " Value Expected : " + expInventoryMenuIsDisplay);
	  		System.out.println("Fixed Assets Menu Display Value Actual    : " + actFixedAssetsMenuIsDisplay +    " Value Expected : " + expFixedAssetsMenuIsDisplay);
	  		System.out.println("Production Menu Display Value Actual      : " + actProductionMenuIsDisplay +     " Value Expected : " + expProductionMenuIsDisplay);
	  		System.out.println("Point Of Sale Menu Display Value Actual   : " + actPointOfSaleMenuIsDisplay +    " Value Expected : " + expPointOfSaleMenuIsDisplay);
	  		System.out.println("Quality Control Menu Display Value Actual : " + actQualityControlMenuIsDisplay + " Value Expected : " + expQualityControlMenuIsDisplay);
	  		System.out.println("Settings Menu Display Value Actual        : " + actSettingsMenuIsDisplay +       " Value Expected : " + expSettingsMenuIsDisplay);
	  	
	  		boolean actMethod = actHomeMenuIsDisplay==actHomeMenuIsDisplay && actFinancialMenuIsDisplay==actFinancialMenuIsDisplay
	  				&& actInventoryMenuIsDisplay==actInventoryMenuIsDisplay && actFixedAssetsMenuIsDisplay==actFixedAssetsMenuIsDisplay
	  				&& actProductionMenuIsDisplay==actProductionMenuIsDisplay && actPointOfSaleMenuIsDisplay==actPointOfSaleMenuIsDisplay
	  				&& actQualityControlMenuIsDisplay==actQualityControlMenuIsDisplay && actSettingsMenuIsDisplay==actSettingsMenuIsDisplay;

	  		String actDisplayingResult =Boolean.toString(actMethod);
	  		String elementsIsDisplay = actDisplayingResult.toUpperCase();
	  		excelReader.setCellData(xlfile, xlSheetName, 7, 8, elementsIsDisplay);
	  		
	  		int actCount					         = menusList.size();
	  		ArrayList<String> actMenusArray          = new ArrayList<String>();
	  		
	  		for (int i = 0; i < actCount; i++) 
	  		{
	  			String data                          = menusList.get(i).getText();
	  			actMenusArray.add(data);
	  		}

	  		String actMenus 					     = actMenusArray.toString();
	  		String expMenus					         = excelReader.getCellData(xlSheetName, 8, 7);
	  		
	  		System.out.println("Menu Names Text Display Value Actual             : " + actMenus);
	  		System.out.println("Menus Names Text Display Value Expected          : " + expMenus);
	  		excelReader.setCellData(xlfile, xlSheetName, 8, 8, actMenus);
	  			
	  		Thread.sleep(4000);
	  		
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
	  		String actUserInfo                      = userNameDisplay.getText();
	  		String expUserInfo                      = excelReader.getCellData(xlSheetName, 9, 7);	
	  		excelReader.setCellData(xlfile, xlSheetName, 9, 8, actUserInfo);
	  		
	  		System.out.println("User Name Display Value Actual                   : " + actUserInfo +             " Value Expected : " + expUserInfo);
	  		
	  		
	  			
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	  		companyLogo.click();
	  		
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
	  		String actGetLoginCompanyNameInformation = companyName.getText();
	  		String actGetLoginCompanyName            = actGetLoginCompanyNameInformation.substring(0, 19);
	  		String expGetLoginCompanyName            = excelReader.getCellData(xlSheetName, 10, 7);
	  		excelReader.setCellData(xlfile, xlSheetName, 10, 8, actGetLoginCompanyName);
	  		companyLogo.click();
	  		
	  		System.out.println("Company Name Display Value Actual                : " + actGetLoginCompanyName +  " Value Expected : " + expGetLoginCompanyName);
	  		
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
	  		String actDashboardLabel                = labelDashboard.getText();
	  		String expDashboardLabel                = excelReader.getCellData(xlSheetName, 11, 7);	
	  		
	  		System.out.println("Dashboard Label Value Actual                     : " + actDashboardLabel +      " Value Expected : " + expDashboardLabel);
	  		
	  		excelReader.setCellData(xlfile, xlSheetName, 11, 8, actDashboardLabel);
	  		
	  		if (actUserInfo.equalsIgnoreCase(expUserInfo) && actGetLoginCompanyName.equalsIgnoreCase(expGetLoginCompanyName)
	  				&& actDashboardLabel.equalsIgnoreCase(expDashboardLabel)
	  				
	  				&& actDisplayingResult.equalsIgnoreCase("true")
	  				&& actMenus.equalsIgnoreCase(expMenus)) 
	  		{
	  			excelReader.setCellData(xlfile, xlSheetName, 7, 9, resPass);
	  			return true;
	  		} 
	  		else 
	  		{
	  			excelReader.setCellData(xlfile, xlSheetName, 7, 9, resFail);
	  			return false;
	  		}
	  	}
	  	catch (Exception e) 
	      {
	  		 String exception =e.getMessage().substring(0, 50);
	          excelReader.setExceptionInExcel(xlfile, xlSheetName, 7, 10, e.getMessage());
	          System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	          return false;
	      }
	   }
	  	
	 
	 @FindBy(xpath="//*[@id='dv_TranHomeHeading']/div[2]/div/div[1]/label")
	 private static WebElement setAsDefaultView;
	 
	 @FindBy(xpath="//a[@id='137']//span[contains(text(),'Transactions')]")
	 private static WebElement  invTransactionsMenu; 
	 
	 @FindBy(xpath="//*[@id='dv_TranHomeHeading']/div[2]/div/div")
	 private static List<WebElement> headingOptions;
	 
	 @FindBy(xpath="//*[@id='tblHeadTransRender']/tr/th")
	 private static List<WebElement> headerTxtList;
	  
	 @FindBy(xpath="//*[@id='cmbDataSet']/option")
	 private static WebElement datasetGetText;
	  
	 @FindBy(xpath="//*[@id='cmbViewType']/option[1]")
	 private static WebElement viewTypeGetText;
	  
	 @FindBy(xpath="//*[@id='cmbCheckStatus']/option[1]")
	 private static WebElement checkStatusGetText;
	  
	 @FindBy(xpath="//*[@id='Suspendstatus']/option[1]")
	 private static WebElement suspendStatusGetText;

	 @FindBy(xpath="//*[@id='cmbAuthorization']/option[1]")
	 private static WebElement authorizationGetText;
	  
	 @FindBy(xpath="//*[@id='cmbQC']/option[1]")
	 private static WebElement qcGetText;
	  
	  @FindBy(xpath="//*[@id='txtViewName']")
	  private static WebElement viewNameGetText;
	 
	  @FindBy(xpath="//a[@id='2033']//span[contains(text(),'Stocks')]")
	  private static WebElement  invTransStocksMenu;
	 
	 
	  public static boolean checkToOpenMaterialRequisitionVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
	  	excelReader=new ExcelReader(POJOUtility.getExcelPath());
	  	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	  	
	  	/*try
	  	{*/
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
	  		inventoryMenu.click();
	  		
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
	  		invTransactionsMenu.click();
	  		
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransStocksMenu));
	  		invTransStocksMenu.click();
	  	
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(materialRequisitionVoucher));
	  		materialRequisitionVoucher.click();
	  		
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			
			int count				                 = openingStocksNewToolBarButtons.size();
			
			ArrayList<String> actbuttons             = new ArrayList<String>();
			
			for (int i = 0; i < count; i++) 
			{
				String data                          = openingStocksNewToolBarButtons.get(i).getText();
				actbuttons.add(data);
			}
			    
			String actopeningStocksNewToolBarButtons = actbuttons.toString();
			String expopeningStocksNewToolBarButtons = "[New,   Edit  , Print, Delete, Authorize, Suspend, Reject, , Print barcode, Export to XML, Setting, Close]";
		
			System.out.println("actopeningStocksNewToolBarButtons  : " + actopeningStocksNewToolBarButtons);
			System.out.println("expopeningStocksNewToolBarButtons  : " + expopeningStocksNewToolBarButtons);
			
			excelReader.setCellData(xlfile, xlSheetName, 19, 8, actopeningStocksNewToolBarButtons);	
			
			String actsetAsDefaultView		=	setAsDefaultView.getText();
			String expsetAsDefaultView		=	excelReader.getCellData(xlSheetName, 20, 7);
			excelReader.setCellData(xlfile, xlSheetName, 20, 8, actsetAsDefaultView);
			
			
			//Ribbon Control Icons
	        int count1		= headingOptions.size();
			
			ArrayList<String> headingButtons = new ArrayList<String>();
			
			for (int i = 1; i < count1; i++) 
			{
				String data=headingOptions.get(i).getAttribute("title");
				
				headingButtons.add(data);
			}
			
			String actHeadingButtons = headingButtons.toString();
			String expHeadingButtons = excelReader.getCellData(xlSheetName, 21, 7);
			excelReader.setCellData(xlfile, xlSheetName, 21, 8, actHeadingButtons);
			
			
			System.out.println("actopeningStocksNewHeadingButtons  : " + actHeadingButtons);
			System.out.println("expopeningStocksNewHeadingButtons  : " + expHeadingButtons);
			
			//Home Page Column Headings
	        int headerTxtCount		= headerTxtList.size();
			
			ArrayList<String> headerTxt = new ArrayList<String>();
			
			for (int i = 2; i < headerTxtCount; i++) 
			{
				String data=headerTxtList.get(i).getText();
				headerTxt.add(data);
			}
			
			String actHeaderTxt = headerTxt.toString();
			String expHeaderTxt = "[Date, Voucher Number, Created by, Modified By, Created Date, Modified Date, Created time, Modified Time, Suspended, Authorization Status, Balance link value, Link status]";
			excelReader.setCellData(xlfile, xlSheetName, 22, 8, actHeaderTxt);
			
			
			System.out.println("actopeningStocksNewHomePageColumnsButtons  : " + actHeaderTxt);
			System.out.println("expopeningStocksNewHomePageColumnsButtons  : " + expHeaderTxt);
			
			if(actopeningStocksNewToolBarButtons.equalsIgnoreCase(expopeningStocksNewToolBarButtons)
					&& actHeadingButtons.equalsIgnoreCase(expHeadingButtons)
					&& actHeaderTxt.equalsIgnoreCase(expHeaderTxt))
			{
				System.out.println("Test Pass : HomeScreen Options Are Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 18, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : HomeScreen Options Are NOT Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 18, 9, resFail);
				return false;
			}
		}

	  	/*catch (Exception e) 
	      {
	  		String exception =e.getMessage().substring(0, 50);
	          excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	          System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	          return false;
	      }
	   }*/
	  	
	  public static void checkUserFriendlyMessage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		try
		{
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(errorMessage));
			String actErrorMessage=errorMessage.getText();
			
			System.out.println("Open Page then Message Display  :  " + actErrorMessage);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
			errorMessageCloseBtn.click();
		}
		catch(Exception e)
		{
			System.err.println("Error Message NOT Found or NOT Clickable");
			System.err.println(e.getMessage());
			
			String Exception=e.getMessage();
		}
	  }
	  
	  
	  @FindBy(xpath="//*[@id='id_transactionentry_toggle']/div")
	  private static List<WebElement> transEntryPageHeaderOptionsList;
		
	  
	  public boolean checkMaterialRequisitionVoucherNewEntryPageHeaderToggleOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
			
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			checkUserFriendlyMessage();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(freeFlowBtn));
			
			 int transEntryPageHeaderOptionsListCount=transEntryPageHeaderOptionsList.size();
			 
			 ArrayList<String> actEntryPageHeader          = new ArrayList<String>();
				
			 for (int i = 0; i < transEntryPageHeaderOptionsListCount; i++) 
		 	 {
				String data=transEntryPageHeaderOptionsList.get(i).getText();
				actEntryPageHeader.add(data);
				
		 	 }
	        
			String actEntryPageHeaderList=actEntryPageHeader.toString();
			String expEntryPageHeaderList=excelReader.getCellData("PurchaseVoucherN", 14, 7);
				
	        excelReader.setCellData(xlfile, "PurchaseVoucherN", 14, 8, actEntryPageHeaderList);
			 
			boolean actEntryPageJumpToHeaderSectionBtn= jumpToHeaderSectionBtn.isDisplayed();
			boolean actEntryPageJumpToBodySectionBtn  = jumpToBodySectionBtn.isDisplayed();
			boolean actEntryPageJumpToFooterBtn       = jumpToFooterBtn.isDisplayed();
			boolean actEntryPageNewBtn                = new_newBtn.isDisplayed();
			boolean actEntryPageSaveBtn               = saveBtn.isDisplayed();
			boolean actEntryPagePreviousBtn           = previousBtn.isDisplayed();
			boolean actEntryPageNextBtn               = nextBtn.isDisplayed();
			boolean actEntryPagePrintBtn              = new_PrintBtn.isDisplayed();
			boolean actEntryPageSuspendBtn            = new_SuspendBtn.isDisplayed();
			boolean actEntryPageCloseBtn              = new_CloseBtn.isDisplayed();
			boolean actEntryPageToggleBtn             = toggleBtn.isDisplayed();

			
			boolean expEntryPageJumpToHeaderSectionBtn= true;
			boolean expEntryPageJumpToBodySectionBtn  = true;
			boolean expEntryPageJumpToFooterBtn       = true;
			boolean expEntryPageNewBtn	              = true;
			boolean expEntryPageSaveBtn               = true;
			boolean expEntryPagePreviousBtn           = true;
			boolean expEntryPageNextBtn               = true;
			boolean expEntryPagePrintBtn              = true;
			boolean expEntryPageSuspendBtn            = true;
			boolean expEntryPageCloseBtn              = true;
			boolean expEntryPageToggleBtn             = true;

			System.out.println("Entry Page RC Jump to Header Value Actual:"+actEntryPageJumpToHeaderSectionBtn +"    Value Expected : "+expEntryPageJumpToHeaderSectionBtn);
			System.out.println("Entry Page RC Jump To Body Value Actual  :"+actEntryPageJumpToBodySectionBtn   +"    Value Expected : "+expEntryPageJumpToBodySectionBtn);
			System.out.println("Entry Page RC Jump To Footer Value Actual:"+actEntryPageJumpToFooterBtn        +"    Value Expected : "+expEntryPageJumpToFooterBtn);
			System.out.println("Entry Page RC New Btn Value Actual       :"+actEntryPageNewBtn                 +"    Value Expected : "+expEntryPageNewBtn);
			System.out.println("Entry Page RC Save Btn Value Actual      :"+actEntryPageSaveBtn                +"    Value Expected : "+expEntryPageSaveBtn);
			System.out.println("Entry Page RC Previous Btn Value Actual  :"+actEntryPagePreviousBtn            +"    Value Expected : "+expEntryPagePreviousBtn);
			System.out.println("Entry Page RC Next Btn Value Actual      :"+actEntryPageNextBtn                +"    Value Expected : "+expEntryPageNextBtn);
			System.out.println("Entry Page RC Print Btn Value Actual     :"+actEntryPagePrintBtn               +"    Value Expected : "+expEntryPagePrintBtn);
			System.out.println("Entry Page RC Suspend Btn Value Actual   :"+actEntryPageSuspendBtn             +"    Value Expected : "+expEntryPageSuspendBtn);
			System.out.println("Entry Page RC Close Btn Value Actual     :"+actEntryPageCloseBtn               +"    Value Expected : "+expEntryPageCloseBtn);
			System.out.println("Entry Page RC Toggle Btn Value Actual    :"+actEntryPageToggleBtn              +"    Value Expected : "+expEntryPageToggleBtn);
			
			
		    boolean actMethod=actEntryPageJumpToHeaderSectionBtn==expEntryPageJumpToHeaderSectionBtn && actEntryPageJumpToBodySectionBtn==expEntryPageJumpToBodySectionBtn 
					&& actEntryPageJumpToFooterBtn==expEntryPageJumpToFooterBtn && actEntryPageJumpToFooterBtn==expEntryPageJumpToFooterBtn
					&& actEntryPageNewBtn==expEntryPageNewBtn && actEntryPageSaveBtn==expEntryPageSaveBtn && actEntryPagePreviousBtn==expEntryPagePreviousBtn
					&& actEntryPageNextBtn==expEntryPageNextBtn && actEntryPagePrintBtn==expEntryPagePrintBtn &&actEntryPageSuspendBtn==expEntryPageSuspendBtn 
					&& actEntryPageCloseBtn==expEntryPageCloseBtn && actEntryPageToggleBtn==expEntryPageToggleBtn && 
					actEntryPageHeaderList.equalsIgnoreCase(expEntryPageHeaderList);
		    
		    
		    String actResult=Boolean.toString(actMethod);
		    String expResult=excelReader.getCellData("PurchaseVoucherN", 15, 7);
				
	        excelReader.setCellData(xlfile, "PurchaseVoucherN", 15, 8, actResult.toUpperCase());
			
			
			if(actMethod==true)		
			{
				System.out.println("Test Pass : entry Page Options Are Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 13, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : entry Page NOT Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 13, 9, resFail);
				return false;
			}
		 }
		 /*catch (Exception e) 
		 {
			  String exception =e.getMessage();
			  excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 13, 10, exception);
			  System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
			  return false;
		 }
	   }*/

	 
	  
	  
	  
	  @FindBy(xpath="//tbody[@id='id_body_23_table_data_body']/tr/td")
	  private static List<WebElement> itemListCount;
	  
	  @FindBy(xpath="//tbody[@id='id_header_3_table_data_body']/tr/td")
	  private static List<WebElement> purchaseAccountListCount; 
	  
	  @FindBy(xpath="//div[3]/div[2]/table/tbody/tr[2]/td/input[2]")
	  private static WebElement purchaseAccountEnter;
	  
	  
	  @FindBy(xpath = "//*[@id='idGlobalError']")
	  public static WebElement validationConfirmationMessage;
	  
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
				
				if(errorMessageCloseBtn.isDisplayed())
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
					errorMessageCloseBtn.click();
				}
				
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
	  
	  public boolean checkBudgetAlertOnMoreQtyInputValue() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/						
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
			purchaseAccountTxt.click();
			
			purchaseAccountTxt.sendKeys(Keys.SPACE);
			
			ArrayList<String>  actPVPurchaseAccount= new ArrayList<String>(); 
							
			int count=purchaseAccountListCount.size();
			
			System.err.println(count);
			
			for(int i=0 ; i < count ;i++)
			{
				String data=purchaseAccountListCount.get(i).getText();
				System.err.println("IN FOR LOOP : "+data);
				actPVPurchaseAccount.add(data);
				
				if(data.equalsIgnoreCase("STD RATE COGS ACC INV"))
				{
					purchaseAccountListCount.get(i).click();	
					break;
				}	
			}
			
            getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			Thread.sleep(2000);
			departmentTxt.sendKeys("DUBAI");
			Thread.sleep(2000);
			departmentTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemcount=itemListCount.size();
			
			System.err.println(itemcount);
				
			for(int i=0 ; i < itemcount ;i++)
			{
				String data=itemListCount.get(i).getText();
				
				if(data.equalsIgnoreCase("STD RATE COGS ITEM"))
				{
					itemListCount.get(i).click();
					break;
				}
			}
				
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
			select1stRow_3rdColumn.click();
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.click();
			enter_Quantity.clear();
			enter_Quantity.sendKeys("5");
			enter_Quantity.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.click();
			enter_Rate.clear();
			enter_Rate.sendKeys("20");
			enter_Rate.sendKeys(Keys.TAB);			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.click();
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			
			String actR1C1Item =select1stRow_1stColumn.getText();
			String actR1C2Units=select1stRow_2ndColumn.getText();
			String actR1C5Qty  =select1stRow_3rdColumn.getText();
			String actR1C7Rate =select1stRow_4thColumn.getText();
			String actR1C8Gross=select1stRow_5thColumn.getText();

			String expR1C1Item ="STD RATE COGS ITEM";
			String expR1C2Units="Dozs";
			String expR1C5Qty  ="5.00";
			String expR1C7Rate ="20.00"; //20
			String expR1C8Gross="100.00";
			
			System.out.println("*********************************************** checkEnteringDatainFirstRow  ********************************************");
			System.out.println("Entry Page Save with Vendor Item Value Actual   : "+actR1C1Item   +" Value Expected : "+expR1C1Item);	
			System.out.println("Entry Page Save with Vendor Units Value Actual  : "+actR1C2Units  +" Value Expected : "+expR1C2Units);
			System.out.println("Entry Page Save with Vendor QTY Value Actual    : "+actR1C5Qty    +" Value Expected : "+expR1C5Qty);
			System.out.println("Entry Page Save with Vendor RATE Value Actual   : "+actR1C7Rate   +" Value Expected : "+expR1C7Rate);
			System.out.println("Entry Page Save with Vendor Gross Value Actual  : "+actR1C8Gross  +" Value Expected : "+expR1C8Gross);			
					
			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();*/
		
			String expSaveValidationMessageOnSaveWithPurchaseInptData="This row is making budget negative.";
			
			String actMessage=checkValidationMessage(expSaveValidationMessageOnSaveWithPurchaseInptData);
			
			String actGetSaveMessage = new String(actMessage);
						
			String actSaveMessage = actGetSaveMessage.substring(0,35);
      
			System.out.println("actSaveMessage"+actSaveMessage);
			
			System.out.println("Entry Page Save with Purchase Input Value Actual  : "+actMessage  +" Value Expected : "+expSaveValidationMessageOnSaveWithPurchaseInptData);			
				
			if(actR1C1Item.equalsIgnoreCase(expR1C1Item) && actR1C2Units.equalsIgnoreCase(expR1C2Units) && actR1C5Qty.equalsIgnoreCase(expR1C5Qty) 
				&& actR1C7Rate.equalsIgnoreCase(expR1C7Rate) && actR1C8Gross.equalsIgnoreCase(expR1C8Gross)
				
					&& actSaveMessage.equalsIgnoreCase(expSaveValidationMessageOnSaveWithPurchaseInptData))
			{
				System.out.println("Test Pass : Data Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resPass);

				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
	    {
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 99, 10, exception);
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	   	 	return false;
	    }	
	  }*/
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  
	  public boolean checkBudgetAlertOnLessQtyInputValue() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/						
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		    select1stRow_3rdColumn.click();
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.click();
			enter_Quantity.clear();
			enter_Quantity.sendKeys("2");
			enter_Quantity.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.click();
			enter_Rate.clear();
			enter_Rate.sendKeys("20");
			enter_Rate.sendKeys(Keys.TAB);			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			
			String actR1C1Item =select1stRow_1stColumn.getText();
			String actR1C2Units=select1stRow_2ndColumn.getText();
			String actR1C5Qty  =select1stRow_3rdColumn.getText();
			String actR1C7Rate =select1stRow_4thColumn.getText();
			String actR1C8Gross=select1stRow_5thColumn.getText();

			String expR1C1Item ="STD RATE COGS ITEM";
			String expR1C2Units="Dozs";
			String expR1C5Qty  ="2.00";
			String expR1C7Rate ="20.00"; //20
			String expR1C8Gross="40.00";
			
			System.out.println("*********************************************** checkEnteringDatainFirstRow  ********************************************");
			System.out.println("Entry Page Save with Vendor Item Value Actual   : "+actR1C1Item   +" Value Expected : "+expR1C1Item);	
			System.out.println("Entry Page Save with Vendor Units Value Actual  : "+actR1C2Units  +" Value Expected : "+expR1C2Units);
			System.out.println("Entry Page Save with Vendor QTY Value Actual    : "+actR1C5Qty    +" Value Expected : "+expR1C5Qty);
			System.out.println("Entry Page Save with Vendor RATE Value Actual   : "+actR1C7Rate   +" Value Expected : "+expR1C7Rate);
			System.out.println("Entry Page Save with Vendor Gross Value Actual  : "+actR1C8Gross  +" Value Expected : "+expR1C8Gross);			
					
		
		   if(actR1C1Item.equalsIgnoreCase(expR1C1Item) && actR1C2Units.equalsIgnoreCase(expR1C2Units) && actR1C5Qty.equalsIgnoreCase(expR1C5Qty) 
				&& actR1C7Rate.equalsIgnoreCase(expR1C7Rate) && actR1C8Gross.equalsIgnoreCase(expR1C8Gross))
			{
				System.out.println("Test Pass : Data Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resPass);

				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resFail);
				return false;
			}	
		}
		/*catch (Exception e) 
	    {
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 99, 10, exception);
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	   	 	return false;
	    }	
	  }*/
	  
	  
	  
	  
	  public boolean checkSaveMaterialRequisitionVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/							
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
	
			String expSaveValidationMessageOnSaveWithPurchaseInptData="Voucher Saved Successfully";
			
			String actMessage=checkValidationMessage(expSaveValidationMessageOnSaveWithPurchaseInptData);
			
			String actGetSaveMessage = new String(actMessage);
			
			String actSaveMessage = actGetSaveMessage.substring(0,26);
      
			System.out.println("actSaveMessage"+actSaveMessage);
			
			System.out.println("Entry Page Save with Purchase Input Value Actual  : "+actMessage  +" Value Expected : "+expSaveValidationMessageOnSaveWithPurchaseInptData);			
		
			
			if( actSaveMessage.equalsIgnoreCase(expSaveValidationMessageOnSaveWithPurchaseInptData))
			{
				System.out.println("Test Pass : Data Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resPass);

				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
	    {
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 99, 10, exception);
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	   	 	return false;
	    }	
	  }*/
	  
	  
	  @FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr[1]/td[12]")
	  private static WebElement openingStocksNewHomeRow1AuthorizationStatus;
	  
	  /*@FindBy(xpath="//*[@id='tdSort_1'_col_'AuthStatus']")
	  private static WebElement openingStocksNewHomeRow1AuthorizationStatus;*/
	  
	
	  
	  public boolean checkMaterailRequsitionHomePageStatus() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		/*try
		{*/
		   
		    getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(new_CloseBtn));
		    new_CloseBtn.click();
		    
		    Thread.sleep(5000);
		    
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksNewHomeRow1AuthorizationStatus));
		    
			String actStatus=openingStocksNewHomeRow1AuthorizationStatus.getText();
			String expStatus="Pending";
			
			System.out.println("************************************* checkVoucherSuspendAuthorizeStatusInAuthorizeUser  *********************************");
			System.out.println("Suspend Status In home Page 	 :  "+actStatus 						+" Value Expected : "+expStatus);
					
			if (actStatus.equalsIgnoreCase(expStatus))
			{
				System.out.println("Test Pass : Entry Page Close Button");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 240, 9, resPass);
				return true;
			}
			
			else
			{
				System.out.println("Test Fail :  Entry Page Close Button");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 240, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
		{
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 240, 10, exception);
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	   	 	return false;
		}
	  }*/
	  
	  

	  @FindBy(xpath="//div[@id='dvHomeTransClose']//span[@class='icon-close icon-font6']")  
	  private static WebElement  voucherhomeCloseBtn;
	  
	  public boolean checkMaterialRequsitionVoucherCloseButtonInHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	  {	
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
			String actUserInfo = userNameDisplay.getText();
			String expUserInfo = excelReader.getCellData(xlSheetName, 453, 7);
			
			excelReader.setCellData(xlfile, xlSheetName, 453, 8, actUserInfo);
			
			System.out.println("User Name Display Value Actual                   : " + actUserInfo +             " Value Expected : " + expUserInfo);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
			companyLogo.click();
			
	        //Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
			String actGetLoginCompanyNameInformation = companyName.getText();
			String actGetLoginCompanyName            = actGetLoginCompanyNameInformation.substring(0, 19);
			String expGetLoginCompanyName            = excelReader.getCellData(xlSheetName, 454, 7);
			companyLogo.click();
			
			excelReader.setCellData(xlfile, xlSheetName, 454, 8, actGetLoginCompanyName);
			
			System.out.println("Company Name Display Value Actual                : " + actGetLoginCompanyName +  " Value Expected : " + expGetLoginCompanyName);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
			
			boolean actlabelDashboard=labelDashboard.isDisplayed();
			boolean actselectDashboard=selectDashboard.isDisplayed();
			boolean actnewAddDashBoard=newAddDashBoard.isDisplayed();
			boolean actdashboardCustomizationSettings=dashboardCustomizationSettings.isDisplayed();
			
			boolean explabelDashboard=true;
			boolean expselectDashboard=true;
			boolean expnewAddDashBoard=true;
			boolean expdashboardCustomizationSettings=true;
			
			System.out.println("labelDashboard Value Actual        : " + actlabelDashboard  +               " Value Expected :  " + explabelDashboard);
			System.out.println("selectDashboard Value Actual       : " + actselectDashboard +               " Value Expected :  " + expselectDashboard);
			System.out.println("newAddDashBoard Value Actual       : " + actnewAddDashBoard +               " Value Expected :  " + expnewAddDashBoard);
			System.out.println("CustomizationSettings Value Actual : " + actdashboardCustomizationSettings+ " Value Expected :  " + expdashboardCustomizationSettings);
			
			boolean actMethod = actlabelDashboard==explabelDashboard && actselectDashboard==expselectDashboard && actnewAddDashBoard==expnewAddDashBoard
								&& actdashboardCustomizationSettings==expdashboardCustomizationSettings;
			
			String actResult = Boolean.toString(actMethod);
			
			excelReader.setCellData(xlfile, xlSheetName, 452, 8, actResult.toUpperCase());
			
			if(actUserInfo.equalsIgnoreCase(expUserInfo) && actGetLoginCompanyName.equalsIgnoreCase(expGetLoginCompanyName) && actMethod==true)
			{
				System.out.println("test pass : Document History Pop Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 452, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("test Fail : Document History Pop NOT Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 452, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
		{
			excelReader.setExceptionInExcel(xlfile, xlSheetName, 452, 10, e.getMessage());
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
			return false;
		}
	  }*/
		
	  
	  
	  // Open Purchase Orders
	  
	  //First Checking Linked Voucher, to display pending Links without Authorization
	  
	  @FindBy(xpath="//*[@id='id_transaction_viewcontainer']/li[1]/a")
	  private static WebElement voucherAll;
	  
	  @FindBy(xpath="//*[@id='id_transaction_viewcontainer']/li[2]/a")
	  private static WebElement pendingAuthorizationProcess;
	  
	  @FindBy(xpath="//*[@id='id_transaction_viewcontainer']/li[3]/a")
	  private static WebElement pendingMaterialRequistion;
	  
	  
	  public static boolean checkOpenPurchasesOrderVoucherToDisplayLinksWithoutAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
	  	excelReader=new ExcelReader(POJOUtility.getExcelPath());
	  	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	  	
	  	/*try
	  	{*/
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
	  		inventoryMenu.click();
	  		
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
	  		invTransactionsMenu.click();
	  		
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
	  		inventoryTransactionsPurchasesMenu.click();
	  	
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
	  		purchasesOrdersVoucher.click();
	  		
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			
			int count				                 = openingStocksNewToolBarButtons.size();
			
			ArrayList<String> actbuttons             = new ArrayList<String>();
			
			for (int i = 0; i < count; i++) 
			{
				String data                          = openingStocksNewToolBarButtons.get(i).getText();
				actbuttons.add(data);
			}
			    
			String actopeningStocksNewToolBarButtons = actbuttons.toString();
			String expopeningStocksNewToolBarButtons = "[New,   Edit  , Print, Delete, Authorize, Suspend, Reject, , Print barcode, Export to XML, Setting, Close]";
		
			System.out.println("actopeningStocksNewToolBarButtons  : " + actopeningStocksNewToolBarButtons);
			System.out.println("expopeningStocksNewToolBarButtons  : " + expopeningStocksNewToolBarButtons);
			
			excelReader.setCellData(xlfile, xlSheetName, 19, 8, actopeningStocksNewToolBarButtons);	
			
			String actsetAsDefaultView		=	setAsDefaultView.getText();
			String expsetAsDefaultView		=	excelReader.getCellData(xlSheetName, 20, 7);
			excelReader.setCellData(xlfile, xlSheetName, 20, 8, actsetAsDefaultView);
			
			
			//Ribbon Control Icons
	        int count1		= headingOptions.size();
			
			ArrayList<String> headingButtons = new ArrayList<String>();
			
			for (int i = 1; i < count1; i++) 
			{
				String data=headingOptions.get(i).getAttribute("title");
				
				headingButtons.add(data);
			}
			
			String actHeadingButtons = headingButtons.toString();
			String expHeadingButtons = excelReader.getCellData(xlSheetName, 21, 7);
			excelReader.setCellData(xlfile, xlSheetName, 21, 8, actHeadingButtons);
			
			
			System.out.println("actopeningStocksNewHeadingButtons  : " + actHeadingButtons);
			System.out.println("expopeningStocksNewHeadingButtons  : " + expHeadingButtons);
			
			//Home Page Column Headings
	        int headerTxtCount		= headerTxtList.size();
			
			ArrayList<String> headerTxt = new ArrayList<String>();
			
			for (int i = 2; i < headerTxtCount; i++) 
			{
				String data=headerTxtList.get(i).getText();
				headerTxt.add(data);
			}
			
			String actHeaderTxt = headerTxt.toString();
			String expHeaderTxt = excelReader.getCellData(xlSheetName, 22, 7);
			excelReader.setCellData(xlfile, xlSheetName, 22, 8, actHeaderTxt);
			
			
			System.out.println("actopeningStocksNewHomePageColumnsButtons  : " + actHeaderTxt);
			System.out.println("expopeningStocksNewHomePageColumnsButtons  : " + expHeaderTxt);
			
			boolean actVoucherAllAuthorizationAndLinks =voucherAll.isDisplayed() && pendingAuthorizationProcess.isDisplayed()
					                                      && pendingMaterialRequistion.isDisplayed();
			
			boolean expVoucherAllAuthorizationAndLinks = true;
			
			System.out.println("Home page All Pending Authorization And Links Value Actual : " + actVoucherAllAuthorizationAndLinks  +   " Value Expected :  " + expVoucherAllAuthorizationAndLinks);
			
			if(actopeningStocksNewToolBarButtons.equalsIgnoreCase(expopeningStocksNewToolBarButtons)
					&& actHeadingButtons.equalsIgnoreCase(expHeadingButtons)
					&& actHeaderTxt.equalsIgnoreCase(expHeaderTxt)
					
					&& actVoucherAllAuthorizationAndLinks==expVoucherAllAuthorizationAndLinks)
			{
				System.out.println("Test Pass : HomeScreen Options Are Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 18, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : HomeScreen Options Are NOT Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 18, 9, resFail);
				return false;
			}
		}
	  	/*catch (Exception e) 
	      {
	  		  String exception =e.getMessage().substring(0, 50);
	          excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	          System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	          return false;
	      }
	   }*/
	  	
	 
		@FindBy(xpath="//*[@id='tblBodyTransRender']")
		private static WebElement  voucherBodyGrid;
		  
	  public boolean checkPurchasesOrderVoucherPendingLinksOptionsInHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {	
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pendingMaterialRequistion));
			pendingMaterialRequistion.click();
			
			boolean actPendingLinksList = voucherBodyGrid.getText().isEmpty();
			boolean expPendingLinksList = true;
			
			System.out.println("Home page click on Pending Links Value Actual : " + actPendingLinksList  +   " Value Expected :  " + expPendingLinksList);
								
			if(actPendingLinksList==expPendingLinksList)		
			{
				System.out.println("Test Pass : entry Page Options Are Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 13, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : entry Page NOT Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 13, 9, resFail);
				return false;
			}
		 }
		 /*catch (Exception e) 
		 {
			  String exception =e.getMessage();
			  excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 13, 10, exception);
			  System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
			  return false;
		 }
	   }*/

	  
	  public boolean checkPurchaseOrderVoucherNewEntryPageHeaderToggleOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			checkUserFriendlyMessage();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(freeFlowBtn));
			
			 int transEntryPageHeaderOptionsListCount=transEntryPageHeaderOptionsList.size();
			 
			 ArrayList<String> actEntryPageHeader          = new ArrayList<String>();
				
			 for (int i = 0; i < transEntryPageHeaderOptionsListCount; i++) 
		 	 {
				String data=transEntryPageHeaderOptionsList.get(i).getText();
				actEntryPageHeader.add(data);
				
		 	 }
	        
			String actEntryPageHeaderList=actEntryPageHeader.toString();
			String expEntryPageHeaderList=excelReader.getCellData("PurchaseVoucherN", 14, 7);
				
	        excelReader.setCellData(xlfile, "PurchaseVoucherN", 14, 8, actEntryPageHeaderList);
			 
			boolean actEntryPageJumpToHeaderSectionBtn= jumpToHeaderSectionBtn.isDisplayed();
			boolean actEntryPageJumpToBodySectionBtn  = jumpToBodySectionBtn.isDisplayed();
			boolean actEntryPageJumpToFooterBtn       = jumpToFooterBtn.isDisplayed();
			boolean actEntryPageNewBtn                = new_newBtn.isDisplayed();
			boolean actEntryPageSaveBtn               = saveBtn.isDisplayed();
			boolean actEntryPagePreviousBtn           = previousBtn.isDisplayed();
			boolean actEntryPageNextBtn               = nextBtn.isDisplayed();
			boolean actEntryPagePrintBtn              = new_PrintBtn.isDisplayed();
			boolean actEntryPageSuspendBtn            = new_SuspendBtn.isDisplayed();
			boolean actEntryPageCloseBtn              = new_CloseBtn.isDisplayed();
			boolean actEntryPageToggleBtn             = toggleBtn.isDisplayed();

			
			boolean expEntryPageJumpToHeaderSectionBtn= true;
			boolean expEntryPageJumpToBodySectionBtn  = true;
			boolean expEntryPageJumpToFooterBtn       = true;
			boolean expEntryPageNewBtn	              = true;
			boolean expEntryPageSaveBtn               = true;
			boolean expEntryPagePreviousBtn           = true;
			boolean expEntryPageNextBtn               = true;
			boolean expEntryPagePrintBtn              = true;
			boolean expEntryPageSuspendBtn            = true;
			boolean expEntryPageCloseBtn              = true;
			boolean expEntryPageToggleBtn             = true;

			System.out.println("Entry Page RC Jump to Header Value Actual:"+actEntryPageJumpToHeaderSectionBtn +"    Value Expected : "+expEntryPageJumpToHeaderSectionBtn);
			System.out.println("Entry Page RC Jump To Body Value Actual  :"+actEntryPageJumpToBodySectionBtn   +"    Value Expected : "+expEntryPageJumpToBodySectionBtn);
			System.out.println("Entry Page RC Jump To Footer Value Actual:"+actEntryPageJumpToFooterBtn        +"    Value Expected : "+expEntryPageJumpToFooterBtn);
			System.out.println("Entry Page RC New Btn Value Actual       :"+actEntryPageNewBtn                 +"    Value Expected : "+expEntryPageNewBtn);
			System.out.println("Entry Page RC Save Btn Value Actual      :"+actEntryPageSaveBtn                +"    Value Expected : "+expEntryPageSaveBtn);
			System.out.println("Entry Page RC Previous Btn Value Actual  :"+actEntryPagePreviousBtn            +"    Value Expected : "+expEntryPagePreviousBtn);
			System.out.println("Entry Page RC Next Btn Value Actual      :"+actEntryPageNextBtn                +"    Value Expected : "+expEntryPageNextBtn);
			System.out.println("Entry Page RC Print Btn Value Actual     :"+actEntryPagePrintBtn               +"    Value Expected : "+expEntryPagePrintBtn);
			System.out.println("Entry Page RC Suspend Btn Value Actual   :"+actEntryPageSuspendBtn             +"    Value Expected : "+expEntryPageSuspendBtn);
			System.out.println("Entry Page RC Close Btn Value Actual     :"+actEntryPageCloseBtn               +"    Value Expected : "+expEntryPageCloseBtn);
			System.out.println("Entry Page RC Toggle Btn Value Actual    :"+actEntryPageToggleBtn              +"    Value Expected : "+expEntryPageToggleBtn);
			
			
		    boolean actMethod=actEntryPageJumpToHeaderSectionBtn==expEntryPageJumpToHeaderSectionBtn && actEntryPageJumpToBodySectionBtn==expEntryPageJumpToBodySectionBtn 
					&& actEntryPageJumpToFooterBtn==expEntryPageJumpToFooterBtn && actEntryPageJumpToFooterBtn==expEntryPageJumpToFooterBtn
					&& actEntryPageNewBtn==expEntryPageNewBtn && actEntryPageSaveBtn==expEntryPageSaveBtn && actEntryPagePreviousBtn==expEntryPagePreviousBtn
					&& actEntryPageNextBtn==expEntryPageNextBtn && actEntryPagePrintBtn==expEntryPagePrintBtn &&actEntryPageSuspendBtn==expEntryPageSuspendBtn 
					&& actEntryPageCloseBtn==expEntryPageCloseBtn && actEntryPageToggleBtn==expEntryPageToggleBtn && 
					actEntryPageHeaderList.equalsIgnoreCase(expEntryPageHeaderList);
		    
		    
		    String actResult=Boolean.toString(actMethod);
		    String expResult=excelReader.getCellData("PurchaseVoucherN", 15, 7);
				
	        excelReader.setCellData(xlfile, "PurchaseVoucherN", 15, 8, actResult.toUpperCase());
			
			
			if(actMethod==true)		
			{
				System.out.println("Test Pass : entry Page Options Are Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 13, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : entry Page NOT Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 13, 9, resFail);
				return false;
			}
		 }
		 /*catch (Exception e) 
		 {
			  String exception =e.getMessage();
			  excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 13, 10, exception);
			  System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
			  return false;
		 }
	   }*/
	  
	  
	  
	
	  @FindBy(xpath="//*[@id='id_transaction_entry_workflow_togglebutton']")
	  private static WebElement plusImageOfLinksBtn;
		
	  public boolean checkPurchasesOrderVoucherPendingLinksOptionsInEntryPage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {	
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(plusImageOfLinksBtn));
			plusImageOfLinksBtn.click();
			
			Thread.sleep(2998);
			
			boolean actPendingLinksList = plusImageOfLinksBtn.isDisplayed();
			boolean expPendingLinksList = true;
			
			System.out.println("Entry page click on Pending Links Value Actual : " + actPendingLinksList  +   " Value Expected :  " + expPendingLinksList);
								
			if(actPendingLinksList==expPendingLinksList)		
			{
				System.out.println("Test Pass : entry Page Options Are Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 13, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : entry Page NOT Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 13, 9, resFail);
				return false;
			}
		 }
		 /*catch (Exception e) 
		 {
			  String exception =e.getMessage();
			  excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 13, 10, exception);
			  System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
			  return false;
		 }
	   }*/
	  
	  
	  

	  public boolean checkPurchaseOrderVoucherCloseButtonInEntryAndHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	  {	
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(new_CloseBtn));
		    new_CloseBtn.click();
		    
		    Thread.sleep(2000);
			
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
			String actUserInfo = userNameDisplay.getText();
			String expUserInfo = excelReader.getCellData(xlSheetName, 453, 7);
			
			excelReader.setCellData(xlfile, xlSheetName, 453, 8, actUserInfo);
			
			System.out.println("User Name Display Value Actual                   : " + actUserInfo +             " Value Expected : " + expUserInfo);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
			companyLogo.click();
			
	        //Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
			String actGetLoginCompanyNameInformation = companyName.getText();
			String actGetLoginCompanyName            = actGetLoginCompanyNameInformation.substring(0, 19);
			String expGetLoginCompanyName            = excelReader.getCellData(xlSheetName, 454, 7);
			companyLogo.click();
			
			excelReader.setCellData(xlfile, xlSheetName, 454, 8, actGetLoginCompanyName);
			
			System.out.println("Company Name Display Value Actual                : " + actGetLoginCompanyName +  " Value Expected : " + expGetLoginCompanyName);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
			
			boolean actlabelDashboard=labelDashboard.isDisplayed();
			boolean actselectDashboard=selectDashboard.isDisplayed();
			boolean actnewAddDashBoard=newAddDashBoard.isDisplayed();
			boolean actdashboardCustomizationSettings=dashboardCustomizationSettings.isDisplayed();
			
			boolean explabelDashboard=true;
			boolean expselectDashboard=true;
			boolean expnewAddDashBoard=true;
			boolean expdashboardCustomizationSettings=true;
			
			System.out.println("labelDashboard Value Actual        : " + actlabelDashboard  +               " Value Expected :  " + explabelDashboard);
			System.out.println("selectDashboard Value Actual       : " + actselectDashboard +               " Value Expected :  " + expselectDashboard);
			System.out.println("newAddDashBoard Value Actual       : " + actnewAddDashBoard +               " Value Expected :  " + expnewAddDashBoard);
			System.out.println("CustomizationSettings Value Actual : " + actdashboardCustomizationSettings+ " Value Expected :  " + expdashboardCustomizationSettings);
			
			boolean actMethod = actlabelDashboard==explabelDashboard && actselectDashboard==expselectDashboard && actnewAddDashBoard==expnewAddDashBoard
								&& actdashboardCustomizationSettings==expdashboardCustomizationSettings;
			
			String actResult = Boolean.toString(actMethod);
			
			excelReader.setCellData(xlfile, xlSheetName, 452, 8, actResult.toUpperCase());
			
			if(actUserInfo.equalsIgnoreCase(expUserInfo) && actGetLoginCompanyName.equalsIgnoreCase(expGetLoginCompanyName) && actMethod==true)
			{
				System.out.println("test pass : Document History Pop Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 452, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("test Fail : Document History Pop NOT Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 452, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
		{
			excelReader.setExceptionInExcel(xlfile, xlSheetName, 452, 10, e.getMessage());
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
			return false;
		}
	  }*/
		
	  
	  
	  
	  
	  // Logout from su and Login From authorize user
	  
	  @FindBy(xpath="//*[@id='userprofile']/li/span[2]")
		private static WebElement logoutOption;

		/*@FindBy(xpath="//*[@id='txtUsername']")
		private static WebElement username;*/

		@FindBy(id="txtPassword")
		private static WebElement password;

		@FindBy(id="ddlCompany")
		private static WebElement companyDropDownList;

		@FindBy(id="btnSignin")
		private static WebElement signIn;
	  
		public boolean checkAlertsOfTransactionAuthorizationLogoutAndLoginCompany() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		    /*try
			{*/
		    	    getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
				    userNameDisplay.click();
				
				    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
				    logoutOption.click();
				 
					boolean verifyFiledsStatus=true;
					
					getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(username));
								
					boolean actusername              = username.isDisplayed();
					boolean actpassword              = password.isDisplayed();
					boolean actcompanyDropDownList   = companyDropDownList.isDisplayed();
					boolean actsignIn                = signIn.isDisplayed();
					
					// Expected
					boolean expusername              = true;
					boolean exppassword              = true;
					boolean expcompanyDropDownList   = true;
					boolean expsignIn                = true;
					
					System.out.println("************************************* checkSignOut  *********************************");
					System.out.println("username             : "+actusername             +" Value Expected : "+expusername);
					System.out.println("password             : "+actpassword             +" Value Expected : "+exppassword);
					System.out.println("companyDropDownList  : "+actcompanyDropDownList  +" Value Expected : "+expcompanyDropDownList);
					System.out.println("signIn               : "+actsignIn               +" Value Expected : "+expsignIn);

			
				
				  getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(username));
				  username.click();
				 
				  username.sendKeys(excelReader.getCellData("PurchaseVoucherN", 252, 6));
				  getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
				 
				  getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(password));
				  password.sendKeys(excelReader.getCellData("PurchaseVoucherN", 253, 6));
				 
				  getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
				  
				  getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(signIn));
				  signIn.click();
			 

			
				Thread.sleep(5000);
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
				String actUserInfo=userNameDisplay.getText();
				
				System.out.println("User Info : "+actUserInfo);
				
				System.out.println("User Info Capture Text :"+userNameDisplay.getText());
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
				
				companyLogo.click();
				
				String actCompanyTxt=companyName.getText();
				String actLoginCompanyName=actCompanyTxt.substring(0, 19);
				System.out.println("company name :"+ actLoginCompanyName);
				companyLogo.click();
							    	
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
				String getDashboard=labelDashboard.getText();
					
				boolean getnoOfAlertsDisplay=noOfAlertsDisplay.isDisplayed();
						
				// Expected
				String expuserInfo="UserAllOptionsST";
				String expLoginCompanyName="Automation Company ";
				String expDashboard="Dashboard";
				boolean expnoOfAlertsDisplay=true;
				
				System.out.println("************************************* checkAlertsToSignIAsAuthorizeUser  *********************************");
				System.out.println("userNameDisplay    : "+actUserInfo              +" Value Expected : "+expuserInfo);
				System.out.println("LoginCompanyName   : "+actLoginCompanyName      +" Value Expected : "+expLoginCompanyName);
				System.out.println("labelDashboard     : "+getDashboard             +" Value Expected : "+expDashboard);
				System.out.println("noOfAlertsDisplay  : "+getnoOfAlertsDisplay     +" Value Expected : "+expnoOfAlertsDisplay);
						
				boolean actMethod=actUserInfo.equalsIgnoreCase(expuserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)
						&& getDashboard.equalsIgnoreCase(expDashboard) && getnoOfAlertsDisplay==expnoOfAlertsDisplay;
				
				String actResult=Boolean.toString(actMethod);
				String expResult=excelReader.getCellData("PurchaseVoucherN", 254, 7);
				
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 254, 8, actResult);
				
				
				if(actResult.equalsIgnoreCase(expResult))
				{
					System.out.println("Test Pass : Transaction Authorization is Displayed");
					excelReader.setCellData(xlfile, "PurchaseVoucherN", 251, 9, resPass);
					return true;
				}
				else
				{
					System.out.println("Test Fail : Transaction Authorization is Displayed");
					excelReader.setCellData(xlfile, "PurchaseVoucherN", 251, 9, resFail);
					return false;
				}
			}
			/*catch (Exception e) 
			{
				String exception =e.getMessage();
				excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 251, 10, exception);
				System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
				return false;
			}
		  }*/
		 
		 
		 
		
	@FindBy(xpath="//span[@id='noofalerts']")
	private static WebElement noOfAlertsDisplay;

	@FindBy(xpath="//li[@class='header']")
	  private static WebElement alertsMainHeader;

	  @FindBy(xpath="//li[@class='treeview-menu']//a[@id='1']")
	  private static WebElement alertTransactionsTreeView;

	  @FindBy(xpath="//*[@id='1']/span")
	  private static WebElement alertTransaction;

	  @FindBy(xpath="//*[@id='7936']/span")
	  private static WebElement transactionAuthorization;
	  
	  @FindBy(xpath="//*[@id='masterChild_7936']/a/span")
	  private static WebElement transactionAuthorizationText;
	  
	  @FindBy(xpath="//*[@id='id_PlanName 1']/label")
	  private static WebElement pendingBudgetInfo;

	  @FindBy(xpath="//*[@id='Grid_ApproveBudgetMore_control_heading_ctrl_1']")
	  private static WebElement checkboxInBudgetAuthorization;
	  
	  @FindBy(xpath="//*[@id='btnAuthorize']")
	  private static WebElement btnAuthorizeInBudgetAuthorization;
	  
	  @FindBy(xpath="//*[@id='btnClose']")
	  private static WebElement btnCloseInBudgetAuthorization;
	
	  public boolean checkClickOnAlertsInWrapperLayoutAuthorizeUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
	   excelReader=new ExcelReader(POJOUtility.getExcelPath());
	   xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
	   /*try
	   {*/
			
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(noOfAlertsDisplay));
		noOfAlertsDisplay.click();
		
		Thread.sleep(2000);
		
		// Actual
		String  actAlertsMainHeader=alertsMainHeader.getText();
		excelReader.setCellData(xlfile, "PurchaseVoucherN", 256, 8, actAlertsMainHeader);
		
		boolean actNoOfAlertsDisplay=alertsMainHeader.isDisplayed();
		boolean expNoOfAlertsDisplay=true;
		
		System.out.println("Alerts is display"+ actAlertsMainHeader);
		
		String actAlertTransactionsTreeView=alertTransaction.getText();
		String expAlertTransactionsTreeView="  Transactions (1)";

		alertTransaction.click();
		
		String actBudgetAuthorization = transactionAuthorization.getText();
		String expBudgetAuthorization = " Material Requisition(s) Require Approval (1)";
		
		transactionAuthorization.click();
		
		String actBudgetText =transactionAuthorizationText.getText();
		String expBudgetText = "Voucher needs your authorization. Doc No: MatReq:1";
		
		transactionAuthorizationText.click();
		
		Thread.sleep(2000);
		
		System.out.println("************************************* checkClickOnAlertsInAuthorizeUser  *********************************");
		System.out.println("noOfAlertsDisplay         : " + actNoOfAlertsDisplay         + " Value Expected  : " + expNoOfAlertsDisplay);
		System.out.println("noOfAlertsDisplay         : " + actAlertTransactionsTreeView + " Value Expected  : " + expAlertTransactionsTreeView);
		System.out.println("noOfAlertsDisplay         : " + actBudgetAuthorization       + " Value Expected  : " + expBudgetAuthorization);
		System.out.println("noOfAlertsDisplay         : " + actBudgetText                + " Value Expected  : " + expBudgetText);
		
		if(actNoOfAlertsDisplay==expNoOfAlertsDisplay && actAlertTransactionsTreeView.equalsIgnoreCase(expAlertTransactionsTreeView)
				&& actBudgetAuthorization.equalsIgnoreCase(expBudgetAuthorization) && actBudgetText.equalsIgnoreCase(expBudgetText))
		{
			System.out.println("Test Pass : Transaction Authorization is Displayed");
			excelReader.setCellData(xlfile, "PurchaseVoucherN", 255, 9, resPass);
			return true;
			
		}
		else
		{
			System.out.println("Test Fail : Transaction Authorization is Displayed");
			excelReader.setCellData(xlfile, "PurchaseVoucherN", 255, 9, resFail);
			return false;
		
		}
		}
		/*catch (Exception e) 
		{
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 255, 10, exception);
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	   	 	return false;
		}
	}*/

		
	
		  
	  public boolean checkMaterialRequisitionVoucherOpensOnAlertClick() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/						
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
			purchaseAccountTxt.click();
			
			String actPurchaseValue = purchaseAccountTxt.getAttribute("value");
			String expPurchaseValue ="STD RATE COGS ACC INV";
			
            getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			
			String actDepartmentValue = departmentTxt.getAttribute("value");
			String expDepartmentValue = "DUBAI";
								
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			
			String actR1C1Item =select1stRow_1stColumn.getText();
			String actR1C2Units=select1stRow_2ndColumn.getText();
			String actR1C5Qty  =select1stRow_3rdColumn.getText();
			String actR1C7Rate =select1stRow_4thColumn.getText();
			String actR1C8Gross=select1stRow_5thColumn.getText();

			String expR1C1Item ="STD RATE COGS ITEM";
			String expR1C2Units="Dozs";
			String expR1C5Qty  ="2.00";
			String expR1C7Rate ="20.00";
			String expR1C8Gross="40.00";
			
			if(errorMessageCloseBtn.isDisplayed())
			{
				errorMessageCloseBtn.click();
			}
			
			System.out.println("Entry Page Save with Vendor Item Value Actual   : "+actPurchaseValue   +" Value Expected : "+expPurchaseValue);	
			System.out.println("Entry Page Save with Vendor Units Value Actual  : "+actDepartmentValue  +" Value Expected : "+actDepartmentValue);
			System.out.println("Entry Page Save with Vendor Gross Value Actual  : "+actR1C1Item  +" Value Expected : "+expR1C1Item);			
			System.out.println("Entry Page Save with Vendor Units Value Actual  : "+actR1C2Units  +" Value Expected : "+expR1C2Units);
			System.out.println("Entry Page Save with Vendor QTY Value Actual    : "+actR1C5Qty    +" Value Expected : "+expR1C5Qty);
			System.out.println("Entry Page Save with Vendor RATE Value Actual   : "+actR1C7Rate   +" Value Expected : "+expR1C7Rate);
			System.out.println("Entry Page Save with Vendor Gross Value Actual  : "+actR1C8Gross  +" Value Expected : "+expR1C8Gross);			
			
		   			
			if(actR1C1Item.equalsIgnoreCase(expR1C1Item) && actR1C2Units.equalsIgnoreCase(expR1C2Units) && actR1C5Qty.equalsIgnoreCase(expR1C5Qty) 
				&& actR1C7Rate.equalsIgnoreCase(expR1C7Rate) && actR1C8Gross.equalsIgnoreCase(expR1C8Gross))
			{
				System.out.println("Test Pass : Data Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resPass);

				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
	    {
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 99, 10, exception);
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	   	 	return false;
	    }	
	  }*/
	  
		  
	 
	 public boolean checInputMoreQtyInAuthorizeUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/						
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		    select1stRow_3rdColumn.click();
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.click();
			enter_Quantity.clear();
			enter_Quantity.sendKeys("4");
			enter_Quantity.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.click();
			enter_Rate.clear();
			enter_Rate.sendKeys("20");
			enter_Rate.sendKeys(Keys.TAB);			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			
			String actR1C1Item =select1stRow_1stColumn.getText();
			String actR1C2Units=select1stRow_2ndColumn.getText();
			String actR1C5Qty  =select1stRow_3rdColumn.getText();
			String actR1C7Rate =select1stRow_4thColumn.getText();
			String actR1C8Gross=select1stRow_5thColumn.getText();

			String expR1C1Item ="STD RATE COGS ITEM";
			String expR1C2Units="Dozs";
			String expR1C5Qty  ="4.00";
			String expR1C7Rate ="20.00"; //20
			String expR1C8Gross="80.00";
			
			System.out.println("*********************************************** checkEnteringDatainFirstRow  ********************************************");
			System.out.println("Entry Page Save with Vendor Item Value Actual   : "+actR1C1Item   +" Value Expected : "+expR1C1Item);	
			System.out.println("Entry Page Save with Vendor Units Value Actual  : "+actR1C2Units  +" Value Expected : "+expR1C2Units);
			System.out.println("Entry Page Save with Vendor QTY Value Actual    : "+actR1C5Qty    +" Value Expected : "+expR1C5Qty);
			System.out.println("Entry Page Save with Vendor RATE Value Actual   : "+actR1C7Rate   +" Value Expected : "+expR1C7Rate);
			System.out.println("Entry Page Save with Vendor Gross Value Actual  : "+actR1C8Gross  +" Value Expected : "+expR1C8Gross);			
					
			
           /* String expSaveValidationMessageOnSaveWithPurchaseInptData="This row is making budget negative.";
			
			String actMessage=checkValidationMessage(expSaveValidationMessageOnSaveWithPurchaseInptData);
			
			String actGetSaveMessage = new String(actMessage);
						
			String actSaveMessage = actGetSaveMessage.substring(0,35);
     
			System.out.println("actSaveMessage"+actSaveMessage);
			
			System.out.println("Entry Page Save with Purchase Input Value Actual  : "+actMessage  +" Value Expected : "+expSaveValidationMessageOnSaveWithPurchaseInptData);*/			
	
		   if(actR1C1Item.equalsIgnoreCase(expR1C1Item) && actR1C2Units.equalsIgnoreCase(expR1C2Units) && actR1C5Qty.equalsIgnoreCase(expR1C5Qty)
				   && actR1C7Rate.equalsIgnoreCase(expR1C7Rate) && actR1C8Gross.equalsIgnoreCase(expR1C8Gross))
			{
				System.out.println("Test Pass : Data Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resFail);
				return false;
			}	
		}
		/*catch (Exception e) 
	    {
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 99, 10, e.getMessage());
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	   	 	return false;
	    }	
	  }*/
	  
	  
	 
	 public boolean checkAuthorizeBtnWithInputMoreQty() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/							
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(authorizeBtn));
		    authorizeBtn.click();
		    
		    boolean actGetAlert = getIsAlertPresent();
		    boolean expGetAlert  = true;
	
		    System.out.println("Entry Page Save with Purchase Input Value Actual  : "+actGetAlert  +" Value Expected : "+expGetAlert);
		    
		    if(getIsAlertPresent())
		    {
			    getAlert().accept();
		    }
		    
		    
		    String expSaveValidationMessageOnSaveWithPurchaseInptData="This Transaction will make the Budget Negative. It cannot be saved";
			
 			String actMessage=checkValidationMessage(expSaveValidationMessageOnSaveWithPurchaseInptData);
 			
 			String actGetSaveMessage = new String(actMessage);
 						
 			String actSaveMessage = actGetSaveMessage.substring(0,66);
      
 			System.out.println("actSaveMessage"+actSaveMessage);
 			
 			System.out.println("Entry Page Save with Purchase Input Value Actual  : "+actSaveMessage  +" Value Expected : "+expSaveValidationMessageOnSaveWithPurchaseInptData);			
 		
			if( actSaveMessage.equalsIgnoreCase(expSaveValidationMessageOnSaveWithPurchaseInptData))
			{
				System.out.println("Test Pass : Data Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resPass);

				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
	    {
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 99, 10, e.getMessage());
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	   	 	return false;
	    }	
	  }*/
	  
	 
	 public boolean checInputLessQtyInAuthorizeUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/						
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		    select1stRow_3rdColumn.click();
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.click();
			enter_Quantity.clear();
			enter_Quantity.sendKeys("2");
			enter_Quantity.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.click();
			enter_Rate.clear();
			enter_Rate.sendKeys("20");
			enter_Rate.sendKeys(Keys.TAB);			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			
			String actR1C1Item =select1stRow_1stColumn.getText();
			String actR1C2Units=select1stRow_2ndColumn.getText();
			String actR1C5Qty  =select1stRow_3rdColumn.getText();
			String actR1C7Rate =select1stRow_4thColumn.getText();
			String actR1C8Gross=select1stRow_5thColumn.getText();

			String expR1C1Item ="STD RATE COGS ITEM";
			String expR1C2Units="Dozs";
			String expR1C5Qty  ="2.00";
			String expR1C7Rate ="20.00";
			String expR1C8Gross="40.00";
			
			System.out.println("*********************************************** checkEnteringDatainFirstRow  ********************************************");
			System.out.println("Entry Page Save with Vendor Item Value Actual   : "+actR1C1Item   +" Value Expected : "+expR1C1Item);	
			System.out.println("Entry Page Save with Vendor Units Value Actual  : "+actR1C2Units  +" Value Expected : "+expR1C2Units);
			System.out.println("Entry Page Save with Vendor QTY Value Actual    : "+actR1C5Qty    +" Value Expected : "+expR1C5Qty);
			System.out.println("Entry Page Save with Vendor RATE Value Actual   : "+actR1C7Rate   +" Value Expected : "+expR1C7Rate);
			System.out.println("Entry Page Save with Vendor Gross Value Actual  : "+actR1C8Gross  +" Value Expected : "+expR1C8Gross);			
					
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(authorizeBtn));
		    authorizeBtn.click();
		    
		    boolean actGetAlert = getIsAlertPresent();
		    boolean expGetAlert  = true;
	
		    System.out.println("Entry Page Save with Purchase Input Value Actual  : "+actGetAlert  +" Value Expected : "+expGetAlert);
		    	    
		    getAlert().accept();
		   		    
		    String expSaveValidationMessageOnSaveWithPurchaseInputData="Voucher saved successfully";
			
 			String actMessage=checkValidationMessage(expSaveValidationMessageOnSaveWithPurchaseInputData);
 			
 			String actGetSaveMessage = new String(actMessage);
 						
 			String actSaveMessage = actGetSaveMessage.substring(0,26);
      
 			System.out.println("actSaveMessage"+actSaveMessage);
 			
 			System.out.println("Entry Page Save with Purchase Input Value Actual  : "+actSaveMessage  +" Value Expected : "+expSaveValidationMessageOnSaveWithPurchaseInputData);			
 		
		    if(actSaveMessage.equalsIgnoreCase(expSaveValidationMessageOnSaveWithPurchaseInputData))
			{
				System.out.println("Test Pass : Data Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resFail);
				return false;
			}	
		}
		/*catch (Exception e) 
	    {
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 99, 10, exception);
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	   	 	return false;
	    }	
	  }*/
	 
	 
	 
	 
	 
	 
	 
	 // Check the Status of Authorization Authorize User Voucher Home Page
	 // Check Close Btn and Login To Su
	 
	 public boolean checkMaterailRequsitionHomePageStatusInAuthUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		/*try
		{*/		   
		    getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(new_CloseBtn));
		    new_CloseBtn.click();
		    
		    Thread.sleep(2000);
		    
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksNewHomeRow1AuthorizationStatus));
		    
			String actStatus=openingStocksNewHomeRow1AuthorizationStatus.getText();
			String expStatus="Authorized";
			
			System.out.println("************************************* checkVoucherSuspendAuthorizeStatusInAuthorizeUser  *********************************");
			System.out.println("Suspend Status In home Page 	 :  "+actStatus 						+" Value Expected : "+expStatus);
					
			if (actStatus.equalsIgnoreCase(expStatus))
			{
				System.out.println("Test Pass : Entry Page Close Button");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 240, 9, resPass);
				return true;
			}
			
			else
			{
				System.out.println("Test Fail :  Entry Page Close Button");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 240, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
		{
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 240, 10, exception);
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	   	 	return false;
		}
	  }*/
	  
	  

	  
	  public boolean checkMaterialRequsitionVoucherCloseButtonInAuthUserHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	  {	
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
			String actUserInfo = userNameDisplay.getText();
			String expUserInfo = "UserAllOptionsST";
			
			excelReader.setCellData(xlfile, xlSheetName, 453, 8, actUserInfo);
			
			System.out.println("User Name Display Value Actual                   : " + actUserInfo +             " Value Expected : " + expUserInfo);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
			companyLogo.click();
			
	        //Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
			String actGetLoginCompanyNameInformation = companyName.getText();
			String actGetLoginCompanyName            = actGetLoginCompanyNameInformation.substring(0, 19);
			String expGetLoginCompanyName            = excelReader.getCellData(xlSheetName, 454, 7);
			companyLogo.click();
			
			excelReader.setCellData(xlfile, xlSheetName, 454, 8, actGetLoginCompanyName);
			
			System.out.println("Company Name Display Value Actual                : " + actGetLoginCompanyName +  " Value Expected : " + expGetLoginCompanyName);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
			
			boolean actlabelDashboard=labelDashboard.isDisplayed();
			boolean actselectDashboard=selectDashboard.isDisplayed();
			boolean actnewAddDashBoard=newAddDashBoard.isDisplayed();
			boolean actdashboardCustomizationSettings=dashboardCustomizationSettings.isDisplayed();
			
			boolean explabelDashboard=true;
			boolean expselectDashboard=true;
			boolean expnewAddDashBoard=true;
			boolean expdashboardCustomizationSettings=true;
			
			System.out.println("labelDashboard Value Actual        : " + actlabelDashboard  +               " Value Expected :  " + explabelDashboard);
			System.out.println("selectDashboard Value Actual       : " + actselectDashboard +               " Value Expected :  " + expselectDashboard);
			System.out.println("newAddDashBoard Value Actual       : " + actnewAddDashBoard +               " Value Expected :  " + expnewAddDashBoard);
			System.out.println("CustomizationSettings Value Actual : " + actdashboardCustomizationSettings+ " Value Expected :  " + expdashboardCustomizationSettings);
			
			boolean actMethod = actlabelDashboard==explabelDashboard && actselectDashboard==expselectDashboard && actnewAddDashBoard==expnewAddDashBoard
								&& actdashboardCustomizationSettings==expdashboardCustomizationSettings;
			
			String actResult = Boolean.toString(actMethod);
			
			excelReader.setCellData(xlfile, xlSheetName, 452, 8, actResult.toUpperCase());
			
			if(actUserInfo.equalsIgnoreCase(expUserInfo) && actGetLoginCompanyName.equalsIgnoreCase(expGetLoginCompanyName) && actMethod==true)
			{
				System.out.println("test pass : Document History Pop Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 452, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("test Fail : Document History Pop NOT Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 452, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
		{
			excelReader.setExceptionInExcel(xlfile, xlSheetName, 452, 10, e.getMessage());
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
			return false;
		}
	  }*/
		
	 
	 
	 
	  public boolean checLogoutAndLoginCompanyToSU() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

	    /*try
		{*/				
    	    getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
		    userNameDisplay.click();
		
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		    logoutOption.click();
		 
			boolean verifyFiledsStatus=true;
			
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(username));
						
			boolean actusername              = username.isDisplayed();
			boolean actpassword              = password.isDisplayed();
			boolean actcompanyDropDownList   = companyDropDownList.isDisplayed();
			boolean actsignIn                = signIn.isDisplayed();
			
			// Expected
			boolean expusername              = true;
			boolean exppassword              = true;
			boolean expcompanyDropDownList   = true;
			boolean expsignIn                = true;
			
			System.out.println("************************************* checkSignOut  *********************************");
			System.out.println("username             : "+actusername             +" Value Expected : "+expusername);
			System.out.println("password             : "+actpassword             +" Value Expected : "+exppassword);
			System.out.println("companyDropDownList  : "+actcompanyDropDownList  +" Value Expected : "+expcompanyDropDownList);
			System.out.println("signIn               : "+actsignIn               +" Value Expected : "+expsignIn);

			
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(username));
			username.click();
			 
			username.sendKeys("SU");
			getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
			 
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(password));
			password.sendKeys("su");
			 
			getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
			  
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(signIn));
			signIn.click();
		
			Thread.sleep(5000);
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
			String actUserInfo=userNameDisplay.getText();
			
			System.out.println("User Info : "+actUserInfo);
			
			System.out.println("User Info Capture Text :"+userNameDisplay.getText());
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
			
			companyLogo.click();
			
			String actCompanyTxt=companyName.getText();
			String actLoginCompanyName=actCompanyTxt.substring(0, 19);
			System.out.println("company name :"+ actLoginCompanyName);
			companyLogo.click();
						    	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
			String getDashboard=labelDashboard.getText();
				
			//boolean getnoOfAlertsDisplay=noOfAlertsDisplay.isDisplayed();
					
			// Expected
			String expuserInfo="SU";
			String expLoginCompanyName="Automation Company ";
			String expDashboard="Dashboard";
			//boolean expnoOfAlertsDisplay=true;
			
			System.out.println("************************************* checkAlertsToSignIAsAuthorizeUser  *********************************");
			System.out.println("userNameDisplay    : "+actUserInfo              +" Value Expected : "+expuserInfo);
			System.out.println("LoginCompanyName   : "+actLoginCompanyName      +" Value Expected : "+expLoginCompanyName);
			System.out.println("labelDashboard     : "+getDashboard             +" Value Expected : "+expDashboard);
			//System.out.println("noOfAlertsDisplay  : "+getnoOfAlertsDisplay     +" Value Expected : "+expnoOfAlertsDisplay);
					
			boolean actMethod=actUserInfo.equalsIgnoreCase(expuserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)
					&& getDashboard.equalsIgnoreCase(expDashboard) ;
			
			String actResult=Boolean.toString(actMethod);
			String expResult=excelReader.getCellData("PurchaseVoucherN", 254, 7);
			
			excelReader.setCellData(xlfile, "PurchaseVoucherN", 254, 8, actResult);
					
			if(actResult.equalsIgnoreCase(expResult))
			{
				System.out.println("Test Pass : Transaction Authorization is Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 251, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Transaction Authorization is Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 251, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
		{
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 251, 10, e.getMessage());
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
			return false;
		}
	  }*/
 
	 
	 
	 
	  public static boolean checkOpenPurchasesOrderVoucherToDisplayLinksOnAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
	  	excelReader=new ExcelReader(POJOUtility.getExcelPath());
	  	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	  	
	  	/*try
	  	{*/
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
	  		inventoryMenu.click();
	  		
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
	  		invTransactionsMenu.click();
	  		
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
	  		inventoryTransactionsPurchasesMenu.click();
	  	
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
	  		purchasesOrdersVoucher.click();
	  		
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			
			int count				                 = openingStocksNewToolBarButtons.size();
			
			ArrayList<String> actbuttons             = new ArrayList<String>();
			
			for (int i = 0; i < count; i++) 
			{
				String data                          = openingStocksNewToolBarButtons.get(i).getText();
				actbuttons.add(data);
			}
			    
			String actopeningStocksNewToolBarButtons = actbuttons.toString();
			String expopeningStocksNewToolBarButtons = "[New,   Edit  , Print, Delete, Authorize, Suspend, Reject, , Print barcode, Export to XML, Setting, Close]";
		
			System.out.println("actopeningStocksNewToolBarButtons  : " + actopeningStocksNewToolBarButtons);
			System.out.println("expopeningStocksNewToolBarButtons  : " + expopeningStocksNewToolBarButtons);
			
			excelReader.setCellData(xlfile, xlSheetName, 19, 8, actopeningStocksNewToolBarButtons);	
			
			String actsetAsDefaultView		=	setAsDefaultView.getText();
			String expsetAsDefaultView		=	excelReader.getCellData(xlSheetName, 20, 7);
			excelReader.setCellData(xlfile, xlSheetName, 20, 8, actsetAsDefaultView);
			
			
			//Ribbon Control Icons
	        int count1		= headingOptions.size();
			
			ArrayList<String> headingButtons = new ArrayList<String>();
			
			for (int i = 1; i < count1; i++) 
			{
				String data=headingOptions.get(i).getAttribute("title");
				
				headingButtons.add(data);
			}
			
			String actHeadingButtons = headingButtons.toString();
			String expHeadingButtons = excelReader.getCellData(xlSheetName, 21, 7);
			excelReader.setCellData(xlfile, xlSheetName, 21, 8, actHeadingButtons);
			
			
			System.out.println("actopeningStocksNewHeadingButtons  : " + actHeadingButtons);
			System.out.println("expopeningStocksNewHeadingButtons  : " + expHeadingButtons);
			
			//Home Page Column Headings
	        int headerTxtCount		= headerTxtList.size();
			
			ArrayList<String> headerTxt = new ArrayList<String>();
			
			for (int i = 2; i < headerTxtCount; i++) 
			{
				String data=headerTxtList.get(i).getText();
				headerTxt.add(data);
			}
			
			String actHeaderTxt = headerTxt.toString();
			String expHeaderTxt = excelReader.getCellData(xlSheetName, 22, 7);
			excelReader.setCellData(xlfile, xlSheetName, 22, 8, actHeaderTxt);
			
			
			System.out.println("actopeningStocksNewHomePageColumnsButtons  : " + actHeaderTxt);
			System.out.println("expopeningStocksNewHomePageColumnsButtons  : " + expHeaderTxt);
			
			boolean actVoucherAllAuthorizationAndLinks =voucherAll.isDisplayed() && pendingAuthorizationProcess.isDisplayed()
					                                      && pendingMaterialRequistion.isDisplayed();
			
			boolean expVoucherAllAuthorizationAndLinks = true;
			
			System.out.println("Home page All Pending Authorization And Links Value Actual : " + actVoucherAllAuthorizationAndLinks  +   " Value Expected :  " + expVoucherAllAuthorizationAndLinks);
			
			if(actopeningStocksNewToolBarButtons.equalsIgnoreCase(expopeningStocksNewToolBarButtons)
					&& actHeadingButtons.equalsIgnoreCase(expHeadingButtons)
					&& actHeaderTxt.equalsIgnoreCase(expHeaderTxt)
					
					&& actVoucherAllAuthorizationAndLinks==expVoucherAllAuthorizationAndLinks)
			{
				System.out.println("Test Pass : HomeScreen Options Are Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 18, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : HomeScreen Options Are NOT Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 18, 9, resFail);
				return false;
			}
		}
	  	/*catch (Exception e) 
	    {
	  		String exception =e.getMessage().substring(0, 50);
	          excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	          System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	          return false;
	     }
	   }*/
	  	
	 
		  
	  
	  
	  
	  
	  
  
	  
     @FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr[1]/td[4]")
	 private static WebElement openingStocksNewHomeRow1VoucherNo;
	 
	 @FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr[1]/td[5]")
	 private static WebElement openingStocksNewHomeRow1CreatedBy;
	 
	 @FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr[1]/td[6]")
	 private static WebElement openingStocksNewHomeRow1ModifiedBy;
	 
	 @FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr[1]/td[11]")
	 private static WebElement openingStocksNewHomeRow1SuspendStatus;
	 
	 @FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr")
	 private static List<WebElement> openingStocksNewHomePageTableRowsCount;
	  
	  public boolean checkPurchasesOrderVoucherPendingLinksOptionsInHomePageOnAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {	
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pendingMaterialRequistion));
			pendingMaterialRequistion.click();
			
			Thread.sleep(2000);
			
			//Home Page Column Headings
			
	        int headerTxtCount		= headerTxtList.size();
			
			ArrayList<String> headerTxt = new ArrayList<String>();
			
			for (int i = 2; i < headerTxtCount; i++) 
			{
				String data=headerTxtList.get(i).getText();
				headerTxt.add(data);
			}
			
			String actHeaderTxt = headerTxt.toString();
			String expHeaderTxt = excelReader.getCellData(xlSheetName, 433, 7);
			excelReader.setCellData(xlfile, xlSheetName, 433, 8, actHeaderTxt);
			
			System.out.println("actopeningStocksNewHomePageColumnsButtons  : " + actHeaderTxt);
			System.out.println("expopeningStocksNewHomePageColumnsButtons  : " + expHeaderTxt);
			
			System.out.println("Count Of Header Columns in Home Page Value Actual    :  " + actHeaderTxt );
			System.out.println("Count Of Header Columns in Home Page Value Expected  :  " + expHeaderTxt );
					
			String actHomePageVoucherNo     = openingStocksNewHomeRow1VoucherNo.getText();
			String actHomePageCreatedBy     = openingStocksNewHomeRow1CreatedBy.getText();
			String actHomePageModifiedBy    = openingStocksNewHomeRow1ModifiedBy.getText();
			String actHomePageSuspendBy     = openingStocksNewHomeRow1SuspendStatus.getText();
			String actHomePageAuthorizedBy  = openingStocksNewHomeRow1AuthorizationStatus.getText();
			
			String expHomePageVoucherNo     = "1";
			String expHomePageCreatedBy     = "SU";
			String expHomePageModifiedBy    = "UserAllOptionsST";
			String expHomePageSuspendBy     = "False";
			String expHomePageAuthorizeBy   = "Authorized";
			
			
			System.out.println("Voucher No in Home Page Value Actual   :  " + actHomePageVoucherNo+    " Value Expected : " + expHomePageVoucherNo);
			System.out.println("Created By in Home Page Value Actual   :  " + actHomePageCreatedBy+    " Value Expected : " + expHomePageCreatedBy);
			System.out.println("Modified By in Home Page Value Actual  :  " + actHomePageModifiedBy+   " Value Expected : " + expHomePageModifiedBy);
			System.out.println("Suspend By in Home Page Value Actual   :  " + actHomePageSuspendBy+    " Value Expected : " + expHomePageSuspendBy);
			System.out.println("Authorize By in Home Page Value Actual :  " + actHomePageAuthorizedBy+ " Value Expected : " + expHomePageAuthorizeBy);
					
			if(actHomePageVoucherNo.equalsIgnoreCase(expHomePageVoucherNo) && actHomePageCreatedBy.equalsIgnoreCase(expHomePageCreatedBy)
					&& actHomePageModifiedBy.equalsIgnoreCase(expHomePageModifiedBy) && actHomePageSuspendBy.equalsIgnoreCase(expHomePageSuspendBy)
					&& actHomePageAuthorizedBy.equalsIgnoreCase(expHomePageAuthorizeBy))		
			{
				System.out.println("Test Pass : entry Page Options Are Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 13, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : entry Page NOT Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 13, 9, resFail);
				return false;
			}
		 }
		/*catch (Exception e) 
	    {
	  		String exception =e.getMessage().substring(0, 50);
	          excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	          System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	          return false;
	    }
	  }	*/
	 
	 
	 
	  public boolean checkPurchasesOrderVoucherOpenInHomePageOnAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {	
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksNewHomeRow1VoucherNo));
			getAction().doubleClick(openingStocksNewHomeRow1VoucherNo).build().perform();
			
			Thread.sleep(2000);
			
            getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
            String actPurchaseAccount = purchaseAccountTxt.getAttribute("value");
            String actDepartment      = departmentValuetxt.getAttribute("value");
                     
			String expPurchaseAccount = "STD RATE COGS ACC INV";
			String expDepartment      = "DUBAI";
           	
			System.out.println("Entry Page Save with Vendor Item Value Actual   : " + actPurchaseAccount   + " Value Expected : " + expPurchaseAccount);	
			System.out.println("Entry Page Save with Vendor Units Value Actual  : " + actDepartment        + " Value Expected : " + expDepartment);
		
			String actR1C1Item =select1stRow_1stColumn.getText();
			String actR1C2Units=select1stRow_2ndColumn.getText();
			String actR1C5Qty  =select1stRow_3rdColumn.getText();
			String actR1C7Link =select1stRow_4thColumn.getText();
			String actR1C8Rate=select1stRow_5thColumn.getText();
			String actR1C9Gross=select1stRow_6thColumn.getText();
			
			String expR1C1Item ="STD RATE COGS ITEM";
			String expR1C2Units="Dozs";
			String expR1C5Qty  ="2.00";
			String expR1C7Link ="MatReq:1";
			String expR1C8Rate="20.00";
			String expR1C9Gross="40.00";
				
			System.out.println("Entry Page Save with Vendor Item Value Actual   : " + actR1C1Item   + " Value Expected : " + expR1C1Item);	
			System.out.println("Entry Page Save with Vendor Units Value Actual  : " + actR1C2Units  + " Value Expected : " + expR1C2Units);
			System.out.println("Entry Page Save with Vendor QTY Value Actual    : " + actR1C5Qty    + " Value Expected : " + expR1C5Qty);
			System.out.println("Entry Page Save with Vendor RATE Value Actual   : " + actR1C7Link   + " Value Expected : " + expR1C7Link);
			System.out.println("Entry Page Save with Vendor Gross Value Actual  : " + actR1C8Rate   + " Value Expected : " + expR1C8Rate);			
			System.out.println("Entry Page Save with Vendor Gross Value Actual  : " + actR1C9Gross  + " Value Expected : " + expR1C9Gross);
			
			if(errorMessageCloseBtn.isDisplayed())
			{
				errorMessageCloseBtn.click();
			}
			
			
			if(actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) && actDepartment.equalsIgnoreCase(expDepartment)
					&& actR1C1Item.equalsIgnoreCase(expR1C1Item) && actR1C2Units.equalsIgnoreCase(expR1C2Units)
					&& actR1C5Qty.equalsIgnoreCase(expR1C5Qty) && actR1C7Link.equalsIgnoreCase(expR1C7Link)
					&& actR1C8Rate.equalsIgnoreCase(expR1C8Rate) && actR1C9Gross.equalsIgnoreCase(expR1C9Gross))		
			{
				System.out.println("Test Pass : entry Page Options Are Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 13, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : entry Page NOT Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 13, 9, resFail);
				return false;
			}
		 }
		/*catch (Exception e) 
	    {
	  		  String exception =e.getMessage().substring(0, 50);
	          excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	          System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	          return false;
	    }
	  }	*/

	  
	  @FindBy(xpath="//*[@id='id_body_35_3']")
	  private static WebElement  linkColumn;
	  
	  
	  public boolean checkPurchaseVoucherInputQtyMoreOnAthorizeVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/								
            getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
            vendorAccountTxt.click();
            vendorAccountTxt.sendKeys(Keys.SPACE);
			Thread.sleep(2000);
			vendorAccountTxt.sendKeys("Vendor A");
			Thread.sleep(2000);
			vendorAccountTxt.sendKeys(Keys.TAB);
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
			select1stRow_3rdColumn.click();
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.click();
			enter_Quantity.clear();
			enter_Quantity.sendKeys("5");
			enter_Quantity.sendKeys(Keys.TAB);
			
			linkColumn.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.click();
			enter_Rate.clear();
			enter_Rate.sendKeys("20");
			enter_Rate.sendKeys(Keys.TAB);			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.click();
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			
			String actR1C1Item =select1stRow_1stColumn.getText();
			String actR1C2Units=select1stRow_2ndColumn.getText();
			String actR1C5Qty  =select1stRow_3rdColumn.getText();
			String actR1C7Rate =select1stRow_4thColumn.getText();
			String actR1C8Gross=select1stRow_5thColumn.getText();

			String expR1C1Item ="STD RATE COGS ITEM";
			String expR1C2Units="Dozs";
			String expR1C5Qty  ="5.00";
			String expR1C7Rate ="MatReq:1";
			String expR1C8Gross="20.00"; //20
			
			System.out.println("*********************************************** checkEnteringDatainFirstRow  ********************************************");
			System.out.println("Entry Page Save with Vendor Item Value Actual   : "+actR1C1Item   +" Value Expected : "+expR1C1Item);	
			System.out.println("Entry Page Save with Vendor Units Value Actual  : "+actR1C2Units  +" Value Expected : "+expR1C2Units);
			System.out.println("Entry Page Save with Vendor QTY Value Actual    : "+actR1C5Qty    +" Value Expected : "+expR1C5Qty);
			System.out.println("Entry Page Save with Vendor RATE Value Actual   : "+actR1C7Rate   +" Value Expected : "+expR1C7Rate);
			System.out.println("Entry Page Save with Vendor Gross Value Actual  : "+actR1C8Gross  +" Value Expected : "+expR1C8Gross);			
			
			/*String expSaveValidationMessageOnSaveWithPurchaseInptData="This row is making budget negative.";
			
			String actMessage=checkValidationMessage(expSaveValidationMessageOnSaveWithPurchaseInptData);
			
			String actGetSaveMessage = new String(actMessage);
						
			String actSaveMessage = actGetSaveMessage.substring(0,35);
      
			System.out.println("actSaveMessage"+actSaveMessage);
			
			System.out.println("Entry Page Save with Purchase Input Value Actual  : "+actMessage  +" Value Expected : "+expSaveValidationMessageOnSaveWithPurchaseInptData);*/			
				
			if(actR1C1Item.equalsIgnoreCase(expR1C1Item) && actR1C2Units.equalsIgnoreCase(expR1C2Units) && actR1C5Qty.equalsIgnoreCase(expR1C5Qty) 
				&& actR1C7Rate.equalsIgnoreCase(expR1C7Rate) && actR1C8Gross.equalsIgnoreCase(expR1C8Gross)
				
					/*&& actSaveMessage.equalsIgnoreCase(expSaveValidationMessageOnSaveWithPurchaseInptData)*/)
			{
				System.out.println("Test Pass : Data Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resPass);

				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
	    {
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 99, 10, exception);
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	   	 	return false;
	    }	
	  }	*/ 
	   

	  
	  
	    //WorkFlow Elements
	
		@FindBy(xpath="//*[@id='id_transactionentry_workflow_popup_searchon']")
		private static WebElement  searchOnDropdown;
		
		@FindBy(xpath="//*[@id='id_transactionentry_workflow_popup_searchtype']")
		private static WebElement  containingDropdown;
			
		@FindBy(xpath="//*[@id='id_transactionentry_workflow_popup_search']")
		private static WebElement  workFlowSearchTxt;
		
		@FindBy(xpath="//*[@id='id_transactionentry_workflow_popup_close']")
		private static WebElement  workFlowCloseBtn;
		
		@FindBy(xpath="//ul[@class='nav navbar-nav navbar-right']//div[@class='toolbar_button_text font-7'][contains(text(),'Customize')]")
		private static WebElement  workFlowCustomizeBtn;
		
		@FindBy(xpath="//*[@id='id_transaction_workflow_popup_heading']/div/ul/li[3]/div/div[2]")
		private static WebElement  workFlowOkBtn;
		
		@FindBy(xpath="//div[@class='toolbar_button_text font-7'][contains(text(),'Ok')]")
		private static WebElement  workFlowOkBtnInSalesVoucher;
		
		
		@FindBy(xpath="//span[@id='id_transaction_entry_workflow_togglebutton']")
		private static WebElement  workFlowMinimizeBtn;
		
		
		
		//WorkFlow Screen with Line wise
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_1-1']//input")
		private static WebElement  workFlowHeaderChkBox;
		
		@FindBy(xpath="//th[@id='id_transaction_entry_detail_workflow_popup_control_heading_6']//div[contains(text(),'Voucher No')]")
		private static WebElement  workFlowHeaderVoucherNo;
		
		@FindBy(xpath="//th[@id='id_transaction_entry_detail_workflow_popup_control_heading_7']//div[contains(text(),'Date')]")
		private static WebElement  workFlowHeaderDate;
		
		@FindBy(xpath="//div[contains(text(),'Name')]")
		private static WebElement  workFlowHeaderName;
		
		@FindBy(xpath="//th[@id='id_transaction_entry_detail_workflow_popup_control_heading_9']//div[contains(text(),'Quantity')]")
		private static WebElement  workFlowHeaderQty;
		
		@FindBy(xpath="//div[contains(text(),'Suspended')]")
		private static WebElement  workFlowHeaderSuspend;
		
		
		//WorkFlow with Line wise
		// 1 Row
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_1-1']//input")
		private static WebElement workFlowRadioBtnRowOne;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_1-6']")
		private static WebElement  workFlowRow1Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_1-7']")
		private static WebElement  workFlowRow1Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_1-8']")
		private static WebElement  workFlowRow1Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_1-9']")
		private static WebElement  workFlowRow1Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_1-10']")
		private static WebElement  workFlowRow1Column5;
		
		// 2 Row
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_2-1']//input")
		private static WebElement workFlowRadioBtnRowTwo;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_2-6']")
		private static WebElement  workFlowRow2Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_2-7']")
		private static WebElement  workFlowRow2Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_2-8']")
		private static WebElement  workFlowRow2Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_2-9']")
		private static WebElement  workFlowRow2Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_2-10']")
		private static WebElement  workFlowRow2Column5;
		
		
		
		
		// 3 Row
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_3-1']//input")
		private static WebElement workFlowRadioBtnRowThree;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_3-6']")
		private static WebElement  workFlowRow3Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_3-7']")
		private static WebElement  workFlowRow3Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_3-8']")
		private static WebElement  workFlowRow3Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_3-9']")
		private static WebElement  workFlowRow3Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_3-10']")
		private static WebElement  workFlowRow3Column5;
		
		
		// 4 Row
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_col_4-1']//input")
		private static WebElement workFlowRadioBtnRowFour;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_4-6']")
		private static WebElement  workFlowRow4Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_4-7']")
		private static WebElement  workFlowRow4Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_4-8']")
		private static WebElement  workFlowRow4Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_4-9']")
		private static WebElement  workFlowRow4Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_4-10']")
		private static WebElement  workFlowRow4Column5;	
		
		
		
	     // 5 Row
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_col_5-1']//input")
		private static WebElement workFlowRadioBtnRowFive;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_5-6']")
		private static WebElement  workFlowRow5Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_5-7']")
		private static WebElement  workFlowRow5Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_5-8']")
		private static WebElement  workFlowRow5Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_5-9']")
		private static WebElement  workFlowRow5Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_5-10']")
		private static WebElement  workFlowRow5Column5;	
			
			
			
	    // 6 Row
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_col_6-1']//input")
		private static WebElement workFlowRadioBtnRowSix;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_6-6']")
		private static WebElement  workFlowRow6Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_6-7']")
		private static WebElement  workFlowRow6Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_6-8']")
		private static WebElement  workFlowRow6Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_6-9']")
		private static WebElement  workFlowRow6Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_6-10']")
		private static WebElement  workFlowRow6Column5;	
		

	    // 7 Row
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_col_7-1']//input")
		private static WebElement workFlowRadioBtnRowSeven;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_7-6']")
		private static WebElement  workFlowRow7Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_7-7']")
		private static WebElement  workFlowRow7Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_7-8']")
		private static WebElement  workFlowRow7Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_7-9']")
		private static WebElement  workFlowRow7Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_7-10']")
		private static WebElement  workFlowRow7Column5;	
		
		

	    // 8 Row
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_col_8-1']//input")
		private static WebElement workFlowRadioBtnRowEight;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_8-6']")
		private static WebElement  workFlowRow8Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_8-7']")
		private static WebElement  workFlowRow8Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_8-8']")
		private static WebElement  workFlowRow8Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_8-9']")
		private static WebElement  workFlowRow8Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_8-10']")
		private static WebElement  workFlowRow8Column5;	
		
		

	    // 9 Row
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_col_9-1']//input")
		private static WebElement workFlowRadioBtnRowNine;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_9-6']")
		private static WebElement  workFlowRow9Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_9-7']")
		private static WebElement  workFlowRow9Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_9-8']")
		private static WebElement  workFlowRow9Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_9-9']")
		private static WebElement  workFlowRow9Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_9-10']")
		private static WebElement  workFlowRow9Column5;	
		
		

	    // 10 Row
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_col_10-1']//input")
		private static WebElement workFlowRadioBtnRowTen;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_10-6']")
		private static WebElement  workFlowRow10Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_10-7']")
		private static WebElement  workFlowRow10Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_10-8']")
		private static WebElement  workFlowRow10Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_10-9']")
		private static WebElement  workFlowRow10Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_10-10']")
		private static WebElement  workFlowRow10Column5;	
		
		

	    // 11 Row
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_col_11-1']//input")
		private static WebElement workFlowRadioBtnRowElevan;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_11-6']")
		private static WebElement  workFlowRow11Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_11-7']")
		private static WebElement  workFlowRow11Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_11-8']")
		private static WebElement  workFlowRow11Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_11-9']")
		private static WebElement  workFlowRow11Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_11-10']")
		private static WebElement  workFlowRow11Column5;	
		
		

	    // 12 Row
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_col_12-1']//input")
		private static WebElement workFlowRadioBtnRowTwevle;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_12-6']")
		private static WebElement  workFlowRow12Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_12-7']")
		private static WebElement  workFlowRow12Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_12-8']")
		private static WebElement  workFlowRow12Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_12-9']")
		private static WebElement  workFlowRow12Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_12-10']")
		private static WebElement  workFlowRow12Column5;	
		

	    // 13 Row
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_col_13-1']//input")
		private static WebElement workFlowRadioBtnRowThirteen;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_13-6']")
		private static WebElement  workFlowRow13Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_13-7']")
		private static WebElement  workFlowRow13Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_13-8']")
		private static WebElement  workFlowRow13Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_13-9']")
		private static WebElement  workFlowRow13Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_12-10']")
		private static WebElement  workFlowRow13Column5;	
		
		

	    // 14 Row
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_col_14-1']//input")
		private static WebElement workFlowRadioBtnRowFourteen;

		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_14-6']")
		private static WebElement  workFlowRow14Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_14-7']")
		private static WebElement  workFlowRow14Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_14-8']")
		private static WebElement  workFlowRow14Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_14-9']")
		private static WebElement  workFlowRow14Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_14-10']")
		private static WebElement  workFlowRow14Column5;	
			
			

	    // 15 Row
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_col_15-1']//input")
		private static WebElement workFlowRadioBtnRowFifteen;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_15-6']")
		private static WebElement  workFlowRow15Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_15-7']")
		private static WebElement  workFlowRow15Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_15-8']")
		private static WebElement  workFlowRow15Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_15-9']")
		private static WebElement  workFlowRow15Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_15-10']")
		private static WebElement  workFlowRow15Column5;	
		
		

		

	    // 16 Row
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_col_16-1']//input")
		private static WebElement workFlowRadioBtnRowSixteen;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_16-6']")
		private static WebElement  workFlowRow16Column1;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_16-7']")
		private static WebElement  workFlowRow16Column2;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_16-8']")
		private static WebElement  workFlowRow16Column3;
				
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_16-9']")
		private static WebElement  workFlowRow16Column4;
		
		@FindBy(xpath="//td[@id='id_transaction_entry_detail_workflow_popup_col_16-10']")
		private static WebElement  workFlowRow16Column5;	
	
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_workflow_popup_body']/tr")
		private static List<WebElement>  LinkRowCount;
		

		@FindBy(xpath="//*[@id='id_transaction_entry_heading2']")
		private static WebElement  purchaseVoucherWorlFlowText;

	  
	  
	  public boolean checkLinksPopUpScreenUserInputMoreQty() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/								
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
			select1stRow_3rdColumn.click();
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.click();
			enter_Quantity.clear();
			enter_Quantity.sendKeys("5");
			enter_Quantity.sendKeys(Keys.TAB);
			
			linkColumn.sendKeys(Keys.SPACE);
		    	
			List<WebElement> rowCountInLinkScreen=LinkRowCount;
			
			int actLinkRowCountValue=rowCountInLinkScreen.size();
			
			System.out.println("the count "+actLinkRowCountValue);
	      			
			boolean actsearchOnDropdown     = searchOnDropdown.isDisplayed();
			boolean actcontainingDropdown   = containingDropdown.isDisplayed();
			boolean actworkFlowSearchTxt    = workFlowSearchTxt.isDisplayed();
			boolean actworkFlowCustomizeBtn = workFlowCustomizeBtn.isDisplayed();
			boolean actworkFlowOkBtn        = workFlowOkBtn.isDisplayed();
					
			String actworkFlowRow11Column1 = workFlowRow1Column1.getText();
			String actworkFlowRow11Column2 = workFlowRow1Column3.getText();
			String actworkFlowRow11Column3 = workFlowRow1Column4.getText();
			String actworkFlowRow11Column4 = workFlowRow1Column5.getText();
			
            int expLinkRowCountValue=1;
			
			boolean expsearchOnDropdown=true;
			boolean expcontainingDropdown=true;
			boolean expworkFlowSearchTxt=true;
			boolean expworkFlowCustomizeBtn=true;
			boolean expworkFlowOkBtn=true;
			
			String expworkFlowRow11Column1="MatReq:1";
			String expworkFlowRow11Column3="STD RATE COGS ITEM";
			String expworkFlowRow11Column4="24.00";
			String expworkFlowRow11Column5="0.00";
			
			Thread.sleep(2000);
			
			workFlowOkBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			
			String actR1C1Item =select1stRow_1stColumn.getText();
			String actR1C2Units=select1stRow_2ndColumn.getText();
			String actR1C5Qty  =select1stRow_3rdColumn.getText();
			String actR1C7Rate =select1stRow_4thColumn.getText();
			String actR1C8Gross=select1stRow_5thColumn.getText();

			String expR1C1Item ="STD RATE COGS ITEM";
			String expR1C2Units="Dozs";
			String expR1C5Qty  ="2.00";
			String expR1C7Rate ="MatReq:1";
			String expR1C8Gross="20.00"; //20
			
			System.out.println("*********************************************** checkEnteringDatainFirstRow  ********************************************");
			System.out.println("Entry Page Save with Vendor Item Value Actual   : "+actR1C1Item   +" Value Expected : "+expR1C1Item);	
			System.out.println("Entry Page Save with Vendor Units Value Actual  : "+actR1C2Units  +" Value Expected : "+expR1C2Units);
			System.out.println("Entry Page Save with Vendor QTY Value Actual    : "+actR1C5Qty    +" Value Expected : "+expR1C5Qty);
			System.out.println("Entry Page Save with Vendor RATE Value Actual   : "+actR1C7Rate   +" Value Expected : "+expR1C7Rate);
			System.out.println("Entry Page Save with Vendor Gross Value Actual  : "+actR1C8Gross  +" Value Expected : "+expR1C8Gross);			
					
			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();*/
		
			
			if(actR1C1Item.equalsIgnoreCase(expR1C1Item) && actR1C2Units.equalsIgnoreCase(expR1C2Units) && actR1C5Qty.equalsIgnoreCase(expR1C5Qty) 
				&& actR1C7Rate.equalsIgnoreCase(expR1C7Rate) && actR1C8Gross.equalsIgnoreCase(expR1C8Gross))
			{
				System.out.println("Test Pass : Data Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resPass);

				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
	    {
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 99, 10, e.getMessage());
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	   	 	return false;
	    }	
	  }	 */

	  
	  
	  public boolean checkSavePurchaseOrdersVoucherInUnAuthorizeUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/							
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
	
			String expSaveValidationMessageOnSaveWithPurchaseInptData="Voucher Saved Successfully";
			
			String actMessage=checkValidationMessage(expSaveValidationMessageOnSaveWithPurchaseInptData);
			
			String actGetSaveMessage = new String(actMessage);
			
			String actSaveMessage = actGetSaveMessage.substring(0,26);
      
			System.out.println("actSaveMessage"+actSaveMessage);
			
			System.out.println("Entry Page Save with Purchase Input Value Actual  : "+actMessage  +" Value Expected : "+expSaveValidationMessageOnSaveWithPurchaseInptData);			
		
			
			if( actSaveMessage.equalsIgnoreCase(expSaveValidationMessageOnSaveWithPurchaseInptData))
			{
				System.out.println("Test Pass : Data Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resPass);

				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT Entered in first Row");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 99, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
	    {
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 99, 10, exception);
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	   	 	return false;
	    }	
	  }*/
	  
	  
	  
	  
	  public boolean checkPurchasesOrderHomePageStatusInUnAuthorizeUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		/*try
		{*/	   
		    getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(new_CloseBtn));
		    new_CloseBtn.click();
		    
		    Thread.sleep(3000);
		    
		    getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(openingStocksNewHomeRow1AuthorizationStatus));
		    
			String actStatus=openingStocksNewHomeRow1AuthorizationStatus.getText();
			String expStatus="Pending";
			
			System.out.println("************************************* checkVoucherSuspendAuthorizeStatusInAuthorizeUser  *********************************");
			System.out.println("Suspend Status In home Page 	 :  "+actStatus 						+" Value Expected : "+expStatus);
					
			if (actStatus.equalsIgnoreCase(expStatus))
			{
				System.out.println("Test Pass : Entry Page Close Button");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 240, 9, resPass);
				return true;
			}		
			else
			{
				System.out.println("Test Fail :  Entry Page Close Button");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 240, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
		{
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 240, 10, exception);
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	   	 	return false;
		}
	  }*/
	  
	  
	  
	  public boolean checkPurchasesOrdersVoucherCloseButtonInHomePageInUnAuthorize() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	  {	
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/*try
		{*/
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
			String actUserInfo = userNameDisplay.getText();
			String expUserInfo = excelReader.getCellData(xlSheetName, 453, 7);
			
			excelReader.setCellData(xlfile, xlSheetName, 453, 8, actUserInfo);
			
			System.out.println("User Name Display Value Actual                   : " + actUserInfo +             " Value Expected : " + expUserInfo);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
			companyLogo.click();
			
	        //Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
			String actGetLoginCompanyNameInformation = companyName.getText();
			String actGetLoginCompanyName            = actGetLoginCompanyNameInformation.substring(0, 19);
			String expGetLoginCompanyName            = excelReader.getCellData(xlSheetName, 454, 7);
			companyLogo.click();
			
			excelReader.setCellData(xlfile, xlSheetName, 454, 8, actGetLoginCompanyName);
			
			System.out.println("Company Name Display Value Actual                : " + actGetLoginCompanyName +  " Value Expected : " + expGetLoginCompanyName);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
			
			boolean actlabelDashboard=labelDashboard.isDisplayed();
			boolean actselectDashboard=selectDashboard.isDisplayed();
			boolean actnewAddDashBoard=newAddDashBoard.isDisplayed();
			boolean actdashboardCustomizationSettings=dashboardCustomizationSettings.isDisplayed();
			
			boolean explabelDashboard=true;
			boolean expselectDashboard=true;
			boolean expnewAddDashBoard=true;
			boolean expdashboardCustomizationSettings=true;
			
			System.out.println("labelDashboard Value Actual        : " + actlabelDashboard  +               " Value Expected :  " + explabelDashboard);
			System.out.println("selectDashboard Value Actual       : " + actselectDashboard +               " Value Expected :  " + expselectDashboard);
			System.out.println("newAddDashBoard Value Actual       : " + actnewAddDashBoard +               " Value Expected :  " + expnewAddDashBoard);
			System.out.println("CustomizationSettings Value Actual : " + actdashboardCustomizationSettings+ " Value Expected :  " + expdashboardCustomizationSettings);
			
			boolean actMethod = actlabelDashboard==explabelDashboard && actselectDashboard==expselectDashboard && actnewAddDashBoard==expnewAddDashBoard
								&& actdashboardCustomizationSettings==expdashboardCustomizationSettings;
			
			String actResult = Boolean.toString(actMethod);
			
			excelReader.setCellData(xlfile, xlSheetName, 452, 8, actResult.toUpperCase());
			
			if(actUserInfo.equalsIgnoreCase(expUserInfo) && actGetLoginCompanyName.equalsIgnoreCase(expGetLoginCompanyName) && actMethod==true)
			{
				System.out.println("test pass : Document History Pop Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 452, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("test Fail : Document History Pop NOT Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 452, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
		{
			excelReader.setExceptionInExcel(xlfile, xlSheetName, 452, 10, e.getMessage());
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
			return false;
		}
	  }*/
	  
	  
	  
	  
	  
	  
	  public boolean checLogoutAndLoginCompanyToAuthorizeUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

	    /*try
		{*/				
    	    getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
		    userNameDisplay.click();
		
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		    logoutOption.click();
		 
			boolean verifyFiledsStatus=true;
			
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(username));
						
			boolean actusername              = username.isDisplayed();
			boolean actpassword              = password.isDisplayed();
			boolean actcompanyDropDownList   = companyDropDownList.isDisplayed();
			boolean actsignIn                = signIn.isDisplayed();
			
			// Expected
			boolean expusername              = true;
			boolean exppassword              = true;
			boolean expcompanyDropDownList   = true;
			boolean expsignIn                = true;
			
			System.out.println("************************************* checkSignOut  *********************************");
			System.out.println("username             : "+actusername             +" Value Expected : "+expusername);
			System.out.println("password             : "+actpassword             +" Value Expected : "+exppassword);
			System.out.println("companyDropDownList  : "+actcompanyDropDownList  +" Value Expected : "+expcompanyDropDownList);
			System.out.println("signIn               : "+actsignIn               +" Value Expected : "+expsignIn);

			
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(username));
			username.click();
			 
			username.sendKeys("UserAllOptionsST");
			getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
			 
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(password));
			password.sendKeys("12345");
			 
			getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
			  
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(signIn));
			signIn.click();
		
			Thread.sleep(5000);
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
			String actUserInfo=userNameDisplay.getText();
			
			System.out.println("User Info : "+actUserInfo);
			
			System.out.println("User Info Capture Text :"+userNameDisplay.getText());
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
			
			companyLogo.click();
			
			String actCompanyTxt=companyName.getText();
			String actLoginCompanyName=actCompanyTxt.substring(0, 19);
			System.out.println("company name :"+ actLoginCompanyName);
			companyLogo.click();
						    	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
			String getDashboard=labelDashboard.getText();
				
			//boolean getnoOfAlertsDisplay=noOfAlertsDisplay.isDisplayed();
					
			// Expected
			String expuserInfo="UserAllOptionsST";
			String expLoginCompanyName="Automation Company ";
			String expDashboard="Dashboard";
			//boolean expnoOfAlertsDisplay=true;
			
			System.out.println("************************************* checkAlertsToSignIAsAuthorizeUser  *********************************");
			System.out.println("userNameDisplay    : "+actUserInfo              +" Value Expected : "+expuserInfo);
			System.out.println("LoginCompanyName   : "+actLoginCompanyName      +" Value Expected : "+expLoginCompanyName);
			System.out.println("labelDashboard     : "+getDashboard             +" Value Expected : "+expDashboard);
			//System.out.println("noOfAlertsDisplay  : "+getnoOfAlertsDisplay     +" Value Expected : "+expnoOfAlertsDisplay);
					
			boolean actMethod=actUserInfo.equalsIgnoreCase(expuserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)
					&& getDashboard.equalsIgnoreCase(expDashboard) ;
			
			String actResult=Boolean.toString(actMethod);
			String expResult=excelReader.getCellData("PurchaseVoucherN", 254, 7);
			
			excelReader.setCellData(xlfile, "PurchaseVoucherN", 254, 8, actResult);
					
			if(actResult.equalsIgnoreCase(expResult))
			{
				System.out.println("Test Pass : Transaction Authorization is Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 251, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Transaction Authorization is Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 251, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
		{
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 251, 10, e.getMessage());
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
			return false;
		}
	  }*/
 
	 
	 
	 
	  public static boolean checkOpenPurchasesOrderVoucherToDisplayPendingAuthorizationInAuthorizeHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
	  	excelReader=new ExcelReader(POJOUtility.getExcelPath());
	  	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	  	
	  	/*try
	  	{*/
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
	  		inventoryMenu.click();
	  		
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
	  		invTransactionsMenu.click();
	  		
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
	  		inventoryTransactionsPurchasesMenu.click();
	  	
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
	  		purchasesOrdersVoucher.click();
	  		
	  		
			boolean actVoucherAllAuthorizationAndLinks =voucherAll.isDisplayed() && pendingAuthorizationProcess.isDisplayed()
					                                      && pendingMaterialRequistion.isDisplayed();
			
			boolean expVoucherAllAuthorizationAndLinks = true;
			
			System.out.println("Home page All Pending Authorization And Links Value Actual : " + actVoucherAllAuthorizationAndLinks  +   " Value Expected :  " + expVoucherAllAuthorizationAndLinks);
			
			pendingAuthorizationProcess.click();			
			
            Thread.sleep(2000);
			
			//Home Page Column Headings
			
	        int headerTxtCount		= headerTxtList.size();
			
			ArrayList<String> headerTxt = new ArrayList<String>();
			
			for (int i = 2; i < headerTxtCount; i++) 
			{
				String data=headerTxtList.get(i).getText();
				headerTxt.add(data);
			}
			
			String actHeaderTxt = headerTxt.toString();
			String expHeaderTxt = excelReader.getCellData(xlSheetName, 433, 7);
			excelReader.setCellData(xlfile, xlSheetName, 433, 8, actHeaderTxt);
			
			System.out.println("actopeningStocksNewHomePageColumnsButtons  : " + actHeaderTxt);
			System.out.println("expopeningStocksNewHomePageColumnsButtons  : " + expHeaderTxt);
			
			System.out.println("Count Of Header Columns in Home Page Value Actual    :  " + actHeaderTxt );
			System.out.println("Count Of Header Columns in Home Page Value Expected  :  " + expHeaderTxt );
					
			String actHomePageVoucherNo     = openingStocksNewHomeRow1VoucherNo.getText();
			String actHomePageCreatedBy     = openingStocksNewHomeRow1CreatedBy.getText();
			String actHomePageModifiedBy    = openingStocksNewHomeRow1ModifiedBy.getText();
			String actHomePageSuspendBy     = openingStocksNewHomeRow1SuspendStatus.getText();
			String actHomePageAuthorizedBy  = openingStocksNewHomeRow1AuthorizationStatus.getText();
			
			String expHomePageVoucherNo     = "1";
			String expHomePageCreatedBy     = "SU";
			String expHomePageModifiedBy    = "SU";
			String expHomePageSuspendBy     = "False";
			String expHomePageAuthorizeBy   = "Pending";
					
			System.out.println("Voucher No in Home Page Value Actual   :  " + actHomePageVoucherNo+    " Value Expected : " + expHomePageVoucherNo);
			System.out.println("Created By in Home Page Value Actual   :  " + actHomePageCreatedBy+    " Value Expected : " + expHomePageCreatedBy);
			System.out.println("Modified By in Home Page Value Actual  :  " + actHomePageModifiedBy+   " Value Expected : " + expHomePageModifiedBy);
			System.out.println("Suspend By in Home Page Value Actual   :  " + actHomePageSuspendBy+    " Value Expected : " + expHomePageSuspendBy);
			System.out.println("Authorize By in Home Page Value Actual :  " + actHomePageAuthorizedBy+ " Value Expected : " + expHomePageAuthorizeBy);
					
			if( actVoucherAllAuthorizationAndLinks==expVoucherAllAuthorizationAndLinks && actHomePageVoucherNo.equalsIgnoreCase(expHomePageVoucherNo) && actHomePageCreatedBy.equalsIgnoreCase(expHomePageCreatedBy)
					&& actHomePageModifiedBy.equalsIgnoreCase(expHomePageModifiedBy) && actHomePageSuspendBy.equalsIgnoreCase(expHomePageSuspendBy)
					&& actHomePageAuthorizedBy.equalsIgnoreCase(expHomePageAuthorizeBy) )
			{
				System.out.println("Test Pass : HomeScreen Options Are Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 18, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : HomeScreen Options Are NOT Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 18, 9, resFail);
				return false;
			}
		}
	  	/*catch (Exception e) 
	    {
	  		String exception =e.getMessage().substring(0, 50);
	          excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	          System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	          return false;
	     }
	   }*/
	  
	  
	  
	  @FindBy(xpath="//div[@id='id_transaction_homescreen_authorize']//div[@class='toolbar_button_image']")
	  private static WebElement authorizeBtnInHomePage;
	  
	  public static boolean checkOpenPurchasesOrderVoucherAuthorizeFromHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
	  	excelReader=new ExcelReader(POJOUtility.getExcelPath());
	  	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	  	
	  	/*try
	  	{*/
	  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
	  		grid_ChkBox1.click();
	  		
			boolean actVoucherAllAuthorizationAndLinks =voucherAll.isDisplayed() && pendingAuthorizationProcess.isDisplayed()
					                                      && pendingMaterialRequistion.isDisplayed();
			
			boolean expVoucherAllAuthorizationAndLinks = true;
			
			System.out.println("Home page All Pending Authorization And Links Value Actual : " + actVoucherAllAuthorizationAndLinks  +   " Value Expected :  " + expVoucherAllAuthorizationAndLinks);
			
			authorizeBtnInHomePage.click();			
            
            Thread.sleep(2000);
            
            voucherAll.click();
            
            Thread.sleep(2000);
			
			//Home Page Column Headings
			
	        int headerTxtCount		= headerTxtList.size();
			
			ArrayList<String> headerTxt = new ArrayList<String>();
			
			for (int i = 2; i < headerTxtCount; i++) 
			{
				String data=headerTxtList.get(i).getText();
				headerTxt.add(data);
			}
			
			String actHeaderTxt = headerTxt.toString();
			String expHeaderTxt = excelReader.getCellData(xlSheetName, 433, 7);
			excelReader.setCellData(xlfile, xlSheetName, 433, 8, actHeaderTxt);
			
			System.out.println("actopeningStocksNewHomePageColumnsButtons  : " + actHeaderTxt);
			System.out.println("expopeningStocksNewHomePageColumnsButtons  : " + expHeaderTxt);
			
			System.out.println("Count Of Header Columns in Home Page Value Actual    :  " + actHeaderTxt );
			System.out.println("Count Of Header Columns in Home Page Value Expected  :  " + expHeaderTxt );
					
			String actHomePageVoucherNo     = openingStocksNewHomeRow1VoucherNo.getText();
			String actHomePageCreatedBy     = openingStocksNewHomeRow1CreatedBy.getText();
			String actHomePageModifiedBy    = openingStocksNewHomeRow1ModifiedBy.getText();
			String actHomePageSuspendBy     = openingStocksNewHomeRow1SuspendStatus.getText();
			String actHomePageAuthorizedBy  = openingStocksNewHomeRow1AuthorizationStatus.getText();
			
			String expHomePageVoucherNo     = "1";
			String expHomePageCreatedBy     = "SU";
			String expHomePageModifiedBy    = "SU";
			String expHomePageSuspendBy     = "False";
			String expHomePageAuthorizeBy   = "Authorized";
					
			System.out.println("Voucher No in Home Page Value Actual   :  " + actHomePageVoucherNo+    " Value Expected : " + expHomePageVoucherNo);
			System.out.println("Created By in Home Page Value Actual   :  " + actHomePageCreatedBy+    " Value Expected : " + expHomePageCreatedBy);
			System.out.println("Modified By in Home Page Value Actual  :  " + actHomePageModifiedBy+   " Value Expected : " + expHomePageModifiedBy);
			System.out.println("Suspend By in Home Page Value Actual   :  " + actHomePageSuspendBy+    " Value Expected : " + expHomePageSuspendBy);
			System.out.println("Authorize By in Home Page Value Actual :  " + actHomePageAuthorizedBy+ " Value Expected : " + expHomePageAuthorizeBy);
		
             //Close Voucher Home Page
			 getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(voucherhomeCloseBtn));
	    	 voucherhomeCloseBtn.click();
			 Thread.sleep(2000);
			
			if(  actVoucherAllAuthorizationAndLinks==expVoucherAllAuthorizationAndLinks && actHomePageVoucherNo.equalsIgnoreCase(expHomePageVoucherNo) && actHomePageCreatedBy.equalsIgnoreCase(expHomePageCreatedBy)
					&& actHomePageModifiedBy.equalsIgnoreCase(expHomePageModifiedBy) && actHomePageSuspendBy.equalsIgnoreCase(expHomePageSuspendBy)
					&& actHomePageAuthorizedBy.equalsIgnoreCase(expHomePageAuthorizeBy) )
			{
				System.out.println("Test Pass : HomeScreen Options Are Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 18, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : HomeScreen Options Are NOT Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 18, 9, resFail);
				return false;
			}
		}
	  	/*catch (Exception e) 
	    {
	  		String exception =e.getMessage().substring(0, 50);
	          excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	          System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	          return false;
	     }
	   }*/
	  
	  
	  public boolean checLogoutAndLoginCompanyToUnAuthorizeUserToCheckPVPendingLinks() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

	    /*try
		{*/				
		    getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
		    userNameDisplay.click();
		
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		    logoutOption.click();
		 		
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(username));
			username.click();
			 
			username.sendKeys("su");
			getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
			 
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(password));
			password.sendKeys("su");
			 
			getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
			  
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(signIn));
			signIn.click();
		
			Thread.sleep(5000);
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
			String actUserInfo=userNameDisplay.getText();
			
			System.out.println("User Info : "+actUserInfo);
			
			System.out.println("User Info Capture Text :"+userNameDisplay.getText());
			
			
					
			// Expected
			String expuserInfo="su";
			
			
			System.out.println("************************************* checkAlertsToSignIAsAuthorizeUser  *********************************");
			System.out.println("userNameDisplay    : "+actUserInfo              +" Value Expected : "+expuserInfo);
			
					
			
					
			if(actUserInfo.equalsIgnoreCase(expuserInfo))
			{
				System.out.println("Test Pass : Transaction Authorization is Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 251, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Transaction Authorization is Displayed");
				excelReader.setCellData(xlfile, "PurchaseVoucherN", 251, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
		{
			String exception =e.getMessage();
			excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 251, 10, e.getMessage());
			System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
			return false;
		}
	  }*/
 

	  
	  
	  
	  
	  
	  public SmokeBudgetVouchersMRPOPage(WebDriver driver)
	   {
	 	  PageFactory.initElements(driver, this);	
	   } 
	 		
	    
	    

}
