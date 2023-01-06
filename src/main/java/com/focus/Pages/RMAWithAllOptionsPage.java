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

import org.apache.commons.lang3.text.translate.AggregateTranslator;
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
	
	public class RMAWithAllOptionsPage extends BaseEngine
	{
		
		private static String xlfile = getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		private static String resPass="Pass";
		private static String resFail="Fail";
		private static ExcelReader excelReader;
		private static String xlSheetName = "RMAOptionsPage";
		
	 
		///// Options Btn  (Display Option)	
	   @FindBy(xpath="//span[@id='transOptions']")
	   private static WebElement  optionsBtn;
	   
	   @FindBy(xpath="//a[contains(text(),'Auto Adjust Width')]")
	   private static WebElement  autoAdjustWidthBtn;
	   
	   @FindBy(xpath="//a[contains(text(),'Fit To Screen')]")
	   private static WebElement  fitToScreenBtn;
	   
	   @FindBy(xpath="//a[contains(text(),'Normal')]")
	   private static WebElement  normalBtn;
	   
	
		
	   // Grid Elements
				
	
					
		//Grid Header CheckBox
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
		
		@FindBy (xpath="//tbody[@id='tblBodyTransRender']/tr/td[2]/input")
		private static List<WebElement> grid_CheckBoxList;
		
		@FindBy (xpath="//tbody[@id='tblBodyTransRender']/tr/td[4]")
		private static List<WebElement> grid_VoucherNoList;
		
		@FindBy (xpath="//tbody[@id='tblBodyTransRender']/tr/td[11]")
		private static List<WebElement> grid_SuspendStatusList;
		
		
					
		////////////////
		
		// Footer Section
		
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
		
		
		////////////////
		
		//New Screen Elements
		
		
	
		// Vouchers_New Header Section
		
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
		
	
		@FindBy(xpath="//*[@id='id_transactionentry_print']")
		private static WebElement  new_PrintBtn;
		
		@FindBy(xpath="//*[@id='id_transactionentry_suspend']")
		private static WebElement  new_SuspendBtn;
		
		@FindBy(xpath="//*[@id='id_transactionentry_close']")
		private static WebElement  new_CloseBtn;
	
		@FindBy(xpath="//div[@id='id_transactionentry_delete']//span[contains(@class,'icon-delete icon-font6')]")
		private static WebElement  new_DeleteBtn;
		
		@FindBy(xpath="//div[@id='id_transaction_homescreen_Delete']")
		private static WebElement salesReturnsVoucherHomeDeleteBtn;
		
		
		
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
		private static WebElement  bomInputOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_autoallocate']")
		private static WebElement  autoAllocateOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_reverseentry']")
		private static WebElement  reverseEntryOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_printbarcode']")
		private static WebElement  printBarcodeOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_postingdetails']")
		private static WebElement  postingDetailsOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_exporttoxml']")
		private static WebElement  exportToXMLOption;
		
		@FindBy(xpath="//*[@id='id_transactionentry_settings']")
		private static WebElement  settingsOption;
		
		
		////// Header Section Completed
		
		// Main Section
		@FindBy(xpath="//input[@id='id_header_1']")
		private static WebElement  documentNumberTxt;
		
		
		
		@FindBy(xpath="//*[@id='id_header_1_input_image']/span")
		private static WebElement  documentNumberDropdown_ExpansionBtn;
		
		@FindBy(xpath="//input[@id='id_header_2']")
		private static WebElement  dateTxt;
		
		@FindBy(xpath="//td[@id='id_header_2_input_image']//span[@class='icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow']")
		private static WebElement  dateTxt_CalenderBtn;
		
		@FindBy(xpath="//tr[@id='id_header_2_day_today']//span[@class='theme_color-inverse'][contains(text(),'Today')]")
		private static WebElement  calender_TodayBtn;
		
		
		@FindBy(xpath="//input[@id='id_header_4']")
		private static WebElement  cashBankAccountTxt;
		
		@FindBy(xpath="//*[@id='id_header_4_input_image']/span")
		private static WebElement  cashBankAccount_ExpansionBtn;
		
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
						
		@FindBy(xpath="//input[@id='id_header_268435459']")
		private static WebElement  departmentTxt;
		
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
								private static WebElement  vendorAccount_Settings_StandardFieldsBtn_AlignmentDropdown;
																	
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
		
		@FindBy(xpath="//input[@id='id_body_536870916']")
		private static WebElement  pvWareHouseTxt;
		
		@FindBy(xpath="//tbody[@id='id_body_23_table_data_body']/tr")
	    private static List<WebElement> itemListCount;

		@FindBy(xpath="//tbody[@id='id_body_536870916_table_data_body']/tr")
	    private static List<WebElement> pvwareHouseListCount;
		
		@FindBy(xpath="//*[@id='id_header_4_input_image']/span")
		private static WebElement  customerAccount_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='id_header_4_input_settings']/span")
		private static WebElement  customerAccount_SettingsBtn;
		
					@FindBy(xpath="//div[@id='id_header_4_customize_popup_container']")
					private static WebElement  customerAccount_Settings_Container;
											
					@FindBy(xpath="//div[@id='id_header_4_customize_popup_footer']//input[1]")
					private static WebElement  customerAccount_Settings_StandardFieldsBtn;
											
								@FindBy(xpath="//select[@id='id_header_4_customize_popup_standardfields_list']")
								private static WebElement  customerAccount_Settings_StandardFieldsBtn_FieldDrpdwn;
																	
								@FindBy(xpath="//input[@id='id_header_4_customize_popup_standardfields_header']")
								private static WebElement  customerAccount_Settings_StandardFieldsBtn_HeaderTxt;
																
								@FindBy(xpath="//select[@id='id_header_4_customize_popup_standardfields_alignment']")
								private static WebElement  customerAccount_Settings_StandardFieldsBtn_AlignmentDrpdwn;
																	
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
								private static WebElement  account_Settings_StandardFieldsBtn_FieldDrpdwn;
																	
								@FindBy(xpath="//input[@id='id_header_4_customize_popup_standardfields_header']")
								private static WebElement  account_Settings_StandardFieldsBtn_HeaderTxt;
																
								@FindBy(xpath="//select[@id='id_header_4_customize_popup_standardfields_alignment']")
								private static WebElement  account_Settings_StandardFieldsBtn_Alignmentdrpdwn;
																	
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
		private static WebElement  warehouseTxt;
		
		@FindBy(xpath="//*[@id='id_header_268435460_input_image']/span")
		private static WebElement  warehouse_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='id_header_268435460_input_settings']/span")
		private static WebElement  warehouse_SettingsBtn;
		
					@FindBy(xpath="//div[@id='id_header_268435460_customize_popup_container']")
					private static WebElement  warehouse_Settings_Container;
											
					@FindBy(xpath="//div[@id='id_header_268435460_customize_popup_footer']//input[1]")
					private static WebElement  warehouse_Settings_StandardFieldsBtn;
											
								@FindBy(xpath="//select[@id='id_header_268435460_customize_popup_standardfields_list']")
								private static WebElement  warehouse_Settings_StandardFieldsBtn_Fielddropdown;
																	
								@FindBy(xpath="//input[@id='id_header_268435460_customize_popup_standardfields_header']")
								private static WebElement  warehouse_Settings_StandardFieldsBtn_HeaderTxt;
																
								@FindBy(xpath="//select[@id='id_header_268435460_customize_popup_standardfields_alignment']")
								private static WebElement  warehouse_Settings_StandardFieldsBtn_Alignmentdropdown;
																	
								@FindBy(xpath="//input[@id='id_header_268435460_customize_popup_standardfields_width']")
								private static WebElement  warehouse_Settings_StandardFieldsBtn_WidthTxt;
																	
								@FindBy(xpath="//div[@class='panel']//input[3]")
								private static WebElement  warehouse_Settings_StandardFieldsBtn_OkBtn;
																	
								@FindBy(xpath="//div[@class='panel']//input[4]")
								private static WebElement  warehouse_Settings_StandardFieldsBtn_CancelBtn;
											
					@FindBy(xpath="//div[@id='id_header_268435460_search_container']//input[2]")
					private static WebElement  warehouse_Settings_DeleteColumnBtn;
											
					@FindBy(xpath="//div[@class='panel']//input[3]")
					private static WebElement  warehouse_Settings_OkBtn;
									
					@FindBy(xpath="//div[@class='panel']//input[4]")
					private static WebElement  warehouse_Settings_CancelBtn;
		
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
		private static WebElement  appropriateBasedOnDrpdwn;
		
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
		private static WebElement  salesQuotations_NarrationTxt;
	
		@FindBy(xpath="//input[@id='id_header_67108887']")
		private static WebElement  deliveryNotes_NarrationTxt;
	
		@FindBy(xpath="//input[@id='id_header_67108917']")
		private static WebElement  posSales_NarrationTxt;
	
		@FindBy(xpath="//input[@id='id_header_67108918']")
		private static WebElement  productionOrder_NarrationTxt;
	
		@FindBy(xpath="//input[@id='id_header_67108886']")
		private static WebElement  salesOrders_NarrationTxt;
	
		@FindBy(xpath="//input[@id='id_header_67108895']")
		private static WebElement  shortageInStock_NarrationTxt;
	
		@FindBy(xpath="//input[@id='id_header_67108896']")
		private static WebElement  excessesInStock_NarrationTxt;
	
		@FindBy(xpath="//input[@id='id_header_67108897']")
		private static WebElement  stockTransfers_NarrationTxt;
	
		@FindBy(xpath="//input[@id='id_header_67108898']")
		private static WebElement  openingStocks_NarrationTxt;
		
		@FindBy(xpath="//input[@id='id_header_67108923']")
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
	
		
		
		@FindBy(xpath="//input[@id='id_body_536870916']")
		private static WebElement  pvWarehouseTxt;
		
		
		
		@FindBy(xpath="//td[@id='id_body_536870916_input_image']//span[@class='icon-down-arrow optioncontrol_arrow_margin']")
		private static WebElement  pvWarehouse_ExpansionBtn;
		
		@FindBy(xpath="//td[@id='id_body_536870916_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement  pvWarehouse_SettingsBtn;
		
					@FindBy(xpath="//div[@id='id_body_536870916_customize_popup_container']")
					private static WebElement  pvWarehouse_Settings_Container;
											
					@FindBy(xpath="//div[@id='id_body_536870916_customize_popup_footer']//input[1]")
					private static WebElement  pvWarehouse_Settings_StandardFieldsBtn;
											
								@FindBy(xpath="//select[@id='id_body_536870916_customize_popup_standardfields_list']")
								private static WebElement  pvWarehouse_Settings_StandardFieldsBtn_Fielddropdown;
																	
								@FindBy(xpath="//input[@id='id_body_536870916_customize_popup_standardfields_header']")
								private static WebElement  pvWarehouse_Settings_StandardFieldsBtn_HeaderTxt;
																
								@FindBy(xpath="//select[@id='id_body_536870916_customize_popup_standardfields_alignment']")
								private static WebElement  pvWarehouse_Settings_StandardFieldsBtn_Alignmentdropdown;
																	
								@FindBy(xpath="//input[@id='id_body_536870916_customize_popup_standardfields_width']")
								private static WebElement  pvWarehouse_Settings_StandardFieldsBtn_WidthTxt;
																	
								@FindBy(xpath="//div[@class='panel']//input[3]")
								private static WebElement  pvWarehouse_Settings_StandardFieldsBtn_OkBtn;
																	
								@FindBy(xpath="//div[@class='panel']//input[4]")
								private static WebElement  pvWarehouse_Settings_StandardFieldsBtn_CancelBtn;
											
					@FindBy(xpath="//div[@id='id_body_536870916_search_container']//input[2]")
					private static WebElement  pvWarehouse_Settings_DeleteColumnBtn;
											
					@FindBy(xpath="//div[@class='panel']//input[3]")
					private static WebElement  pvWarehouse_Settings_OkBtn;
									
					@FindBy(xpath="//div[@class='panel']//input[4]")
					private static WebElement  pvWarehouse_Settings_CancelBtn;
		
		// Grid Elements
		
						
		@FindBy(xpath="//th[@id='id_transaction_entry_detail_table_control_heading_1']")
		private static WebElement  grid_Header_Account;
		
		@FindBy(xpath="//th[@id='id_transaction_entry_detail_table_control_heading_1']")
		private static WebElement  grid_Header_Item;
		
		@FindBy(xpath="//th[@id='id_transaction_entry_detail_table_control_heading_1']")
		private static WebElement  grid_Header_DebitAC;
		
		@FindBy(xpath="//th[@id='id_transaction_entry_detail_table_control_heading_2']")
		private static WebElement  Grid_Header_CreditAC;
		
		//
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[1]")
		private static WebElement  firstRowIndex;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[1]")
		private static WebElement  secondRowIndex;
		
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
	
		//
		
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
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[20]")
		private static WebElement  select1stRow_19thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[21]")
		private static WebElement  select1stRow_20thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[22]")
		private static WebElement  select1stRow_21stColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[23]")
		private static WebElement  select1stRow_22ndColumn;
		
		// Batch Popup Elements
		
		@FindBy(xpath="//label[@id='div_Ok']")
		private static WebElement batchOkIcon;

				@FindBy(xpath="//label[@id='lblitem']")
				private static WebElement batchItemName;
				
				@FindBy(xpath="//label[@id='lblTotalQty']")
				private static WebElement batchTotalQtyAmt;
				
				@FindBy(xpath="//label[@id='lblAdjustQty']")
				private static WebElement batchAdjustQtyAmt;
				
				
				@FindBy(xpath="//label[@id='lblBalnceQty']")
				private static WebElement batchBalanceAmt;
				
				@FindBy(xpath="//div[contains(@class,'col-xs-12 col-sm-12 col-md-12 col-lg-12 P_ingBottom')]//div[2]//input[1]")
				private static WebElement batchExipiringInRadioBtn;
				             
				@FindBy(xpath="//section[@id='page_Content']//div[@class='row']//div//div[3]//input[1]")
				private static WebElement batchExipiringAfetrRadioBtn;
				
				@FindBy(xpath="//input[@id='txtWithin']")
				private static WebElement batchWithInTxt;
				
				@FindBy(xpath="//*[@id='Id_MRPBatch_Grid_control_heading_6']")
				private static WebElement batchTableHeaderRateColumn;
				
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_1-1']")
				private static WebElement batchTable1stRow1stCol;
				
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_1-2']")
				private static WebElement batchTable1stRow2ndCol;
				
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_1-5']")
				private static WebElement batchTable1stRow3rdCol;
				
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_1-6']")
				private static WebElement batchTable1stRow4thCol;
				               
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_1-8']")
				private static WebElement batchTable1stRow5thCol;
				
				@FindBy(xpath="//input[@id='MRPBatch_AdjBatchQty']")
				private static WebElement enter_BatchTableQtyAdjusted;
				
				
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_1-13']")
				private static WebElement batchTable1stRow6thCol;
				
				
				@FindBy(xpath="//*[@id='Id_MRPBatch_Grid_col_2-1']")
				private static WebElement batchTable2ndRow1stCol;
				
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_2-2']")
				private static WebElement batchTable2ndRow2ndCol;
				
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_2-5']")
				private static WebElement batchTable2ndRow3rdCol;
				
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_2-6']")
				private static WebElement batchTable2ndRow4thCol;
				
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_2-7']")
				private static WebElement batchTable2ndRow5thCol;
				
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_2-13']")
				private static WebElement batchTable2ndRow6thCol;
				
				
				@FindBy(xpath="//*[@id='Id_MRPBatch_Grid_col_3-1']")
				private static WebElement batchTable3rdRow1stCol;
				
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_3-2']")
				private static WebElement batchTable3rdRow2ndCol;
				
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_3-5']")
				private static WebElement batchTable3rdRow3rdCol;
				
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_3-6']")
				private static WebElement batchTable3rdRow4thCol;
				
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_3-7']")
				private static WebElement batchTable3rdRow5thCol;
				
				@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_3-13']")
				private static WebElement batchTable3rdRow6thCol;
				
				
		
		@FindBy(xpath="//tbody[@id='Id_MRPBatch_Grid_body']/tr/td[2]")
		private static List<WebElement> batchPopupBatchNoList;
		
		@FindBy(xpath="//tbody[@id='Id_MRPBatch_Grid_body']/tr/td[1]")
		private static List<WebElement> batchPopupRowNosList;
		
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[1]")
		private static WebElement selectSerialNo2;
		
		@FindBy (xpath="//div[contains(text(),'Delete Row')]")
		private static WebElement selectSerialNoDeleteRow; 
		
		@FindBy (xpath="//div[contains(text(),'Insert Row')]")
		private static WebElement selectSerialNoInsertRow;
		
		@FindBy (xpath="//div[contains(text(),'Select Row')]")
		private static WebElement selectSerialNoSelectRow;
		
		@FindBy (xpath="//div[contains(text(),'Duplicate Row')]")
		private static WebElement selectSerialNoDuplicateRow;
		
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
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[11]")
		private static WebElement  select2ndRow_10thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[12]")
		private static WebElement  select2ndRow_11thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[13]")
		private static WebElement  select2ndRow_12thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[14]")
		private static WebElement  select2ndRow_13thColumn;
		
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
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[21]")
		private static WebElement  select2ndRow_20thColumn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[22]")
		private static WebElement  select2ndRow_21stColumn;
		
		
		
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
		
		
		
		
		
		@FindBy(xpath="//input[@id='MasterGroup__101']")
		private static WebElement  subLedgerAccountTxt;
		
		@FindBy(xpath="//span[@class='optioncontrol_arrow_margin icon-down-arrow']")
		private static WebElement  subLedgerAccount_ExpansionBtn;
		
		@FindBy(xpath="//td[@id='MasterGroup__101_input_settings']//span[@class='icon-settings optioncontrol_settings_margin']")
		private static WebElement  subLedgerAccount_SettingsBtn;
		
					@FindBy(xpath="//div[@id='MasterGroup__101_customize_popup_container']")
					private static WebElement  subLedgerAccount_Settings_Container;
											
					@FindBy(xpath="//div[@id='MasterGroup__101_customize_popup_footer']//input[1]")
					private static WebElement  subLedgerAccount_Settings_StandardFieldsBtn;
											
								@FindBy(xpath="//select[@id='MasterGroup__101_customize_popup_standardfields_list']")
								private static WebElement  subLedgerAccount_Settings_StandardFieldsBtn_Fielddropdown;
																	
								@FindBy(xpath="//input[@id='MasterGroup__101_customize_popup_standardfields_header']")
								private static WebElement  subLedgerAccount_Settings_StandardFieldsBtn_HeaderTxt;
																
								@FindBy(xpath="//select[@id='MasterGroup__101_customize_popup_standardfields_alignment']")
								private static WebElement  subLedgerAccount_Settings_StandardFieldsBtn_Alignmentdropdown;
																	
								@FindBy(xpath="//input[@id='MasterGroup__101_customize_popup_standardfields_width']")
								private static WebElement  subLedgerAccount_Settings_StandardFieldsBtn_WidthTxt;
																	
								@FindBy(xpath="//section[@id='reportView']//input[3]")
								private static WebElement  subLedgerAccount_Settings_StandardFieldsBtn_OkBtn;
																	
								@FindBy(xpath="//section[@id='reportView']//input[4]")
								private static WebElement  subLedgerAccount_Settings_StandardFieldsBtn_CancelBtn;
											
											
					@FindBy(xpath="//div[@id='MasterGroup__101_search_container']//input[2]")
					private static WebElement  subLedgerAccount_Settings_DeleteColumnBtn;
											
					@FindBy(xpath="//section[@id='reportView']//input[3]")
					private static WebElement  subLedgerAccount_Settings_OkBtn;
									
					@FindBy(xpath="//section[@id='reportView']//input[4]")
					private static WebElement  subLedgerAccount_Settings_CancelBtn;
		
		
		
		
		
		
		//
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
								private static WebElement  enter_Account_Settings_StandardFieldsBtn_FieldDrpdwn;
																	
								@FindBy(xpath="//input[@id='id_header_12_customize_popup_standardfields_header']")
								private static WebElement  enter_Account_Settings_StandardFieldsBtn_HeaderTxt;
																
								@FindBy(xpath="//select[@id='id_header_12_customize_popup_standardfields_alignment']")
								private static WebElement  enter_Account_Settings_StandardFieldsBtn_AlignmentDrpdwn;
																	
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
								private static WebElement  item_Settings_StandardFieldsBtn_FieldDrpdwn;
																	
								@FindBy(xpath="//input[@id='id_body_23_customize_popup_standardfields_header']")
								private static WebElement  item_Settings_StandardFieldsBtn_HeaderTxt;
																
								@FindBy(xpath="//select[@id='id_body_23_customize_popup_standardfields_alignment']")
								private static WebElement  item_Settings_StandardFieldsBtn_AlignmentDrpdwn;
																	
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
	
		
					
	
		@FindBy(xpath="//input[@id='id_body_12']")
		private static WebElement  enter_SalesAccount;
					
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
								private static WebElement  unit_Settings_StandardFieldsBtn_FieldDrpdwn;
																	
								@FindBy(xpath="//input[@id='id_body_24_customize_popup_standardfields_header']")
								private static WebElement  unit_Settings_StandardFieldsBtn_HeaderTxt;
																
								@FindBy(xpath="//select[@id='id_body_24_customize_popup_standardfields_alignment']")
								private static WebElement  unit_Settings_StandardFieldsBtn_AlignmentDrpdwn;
																	
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
	
					
					
		@FindBy(xpath="//input[@id='id_body_16777328']")
		private static WebElement  enter_TaxCode;
					
		@FindBy(xpath="//input[@id='id_body_26']")
		private static WebElement  enter_Quantity;
		
		@FindBy(xpath="//input[@id='id_body_41']")
		private static WebElement enter_AlternateQuantity;
		
		@FindBy(xpath="//input[@id='id_body_27']")
		private static WebElement  enter_Rate;
		
		@FindBy(xpath="//input[@id='id_body_28']")
		private static WebElement  enter_Gross;
		
		@FindBy (xpath="//input[@id='id_body_38']")
		private static WebElement enter_RMA;
		
		
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
								private static WebElement  enter_DebitAC_Settings_StandardFieldsBtn_FieldDrpdwn;
																	
								@FindBy(xpath="//input[@id='id_header_12_customize_popup_standardfields_header']")
								private static WebElement  enter_DebitAC_Settings_StandardFieldsBtn_HeaderTxt;
																
								@FindBy(xpath="//select[@id='id_header_12_customize_popup_standardfields_alignment']")
								private static WebElement  enter_DebitAC_Settings_StandardFieldsBtn_AlignmentDrpdwn;
																	
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
		
		
		////////////
		
		// Right Panel*** InfoSide Bar
					
	
		@FindBy(xpath="//*[@id='id_transactionentry_infopanel_container']/div[1]/div[2]/span")
		private static WebElement  infoSideBarCustomizeBtn;
		
					@FindBy(xpath="//*[@id='id_transactionentry_infopanel_container']/div[1]/div[2]/span")
					private static WebElement  customizeCloseBtn;
					
					@FindBy(xpath="//div[@id='id_transactionentry_infopanel_customize_container']//li[@id='Dash_Save']")
					private static WebElement  customizeSaveBtn;
					
					@FindBy(xpath="//div[@id='id_transactionentry_infopanel_customize_container']//li[@id='Dash_Close']")
					private static WebElement  customizeCancelBtn;
					
					@FindBy(xpath="//input[@id='searchBoxTrans']")
					private static WebElement  customizeSearchTxt;
								
					@FindBy(xpath="//*[@id='Dashboard_Graph_panelID_Trans']/div[1]/i")
					private static WebElement  customizeSearchBtn;
								
					@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//a[@id='GraphOption']")
					private static WebElement  graphBtn;
								
								@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//li[@id='Newgraph']")
								private static WebElement  newGraphOption;
								
					@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//a[@id='ReportOption']")
					private static WebElement  reportBtn;
								
								@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//li[@id='NewReport']")
								private static WebElement  newReportOption;
								
					@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//a[@id='InfopanelOption']")
					private static WebElement  infoPanelBtn;
								
								@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//li[@id='NewInfopanel']")
								private static WebElement  newInfoPanelOption;
								
					@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//a[@id='workflow']")
					private static WebElement  workFlowBtn;
								
								@FindBy(xpath="//div[@id='Dashboard_Graph_panelID_Trans']//li[@id='NewWorkflow']")
								private static WebElement  newWorkFlowOption;
								
					@FindBy(xpath="//a[@id='DocumentInfoOption']")
					private static WebElement  docInfoBtn;
								
								@FindBy(xpath="//li[@id='id_Dashlet0']//a")
								private static WebElement  newDocInfoOption;
		
		@FindBy(xpath="//span[@class='icon-left-and-right-panel-icon icon-font6 no_padding_left_right']")
		private static WebElement  infoSideBarExpandBtn;
		
		
		
		
		
		
		// Footer Section
		
	
		@FindBy(xpath="//label[@id='id_transactionentry_footer_panel_summary_value_net']")
		private static WebElement  netLabel;
		
	
		@FindBy(xpath="//*[@id='id_transactionentry_footer_panel_summary_value_net']/span[2]")
		private static WebElement  net_Amount;
	
		
	/////////////////////////////////////////////////////		
		
		
	    @FindBy(xpath="//*[@id='81']/span")
	    private static WebElement  financialsReportsMenu; 
	    
		    @FindBy(xpath="//a[@id='500']//span[contains(text(),'Ledger')]")
			private static WebElement  ledger;
	
			@FindBy(xpath="//*[@id='658']/span")
			private static WebElement  ledgerDetail;
	
			@FindBy(xpath="//*[@id='501']/span")
			private static WebElement  subLedger;
			
			@FindBy(xpath="//*[@id='82']/span")
			private static WebElement  cashAndBankBooksMenu;
			
				@FindBy(xpath="//*[@id='510']/span")
				private static WebElement  cashBookReport;
				
				@FindBy(xpath="//*[@id='511']/span")
				private static WebElement  bankBookReport;
				
				@FindBy(xpath="//*[@id='512']/span")
				private static WebElement  pettyCashBokReport;
				
				@FindBy(xpath="//*[@id='513']/span")
				private static WebElement  dayBookReport;
				
				@FindBy(xpath="//*[@id='83']/span")
				private static WebElement  bankReconciliationReport;
				
				@FindBy(xpath="//*[@id='84']/span")
				private static WebElement  bankReconciliationImport;
				
				@FindBy(xpath="//*[@id='85']/span")
				private static WebElement  customerVendorReconciliation;
			
				@FindBy(xpath="//*[@id='537']/span")
				private static WebElement  bankReconciliationStatement;
					
				@FindBy(xpath="//*[@id='162']/span")
				private static WebElement chequeDiscountingMenu;
			
			@FindBy(xpath="//*[@id='502']/span")
			private static WebElement  salesReportsMenu;
			
				@FindBy(xpath="//*[@id='503']/span")
				private static WebElement  salesRegisterReport;
	
				@FindBy(xpath="//*[@id='504']/span")
				private static WebElement  salesReturnRegisterReport;
				
				@FindBy(xpath="//*[@id='505']/span")
				private static WebElement summarySalesBookReport;
				
				@FindBy(xpath="//*[@id='628']/span")
				private static WebElement monthlySalesBookReport;
				
				@FindBy(xpath="//*[@id='629']/span")
				private static WebElement  topCustomersListReport;
			
			@FindBy(xpath="//*[@id='506']/span")
			private static WebElement  purchaseReportsMenu;
			
				@FindBy(xpath="//*[@id='507']/span")
				private static WebElement  purchaseRegisterReport;
				
				@FindBy(xpath="//*[@id='508']/span")
				private static WebElement  purchaseReturnRegisteReport;
				
				@FindBy(xpath="//*[@id='509']/span")
				private static WebElement  summaryPurchaseBookReport;
						
			@FindBy(xpath="//*[@id='514']/span")
			private static WebElement  registersReportMenu;
			
				@FindBy(xpath="//*[@id='516']/span")
				private static WebElement  openingBalanceRegisterReport;
				
				@FindBy(xpath="//*[@id='515']/span")
				private static WebElement journalEntriesRegisterReport;
				
				@FindBy(xpath="//span[contains(text(),'Credit Note Register')]")
				private static WebElement creditNoteRegisterReport;
				
				@FindBy(xpath="//*[@id='610']/span")
				private static WebElement  debitNoteRegisterReport;
				
				@FindBy(xpath="//*[@id='598']/span")
				private static WebElement  receiptsRegisterReport;
				
				@FindBy(xpath="//*[@id='599']/span")
				private static WebElement  paymentRegisterReport;
				
				@FindBy(xpath="//*[@id='518']/span")
				private static WebElement  pdcReceiptsRegisterReport;
				
				@FindBy(xpath="//*[@id='519']/span")
				private static WebElement pdcPaymentsRegisterReport;
				
				@FindBy(xpath="//*[@id='517']/span")
				private static WebElement  entryJournalRegisterMenu;
						
			@FindBy(xpath="//*[@id='523']/span")
			private static WebElement  salesAndPurchasesReportMenu;
			
				@FindBy(xpath="//*[@id='524']/span")
				private static WebElement  salesGroupedByCustomerReport;
				
				@FindBy(xpath="//*[@id='525']/span")
				private static WebElement  salesGroupedByProductReport;
				
				@FindBy(xpath="//*[@id='526']/span")
				private static WebElement  salesGroupedByDepartmentReport;
				
				@FindBy(xpath="//*[@id='527']/span")
				private static WebElement  purchasesGroupedByVendorReport;
				
				@FindBy(xpath="//*[@id='528']/span")
				private static WebElement  purchasesGroupedByProductReport;
				
				@FindBy(xpath="//*[@id='529']/span")
				private static WebElement  purchasesGroupedByDepartmentReport;		
			
			@FindBy(xpath="//*[@id='568']/span")
			private static WebElement  MiscellaneousAnalysisReportMenu;
			
				@FindBy(xpath="//*[@id='569']/span")
				private static WebElement  peakAndLowBalancesReport;
				
				@FindBy(xpath="//*[@id='570']/span")
				private static WebElement  comprativeAnalysisReport;
				
				@FindBy(xpath="//*[@id='571']/span")
				private static WebElement  transactionsTypeAnalysisReport;
				
				@FindBy(xpath="//*[@id='572']/span")
				private static WebElement  abcAnalysisReport;
				
				@FindBy(xpath="//*[@id='597']/span")
				private static WebElement  interestCalculationReport;
	
			@FindBy(xpath="//*[@id='536']/span")
			private static WebElement  mastersInformationMenu;
			
		     	@FindBy(xpath="//*[@id='538']/span")
				private static WebElement  masterInfoMenu;
				
				@FindBy(xpath="//*[@id='144']/span")
				private static WebElement  productLabelMenu;
				
				@FindBy(xpath="//*[@id='573']/span")
				private static WebElement  unitConversionMenu;
					
				@FindBy(xpath="//*[@id='147']/span")
				private static WebElement  skidDefinitionMenu;
								
			@FindBy(xpath="//*[@id='520']/span")
			private static WebElement  auditTrailMenu;
			
				@FindBy(xpath="//*[@id='521']/span")
				private static WebElement  auditTrailtransactionsReport;
				
				@FindBy(xpath="//*[@id='522']/span")
				private static WebElement  aAuditTrailloginsReport;
			
			@FindBy(xpath="//*[@id='204']/span")
			private static WebElement  accountQuery;
		
		
		
		@FindBy (xpath="//select[@id='RITCombobox__1']")
		private static WebElement includePDCDrpDwn;
		
		@FindBy (xpath="//select[@id='RITCombobox__1']")
		private static WebElement sl_StockTypeDropdown;
		
		@FindBy (xpath="//select[@id='RITCombobox__1']")
		private static WebElement stockTypeOrValidation;
	
		@FindBy (xpath="//select[@id='RITCombobox__10']")
		private static WebElement includeNonAuthorizeDataDrpDwn;
		
		@FindBy (xpath="//select[@id='RITCombobox__5']")
		private static WebElement sortingDrpDwn;
		
		@FindBy (xpath="//input[@id='RITCheckbox__2']")
		private static WebElement movedAccountOnlyChkBox;
		
		@FindBy (xpath="//input[@id='RITCheckbox__3']")
		private static WebElement printAccountIndexChkBox;
		
		@FindBy (xpath="//input[@id='RITCheckbox__9']")
		private static WebElement freshAccountForEachPageChkBox;
		
		@FindBy (xpath="//input[@id='RITCheckbox__11']")
		private static WebElement printAsStatmentOfAccountChkBox;
		
		@FindBy (xpath="//input[@id='RITCheckbox__6']")
		private static WebElement displayUnrealizedLossOrGainChkBox;
	
	

		//Dash board Page
		@FindBy(xpath="//*[@id='dashName']")
	    private static WebElement labelDashboard ;
		
		@FindBy(xpath="//*[@id='Select_dash']")
		private static WebElement selectDashboard ;
			
		@FindBy(xpath="//*[@id='Dashboard_AddDash']")
	    private static WebElement newAddDashBoard;
		 
	    @FindBy(xpath="//*[@id='Dashboard_Dash_Config']")
	    private static WebElement dashboardCustomizationSettings;
		
	    
	    /////////////////////
	    
	    //Settings Screen
	  
		/*@FindBy(xpath="//span[@id='updateButton']")
		private static WebElement  updateBtn;
		
		@FindBy(xpath="//i[@class='icon-close icon-font6']")
		private static WebElement  settings_closeBtn;
		*/
		
		//Documents Tab
	
	
		@FindBy(xpath="//div[@class='font-5 theme_background-color-inverse theme_color-inverse']")
		private static WebElement  documentsTab;
		
		@FindBy(xpath="//li[@id='navigationtab1']")
		private static WebElement  documentsTabForMaterialReceiptNotes;
		
	
		@FindBy(xpath="//span[@id='DocumentLoadMasterPopupBtn']")
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
		
		@FindBy(xpath="//span[@id='DocumentAddGroupPopupBtn']")
		private static WebElement  addGroupBtn;
		
					@FindBy(xpath="//input[@id='doc_GroupName']")
					private static WebElement  groupNameTXt;
					
					@FindBy(xpath="//button[contains(text(),'Add')]")
					private static WebElement  addBtn;
					
					@FindBy(xpath="//div[@class='col-xs-12 form-group']//button[@id='btnCancel']")
					private static WebElement  cancelBtn;
	
		@FindBy(xpath="//div[@id='voucherTabContent']//span[2]")
		private static WebElement  deletegroupBtn;
		
		@FindBy(xpath="//input[@id='doc_title']")
		private static WebElement  titleTxt;
		
		
		/*@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/input[1]")
		private static WebElement  titleTxt;*/
		
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
		
		
		// Documents Grid Elements
		
			// first row 
		
		@FindBy(xpath="//td[@id='doc_TagsTable_col_1-0']")
		private static WebElement  firstrow;
		
		@FindBy(xpath="//div[contains(text(),'Delete Row')]")
		private static WebElement  Doc_deleterowBtn;
		
		@FindBy(xpath="//div[contains(text(),'Insert Row')]")
		private static WebElement  Doc_insertRowBtn;
		
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
					
		/*@FindBy(xpath="//*[@id='id_menu_tree_60']/a/i")
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
					private static WebElement  onlinepaymentsBtn;*/
					
					
	
							@FindBy(xpath="//*[@id='2001']")
							private static WebElement  cashAndBankMenu; 
							
					     		@FindBy(xpath="//*[@id='2002']/span")
								private static WebElement  receiptsVoucher;
								
								@FindBy(xpath="//*[@id='2003']/span")
								private static WebElement  paymentsVoucher;
								
								@FindBy(xpath="//*[@id='2004']/span")
								private static WebElement  pettyVoucher;
								
								@FindBy(xpath="//*[@id='2005']/span")
								private static WebElement  postDatedReceiptsVoucher;
											
								@FindBy(xpath="//*[@id='2006']/span")
								private static WebElement  postDatedPaymentsVoucher;
								
								@FindBy(xpath="//*[@id='108']/span")
								private static WebElement  defineChequeSeries;
								
								@FindBy(xpath="//*[@id='109']/span")
								private static WebElement  cancelCheque;
							
	
							@FindBy(xpath="//*[@id='2007']/span")
							private static WebElement  financialsTransactionsPurchaseMenu; 
							
								@FindBy(xpath="//*[@id='2008']/span")
								private static WebElement  purchaseVoucher;
								
								@FindBy(xpath="//*[@id='2009']/span")
								private static WebElement  purchaseReturnsVoucher;
								
								@FindBy(xpath="//*[@id='2050']/span")
								private static WebElement  purchaseVoucherN;
								
										
							@FindBy(xpath="//*[@id='2014']/span")
							private static WebElement  financialsTransactionsSalesMenu;
								
								@FindBy(xpath="//*[@id='2015']/span")
								private static WebElement  salesInvoicesVoucher;
								
								@FindBy(xpath="//*[@id='2016']/span")
								private static WebElement  salesReturnsVoucher;
								
								@FindBy(xpath="//*[@id='2020']/span")
								private static WebElement  cashSales;
										
								@FindBy(xpath="//*[@id='2021']/span")
								private static WebElement  hirePurchaseSales;
							
							@FindBy(xpath="//*[@id='2023']/span")
							private static WebElement  financialsTransactionsJournalsMenu;
							
							    @FindBy(xpath="//*[@id='2032']/span")
								private static WebElement  journalsEntriesVoucher;
								
								@FindBy(xpath="//*[@id='2025']/span")
								private static WebElement  forexJVVoucher;
								
								@FindBy(xpath="//*[@id='2029']/span")
								private static WebElement  interdepartmentaljvVoucher;
								
								@FindBy(xpath="//*[@id='2024']/span")
								private static WebElement  oldJouranlEntriesVoucher;
								
								@FindBy(xpath="//*[@id='2026']/span")
								private static WebElement  debitNotesVoucher;
								
								@FindBy(xpath="//*[@id='2030']/span")
								private static WebElement  debitNotesLinewiseVoucher;
								
								@FindBy(xpath="//*[@id='2027']/span")
								private static WebElement  creditNotesVoucher;
								
								@FindBy(xpath="//*[@id='2031']/span")
								private static WebElement  creditNotesLinewiseVoucher;
								
								@FindBy(xpath="//*[@id='2028']/span")
								private static WebElement  openingBalancesVoucher;
								
								@FindBy(xpath="//*[@id='2046']/span")
								private static WebElement  fixedAssetsDepreciationVoucher;
							
							@FindBy(xpath="//*[@id='219']/span")
							private static WebElement  financialsTransactionsAutoPostingsMenu;
							
								@FindBy(xpath="//*[@id='77']/span")
								private static WebElement  recurringJournalVoucher;
								
								@FindBy(xpath="//*[@id='79']/span")
								private static WebElement  postRecurringJournal;
								
								@FindBy(xpath="//*[@id='92']/span")
								private static WebElement  recurringJournalTemplate;
								
								@FindBy(xpath="//*[@id='91']/span")
								private static WebElement  postInterest;
								
								@FindBy(xpath="//*[@id='90']/span")
								private static WebElement  convertMaturedPDCs;
								
								@FindBy(xpath="//*[@id='161']/span")
								private static WebElement assignPDCLimit;
								
								@FindBy(xpath="//*[@id='166']/span")
								private static WebElement  stockReplenishment;
	
						@FindBy(xpath="//*[@id='62']/span")
						private static WebElement  financialsCreditManagementMenu; 
						
							@FindBy(xpath="//*[@id='63']/span")
							private static WebElement  paymentTerms;
							
							@FindBy(xpath="//*[@id='64']/span")
							private static WebElement  financeTerms;
							
							@FindBy(xpath="//*[@id='65']/span")
							private static WebElement  reminderTerms;		
							
							@FindBy(xpath="//*[@id='89']/span")
							private static WebElement  sendReminder;
							
							@FindBy(xpath="//*[@id='217']/span")
							private static WebElement  delinquency;
							
								@FindBy(xpath="//*[@id='73']/span")
								private static WebElement  questionnarie;
								
								@FindBy(xpath="//*[@id='218']/span")
								private static WebElement  delinquencyMgmt;
								
							@FindBy(xpath="//*[@id='35']/span")
							private static WebElement  creidtApproval;
					
				
					
					
	//Inventory   *************************
					
		@FindBy(xpath="//*[@id='id_menu_tree_135']/a/i")
		private static WebElement  inventoryExpandBtn;
		
					
					
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
			
		
		
		
	////////////////////
		
		@FindBy(xpath="//*[@id='135']")
		private static WebElement  inventoryMenu; 
		
			@FindBy(xpath="//a[@id='137']//span[contains(text(),'Transactions')]")
			private static WebElement  invTransactionsMenu; 
				
				@FindBy(xpath="//*[@id='139']/span")
				private static WebElement  invTransPurchasesMenu; 
			
					@FindBy(xpath="//*[@id='2010']/span")
					private static WebElement  requestForQuoteVoucher;
					
					@FindBy(xpath="//*[@id='2011']/span")
					private static WebElement  purchasesQuotationsVoucher;
						
					@FindBy(xpath="//*[@id='2012']/span")
					private static WebElement  purchasesOrdersVoucher;
				
					@FindBy(xpath="//*[@id='2013']/span")
					private static WebElement  materialReceiptNotesVoucher;
				
				@FindBy(xpath="//*[@id='140']/span")
				private static WebElement  invTransSalesMenu; 
			
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
						
				
					
					
				
				@FindBy(xpath="//*[@id='164']/span")
				private static WebElement  holdAndUnholdStockVoucher;
				
				@FindBy(xpath="//*[@id='93']/span")
				private static WebElement  stockReconciliation;
				
				@FindBy(xpath="//*[@id='99']/span")
				private static WebElement  stockAllocation;
				
			@FindBy(xpath="//a[@id='200']//span[contains(text(),'Reports')]")
			private static WebElement  inventoryReportsMenu; 
			
				@FindBy(xpath="//a[@id='202']//span[contains(text(),'Item Query')]")
				private static WebElement itemQuery;
				
				@FindBy(xpath="//span[contains(text(),'Stock Ledger')]")
				private static WebElement  stockLedger;
				
				@FindBy(xpath="//span[contains(text(),'Opening Stocks Register')]")
				private static WebElement  openingStocksRegister;
				
				@FindBy(xpath="//span[contains(text(),'Stock Statement')]")
				private static WebElement  stockStatement;
							   
				@FindBy(xpath="//a[@id='575']//span[contains(text(),'Stock movement')]")
				private static WebElement  stockMovement;
				
				@FindBy(xpath="//span[contains(text(),'Multi Level Stock Movement')]")
				private static WebElement  multiLevelStockMovement;
				
				@FindBy(xpath="//span[contains(text(),'Virtual Stock Analysis')]")
				private static WebElement  virtualStockAnalysis;
				
				@FindBy(xpath="//a[@id='578']")
				private static WebElement  stockValuation;
				
				@FindBy(xpath="//a[@id='656']")
				private static WebElement  stockReportByTag;
				
				@FindBy(xpath="//a[@id='579']//span[contains(text(),'ABC Analysis')]")
				private static WebElement  abcAnalysis;
				
				@FindBy(xpath="//span[contains(text(),'Stock Analysis by Batch')]")
				private static WebElement  stockAnalysisByBatch;
				
				@FindBy(xpath="//span[contains(text(),'Stock Analysis by RMA')]")
				private static WebElement  stockAnalysisByRMA;
				
				@FindBy(xpath="//span[contains(text(),'Bins Report')]")
				private static WebElement  binsReportMenu;
				
					@FindBy(xpath="//*[@id='587']/span")
					private static WebElement  stockDetailsByBins;
				
					@FindBy(xpath="//*[@id='588']/span")
					private static WebElement  stockBalancesByBins;
					
					@FindBy(xpath="//*[@id='589']/span")
					private static WebElement  stockBalancesByItemByBins;
					
					@FindBy(xpath="//*[@id='590']/span")
					private static WebElement  expiredStockByBins;		
				
				@FindBy(xpath="//span[contains(text(),'Stock balance by warehouse')]")
				private static WebElement  stockBalanceByWarehouse;
				
				@FindBy(xpath="//span[contains(text(),'Stock ageing analysis')]")
				private static WebElement  stockAgeingAnalysisMenu;
					
					@FindBy(xpath="//*[@id='582']/span")
					private static WebElement  ageingAnalysis;
					
					@FindBy(xpath="//*[@id='622']/span")
					private static WebElement  ageingAnalysisByBatch;
					
					@FindBy(xpath="//*[@id='623']/span")
					private static WebElement  ageingAnalysisByRMA;		
					
					@FindBy(xpath="//*[@id='591']/span")
					private static WebElement  ageingStockByBins;
							
					@FindBy(xpath="//*[@id='646']/span")
					private static WebElement  ageingByItemByBins;		
	
				@FindBy(xpath="//span[contains(text(),'Reorder Report')]")
				private static WebElement reorderReport;
				
				@FindBy(xpath="//a[@id='584']//span[contains(text(),'Fast Moving Item')]")
				private static WebElement  fastMovingItem;
						
				@FindBy(xpath="//a[@id='585']//span[contains(text(),'Slow Moving Item')]")
				private static WebElement  slowMovingItem;
				
				@FindBy(xpath="//span[contains(text(),'Peak/Low Balances')]")
				private static WebElement  peakORLowBalances;
				
				@FindBy(xpath="//span[contains(text(),'Best Selling Item')]")
				private static WebElement  bestSellingItem;
			
				@FindBy(xpath="//span[contains(text(),'Raise Indent Item with Low Stock')]")
				private static WebElement  raiseIndentItemWithLowStock;
				
				@FindBy(xpath="//span[contains(text(),'Stock transfer report')]")
				private static WebElement  stockTransferReport;
				
				@FindBy(xpath="//span[contains(text(),'Raise request for quote')]")
				private static WebElement  raiseRequestForQuote;
				
				@FindBy(xpath="//span[contains(text(),'Stock reservation report')]")
				private static WebElement  stockReservationReport;
				
				@FindBy(xpath="//span[contains(text(),'Transaction type-wise stock report')]")
				private static WebElement  transactionTypeWiseStockReport;
			
			@FindBy(xpath="//span[contains(text(),'Order Management')]")
			private static WebElement  orderManagementMenu; 
		
				@FindBy(xpath="//*[@id='600']/span")
				private static WebElement  analysisOfLinkedORUnlinkedDocumentsMenu;
						
					@FindBy(xpath="//*[@id='601']/span")
					private static WebElement  listingOfDocuments;
					
					@FindBy(xpath="//*[@id='602']/span")
					private static WebElement listingOfDocumentsByDueDate;
							
					@FindBy(xpath="//*[@id='603']/span")
					private static WebElement  ageingOfPendingDocuments;
					
					@FindBy(xpath="//*[@id='604']/span")
					private static WebElement  ageingofPendingDocumentsByDueDate;
					
					@FindBy(xpath="//*[@id='624']/span")
					private static WebElement  linkChainAnalysis;
					
					@FindBy(xpath="//*[@id='657']/span")
					private static WebElement  pendingVoucherReport;
				
				@FindBy(xpath="//*[@id='201']/span")
				private static WebElement  quotationAnalysis;
				
				@FindBy(xpath="//*[@id='555']/span")
				private static WebElement  purchasesQuotation;
				
				@FindBy(xpath="//*[@id='556']/span")
				private static WebElement  salesQuotation;
				
				@FindBy(xpath="//*[@id='574']/span")
				private static WebElement  lowestQuotation;
				
				@FindBy(xpath="//*[@id='553']/span")
				private static WebElement  purchasesOrders;		
				
				@FindBy(xpath="//*[@id='554']/span")
				private static WebElement  salesOrders;
				
		///////////////////////////////////////////////////////////////////
				
		
		// Item Query Elements
				
		@FindBy (xpath="//div[@id='Dv_panelHeader']/nav/div/div[1]/a/div/span[2]")
		private static WebElement itemQueryLabel;
	
		@FindBy(xpath="//input[@id='productQuery']")
		private static WebElement IQ_ItemTxt;
		
		@FindBy(xpath="//button[contains(text(),'GetStock')]")
		private static WebElement IQ_getStockBtn;
		
		@FindBy(xpath="//button[contains(text(),'Customize')]")
		private static WebElement IQ_customizeBtn;
		
		@FindBy(xpath="//span[@id='id_productquery_button_clear']")
		private static WebElement IQ_clearBtn;
		
		@FindBy(xpath="//span[@id='id_productquery_button_close']")
		private static WebElement IQ_closeBtn;
		
		@FindBy(xpath="//label[contains(text(),'Opening Stock Quantity:')]")
		private static WebElement IQ_OpeningStockQtyLable;
		
		@FindBy(xpath="//label[contains(text(),'Pending purchases orders :')]")
		private static WebElement IQ_PendingPurchaseOrdersLable;
		
		@FindBy(xpath="//label[contains(text(),'Current Stock :')]")
		private static WebElement IQ_CurrentStockLable;
		
		@FindBy(xpath="//label[contains(text(),'Pending Sales Orders :')]")
		private static WebElement IQ_PendingSalesOrdersLable;
		
		@FindBy(xpath="//label[contains(text(),'Avg. Stock Rate :')]")
		private static WebElement IQ_AvgStockRateLable;
		
		@FindBy(xpath="//label[contains(text(),'Qty to be Ordered :')]")
		private static WebElement IQ_QtytobeOrderedLable;
		
		@FindBy(xpath="//label[contains(text(),'Base Unit :')]")
		private static WebElement IQ_BaseUnitLable;
		
		@FindBy(xpath="//label[contains(text(),'Value :')]")
		private static WebElement IQ_ValueLable;
		
		@FindBy(xpath="//table[@id='purchaseTable']//th[contains(text(),'Date')]")
		private static WebElement IQ_PurDate;
		
		@FindBy(xpath="//table[@id='purchaseTable']//th[contains(text(),'Voucher No')]")
		private static WebElement IQ_PurVoucherNo;
		
		@FindBy(xpath="//th[contains(text(),'Vendor')]")
		private static WebElement IQ_PurVendor;
		
		@FindBy(xpath="//table[@id='purchaseTable']//th[contains(text(),'Quantity')]")
		private static WebElement IQ_PurQty;
		
		@FindBy(xpath="//table[@id='purchaseTable']//th[contains(text(),'Rate')]")
		private static WebElement IQ_PurRate;
		
		@FindBy(xpath="//table[@id='purchaseTable']//th[contains(text(),'Currency')]")
		private static WebElement IQ_PurCurrency;
		
		@FindBy(xpath="//table[@id='salesTable']//th[contains(text(),'Date')]")
		private static WebElement IQ_SaleDate;
		
		@FindBy(xpath="//table[@id='salesTable']//th[contains(text(),'Voucher No')]")
		private static WebElement IQ_SaleVoucherNo;
		
		@FindBy(xpath="//th[contains(text(),'Customer')]")
		private static WebElement IQ_SaleCustomer;
		
		@FindBy(xpath="//table[@id='salesTable']//th[contains(text(),'Quantity')]")
		private static WebElement IQ_SaleQty;
		
		@FindBy(xpath="//table[@id='salesTable']//th[contains(text(),'Rate')]")
		private static WebElement IQ_SaleRate;
		
		@FindBy(xpath="//table[@id='salesTable']//th[contains(text(),'Currency')]")
		private static WebElement IQ_SaleCurrency;
		
		@FindBy(xpath="//span[@id='openingStock']")
		private static WebElement IQ_OpeningStockQuantity;
		
		@FindBy(xpath="//span[@id='currentStock']")
		private static WebElement IQ_CurrentStock;
		
		@FindBy(xpath="//span[@id='avgStockRate']")
		private static WebElement IQ_AvgStockRate;
		
		@FindBy(xpath="//span[@id='value']")
		private static WebElement IQ_Value;
		
		@FindBy(xpath="//span[@id='iBaseUnit']")
		private static WebElement IQ_BaseUnit;
		
		@FindBy(xpath="//span[@id='purchaseOrdered']")
		private static WebElement IQ_PendingPurchaseOrders;
		
		@FindBy(xpath="//span[@id='salesOrdered']")
		private static WebElement IQ_PendingSalesOrders;
		
		@FindBy(xpath="//span[@id='quantityToBeOrdered']")
		private static WebElement IQ_QtyToBeOrdered;
		
		
		// Purchase Table
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[1]")
		private static WebElement Pur1strow_1stColumn;
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[2]")
		private static WebElement Pur1strow_2ndColumn;
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[3]")
		private static WebElement Pur1strow_3rdColumn;
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[4]")
		private static WebElement Pur1strow_4thColumn;
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[5]")
		private static WebElement Pur1strow_5thColumn;
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[6]")
		private static WebElement Pur1strow_6thColumn;
		
		
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[1]")
		private static WebElement Pur2ndrow_1stColumn;
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[2]")
		private static WebElement Pur2ndrow_2ndColumn;
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[3]")
		private static WebElement Pur2ndrow_3rdColumn;
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[4]")
		private static WebElement Pur2ndrow_4thColumn;
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[5]")
		private static WebElement Pur2ndrow_5thColumn;
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[6]")
		private static WebElement Pur2ndrow_6thColumn;
		
		
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[1]")
		private static WebElement Pur3rdrow_1stColumn;
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[2]")
		private static WebElement Pur3rdrow_2ndColumn;
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[3]")
		private static WebElement Pur3rdrow_3rdColumn;
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[4]")
		private static WebElement Pur3rdrow_4thColumn;
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[5]")
		private static WebElement Pur3rdrow_5thColumn;
		
		@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[6]")
		private static WebElement Pur3rdrow_6thColumn;
		
		
		
		@FindBy(xpath="//td[contains(text(),'Quantity Purchased Till Date:')]")
		private static WebElement QtyPurchasedTillDate;
		
		@FindBy(xpath="//div[@class='form-group col-xs-12']//tr[4]//td[2]")
		private static WebElement PurTotal;
				
				
		
		
		// Stock Ledger Elements 
		
		@FindBy(xpath="//div[@id='spnhiererchy']//select[@id='iTreeId']")
		private static WebElement sl_Dropdown1;
		
		@FindBy(xpath="//div[@id='spnhiererchy']//select[@id='iTreeViewId']")
		private static WebElement sl_Dropdown2;
		
		@FindBy(xpath="//label[@id='oncheckaccorderan']")
		private static WebElement sl_oncheckBtn;
		
		@FindBy(xpath="//input[@id='txtsrch-term']")
		private static WebElement sl_SearchItemTxt;
		
		@FindBy(xpath="//i[contains(@class,'icon-search theme_color-inverse')]")
		private static WebElement sl_SearchBtn;
		
		@FindBy(xpath="//input[@id='cmbUserTypeMaster']")
		private static WebElement sl_MasterTypeTxt;
		
		@FindBy(xpath="//span[contains(@class,'icon-font6 icon-external-module theme_icon-color')]")
		private static WebElement sl_AutoAdjustColumnsBtn;
		
		/*@FindBy(xpath="//a[contains(@class,'clsMasterName')][contains(text(),'ITEMS GROUP')]")
		private static WebElement sl_SearchItemTxt;*/
		
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
		
		@FindBy(xpath="//input[@id='id_starting_date_']")
		private static WebElement sl_StartDateTxt;
		
		@FindBy(xpath="//td[@id='id_starting_date__input_image']//span[contains(@class,'icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow')]")
		private static WebElement sl_StartDateCalanderBtn;
		
		@FindBy(xpath="//input[@id='id_ending_date_']")
		private static WebElement sl_EndDateTxt;
		
		@FindBy(xpath="//td[@id='id_ending_date__input_image']//span[contains(@class,'icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow')]")
		private static WebElement sl_EndDateCalanderBtn;
		
		@FindBy(xpath="//*[@id='MasterSingle__101']")
		private static WebElement sl_WarehouseTxt;
		
		@FindBy(xpath="//*[@id='RITCheckbox__6']")
		private static WebElement sl_IgnoreInternalTransferForAgeingChkbox;
		
		@FindBy(xpath="//*[@id='RITCombobox__2']")
		private static WebElement sl_ReportingLevelDropdown;
		
		@FindBy(xpath="//*[@id='RITCheckbox__4']")
		private static WebElement sl_DisplayLinearReportChkbox;
		
		@FindBy(xpath="//*[@id='RITCombobox__1']")
		private static WebElement sl_StockValuationDropdown;
		
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
		
		@FindBy(xpath="//tr[1]//td[8]//div[1]//label[1]//input")
		private static WebElement sl_1stRowChkBox;
		
		@FindBy(xpath="//tr[2]//td[8]//div[1]//label[1]//input")
		private static WebElement sl_2ndRowChkBox;
		
		@FindBy(xpath="//tr[3]//td[8]//div[1]//label[1]//input")
		private static WebElement sl_3rdRowChkBox;
		
		
		
		////////
		
		@FindBy(xpath="//span[@id='reportSort']")
		private static WebElement report_sortingBtn;
				
		@FindBy(xpath="//span[@id='reportRefresh']")
		private static WebElement report_RefreshBtn;
				
		@FindBy(xpath="//span[@id='print_report_']")
		private static WebElement report_PrintBtn;
				
		@FindBy(xpath="//span[text()='Export']")
		private static WebElement report_ExportBtn;
				
		@FindBy(xpath="//span[@id='GraphIcon']")
		private static WebElement report_GraphBtn;
				
		@FindBy(xpath="//span[@id='BackTrackIcon']")
		private static WebElement report_BackTrackBtn;
				
		@FindBy(xpath="//span[@id='reportCustomize']")
		private static WebElement report_ReportCustomizeBtn;
				
		@FindBy(xpath="//span[@id='filterIcon']")
		private static WebElement report_FilterBtn;
				
		@FindBy(xpath="//div[@id='dvfilter']//span[contains(@class,'icon-down-arrow icon-font8 dropdown-toggle theme_button_color')]")
		private static WebElement report_FilterDownBtn;
				
		@FindBy(xpath="//span[@id='analyzeIcon']")
		private static WebElement report_AnalyzeBtn;
				
		@FindBy(xpath="//span[@id='CrossrefrenceIcon']")
		private static WebElement report_CrossReferenceBtn;
				
		@FindBy(xpath="//span[contains(text(),'Options')]")
		private static WebElement report_OptionsBtn;
				
		@FindBy(xpath="//div[@id='dvoptions']//span[contains(@class,'icon-down-arrow icon-font8 dropdown-toggle theme_button_color')]")
		private static WebElement report_OptionsDownBtn;
				
		@FindBy(xpath="//span[@id='reportClose_']")
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
		
		
		@FindBy(xpath="//*[@id='trRender_1']/td[2]")
		private static WebElement sl_1stRow1stCol;
				
		@FindBy(xpath="//*[@id='trRender_1']/td[3]")
		private static WebElement sl_1stRow2ndCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[4]")
		private static WebElement sl_1stRow3rdCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[5]")
		private static WebElement sl_1stRow4thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[6]")
		private static WebElement sl_1stRow5thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[7]")
		private static WebElement sl_1stRow6thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[8]")
		private static WebElement sl_1stRow7thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[9]")
		private static WebElement sl_1stRow8thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[10]")
		private static WebElement sl_1stRow9thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[11]")
		private static WebElement sl_1stRow10thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[12]")
		private static WebElement sl_1stRow11thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[13]")
		private static WebElement sl_1stRow12thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[14]")
		private static WebElement sl_1stRow13thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[15]")
		private static WebElement sl_1stRow14thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[16]")
		private static WebElement sl_1stRow15thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[17]")
		private static WebElement sl_1stRow16thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[18]")
		private static WebElement sl_1stRow17thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[19]")
		private static WebElement sl_1stRow18thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[20]")
		private static WebElement sl_1stRow19thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[21]")
		private static WebElement sl_1stRow20thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[22]")
		private static WebElement sl_1stRow21thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[23]")
		private static WebElement sl_1stRow22thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[24]")
		private static WebElement sl_1stRow23thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[25]")
		private static WebElement sl_1stRow24thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[26]")
		private static WebElement sl_1stRow25thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[27]")
		private static WebElement sl_1stRow26thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[28]")
		private static WebElement sl_1stRow27thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[29]")
		private static WebElement sl_1stRow28thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[30]")
		private static WebElement sl_1stRow29thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[31]")
		private static WebElement sl_1stRow30thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[32]")
		private static WebElement sl_1stRow31thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[33]")
		private static WebElement sl_1stRow32thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[34]")
		private static WebElement sl_1stRow33thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[35]")
		private static WebElement sl_1stRow34thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[36]")
		private static WebElement sl_1stRow35thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[37]")
		private static WebElement sl_1stRow36thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[38]")
		private static WebElement sl_1stRow37thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[39]")
		private static WebElement sl_1stRow38thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[40]")
		private static WebElement sl_1stRow39thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[41]")
		private static WebElement sl_1stRow40thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[42]")
		private static WebElement sl_1stRow41thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[43]")
		private static WebElement sl_1stRow42thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[44]")
		private static WebElement sl_1stRow43thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[45]")
		private static WebElement sl_1stRow44thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[46]")
		private static WebElement sl_1stRow45thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[47]")
		private static WebElement sl_1stRow46thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[48]")
		private static WebElement sl_1stRow47thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[49]")
		private static WebElement sl_1stRow48thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[50]")
		private static WebElement sl_1stRow49thCol;
		
		@FindBy(xpath="//*[@id='trRender_1']/td[51]")
		private static WebElement sl_1stRow50thCol;
		
		
		@FindBy(xpath="//*[@id='trRender_2']/td[2]")
		private static WebElement sl_2ndRow1stCol;
				
		@FindBy(xpath="//*[@id='trRender_2']/td[3]")
		private static WebElement sl_2ndRow2ndCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[4]")
		private static WebElement sl_2ndRow3rdCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[5]")
		private static WebElement sl_2ndRow4thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[6]")
		private static WebElement sl_2ndRow5thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[7]")
		private static WebElement sl_2ndRow6thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[8]")
		private static WebElement sl_2ndRow7thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[9]")
		private static WebElement sl_2ndRow8thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[10]")
		private static WebElement sl_2ndRow9thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[11]")
		private static WebElement sl_2ndRow10thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[12]")
		private static WebElement sl_2ndRow11thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[13]")
		private static WebElement sl_2ndRow12thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[14]")
		private static WebElement sl_2ndRow13thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[15]")
		private static WebElement sl_2ndRow14thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[16]")
		private static WebElement sl_2ndRow15thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[17]")
		private static WebElement sl_2ndRow16thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[18]")
		private static WebElement sl_2ndRow17thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[19]")
		private static WebElement sl_2ndRow18thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[20]")
		private static WebElement sl_2ndRow19thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[21]")
		private static WebElement sl_2ndRow20thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[22]")
		private static WebElement sl_2ndRow21thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[23]")
		private static WebElement sl_2ndRow22thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[24]")
		private static WebElement sl_2ndRow23thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[25]")
		private static WebElement sl_2ndRow24thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[26]")
		private static WebElement sl_2ndRow25thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[27]")
		private static WebElement sl_2ndRow26thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[28]")
		private static WebElement sl_2ndRow27thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[29]")
		private static WebElement sl_2ndRow28thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[30]")
		private static WebElement sl_2ndRow29thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[31]")
		private static WebElement sl_2ndRow30thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[32]")
		private static WebElement sl_2ndRow31thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[33]")
		private static WebElement sl_2ndRow32thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[34]")
		private static WebElement sl_2ndRow33thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[35]")
		private static WebElement sl_2ndRow34thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[36]")
		private static WebElement sl_2ndRow35thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[37]")
		private static WebElement sl_2ndRow36thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[38]")
		private static WebElement sl_2ndRow37thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[39]")
		private static WebElement sl_2ndRow38thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[40]")
		private static WebElement sl_2ndRow39thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[41]")
		private static WebElement sl_2ndRow40thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[42]")
		private static WebElement sl_2ndRow41thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[43]")
		private static WebElement sl_2ndRow42thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[44]")
		private static WebElement sl_2ndRow43thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[45]")
		private static WebElement sl_2ndRow44thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[46]")
		private static WebElement sl_2ndRow45thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[47]")
		private static WebElement sl_2ndRow46thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[48]")
		private static WebElement sl_2ndRow47thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[49]")
		private static WebElement sl_2ndRow48thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[50]")
		private static WebElement sl_2ndRow49thCol;
		
		@FindBy(xpath="//*[@id='trRender_2']/td[51]")
		private static WebElement sl_2ndRow50thCol;
		
		
		
		
		
		@FindBy(xpath="//*[@id='trRender_3']/td[2]")
		private static WebElement sl_3rdRow1stCol;
				
		@FindBy(xpath="//*[@id='trRender_3']/td[3]")
		private static WebElement sl_3rdRow2ndCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[4]")
		private static WebElement sl_3rdRow3rdCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[5]")
		private static WebElement sl_3rdRow4thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[6]")
		private static WebElement sl_3rdRow5thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[7]")
		private static WebElement sl_3rdRow6thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[8]")
		private static WebElement sl_3rdRow7thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[9]")
		private static WebElement sl_3rdRow8thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[10]")
		private static WebElement sl_3rdRow9thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[11]")
		private static WebElement sl_3rdRow10thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[12]")
		private static WebElement sl_3rdRow11thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[13]")
		private static WebElement sl_3rdRow12thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[14]")
		private static WebElement sl_3rdRow13thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[15]")
		private static WebElement sl_3rdRow14thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[16]")
		private static WebElement sl_3rdRow15thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[17]")
		private static WebElement sl_3rdRow16thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[18]")
		private static WebElement sl_3rdRow17thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[19]")
		private static WebElement sl_3rdRow18thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[20]")
		private static WebElement sl_3rdRow19thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[21]")
		private static WebElement sl_3rdRow20thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[22]")
		private static WebElement sl_3rdRow21thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[23]")
		private static WebElement sl_3rdRow22thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[24]")
		private static WebElement sl_3rdRow23thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[25]")
		private static WebElement sl_3rdRow24thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[26]")
		private static WebElement sl_3rdRow25thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[27]")
		private static WebElement sl_3rdRow26thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[28]")
		private static WebElement sl_3rdRow27thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[29]")
		private static WebElement sl_3rdRow28thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[30]")
		private static WebElement sl_3rdRow29thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[31]")
		private static WebElement sl_3rdRow30thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[32]")
		private static WebElement sl_3rdRow31thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[33]")
		private static WebElement sl_3rdRow32thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[34]")
		private static WebElement sl_3rdRow33thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[35]")
		private static WebElement sl_3rdRow34thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[36]")
		private static WebElement sl_3rdRow35thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[37]")
		private static WebElement sl_3rdRow36thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[38]")
		private static WebElement sl_3rdRow37thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[39]")
		private static WebElement sl_3rdRow38thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[40]")
		private static WebElement sl_3rdRow39thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[41]")
		private static WebElement sl_3rdRow40thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[42]")
		private static WebElement sl_3rdRow41thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[43]")
		private static WebElement sl_3rdRow42thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[44]")
		private static WebElement sl_3rdRow43thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[45]")
		private static WebElement sl_3rdRow44thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[46]")
		private static WebElement sl_3rdRow45thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[47]")
		private static WebElement sl_3rdRow46thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[48]")
		private static WebElement sl_3rdRow47thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[49]")
		private static WebElement sl_3rdRow48thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[50]")
		private static WebElement sl_3rdRow49thCol;
		
		@FindBy(xpath="//*[@id='trRender_3']/td[51]")
		private static WebElement sl_3rdRow50thCol;
		
		
		
		
		
		@FindBy(xpath="//*[@id='trRender_4']/td[2]")
		private static WebElement sl_4thRow1stCol;
				
		@FindBy(xpath="//*[@id='trRender_4']/td[3]")
		private static WebElement sl_4thRow2ndCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[4]")
		private static WebElement sl_4thRow3rdCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[5]")
		private static WebElement sl_4thRow4thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[6]")
		private static WebElement sl_4thRow5thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[7]")
		private static WebElement sl_4thRow6thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[8]")
		private static WebElement sl_4thRow7thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[9]")
		private static WebElement sl_4thRow8thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[10]")
		private static WebElement sl_4thRow9thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[11]")
		private static WebElement sl_4thRow10thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[12]")
		private static WebElement sl_4thRow11thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[13]")
		private static WebElement sl_4thRow12thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[14]")
		private static WebElement sl_4thRow13thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[15]")
		private static WebElement sl_4thRow14thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[16]")
		private static WebElement sl_4thRow15thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[17]")
		private static WebElement sl_4thRow16thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[18]")
		private static WebElement sl_4thRow17thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[19]")
		private static WebElement sl_4thRow18thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[20]")
		private static WebElement sl_4thRow19thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[21]")
		private static WebElement sl_4thRow20thCol;
		

		@FindBy(xpath="//*[@id='trRender_4']/td[22]")
		private static WebElement sl_4thRow21thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[23]")
		private static WebElement sl_4thRow22thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[24]")
		private static WebElement sl_4thRow23thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[25]")
		private static WebElement sl_4thRow24thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[26]")
		private static WebElement sl_4thRow25thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[27]")
		private static WebElement sl_4thRow26thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[28]")
		private static WebElement sl_4thRow27thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[29]")
		private static WebElement sl_4thRow28thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[30]")
		private static WebElement sl_4thRow29thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[31]")
		private static WebElement sl_4thRow30thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[32]")
		private static WebElement sl_4thRow31thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[33]")
		private static WebElement sl_4thRow32thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[34]")
		private static WebElement sl_4thRow33thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[35]")
		private static WebElement sl_4thRow34thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[36]")
		private static WebElement sl_4thRow35thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[37]")
		private static WebElement sl_4thRow36thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[38]")
		private static WebElement sl_4thRow37thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[39]")
		private static WebElement sl_4thRow38thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[40]")
		private static WebElement sl_4thRow39thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[41]")
		private static WebElement sl_4thRow40thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[42]")
		private static WebElement sl_4thRow41thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[43]")
		private static WebElement sl_4thRow42thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[44]")
		private static WebElement sl_4thRow43thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[45]")
		private static WebElement sl_4thRow44thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[46]")
		private static WebElement sl_4thRow45thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[47]")
		private static WebElement sl_4thRow46thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[48]")
		private static WebElement sl_4thRow47thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[49]")
		private static WebElement sl_4thRow48thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[50]")
		private static WebElement sl_4thRow49thCol;
		
		@FindBy(xpath="//*[@id='trRender_4']/td[51]")
		private static WebElement sl_4thRow50thCol;
		
		

		
		
		
		
		
		@FindBy(xpath="//*[@id='trRender_5']/td[2]")
		private static WebElement sl_5thRow1stCol;
				
		@FindBy(xpath="//*[@id='trRender_5']/td[3]")
		private static WebElement sl_5thRow2ndCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[4]")
		private static WebElement sl_5thRow3rdCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[5]")
		private static WebElement sl_5thRow4thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[6]")
		private static WebElement sl_5thRow5thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[7]")
		private static WebElement sl_5thRow6thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[8]")
		private static WebElement sl_5thRow7thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[9]")
		private static WebElement sl_5thRow8thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[10]")
		private static WebElement sl_5thRow9thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[11]")
		private static WebElement sl_5thRow10thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[12]")
		private static WebElement sl_5thRow11thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[13]")
		private static WebElement sl_5thRow12thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[14]")
		private static WebElement sl_5thRow13thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[15]")
		private static WebElement sl_5thRow14thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[16]")
		private static WebElement sl_5thRow15thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[17]")
		private static WebElement sl_5thRow16thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[18]")
		private static WebElement sl_5thRow17thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[19]")
		private static WebElement sl_5thRow18thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[20]")
		private static WebElement sl_5thRow19thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[21]")
		private static WebElement sl_5thRow20thCol;
		

		@FindBy(xpath="//*[@id='trRender_5']/td[22]")
		private static WebElement sl_5thRow21thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[23]")
		private static WebElement sl_5thRow22thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[24]")
		private static WebElement sl_5thRow23thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[25]")
		private static WebElement sl_5thRow24thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[26]")
		private static WebElement sl_5thRow25thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[27]")
		private static WebElement sl_5thRow26thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[28]")
		private static WebElement sl_5thRow27thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[29]")
		private static WebElement sl_5thRow28thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[30]")
		private static WebElement sl_5thRow29thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[31]")
		private static WebElement sl_5thRow30thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[32]")
		private static WebElement sl_5thRow31thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[33]")
		private static WebElement sl_5thRow32thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[34]")
		private static WebElement sl_5thRow33thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[35]")
		private static WebElement sl_5thRow34thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[36]")
		private static WebElement sl_5thRow35thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[37]")
		private static WebElement sl_5thRow36thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[38]")
		private static WebElement sl_5thRow37thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[39]")
		private static WebElement sl_5thRow38thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[40]")
		private static WebElement sl_5thRow39thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[41]")
		private static WebElement sl_5thRow40thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[42]")
		private static WebElement sl_5thRow41thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[43]")
		private static WebElement sl_5thRow42thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[44]")
		private static WebElement sl_5thRow43thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[45]")
		private static WebElement sl_5thRow44thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[46]")
		private static WebElement sl_5thRow45thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[47]")
		private static WebElement sl_5thRow46thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[48]")
		private static WebElement sl_5thRow47thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[49]")
		private static WebElement sl_5thRow48thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[50]")
		private static WebElement sl_5thRow49thCol;
		
		@FindBy(xpath="//*[@id='trRender_5']/td[51]")
		private static WebElement sl_5thRow50thCol;
		
		
		
		

		
		@FindBy(xpath="//*[@id='trRender_6']/td[2]")
		private static WebElement sl_6thRow1stCol;
				
		@FindBy(xpath="//*[@id='trRender_6']/td[3]")
		private static WebElement sl_6thRow2ndCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[4]")
		private static WebElement sl_6thRow3rdCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[5]")
		private static WebElement sl_6thRow4thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[6]")
		private static WebElement sl_6thRow5thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[7]")
		private static WebElement sl_6thRow6thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[8]")
		private static WebElement sl_6thRow7thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[9]")
		private static WebElement sl_6thRow8thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[10]")
		private static WebElement sl_6thRow9thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[11]")
		private static WebElement sl_6thRow10thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[12]")
		private static WebElement sl_6thRow11thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[13]")
		private static WebElement sl_6thRow12thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[14]")
		private static WebElement sl_6thRow13thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[15]")
		private static WebElement sl_6thRow14thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[16]")
		private static WebElement sl_6thRow15thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[17]")
		private static WebElement sl_6thRow16thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[18]")
		private static WebElement sl_6thRow17thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[19]")
		private static WebElement sl_6thRow18thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[20]")
		private static WebElement sl_6thRow19thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[21]")
		private static WebElement sl_6thRow20thCol;
		

		@FindBy(xpath="//*[@id='trRender_6']/td[22]")
		private static WebElement sl_6thRow21thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[23]")
		private static WebElement sl_6thRow22thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[24]")
		private static WebElement sl_6thRow23thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[25]")
		private static WebElement sl_6thRow24thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[26]")
		private static WebElement sl_6thRow25thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[27]")
		private static WebElement sl_6thRow26thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[28]")
		private static WebElement sl_6thRow27thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[29]")
		private static WebElement sl_6thRow28thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[30]")
		private static WebElement sl_6thRow29thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[31]")
		private static WebElement sl_6thRow30thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[32]")
		private static WebElement sl_6thRow31thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[33]")
		private static WebElement sl_6thRow32thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[34]")
		private static WebElement sl_6thRow33thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[35]")
		private static WebElement sl_6thRow34thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[36]")
		private static WebElement sl_6thRow35thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[37]")
		private static WebElement sl_6thRow36thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[38]")
		private static WebElement sl_6thRow37thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[39]")
		private static WebElement sl_6thRow38thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[40]")
		private static WebElement sl_6thRow39thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[41]")
		private static WebElement sl_6thRow40thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[42]")
		private static WebElement sl_6thRow41thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[43]")
		private static WebElement sl_6thRow42thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[44]")
		private static WebElement sl_6thRow43thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[45]")
		private static WebElement sl_6thRow44thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[46]")
		private static WebElement sl_6thRow45thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[47]")
		private static WebElement sl_6thRow46thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[48]")
		private static WebElement sl_6thRow47thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[49]")
		private static WebElement sl_6thRow48thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[50]")
		private static WebElement sl_6thRow49thCol;
		
		@FindBy(xpath="//*[@id='trRender_6']/td[51]")
		private static WebElement sl_6thRow50thCol;
		
		
		
		
		
		
		
		
		@FindBy(xpath="//*[@id='trRender_7']/td[2]")
		private static WebElement sl_7thRow1stCol;
				
		@FindBy(xpath="//*[@id='trRender_7']/td[3]")
		private static WebElement sl_7thRow2ndCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[4]")
		private static WebElement sl_7thRow3rdCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[5]")
		private static WebElement sl_7thRow4thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[6]")
		private static WebElement sl_7thRow5thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[7]")
		private static WebElement sl_7thRow6thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[8]")
		private static WebElement sl_7thRow7thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[9]")
		private static WebElement sl_7thRow8thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[10]")
		private static WebElement sl_7thRow9thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[11]")
		private static WebElement sl_7thRow10thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[12]")
		private static WebElement sl_7thRow11thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[13]")
		private static WebElement sl_7thRow12thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[14]")
		private static WebElement sl_7thRow13thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[15]")
		private static WebElement sl_7thRow14thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[16]")
		private static WebElement sl_7thRow15thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[17]")
		private static WebElement sl_7thRow16thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[18]")
		private static WebElement sl_7thRow17thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[19]")
		private static WebElement sl_7thRow18thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[20]")
		private static WebElement sl_7thRow19thCol;
		
		@FindBy(xpath="//*[@id='trRender_7']/td[21]")
		private static WebElement sl_7thRow20thCol;
		
		
		
		
		
		@FindBy(xpath="//*[@id='trRender_8']/td[2]")
		private static WebElement sl_8thRow1stCol;
				
		@FindBy(xpath="//*[@id='trRender_8']/td[3]")
		private static WebElement sl_8thRow2ndCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[4]")
		private static WebElement sl_8thRow3rdCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[5]")
		private static WebElement sl_8thRow4thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[6]")
		private static WebElement sl_8thRow5thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[7]")
		private static WebElement sl_8thRow6thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[8]")
		private static WebElement sl_8thRow7thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[9]")
		private static WebElement sl_8thRow8thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[10]")
		private static WebElement sl_8thRow9thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[11]")
		private static WebElement sl_8thRow10thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[12]")
		private static WebElement sl_8thRow11thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[13]")
		private static WebElement sl_8thRow12thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[14]")
		private static WebElement sl_8thRow13thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[15]")
		private static WebElement sl_8thRow14thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[16]")
		private static WebElement sl_8thRow15thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[17]")
		private static WebElement sl_8thRow16thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[18]")
		private static WebElement sl_8thRow17thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[19]")
		private static WebElement sl_8thRow18thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[20]")
		private static WebElement sl_8thRow19thCol;
		
		@FindBy(xpath="//*[@id='trRender_8']/td[21]")
		private static WebElement sl_8thRow20thCol;
		
		
		
		@FindBy(xpath="//*[@id='trRender_9']/td[2]")
		private static WebElement sl_9thRow1stCol;
				
		@FindBy(xpath="//*[@id='trRender_9']/td[3]")
		private static WebElement sl_9thRow2ndCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[4]")
		private static WebElement sl_9thRow3rdCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[5]")
		private static WebElement sl_9thRow4thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[6]")
		private static WebElement sl_9thRow5thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[7]")
		private static WebElement sl_9thRow6thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[8]")
		private static WebElement sl_9thRow7thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[9]")
		private static WebElement sl_9thRow8thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[10]")
		private static WebElement sl_9thRow9thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[11]")
		private static WebElement sl_9thRow10thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[12]")
		private static WebElement sl_9thRow11thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[13]")
		private static WebElement sl_9thRow12thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[14]")
		private static WebElement sl_9thRow13thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[15]")
		private static WebElement sl_9thRow14thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[16]")
		private static WebElement sl_9thRow15thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[17]")
		private static WebElement sl_9thRow16thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[18]")
		private static WebElement sl_9thRow17thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[19]")
		private static WebElement sl_9thRow18thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[20]")
		private static WebElement sl_9thRow19thCol;
		
		@FindBy(xpath="//*[@id='trRender_9']/td[21]")
		private static WebElement sl_9thRow20thCol;
		
		
		
		@FindBy(xpath="//*[@id='trRender_10']/td[2]")
		private static WebElement sl_10thRow1stCol;
				
		@FindBy(xpath="//*[@id='trRender_10']/td[3]")
		private static WebElement sl_10thRow2ndCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[4]")
		private static WebElement sl_10thRow3rdCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[5]")
		private static WebElement sl_10thRow4thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[6]")
		private static WebElement sl_10thRow5thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[7]")
		private static WebElement sl_10thRow6thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[8]")
		private static WebElement sl_10thRow7thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[9]")
		private static WebElement sl_10thRow8thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[10]")
		private static WebElement sl_10thRow9thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[11]")
		private static WebElement sl_10thRow10thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[12]")
		private static WebElement sl_10thRow11thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[13]")
		private static WebElement sl_10thRow12thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[14]")
		private static WebElement sl_10thRow13thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[15]")
		private static WebElement sl_10thRow14thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[16]")
		private static WebElement sl_10thRow15thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[17]")
		private static WebElement sl_10thRow16thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[18]")
		private static WebElement sl_10thRow17thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[19]")
		private static WebElement sl_10thRow18thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[20]")
		private static WebElement sl_10thRow19thCol;
		
		@FindBy(xpath="//*[@id='trRender_10']/td[21]")
		private static WebElement sl_10thRow20thCol;
		
		
		@FindBy(xpath="//*[@id='trRender_11']/td[2]")
		private static WebElement sl_11thRow1stCol;
				
		@FindBy(xpath="//*[@id='trRender_11']/td[3]")
		private static WebElement sl_11thRow2ndCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[4]")
		private static WebElement sl_11thRow3rdCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[5]")
		private static WebElement sl_11thRow4thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[6]")
		private static WebElement sl_11thRow5thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[7]")
		private static WebElement sl_11thRow6thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[8]")
		private static WebElement sl_11thRow7thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[9]")
		private static WebElement sl_11thRow8thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[10]")
		private static WebElement sl_11thRow9thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[11]")
		private static WebElement sl_11thRow10thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[12]")
		private static WebElement sl_11thRow11thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[13]")
		private static WebElement sl_11thRow12thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[14]")
		private static WebElement sl_11thRow13thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[15]")
		private static WebElement sl_11thRow14thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[16]")
		private static WebElement sl_11thRow15thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[17]")
		private static WebElement sl_11thRow16thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[18]")
		private static WebElement sl_11thRow17thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[19]")
		private static WebElement sl_11thRow18thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[20]")
		private static WebElement sl_11thRow19thCol;
		
		@FindBy(xpath="//*[@id='trRender_11']/td[21]")
		private static WebElement sl_11thRow20thCol;
		
		
		@FindBy(xpath="//*[@id='trRender_12']/td[2]")
		private static WebElement sl_12thRow1stCol;
				
		@FindBy(xpath="//*[@id='trRender_12']/td[3]")
		private static WebElement sl_12thRow2ndCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[4]")
		private static WebElement sl_12thRow3rdCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[5]")
		private static WebElement sl_12thRow4thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[6]")
		private static WebElement sl_12thRow5thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[7]")
		private static WebElement sl_12thRow6thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[8]")
		private static WebElement sl_12thRow7thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[9]")
		private static WebElement sl_12thRow8thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[10]")
		private static WebElement sl_12thRow9thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[11]")
		private static WebElement sl_12thRow10thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[12]")
		private static WebElement sl_12thRow11thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[13]")
		private static WebElement sl_12thRow12thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[14]")
		private static WebElement sl_12thRow13thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[15]")
		private static WebElement sl_12thRow14thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[16]")
		private static WebElement sl_12thRow15thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[17]")
		private static WebElement sl_12thRow16thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[18]")
		private static WebElement sl_12thRow17thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[19]")
		private static WebElement sl_12thRow18thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[20]")
		private static WebElement sl_12thRow19thCol;
		
		@FindBy(xpath="//*[@id='trRender_12']/td[21]")
		private static WebElement sl_12thRow20thCol;
		
		
		
		@FindBy(xpath="//*[@id='trRender_13']/td[2]")
		private static WebElement sl_13thRow1stCol;
				
		@FindBy(xpath="//*[@id='trRender_13']/td[3]")
		private static WebElement sl_13thRow2ndCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[4]")
		private static WebElement sl_13thRow3rdCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[5]")
		private static WebElement sl_13thRow4thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[6]")
		private static WebElement sl_13thRow5thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[7]")
		private static WebElement sl_13thRow6thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[8]")
		private static WebElement sl_13thRow7thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[9]")
		private static WebElement sl_13thRow8thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[10]")
		private static WebElement sl_13thRow9thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[11]")
		private static WebElement sl_13thRow10thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[12]")
		private static WebElement sl_13thRow11thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[13]")
		private static WebElement sl_13thRow12thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[14]")
		private static WebElement sl_13thRow13thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[15]")
		private static WebElement sl_13thRow14thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[16]")
		private static WebElement sl_13thRow15thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[17]")
		private static WebElement sl_13thRow16thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[18]")
		private static WebElement sl_13thRow17thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[19]")
		private static WebElement sl_13thRow18thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[20]")
		private static WebElement sl_13thRow19thCol;
		
		@FindBy(xpath="//*[@id='trRender_13']/td[21]")
		private static WebElement sl_13thRow20thCol;
		
		
		
		@FindBy(xpath="//*[@id='trRender_14']/td[2]")
		private static WebElement sl_14thRow1stCol;
				
		@FindBy(xpath="//*[@id='trRender_14']/td[3]")
		private static WebElement sl_14thRow2ndCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[4]")
		private static WebElement sl_14thRow3rdCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[5]")
		private static WebElement sl_14thRow4thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[6]")
		private static WebElement sl_14thRow5thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[7]")
		private static WebElement sl_14thRow6thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[8]")
		private static WebElement sl_14thRow7thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[9]")
		private static WebElement sl_14thRow8thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[10]")
		private static WebElement sl_14thRow9thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[11]")
		private static WebElement sl_14thRow10thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[12]")
		private static WebElement sl_14thRow11thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[13]")
		private static WebElement sl_14thRow12thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[14]")
		private static WebElement sl_14thRow13thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[15]")
		private static WebElement sl_14thRow14thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[16]")
		private static WebElement sl_14thRow15thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[17]")
		private static WebElement sl_14thRow16thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[18]")
		private static WebElement sl_14thRow17thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[19]")
		private static WebElement sl_14thRow18thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[20]")
		private static WebElement sl_14thRow19thCol;
		
		@FindBy(xpath="//*[@id='trRender_14']/td[21]")
		private static WebElement sl_14thRow20thCol;
		
		
		
		@FindBy(xpath="//*[@id='trRender_15']/td[2]")
		private static WebElement sl_15thRow1stCol;
				
		@FindBy(xpath="//*[@id='trRender_15']/td[3]")
		private static WebElement sl_15thRow2ndCol;
		
		@FindBy(xpath="//*[@id='trRender_15']/td[4]")
		private static WebElement sl_15thRow3rdCol;
		
		@FindBy(xpath="//*[@id='trRender_15']/td[5]")
		private static WebElement sl_15thRow4thCol;
		
		@FindBy(xpath="//*[@id='trRender_15']/td[6]")
		private static WebElement sl_15thRow5thCol;
		
		@FindBy(xpath="//*[@id='trRender_15']/td[7]")
		private static WebElement sl_15thRow6thCol;
		
		@FindBy(xpath="//*[@id='trRender_15']/td[8]")
		private static WebElement sl_15thRow7thCol;
		
		@FindBy(xpath="//*[@id='trRender_15']/td[9]")
		private static WebElement sl_15thRow8thCol;
		
		@FindBy(xpath="//*[@id='trRender_15']/td[10]")
		private static WebElement sl_15thRow9thCol;
		
		@FindBy(xpath="//*[@id='trRender_15']/td[11]")
		private static WebElement sl_15thRow10thCol;
		
		@FindBy(xpath="//*[@id='trRender_15']/td[12]")
		private static WebElement sl_15thRow11thCol;
		
		@FindBy(xpath="//*[@id='trRender_15']/td[13]")
		private static WebElement sl_15thRow12thCol;
		
		
		
		
		/////////////////////////////////
		
		// Opening Stock Register Elements
		
		@FindBy (xpath="//li[@id='154']/a[2]")
		private static WebElement osr_ItemsGroupBtn;
		
		@FindBy(xpath="//select[@id='DateOptions_']")
		private static WebElement osr_dateOptionDropdown;
		
		@FindBy(xpath="//input[@id='id_starting_date_']")
		private static WebElement osr_StartDateTxt;
		
		@FindBy(xpath="//td[@id='id_starting_date__input_image']//span[@class='icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow']")
		private static WebElement osr_startdateCalanderBtn;
		
		@FindBy(xpath="//input[@id='id_ending_date_']")
		private static WebElement osr_EndDateTxt;
		
		@FindBy(xpath="//td[@id='id_ending_date__input_image']//span[@class='icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow']")
		private static WebElement osr_EndDateCalanderBtn;
		
		@FindBy(xpath="//select[@id='RITLayout_']")
		private static WebElement osr_layoutDropdown;
		
		@FindBy(xpath="//select[@id='RITOutput_']")
		private static WebElement osr_outputDropdown;
		
		@FindBy(xpath="//i[@class='icon-font6 icon-schedule']")
		private static WebElement osr_ScheduleBtn;
		
		@FindBy(xpath="//i[@class='icon-font6 icon-custamize']")
		private static WebElement osr_CustomizeBtn;
		
		@FindBy(xpath="//i[@class='icon-font6 icon-filter']")
		private static WebElement osr_FilterBtn;
		
		@FindBy(xpath="//i[@class='icon-font6 icon-ok']")
		private static WebElement osr_OkBtn;
		
		@FindBy(xpath="//i[@class='icon-font6 icon-close']")
		private static WebElement osr_CloseBtn;
		
		///////
		
		@FindBy(xpath="//input[@id='id_body_37']")
		private static WebElement Enter_Expirydate;
		
		////////////////////////
		
		// Finance Reports Almost Common Some Fields Are Added Below
		
		@FindBy (xpath="//li[@id='18']/a[1]")
		private static WebElement assetsExpandOrCollepseBtn;
		
		@FindBy(xpath="//li[@id='18']/a[2]")
		private static WebElement assets;
		
		@FindBy(xpath="//li[@id='17']/a[2]")
		private static WebElement fixedAssets;
		
		@FindBy (xpath="//li[@id='23']/a[1]")
		private static WebElement currentAssetsExpandOrCollepseBtn;	
				
		@FindBy(xpath="//li[@id='23']/a[2]")
		private static WebElement currentAssets;
		
		@FindBy(xpath="//li[@id='8']/a[2]")
		private static WebElement cashAndBank;
		
		@FindBy(xpath="//li[@id='13']/a[2]")
		private static WebElement inventories;
		
		@FindBy(xpath="//li[@id='22']/a[2]")
		private static WebElement accountsReceivable;
		
		@FindBy (xpath="//li[@id='83']/a[2]")
		private static WebElement customerUpdateGroup;
		
		@FindBy(xpath="//li[@id='72']/a[2]")
		private static WebElement investments;
		
		@FindBy (xpath="//li[@id='25']/a[1]")
		private static WebElement expensesExpensionOrCollepseBtn;
		
		@FindBy(xpath="//li[@id='25']/a[2]")
		private static WebElement expenses;
		
		@FindBy(xpath="//li[@id='24']/a[2]")
		private static WebElement directExpenses;
		
		@FindBy (xpath="//li[@id='73']/a[1]")
		private static WebElement indirectExpensesExpandOrCollepseBtn; 
		
		@FindBy(xpath="//li[@id='73']/a[2]")
		private static WebElement indirectExpenses;
		
		@FindBy(xpath="//li[@id='54']/a[2]")
		private static WebElement administrativeExpenses;
		
		@FindBy(xpath="//li[@id='60']/a[2]")
		private static WebElement employeeBenefits;
		
		@FindBy(xpath="//li[@id='64']/a[2]")
		private static WebElement financialCharges;
		
		@FindBy(xpath="//li[@id='69']/a[2]")
		private static WebElement gainAndLoss;
		
		@FindBy(xpath="//li[@id='26']/a[2]")
		private static WebElement controlAccounts;
		
		@FindBy (xpath="//li[@id='29']/a[1]")
		private static WebElement revenueExpandOrCollepseBtn;
		
		@FindBy(xpath="//li[@id='29']/a[2]")
		private static WebElement revenue;
		
		@FindBy (xpath="//li[@id='28']/a[2]")
		private static WebElement sales;
		
		@FindBy (xpath="//li[@id='33']/a[1]")
		private static WebElement equitiesExpandOrCollepseBtn;
		
		@FindBy(xpath="//li[@id='33']/a[2]")
		private static WebElement equities;
		
		@FindBy (xpath="//li[@id='32']/a[2]")
		private static WebElement capital;
		
		@FindBy (xpath="//li[@id='39']/a[1]")
		private static WebElement liabilitiesExpandOrCollepseBtn;
		
		@FindBy (xpath="//li[@id='39']/a[2]")
		private static WebElement liabilities;
		
		@FindBy (xpath="//li[@id='38']/a[1]")
		private static WebElement loansAndBorrowingsExpandOrCollepseBtn;
		
		@FindBy (xpath="//li[@id='38']/a[2]")
		private static WebElement loansAndBorrowings;
		
		@FindBy (xpath="//li[@id='37']/a[2]")
		private static WebElement loans;
		
		@FindBy (xpath="//li[@id='42']/a[2]")
		private static WebElement accruedLiabilities;
		
		@FindBy (xpath="//li[@id='45']/a[2]")
		private static WebElement tradePayable;
		
		@FindBy (xpath="//li[@id='48']/a[2]")
		private static WebElement provisions;
		
		/////////////////////////////////////////////////////////////////////////
		
		
		@FindBy (xpath="//td[@id='id_header_2_day34']")
		private static WebElement calendarDate30;
		
		@FindBy (xpath="//td[@id='id_header_2_day34']")
		private static WebElement dec2016D30;
			
		@FindBy (xpath="//input[@id='txtSerialNo']")
		private static WebElement RMAPopupSerialNumberTxtField;
	
		@FindBy (xpath="//input[@id='txtQuantity']")
		private static WebElement RMAPopupQuantityTxtField;
		
		/*@FindBy (xpath="//div[@id='img_add']/span")
		private static WebElement rmaAddBtn;
			*/
		
		
		
		@FindBy (xpath="//div[@id='Modal_Search_Body']/div[3]/label")
		private static WebElement RMAPopupNumberofItemsLabel;
		
		@FindBy (xpath="//label[@id='lblItems']")
		private static WebElement RMAPopupNumberOfItems;
		
		@FindBy (xpath="//th[@id='RMA_Table_control_heading_1']/div")
		private static WebElement RMAPopupTableHeadingRMA;
		
		@FindBy (xpath="//th[@id='RMA_Table_control_heading_2']/div")
		private static WebElement RMAPopupTableHeadingDocumentNo;
		
		@FindBy (xpath="//th[@id='RMA_Table_control_heading_3']/div")
		private static WebElement RMAPopupTableHeadingDocumentDate;
		
		@FindBy (xpath="//div[@id='RMAModel_Bottom']/div[1]/label")
		private static WebElement RMAPopupClearBtn;
		
		
	
		
		@FindBy (xpath="//td[@id='RMA_Table_col_1-1']")
		private static WebElement RMAPopupTableRow1Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_2-1']")
		private static WebElement RMAPopupTableRow2Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_3-1']")
		private static WebElement RMAPopupTableRow3Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_4-1']")
		private static WebElement RMAPopupTableRow4Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_5-1']")
		private static WebElement RMAPopupTableRow5Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_6-1']")
		private static WebElement RMAPopupTableRow6Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_7-1']")
		private static WebElement RMAPopupTableRow7Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_8-1']")
		private static WebElement RMAPopupTableRow8Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_9-1']")
		private static WebElement RMAPopupTableRow9Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_10-1']")
		private static WebElement RMAPopupTableRow10Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_11-1']")
		private static WebElement RMAPopupTableRow11Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_12-1']")
		private static WebElement RMAPopupTableRow12Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_13-1']")
		private static WebElement RMAPopupTableRow13Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_14-1']")
		private static WebElement RMAPopupTableRow14Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_15-1']")
		private static WebElement RMAPopupTableRow15Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_16-1']")
		private static WebElement RMAPopupTableRow16Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_17-1']")
		private static WebElement RMAPopupTableRow17Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_18-1']")
		private static WebElement RMAPopupTableRow18Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_19-1']")
		private static WebElement RMAPopupTableRow19Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_20-1']")
		private static WebElement RMAPopupTableRow20Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_21-1']")
		private static WebElement RMAPopupTableRow21Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_22-1']")
		private static WebElement RMAPopupTableRow22Column1;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_1-2']")
		private static WebElement RMAPopupTableRow1Column2;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_1-3']")
		private static WebElement RMAPopupTableRow1Column3;
		
		@FindBy (xpath="//td[@id='id_body_38_input_image']/img")
		private static WebElement RMAPopupExpansionBtn;
		
		@FindBy (xpath="//*[@id='id_transactionentry_infopanel_container']/div[1]/div[3]/span")
		private static WebElement infoSideBarBtn;
		
		@FindBy (xpath="//tr[@id='trRender_1']/td[4]")
		private static WebElement savedVoucher1stRowVoucherNumer;
		
		@FindBy (xpath="//span[@id='reportClose']")
		private static WebElement sl_ReportsCloseBtn;
		
	
		//Stock Analysis by RMA Reports Table
		//RMA Number
		@FindBy (xpath="//tr[@id='trRender_1']/td[1]")
		private static WebElement rmaNumber1stRow;
		
		@FindBy (xpath="//tr[@id='trRender_2']/td[1]")
		private static WebElement rmaNumber2ndRow;
		
		@FindBy (xpath="//tr[@id='trRender_3']/td[1]")
		private static WebElement rmaNumber3rdRow;
		
		@FindBy (xpath="//tr[@id='trRender_4']/td[1]")
		private static WebElement rmaNumber4thRow;
		
		@FindBy (xpath="//tr[@id='trRender_5']/td[1]")
		private static WebElement rmaNumber5thRow;
		
		@FindBy (xpath="//tr[@id='trRender_6']/td[1]")
		private static WebElement rmaNumber6thRow;
		
		@FindBy (xpath="//tr[@id='trRender_7']/td[1]")
		private static WebElement rmaNumber7thRow;
		
		@FindBy (xpath="//tr[@id='trRender_8']/td[1]")
		private static WebElement rmaNumber8thRow;
		
		@FindBy (xpath="//tr[@id='trRender_9']/td[1]")
		private static WebElement rmaNumber9thRow;
		
		@FindBy (xpath="//tr[@id='trRender_10']/td[1]")
		private static WebElement rmaNumber10thRow;
		
		@FindBy (xpath="//tr[@id='trRender_11']/td[1]")
		private static WebElement rmaNumber11thRow;
		
		@FindBy (xpath="//tr[@id='trRender_12']/td[1]")
		private static WebElement rmaNumber12thRow;
		
		@FindBy (xpath="//tr[@id='trRender_13']/td[1]")
		private static WebElement rmaNumber13thRow;
		
		@FindBy (xpath="//tr[@id='trRender_14']/td[1]")
		private static WebElement rmaNumber14thRow;
		
		@FindBy (xpath="//tr[@id='trRender_1']/td[1]")
		private static WebElement rmaNumberPg2_1stRow;
		
		@FindBy (xpath="//tr[@id='trRender_2']/td[1]")
		private static WebElement rmaNumberPg2_2ndRow;
		
		@FindBy (xpath="//tr[@id='trRender_3']/td[1]")
		private static WebElement rmaNumberPg2_3rdRow;
		
		@FindBy (xpath="//tr[@id='trRender_4']/td[1]")
		private static WebElement rmaNumberPg2_4thRow;
		
		@FindBy (xpath="//tr[@id='trRender_5']/td[1]")
		private static WebElement rmaNumberPg2_5thRow;
		
		@FindBy (xpath="//tr[@id='trRender_6']/td[1]")
		private static WebElement rmaNumberPg2_6thRow;
		
		
		//Quantity
		@FindBy (xpath="//tr[@id='trRender_1']/td[2]")
		private static WebElement quantity1stRow;
		
		@FindBy (xpath="//tr[@id='trRender_2']/td[2]")
		private static WebElement quantity2ndRow;
		
		@FindBy (xpath="//tr[@id='trRender_3']/td[2]")
		private static WebElement quantity3rdRow;
		
		@FindBy (xpath="//tr[@id='trRender_4']/td[2]")
		private static WebElement quantity4thRow;
		
		@FindBy (xpath="//tr[@id='trRender_5']/td[2]")
		private static WebElement quantity5thRow;
		
		@FindBy (xpath="//tr[@id='trRender_6']/td[2]")
		private static WebElement quantity6thRow;
		
		@FindBy (xpath="//tr[@id='trRender_7']/td[2]")
		private static WebElement quantity7thRow;
		
		@FindBy (xpath="//tr[@id='trRender_8']/td[2]")
		private static WebElement quantity8thRow;
		
		@FindBy (xpath="//tr[@id='trRender_9']/td[2]")
		private static WebElement quantity9thRow;
		
		@FindBy (xpath="//tr[@id='trRender_10']/td[2]")
		private static WebElement quantity10thRow;
		
		@FindBy (xpath="//tr[@id='trRender_11']/td[2]")
		private static WebElement quantity11thRow;
		
		@FindBy (xpath="//tr[@id='trRender_12']/td[2]")
		private static WebElement quantity12thRow;
		
		@FindBy (xpath="//tr[@id='trRender_13']/td[2]")
		private static WebElement quantity13thRow;
		
		@FindBy (xpath="//tr[@id='trRender_14']/td[2]")
		private static WebElement quantity14thRow;
		
		@FindBy (xpath="//tr[@id='trRender_1']/td[2]")
		private static WebElement quantityPg2_1stRow;
		
		@FindBy (xpath="//tr[@id='trRender_2']/td[2]")
		private static WebElement quantityPg2_2ndRow;
		
		@FindBy (xpath="//tr[@id='trRender_3']/td[2]")
		private static WebElement quantityPg2_3rdRow;
		
		@FindBy (xpath="//tr[@id='trRender_4']/td[2]")
		private static WebElement quantityPg2_4thRow;
		
		@FindBy (xpath="//tr[@id='trRender_5']/td[2]")
		private static WebElement quantityPg2_5thRow;
		
		@FindBy (xpath="//tr[@id='trRender_6']/td[2]")
		private static WebElement quantityPg2_6thRow;
		
		//Rate
		@FindBy (xpath="//tr[@id='trRender_1']/td[3]")
		private static WebElement rate1stRow;
		
		@FindBy (xpath="//tr[@id='trRender_2']/td[3]")
		private static WebElement rate2ndRow;
		
		@FindBy (xpath="//tr[@id='trRender_3']/td[3]")
		private static WebElement rate3rdRow;
		
		@FindBy (xpath="//tr[@id='trRender_4']/td[3]")
		private static WebElement rate4thRow;
		
		@FindBy (xpath="//tr[@id='trRender_5']/td[3]")
		private static WebElement rate5thRow;
		
		@FindBy (xpath="//tr[@id='trRender_6']/td[3]")
		private static WebElement rate6thRow;
		
		@FindBy (xpath="//tr[@id='trRender_7']/td[3]")
		private static WebElement rate7thRow;
		
		@FindBy (xpath="//tr[@id='trRender_8']/td[3]")
		private static WebElement rate8thRow;
		
		@FindBy (xpath="//tr[@id='trRender_9']/td[3]")
		private static WebElement rate9thRow;
		
		@FindBy (xpath="//tr[@id='trRender_10']/td[3]")
		private static WebElement rate10thRow;
		
		@FindBy (xpath="//tr[@id='trRender_11']/td[3]")
		private static WebElement rate11thRow;
		
		@FindBy (xpath="//tr[@id='trRender_12']/td[3]")
		private static WebElement rate12thRow;
		
		@FindBy (xpath="//tr[@id='trRender_13']/td[3]")
		private static WebElement rate13thRow;
		
		@FindBy (xpath="//tr[@id='trRender_14']/td[3]")
		private static WebElement rate14thRow;
		
		@FindBy (xpath="//tr[@id='trRender_1']/td[3]")
		private static WebElement ratePg2_1stRow;
		
		@FindBy (xpath="//tr[@id='trRender_2']/td[3]")
		private static WebElement ratePg2_2ndRow;
		
		@FindBy (xpath="//tr[@id='trRender_3']/td[3]")
		private static WebElement ratePg2_3rdRow;
		
		@FindBy (xpath="//tr[@id='trRender_4']/td[3]")
		private static WebElement ratePg2_4thRow;
		
		@FindBy (xpath="//tr[@id='trRender_5']/td[3]")
		private static WebElement ratePg2_5thRow;
		
		@FindBy (xpath="//tr[@id='trRender_6']/td[3]")
		private static WebElement ratePg2_6thRow;
		
		//Stock Value
		@FindBy (xpath="//tr[@id='trRender_1']/td[4]")
		private static WebElement stockValue1stRow;
		
		@FindBy (xpath="//tr[@id='trRender_2']/td[4]")
		private static WebElement stockValue2ndRow;
		
		@FindBy (xpath="//tr[@id='trRender_3']/td[4]")
		private static WebElement stockValue3rdRow;
		
		@FindBy (xpath="//tr[@id='trRender_4']/td[4]")
		private static WebElement stockValue4thRow;
		
		@FindBy (xpath="//tr[@id='trRender_5']/td[4]")
		private static WebElement stockValue5thRow;
		
		@FindBy (xpath="//tr[@id='trRender_6']/td[4]")
		private static WebElement stockValue6thRow;
		
		@FindBy (xpath="//tr[@id='trRender_7']/td[4]")
		private static WebElement stockValue7thRow;
		
		@FindBy (xpath="//tr[@id='trRender_8']/td[4]")
		private static WebElement stockValue8thRow;
		
		@FindBy (xpath="//tr[@id='trRender_9']/td[4]")
		private static WebElement stockValue9thRow;
		
		@FindBy (xpath="//tr[@id='trRender_10']/td[4]")
		private static WebElement stockValue10thRow;
		
		@FindBy (xpath="//tr[@id='trRender_11']/td[4]")
		private static WebElement stockValue11thRow;
		
		@FindBy (xpath="//tr[@id='trRender_12']/td[4]")
		private static WebElement stockValue12thRow;
		
		@FindBy (xpath="//tr[@id='trRender_13']/td[4]")
		private static WebElement stockValue13thRow;
		
		@FindBy (xpath="//tr[@id='trRender_14']/td[4]")
		private static WebElement stockValue14thRow;
		
		@FindBy (xpath="//tr[@id='trRender_1']/td[4]")
		private static WebElement stockValuePg2_1stRow;
		
		@FindBy (xpath="//tr[@id='trRender_2']/td[4]")
		private static WebElement stockValuePg2_2ndRow;
		
		@FindBy (xpath="//tr[@id='trRender_3']/td[4]")
		private static WebElement stockValuePg2_3rdRow;
		
		@FindBy (xpath="//tr[@id='trRender_4']/td[4]")
		private static WebElement stockValuePg2_4thRow;
		
		@FindBy (xpath="//tr[@id='trRender_5']/td[4]")
		private static WebElement stockValuePg2_5thRow;
		
		@FindBy (xpath="//tr[@id='trRender_6']/td[4]")
		private static WebElement stockValuePg2_6thRow;
		
		@FindBy (xpath="//button[@id='pgnumber_']")
		private static WebElement rmaReportsCurrentPageBtn;
		
		@FindBy (xpath="//td[@id='RMA_Table_col_8-4']")
		private static WebElement RMAPopupTable8thRowClrBtn;
		
		@FindBy (xpath="//tr[@id='trRender_1']/td[2]")
		private static WebElement savedVoucher1stRowChkBox;
		
			
		
				
		@FindBy (xpath="//input[@id='MasterSingle__101']")
		private static WebElement ageingAnalysisByRMAWarehouseDrpdwn;
		
		@FindBy (xpath="//input[@id='RITNumber__11']")
		private static WebElement slab1;
		
		@FindBy (xpath="//input[@id='RITNumber__12']")
		private static WebElement slab2;
		
		@FindBy (xpath="//input[@id='RITNumber__13']")
		private static WebElement slab3;
		
		@FindBy (xpath="//input[@id='RITNumber__14']")
		private static WebElement slab4;
		
		@FindBy (xpath="//input[@id='RITNumber__15']")
		private static WebElement slab5;
		
		@FindBy (xpath="//input[@id='RITNumber__16']")
		private static WebElement slab6;
		
		@FindBy (xpath="//input[@id='RITNumber__17']")
		private static WebElement slab7;
		
		@FindBy (xpath="//select[@id='RITCombobox__8']")
		private static WebElement slabUptoDrpdwn;
		
		
		
		@FindBy(xpath="//tr[@id='id_header_2_day_grid']")
		private static WebElement cal;
		
		
		@FindBy(xpath="//select[@id='id_header_2_month']")
		private static WebElement month;
		
		@FindBy(xpath="//select[@id='id_header_2_year']")
		private static WebElement year;
		
		
		
		
		@FindBy (xpath="//*[@id='sp_search']")
		private static WebElement RMAPopupSerialNoExpansionBtn;
		
		@FindBy (xpath="//input[@id='Search_txtSearchBox']")
		private static WebElement searchRMASearchTxtField;
		
		@FindBy (xpath="//input[@id='Search_txtQuantity']")
		private static WebElement searchRMAQuantityTxtField;
		
		@FindBy (xpath="//label[@id='Search_lblItems']")
		private static WebElement searchRMANoOfItemsTxt;
		
		@FindBy (xpath="//input[@id='chkSelectALL']")
		private static WebElement searchRMASelectAllChkBox;
		
		@FindBy (xpath="//*[@id='RMASearchTable']/tbody/tr[4]/td[1]/input")
		private static WebElement searchRMA4ChkBox;
		
		@FindBy (xpath="//*[@id='RMASearchTable']/tbody/tr[101]/td[1]/input")
		private static WebElement searchRMA101ChkBox;
		
		@FindBy (xpath="//*[@id='RMASearchTable']/tbody/tr[102]/td[1]/input")
		private static WebElement searchRMA102ChkBox;
		
		@FindBy (xpath="//*[@id='RMASearchTable']/tbody/tr[103]/td[1]/input")
		private static WebElement searchRMA103ChkBox;
		
		@FindBy (xpath="//*[@id='RMASearchTable']/tbody/tr[2]/td[1]/input")
		private static WebElement searchRMA2ndRowChkBox;
		
		@FindBy (xpath="//table[@id='RMASearchTable']/thead/tr/td[2]")
		private static WebElement searchRMATableHeadingSerialNo;
		
		@FindBy (xpath="//table[@id='RMASearchTable']/thead/tr/td[3]")
		private static WebElement searchRMATableHeadingDocNo;
		
		@FindBy (xpath="//table[@id='RMASearchTable']/thead/tr/td[4]")
		private static WebElement searchRMATableHeadingDate;
		
		@FindBy (xpath="(//button[@id='div_OK'])[2]")
		private static WebElement searchRMAOkBtn;
		
		@FindBy (xpath="//*[@id='div_Cancel']")
		private static WebElement searchRMACancelBtn;
		
		@FindBy (xpath="//*[@id='Modal_Header']/span/i")
		private static WebElement searchRMACloseBtn;
		
		
		
		@FindBy(xpath="//*[@id='16']")
		private static WebElement  SettingsmenuBtn;
		
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
		
		@FindBy(xpath="//input[@id='MaintainInventoryInAlternateUnit']")
		private static WebElement  MaintainInventoryInAlternateUnitChkBoxIsSelected;
		
		
		@FindBy(xpath="//input[@id='MaintainInventoryInAlternateUnit']/following-sibling::span")
		private static WebElement  MaintainInventoryInAlternateUnitChkBox;
		
		@FindBy(xpath="//ul[@id='preferenceUL']//ul//li//a[@id='7']")
		private static WebElement  BatchBtn;
		
		@FindBy(xpath="(//*[@id='30'])[2]")
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
		
		@FindBy(xpath="//*[@id='btnCustomizeClose']")
		private static WebElement  CloseBtn;
	
	

			@FindBy(xpath="//*[@id='dvRMASupports']/div/label/span")
			private static WebElement  RmaSupportChkBox;
			
			
			
			
			@FindBy(xpath="//*[@id='chkRMAbasedalternateQuantity']/following-sibling::span")
			private static WebElement  RMAbasedalternateQuantityChkBox;
			
			@FindBy(xpath="//*[@id='chkRMAbasedalternateQuantity']")
			private static WebElement  RMAbasedalternateQuantityChkBoxIsSelected;
			
			
			
			
			@FindBy(xpath="//*[@id='chkRMANumberAlwaysnumeric']/following-sibling::span")
			private static WebElement  RMANumberAlwaysnumericChkBox;
			
			@FindBy(xpath="//*[@id='chkRMANumberAlwaysnumeric']")
			private static WebElement  RMANumberAlwaysnumericChkBoxIsSelected;
			
			
			
			
			
			
			@FindBy(xpath="//*[@id='chkRMANumberUniqueAcrossAllItems']/following-sibling::span")
			private static WebElement  RMANumberUniqueAcrossAllItemsChkBox;

			@FindBy(xpath="//*[@id='chkRMANumberUniqueAcrossAllItems']")
			private static WebElement  RMANumberUniqueAcrossAllItemsChkBoxIsSelected;
			
			
			
			@FindBy(xpath="//*[@id='chkInputRMANo']/following-sibling::span")
			private static WebElement  InputRMANoBeforeQuantityChkBox;
			
			@FindBy(xpath="//*[@id='chkInputRMANo']")
			private static WebElement  InputRMANoBeforeQuantityChkBoxIsSelected;
			
			
			@FindBy(xpath="//*[@id='chkRMANumberCantBeUsed']/following-sibling::span")
			private static WebElement  RMANumberCantBeUsedChkBox;
			
			@FindBy(xpath="//*[@id='chkRMANumberCantBeUsed']")
			private static WebElement  RMANumberCantBeUsedChkBoxIsSelected;
			
			@FindBy(xpath="//*[@id='PrefixVoucherNo']/following-sibling::span")
			private static WebElement  PrefixVoucherNoChkBox;
	
			@FindBy(xpath="//*[@id='PrefixVoucherNo']")
			private static WebElement  PrefixVoucherNoChkBoxIsSelected;
			
			
			
			
		
			
			@FindBy (xpath="//*[@id='27']/span")
			private static WebElement restore;
			
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
			
			
			
			@FindBy(xpath="//*[@id='txtUsername']")
			private static WebElement username;
	
			@FindBy(id="txtPassword")
			private static WebElement password;
	
			@FindBy(id="btnSignin")
			private static WebElement signIn;
	
		    @FindBy(id="ddlCompany")
		    private static WebElement companyDropDownList;
	
			
		    @FindBy(xpath = "//*[@id='id_footer']/strong")
			public static WebElement dashbordPageFooterFiledTxt;
		    
		    @FindBy (xpath="//*[@id='RMA_Table_body']/tr[1]/td[2]")
			private static WebElement stockTransferRMA1stRow1stColumn;
			
			@FindBy (xpath="//*[@id='RMA_Table_body']/tr[2]/td[2]")
			private static WebElement stockTransferRMA2ndRow1stColumn;
			
			@FindBy (xpath="//*[@id='RMA_Table_body']/tr[3]/td[2]")
			private static WebElement stockTransferRMA3rdRow1stColumn;
			
			@FindBy (xpath="//*[@id='RMA_Table_body']/tr[4]/td[2]")
			private static WebElement stockTransferRMA4thRow1stColumn;
			
			@FindBy (xpath="//*[@id='RMA_Table_body']/tr[5]/td[2]")
			private static WebElement stockTransferRMA5thRow1stColumn;
			
			@FindBy (xpath="//*[@id='RMA_Table_body']/tr[6]/td[2]")
			private static WebElement stockTransferRMA6thRow1stColumn;
			
			@FindBy (xpath="//*[@id='RMA_Table_body']/tr[7]/td[2]")
			private static WebElement stockTransferRMA7thRow1stColumn;
			
			@FindBy (xpath="//*[@id='RMA_Table_body']/tr[8]/td[2]")
			private static WebElement stockTransferRMA8thRow1stColumn;
			
			@FindBy (xpath="//*[@id='RMA_Table_body']/tr[9]/td[2]")
			private static WebElement stockTransferRMA9thRow1stColumn;
			
			@FindBy (xpath="//*[@id='RMA_Table_body']/tr[10]/td[2]")
			private static WebElement stockTransferRMA10thRow1stColumn;
			
	
			@FindBy (xpath="//*[@id='RMA_Table_body']/tr[1]/td[3]")
			private static WebElement stockTransferRMA1stRow2ndColumn;
			
			@FindBy (xpath="//*[@id='RMA_Table_body']/tr[2]/td[3]")
			private static WebElement stockTransferRMA2ndRow2ndColumn;
			
			@FindBy (xpath="//*[@id='RMA_Table_body']/tr[1]/td[4]")
			private static WebElement stockTransferRMA1stRow3rdColumn;
			
			@FindBy (xpath="//*[@id='RMA_Table_body']/tr[2]/td[4]")
			private static WebElement stockTransferRMA2ndRow3rdColumn;
			
	
			
		
	        @FindBy(xpath="//input[@id='donotshow']")
	        private static WebElement DontShowChkBox;
	        
	        @FindBy(xpath="//*[@id='2053']")
	        private static WebElement salesInvoiceNvoucher;
	        
	        @FindBy(xpath="//a[@id='137']//span[contains(text(),'Transactions')]")
			private static WebElement  inventoryTransactionsMenu; 
	
			
		     
			@FindBy(xpath="//tbody[@id='tblBodyReportRender']/tr")
			private static List<WebElement> reportsBodyGridRowsCount;
			
			
			@FindBy(xpath="//tbody[@id='id_body_23_table_data_body']/tr")
			private static List<WebElement> voucherItem;
			
			@FindBy(xpath="//div[@id='idGlobalError']")
			private static WebElement  globalError; 
			
			
			
			@FindBy(xpath="//div[@id='ShowLoginErrId']/div/div/div[2]")
			private static WebElement loginWarningPopup;
			
			@FindBy(xpath="//div[@id='ShowLoginErrId']/div/div/div[3]/button")
			private static WebElement loginWarningPopupOkBtn;
			
			
			
			
			
			
			
			
			@FindBy(xpath="//tbody[@id='id_header_268435460_table_body']/tr/td[2]")
			private static List<WebElement> warehouseHeaderComboList;
			
			@FindBy(xpath="//tbody[@id='id_body_23_table_body']/tr/td[2]")
			private static List<WebElement> itemComboList;
			
			@FindBy(xpath="//tbody[@id='productQuery_table_body']/tr/td[2]")
			private static List<WebElement> iqItemComboList;
			
			
			
			
			@FindBy(xpath="//tbody[@id='id_header_4_table_body']/tr/td[2]")
			private static List<WebElement> customerAccountComboList;
			
			@FindBy(xpath="//tbody[@id='id_header_10_table_body']/tr/td[2]")
			private static List<WebElement> currencyComboList;
			
			@FindBy(xpath="//tbody[@id='id_header_268435459_table_body']/tr/td[2]")
			private static List<WebElement> departmentComboList;
			
			@FindBy(xpath="//tbody[@id='id_body_536870916_table_body']/tr/td[2]")
			private static List<WebElement> warehouseBodyComboList;
			
			@FindBy(xpath="//tbody[@id='id_header_86_table_body']/tr/td[2]")
			private static List<WebElement> st_warehouse1HeaderComboList;
			
			@FindBy(xpath="//tbody[@id='id_body_87_table_body']/tr/td[2]")
			private static List<WebElement> st_warehouse2BodyComboList;
			
			
			@FindBy(xpath="//input[@id='id_header_268435470']")
			private static WebElement  salesInvoiceVATPlaceOFSupply;
			
			
			@FindBy(xpath="//*[@id='id_body_33554511']")
			private static WebElement enter_AQ;
			
			@FindBy(xpath="//*[@id='id_body_33554512']")
			private static WebElement enter_FQ;
			
			@FindBy(xpath="//*[@id='id_body_33554513']")
			private static WebElement enter_RD;
			
			@FindBy (xpath="(//*[@id='div_OK'])[1]")
			private static WebElement salesInvoiceRMAPopupOkBtn;
			
			@FindBy (xpath="//*[@id='div_Clear']")
			private static WebElement salesInvoiceRMAPopupCancelBtn;
			
			@FindBy(xpath="//input[@id='id_header_6']")
			private static WebElement  voucherHeaderDueDate;
			
			@FindBy(xpath="//input[@id='id_header_10']")
			private static WebElement  voucherHeaderCurrency;
			
			@FindBy(xpath="//input[@id='id_header_11']")
			private static WebElement  voucherHeaderExchangeRate;
			
			@FindBy(xpath="//input[@id='id_header_145']")
			private static WebElement  voucherHeaderLocalExchangeRate;
			
			
			
			@FindBy(xpath="//tbody[@id='id_body_12_table_body']/tr/td[2]")
			private static List<WebElement>  salesAccountBodyComboList;
			
			@FindBy(xpath="//input[@id='id_body_33554503']")
			private static WebElement enter_Discount;
			
			@FindBy(xpath="//input[@id='id_body_33554490']")
			private static WebElement enter_PVDiscount;
			
			
			
			@FindBy(xpath="//input[@id='txtNewReference']")
			private static WebElement newReferenceTxt;
			
			
			
			@FindBy(xpath="//input[@id='id_body_37']")
			private static WebElement enter_Expirydate;
			
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[1]/td[1]")
			private static WebElement iq_stockTransaction1stRow1stColumn;
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[1]/td[2]")
			private static WebElement iq_stockTransaction1stRow2ndColumn;
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[1]/td[3]")
			private static WebElement iq_stockTransaction1stRow3rdColumn;
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[1]/td[4]")
			private static WebElement iq_stockTransaction1stRow4thColumn;
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[1]/td[5]")
			private static WebElement iq_stockTransaction1stRow5thColumn;
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[1]/td[6]")
			private static WebElement iq_stockTransaction1stRow6thColumn;
			
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[2]/td[1]")
			private static WebElement iq_stockTransaction2ndRow1stColumn;
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[2]/td[2]")
			private static WebElement iq_stockTransaction2ndRow2ndColumn;
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[2]/td[3]")
			private static WebElement iq_stockTransaction2ndRow3rdColumn;
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[2]/td[4]")
			private static WebElement iq_stockTransaction2ndRow4thColumn;
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[2]/td[5]")
			private static WebElement iq_stockTransaction2ndRow5thColumn;
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[2]/td[6]")
			private static WebElement iq_stockTransaction2ndRow6thColumn;
			
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[3]/td[1]")
			private static WebElement iq_stockTransaction3rdRow1stColumn;
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[3]/td[2]")
			private static WebElement iq_stockTransaction3rdRow2ndColumn;
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[3]/td[3]")
			private static WebElement iq_stockTransaction3rdRow3rdColumn;
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[3]/td[4]")
			private static WebElement iq_stockTransaction3rdRow4thColumn;
			
			@FindBy (xpath="//*[@id='salesTable']/tbody/tr[3]/td[5]")
			private static WebElement iq_stockTransaction3rdRow5thColumn;
			
			
			@FindBy (xpath="//tr[@id='id_ending_date__day_today']//span[@class='theme_color-inverse'][contains(text(),'Today')]")
			private static WebElement sl_EndDateCalender_TodayBtn;
			
			
			
			@FindBy(xpath="//tbody[@id='id_header_3_table_body']/tr/td[2]")
			private static List<WebElement> purchaseAccountComboList;
			
			@FindBy(xpath="//tbody[@id='id_header_3_table_body']/tr/td[2]")
			private static List<WebElement> salesAccountHeaderComboList;
			
			@FindBy(xpath="//tbody[@id='id_header_4_table_body']/tr/td[2]")
			private static List<WebElement> vendorAccountComboList;
			
			
			@FindBy(xpath="//a[@id='2057']//span[contains(text(),'Purchases Voucher VAT')]")
			private static WebElement  purchaseVouchersVat;
			
			
			
			@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[10]")
			private static List<WebElement> nameList;

			@FindBy(xpath="//span[contains(text(),'Place of supply')]")
			private static WebElement placeOfSupply;

			@FindBy(xpath="//a[@id='1178']//span[contains(text(),'Jurisdiction')]")
			private static WebElement jurisdiction;

			@FindBy(xpath ="//*[@id='LandingGridBody']/tr/td[10]")
			public static List<WebElement> masterTaxCodeList;

			@FindBy (xpath="//*[@id='1161']/span")
			private static WebElement taxCode;

			@FindBy(xpath = "//span[@id='btnClose']")
			private static WebElement cancelIcon;

			@FindBy(xpath="//td[@id='editScreen_bodyrow_4 _8']")
			private static WebElement  editScreenbody4thRowFormula;

			@FindBy(xpath="//td[@id='editScreen_bodyrow_4 _9']")
			private static WebElement  editScreenbody4thRowPreload;

			@FindBy(xpath="//td[@id='editScreen_bodyrow_5 _9']")
			private static WebElement  editScreenbody5thRowPreload;

			@FindBy(xpath="//td[@id='editScreen_bodyrow_4 _1']//span[@id='editIcon']")
			private static WebElement  editScreenbody4thRowEditBtn;

			@FindBy(xpath="//td[@id='editScreen_bodyrow_5 _1']//span[@id='editIcon']")
			private static WebElement  editScreenbody5thRowEditBtn;
			
			@FindBy(xpath="//td[@id='editScreen_bodyrow_3 _8']")
			private static WebElement  editScreenbody3rdRowFormula;

			@FindBy(xpath="//td[@id='editScreen_bodyrow_3 _9']")
			private static WebElement  editScreenbody3rdRowPreload;

			@FindBy(xpath="//td[@id='editScreen_bodyrow_3 _1']//span[@id='editIcon']")
			private static WebElement  editScreenbody3rdRowEditBtn;


			@FindBy(xpath="//*[@id='navigationtab2']/div/span[2]")
			private static WebElement  miscellaneousTab;

			@FindBy(xpath="//input[@id='misc_PostingUpdateFA']")
			private static WebElement  postingUpdateChkBox;

			@FindBy(xpath="//input[@id='misc_PostingSelectAccInLine']")
			private static WebElement  selectAccountInEveryLineChkBox;

			@FindBy(xpath="//input[@id='misc_chkUpdateStock']")
			private static WebElement  updateStockChkBox;

			@FindBy(xpath="//input[@id='misc_PostingPostVAT']")
			private static WebElement  postVatChkBox;

			@FindBy(xpath="//input[@id='misc_chkDetailedPosting']")
			private static WebElement detailedPostingChkbox;
			
			
			
			@FindBy(xpath="//input[@id='id_body_536870916']")
		    private static WebElement enter_WarehouseTxt;
				
		    @FindBy(xpath="//*[@id='id_body_536870916_table_body']/tr")
		    private static List<WebElement> pvvGridWarehouseList;
			
		    @FindBy(xpath="//*[@id='id_body_23_table_body']/tr")
		    private static List<WebElement> pvvGridItemList;
		    
		    
		    
		    @FindBy(xpath="//input[@id='id_body_33554486']")
			private static WebElement enter_PvTaxable;
			
			@FindBy(xpath="//tbody[@id='id_body_16777307_table_body']/tr/td")
			private static List<WebElement> pvvGridTaxCodeList;
			
			@FindBy(xpath="//input[@id='id_body_16777323']")
			private static WebElement enter_PvTaxCode;
			
			
			
			@FindBy (xpath="//*[@id='tblBodyTransRender']/tr[1]/td[11]")
			private static WebElement grid_1stRowSuspendStatus;
			
			@FindBy (xpath="//*[@id='tblBodyTransRender']/tr[2]/td[11]")
			private static WebElement grid_2ndRowSuspendStatus;
			
			
			
			@FindBy (xpath="//*[@id='id_transaction_entry_detail_table_control_heading_1']/div[1]")
			private static WebElement stockTransferWarehouse2TableHeader;
			
			
			@FindBy (xpath="//input[@id='id_header_86']")
			private static WebElement stockTransferWarehouse1;
			
			
			@FindBy(xpath="//span[@id='id_transaction_entry_workflow_togglebutton']")
			private static WebElement workFlowExpandBtn;
			
			@FindBy(xpath="//input[@id='id_transaction_entry_detail_workflow_control_heading_ctrl_1']")
			private static WebElement  workFlowHeaderChkBoxCL;
			
			@FindBy(xpath="//span[@class='icon-ok icon-font7']")
			private static WebElement  workFlowOkBtnCL;
			
			
			
			// Item Query Elements
			
			@FindBy(xpath="//input[@id='productQuery']")
			private static WebElement iq_ItemTxt;
			
			@FindBy(xpath="//button[contains(text(),'GetStock')]")
			private static WebElement iq_getStockBtn;
			
			@FindBy(xpath="//button[contains(text(),'Customize')]")
			private static WebElement iq_customizeBtn;
			
			@FindBy(xpath="//span[@id='id_productquery_button_clear']")
			private static WebElement iq_clearBtn;
			
			@FindBy(xpath="//span[@id='id_productquery_button_close']")
			private static WebElement iq_closeBtn;
			
			@FindBy(xpath="//label[contains(text(),'Opening Stock Quantity:')]")
			private static WebElement iq_OpeningStockQtyLable;
			
			@FindBy(xpath="//label[contains(text(),'Pending purchases orders :')]")
			private static WebElement iq_PendingPurchaseOrdersLable;
			
			@FindBy(xpath="//label[contains(text(),'Current Stock :')]")
			private static WebElement iq_CurrentStockLable;
			
			@FindBy(xpath="//label[contains(text(),'Pending Sales Orders :')]")
			private static WebElement iq_PendingSalesOrdersLable;
			
			@FindBy(xpath="//label[contains(text(),'Avg. Stock Rate :')]")
			private static WebElement iq_AvgStockRateLable;
			
			@FindBy(xpath="//label[contains(text(),'Qty to be Ordered :')]")
			private static WebElement iq_QtytobeOrderedLable;
			
			@FindBy(xpath="//label[contains(text(),'Base Unit :')]")
			private static WebElement iq_BaseUnitLable;
			
			@FindBy(xpath="//label[contains(text(),'Value :')]")
			private static WebElement iq_ValueLable;
			
			@FindBy(xpath="//table[@id='purchaseTable']//th[contains(text(),'Date')]")
			private static WebElement iq_PurDate;
			
			@FindBy(xpath="//table[@id='purchaseTable']//th[contains(text(),'Voucher No')]")
			private static WebElement iq_PurVoucherNo;
			
			@FindBy(xpath="//th[contains(text(),'Vendor')]")
			private static WebElement iq_PurVendor;
			
			@FindBy(xpath="//table[@id='purchaseTable']//th[contains(text(),'Quantity')]")
			private static WebElement iq_PurQty;
			
			@FindBy(xpath="//table[@id='purchaseTable']//th[contains(text(),'Unit')]")
			private static WebElement iq_PurUnit;
			
			@FindBy(xpath="//table[@id='purchaseTable']//th[contains(text(),'Rate')]")
			private static WebElement iq_PurRate;
			
			@FindBy(xpath="//table[@id='purchaseTable']//th[contains(text(),'Currency')]")
			private static WebElement iq_PurCurrency;
			
			@FindBy(xpath="//table[@id='purchaseTable']//th[contains(text(),'Quantity In Base')]")
			private static WebElement iq_PurQtyInBase;
			
			@FindBy(xpath="//table[@id='salesTable']//th[contains(text(),'Date')]")
			private static WebElement iq_SaleDate;
			
			@FindBy(xpath="//table[@id='salesTable']//th[contains(text(),'Voucher No')]")
			private static WebElement iq_SaleVoucherNo;
			
			@FindBy(xpath="//th[contains(text(),'Customer')]")
			private static WebElement iq_SaleCustomer;
			
			@FindBy(xpath="//table[@id='salesTable']//th[contains(text(),'Quantity')]")
			private static WebElement iq_SaleQty;
			
			@FindBy(xpath="//table[@id='salesTable']//th[contains(text(),'Unit')]")
			private static WebElement iq_SaleUnit;
			
			@FindBy(xpath="//table[@id='salesTable']//th[contains(text(),'Rate')]")
			private static WebElement iq_SaleRate;
			
			@FindBy(xpath="//table[@id='salesTable']//th[contains(text(),'Currency')]")
			private static WebElement iq_SaleCurrency;
			
			@FindBy(xpath="//table[@id='salesTable']//th[contains(text(),'Quantity In Base')]")
			private static WebElement iq_SaleQtyInBase;
			
			
			@FindBy(xpath="//span[@id='openingStock']")
			private static WebElement iq_OpeningStocks;
			
			@FindBy(xpath="//span[@id='currentStock']")
			private static WebElement iq_CurrentStock;
			
			@FindBy(xpath="//span[@id='avgStockRate']")
			private static WebElement iq_AvgStockRate;
			
			@FindBy(xpath="//span[@id='value']")
			private static WebElement iq_Value;
			
			@FindBy(xpath="//span[@id='iBaseUnit']")
			private static WebElement iq_BaseUnit;
			
			@FindBy(xpath="//span[@id='purchaseOrdered']")
			private static WebElement iq_PendingPurchaseOrders;
			
			@FindBy(xpath="//span[@id='salesOrdered']")
			private static WebElement iq_PendingSalesOrders;
			
			@FindBy(xpath="//span[@id='quantityToBeOrdered']")
			private static WebElement iq_QtyToBeOrdered;
			
			
			// Purchase Table
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[1]")
			private static WebElement pur1strow_1stColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[2]")
			private static WebElement pur1strow_2ndColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[3]")
			private static WebElement pur1strow_3rdColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[4]")
			private static WebElement pur1strow_4thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[5]")
			private static WebElement pur1strow_5thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[6]")
			private static WebElement pur1strow_6thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[7]")
			private static WebElement pur1strow_7thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[8]")
			private static WebElement pur1strow_8thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[9]")
			private static WebElement pur1strow_9thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[10]")
			private static WebElement pur1strow_10thColumn;
			
			
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[1]")
			private static WebElement pur2ndrow_1stColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[2]")
			private static WebElement pur2ndrow_2ndColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[3]")
			private static WebElement pur2ndrow_3rdColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[4]")
			private static WebElement pur2ndrow_4thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[5]")
			private static WebElement pur2ndrow_5thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[6]")
			private static WebElement pur2ndrow_6thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[7]")
			private static WebElement pur2ndrow_7thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[8]")
			private static WebElement pur2ndrow_8thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[9]")
			private static WebElement pur2ndrow_9thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[10]")
			private static WebElement pur2ndrow_10thColumn;
			
			
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[1]")
			private static WebElement pur7thRow_1stColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[2]")
			private static WebElement pur3rdrow_2ndColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[3]")
			private static WebElement pur3rdrow_3rdColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[4]")
			private static WebElement pur3rdrow_4thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[5]")
			private static WebElement pur3rdrow_5thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[6]")
			private static WebElement pur3rdrow_6thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[7]")
			private static WebElement pur3rdrow_7thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[8]")
			private static WebElement pur3rdrow_8thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[9]")
			private static WebElement pur3rdrow_9thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[10]")
			private static WebElement pur3rdrow_10thColumn;
			
			
			// 4TH
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[1]")
			private static WebElement pur4throw_1stColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[2]")
			private static WebElement pur4throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[3]")
			private static WebElement pur4throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[4]")
			private static WebElement pur4throw_4thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[5]")
			private static WebElement pur4throw_5thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[6]")
			private static WebElement pur4throw_6thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[7]")
			private static WebElement pur4throw_7thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[8]")
			private static WebElement pur4throw_8thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[9]")
			private static WebElement pur4throw_9thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[10]")
			private static WebElement pur4throw_10thColumn;
			
			// 5TH
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[1]")
			private static WebElement pur5throw_1stColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[2]")
			private static WebElement pur5throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[3]")
			private static WebElement pur5throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[4]")
			private static WebElement pur5throw_4thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[5]")
			private static WebElement pur5throw_5thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[6]")
			private static WebElement pur5throw_6thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[7]")
			private static WebElement pur5throw_7thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[8]")
			private static WebElement pur5throw_8thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[9]")
			private static WebElement pur5throw_9thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[10]")
			private static WebElement pur5throw_10thColumn;
			
			// 6TH
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[1]")
			private static WebElement pur6throw_1stColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[2]")
			private static WebElement pur6throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[3]")
			private static WebElement pur6throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[4]")
			private static WebElement pur6throw_4thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[5]")
			private static WebElement pur6throw_5thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[6]")
			private static WebElement pur6throw_6thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[7]")
			private static WebElement pur6throw_7thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[8]")
			private static WebElement pur6throw_8thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[9]")
			private static WebElement pur6throw_9thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[10]")
			private static WebElement pur6throw_10thColumn;
			
			
			@FindBy(xpath="//td[contains(text(),'Quantity Purchased Till Date:')]")
			private static WebElement qtyPurchasedTillDate;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[2]")
			private static WebElement purTotalWa;
					
					
			// Sales Table
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[1]")
			private static WebElement sal1strow_1stColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[2]")
			private static WebElement sal1strow_2ndColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[3]")
			private static WebElement sal1strow_3rdColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[4]")
			private static WebElement sal1strow_4thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[5]")
			private static WebElement sal1strow_5thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[6]")
			private static WebElement sal1strow_6thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[7]")
			private static WebElement sal1strow_7thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[8]")
			private static WebElement sal1strow_8thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[9]")
			private static WebElement sal1strow_9thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[10]")
			private static WebElement sal1strow_10thColumn;
			
			// 2nd
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[1]")
			private static WebElement sal2ndrow_1stColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[2]")
			private static WebElement sal2ndrow_2ndColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[3]")
			private static WebElement sal2ndrow_3rdColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[4]")
			private static WebElement sal2ndrow_4thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[5]")
			private static WebElement sal2ndrow_5thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[6]")
			private static WebElement sal2ndrow_6thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[7]")
			private static WebElement sal2ndrow_7thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[8]")
			private static WebElement sal2ndrow_8thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[9]")
			private static WebElement sal2ndrow_9thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[10]")
			private static WebElement sal2ndrow_10thColumn;
			
			// 3rd
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[1]")
			private static WebElement sal3rdrow_1stColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[2]")
			private static WebElement sal3rdrow_2ndColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[3]")
			private static WebElement sal3rdrow_3rdColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[4]")
			private static WebElement sal3rdrow_4thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[5]")
			private static WebElement sal3rdrow_5thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[6]")
			private static WebElement sal3rdrow_6thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[7]")
			private static WebElement sal3rdrow_7thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[8]")
			private static WebElement sal3rdrow_8thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[9]")
			private static WebElement sal3rdrow_9thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[10]")
			private static WebElement sal3rdrow_10thColumn;
			
			// 4th
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[1]")
			private static WebElement sal4throw_1stColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[2]")
			private static WebElement sal4throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[3]")
			private static WebElement sal4throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[4]")
			private static WebElement sal4throw_4thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[5]")
			private static WebElement sal4throw_5thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[6]")
			private static WebElement sal4throw_6thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[7]")
			private static WebElement sal4throw_7thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[8]")
			private static WebElement sal4throw_8thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[9]")
			private static WebElement sal4throw_9thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[10]")
			private static WebElement sal4throw_10thColumn;
			
			//5th
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[1]")
			private static WebElement sal5throw_1stColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[2]")
			private static WebElement sal5throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[3]")
			private static WebElement sal5throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[4]")
			private static WebElement sal5throw_4thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[5]")
			private static WebElement sal5throw_5thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[6]")
			private static WebElement sal5throw_6thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[7]")
			private static WebElement sal5throw_7thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[8]")
			private static WebElement sal5throw_8thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[9]")
			private static WebElement sal5throw_9thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[10]")
			private static WebElement sal5throw_10thColumn;
			
			//6th
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[1]")
			private static WebElement sal6throw_1stColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[2]")
			private static WebElement sal6throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[3]")
			private static WebElement sal6throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[4]")
			private static WebElement sal6throw_4thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[5]")
			private static WebElement sal6throw_5thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[6]")
			private static WebElement sal6throw_6thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[7]")
			private static WebElement sal6throw_7thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[8]")
			private static WebElement sal6throw_8thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[9]")
			private static WebElement sal6throw_9thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[10]")
			private static WebElement sal6throw_10thColumn;
			
			// 7th
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[1]")
			private static WebElement sal7throw_1stColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[2]")
			private static WebElement sal7throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[3]")
			private static WebElement sal7throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[4]")
			private static WebElement sal7throw_4thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[5]")
			private static WebElement sal7throw_5thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[6]")
			private static WebElement sal7throw_6thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[7]")
			private static WebElement sal7throw_7thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[8]")
			private static WebElement sal7throw_8thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[9]")
			private static WebElement sal7throw_9thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[10]")
			private static WebElement sal7throw_10thColumn;

			
			@FindBy(xpath="//td[contains(text(),'Quantity Sold Till Date:')]")
			private static WebElement qtySoldTillDate;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[2]")
			private static WebElement salTotalWa;
					
			
			
			// Warehouse Table
			
			@FindBy(xpath="//div[contains(text(),'Warehouse')]")
			private static WebElement BalancebyWareHouselable;
			
			@FindBy(xpath="//div[@id='PrdtDetails']//div[1]//div[1]//span[1]")
			private static WebElement BalancebyWareHouseExpandBtn;
			
			@FindBy(xpath="//tr[@id='1']/td[2]")
			private static WebElement wh1strow_2ndColumn;
			
			@FindBy(xpath="//tr[@id='1']/td[3]")
			private static WebElement wh1strow_3rdColumn;
			
			@FindBy(xpath="//tr[@id='1']/td[4]")
			private static WebElement wh1strow_4thColumn;
			
			@FindBy(xpath="//tr[@id='1']/td[5]")
			private static WebElement wh1strow_5thColumn;
			
			@FindBy(xpath="//tr[@id='1']/td[6]")
			private static WebElement wh1strow_6thColumn;
			
			
			
			@FindBy(xpath="//*[@id='3']/td[2]")
			private static WebElement wh2ndrow_2ndColumn;
			
			@FindBy(xpath="//*[@id='3']/td[3]")
			private static WebElement wh2ndrow_3rdColumn;
			
			@FindBy(xpath="//*[@id='3']/td[4]")
			private static WebElement wh2ndrow_4thColumn;
			
			@FindBy(xpath="//*[@id='3']/td[5]")
			private static WebElement wh2ndrow_5thColumn;
			
			@FindBy(xpath="//*[@id='3']/td[6]")
			private static WebElement wh2ndrow_6thColumn;
			
			
			
			@FindBy(xpath="//*[@id='4']/td[2]")
			private static WebElement wh3rdrow_2ndColumn;
			
			@FindBy(xpath="//*[@id='4']/td[3]")
			private static WebElement wh3rdrow_3rdColumn;
			
			@FindBy(xpath="//*[@id='4']/td[4]")
			private static WebElement wh3rdrow_4thColumn;
			
			@FindBy(xpath="//*[@id='4']/td[5]")
			private static WebElement wh3rdrow_5thColumn;
			
			@FindBy(xpath="//*[@id='4']/td[6]")
			private static WebElement wh3rdrow_6thColumn;
			
			
			@FindBy(xpath="//*[@id='8']/td[2]")
			private static WebElement wh4throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='8']/td[3]")
			private static WebElement wh4throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='8']/td[4]")
			private static WebElement wh4throw_4thColumn;
			
			@FindBy(xpath="//*[@id='8']/td[5]")
			private static WebElement wh4throw_5thColumn;
			
			@FindBy(xpath="//*[@id='8']/td[6]")
			private static WebElement wh4throw_6thColumn;
			
			
			@FindBy(xpath="//*[@id='9']/td[2]")
			private static WebElement wh5throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='9']/td[3]")
			private static WebElement wh5throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='9']/td[4]")
			private static WebElement wh5throw_4thColumn;
			
			@FindBy(xpath="//*[@id='9']/td[5]")
			private static WebElement wh5throw_5thColumn;
			
			@FindBy(xpath="//*[@id='9']/td[6]")
			private static WebElement wh5throw_6thColumn;
			
			
			
			@FindBy(xpath="//*[@id='tagWiseData']/tbody/tr[2]/td[2]")
			private static WebElement whStockTotal;
			
			@FindBy(xpath="//*[@id='tagWiseData']/tbody/tr[2]/td[4]")
			private static WebElement whValueTotal;
			
			@FindBy(xpath="//table[@id='tagWiseData']/tbody")
			private static WebElement itemQueryWHTable;
			
			
			// Selling rate and buying rates
			
			@FindBy(xpath="//div[contains(text(),'Selling rates and Buying rates')]")
			private static WebElement sellingAndBuyingRateslable;
			
			@FindBy(xpath="//span[@id='ProdQuery_ShowSellandBuyRatesPopup']")
			private static WebElement sellingAndBuyingRatesExpandBtn;
			
			
			// 1st 
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[1]/td[2]")
			private static WebElement sbr1strow_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[1]/td[3]")
			private static WebElement sbr1strow_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[1]/td[4]")
			private static WebElement sbr1strow_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[1]/td[5]")
			private static WebElement sbr1strow_5thColumn;
			
			// 2nd 
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[2]/td[2]")
			private static WebElement sbr2ndrow_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[2]/td[3]")
			private static WebElement sbr2ndrow_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[2]/td[4]")
			private static WebElement sbr2ndrow_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[2]/td[5]")
			private static WebElement sbr2ndrow_5thColumn;
			
			//3rd
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[3]/td[2]")
			private static WebElement sbr3rdrow_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[3]/td[3]")
			private static WebElement sbr3rdrow_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[3]/td[4]")
			private static WebElement sbr3rdrow_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[3]/td[5]")
			private static WebElement sbr3rdrow_5thColumn;
			
			// 4th
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[4]/td[2]")
			private static WebElement sbr4throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[4]/td[3]")
			private static WebElement sbr4throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[4]/td[4]")
			private static WebElement sbr4throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[4]/td[5]")
			private static WebElement sbr4throw_5thColumn;
			
			
			
			// 5th
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[5]/td[2]")
			private static WebElement sbr5throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[5]/td[3]")
			private static WebElement sbr5throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[5]/td[4]")
			private static WebElement sbr5throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[5]/td[5]")
			private static WebElement sbr5throw_5thColumn;
				
				
			// 6th
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[6]/td[2]")
			private static WebElement sbr6throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[6]/td[3]")
			private static WebElement sbr6throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[6]/td[4]")
			private static WebElement sbr6throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[6]/td[5]")				
			private static WebElement sbr6throw_5thColumn;
				
			
			
			// 7th

			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[7]/td[2]")
			private static WebElement sbr7throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[7]/td[3]")
			private static WebElement sbr7throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[7]/td[4]")
			private static WebElement sbr7throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[7]/td[5]")
			private static WebElement sbr7throw_5thColumn;
				
				
			// 8th
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[8]/td[2]")
			private static WebElement sbr8throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[8]/td[3]")
			private static WebElement sbr8throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[8]/td[4]")
			private static WebElement sbr8throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[8]/td[5]")
			private static WebElement sbr8throw_5thColumn;
				
				
			
			// 9th
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[9]/td[2]")
			private static WebElement sbr9throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[9]/td[3]")
			private static WebElement sbr9throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[9]/td[4]")
			private static WebElement sbr9throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[9]/td[5]")
			private static WebElement sbr9throw_5thColumn;
				
				
			// 10th
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[10]/td[2]")
			private static WebElement sbr10throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[10]/td[3]")
			private static WebElement sbr10throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[10]/td[4]")
			private static WebElement sbr10throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[10]/td[5]")
			private static WebElement sbr10throw_5thColumn;
			
			
			// 11th
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[11]/td[2]")
			private static WebElement sbr11throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[11]/td[3]")
			private static WebElement sbr11throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[11]/td[4]")
			private static WebElement sbr11throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[11]/td[5]")
			private static WebElement sbr11throw_5thColumn;
			
			
			
			// 12th
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[12]/td[2]")
			private static WebElement sbr12throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[12]/td[3]")
			private static WebElement sbr12throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[12]/td[4]")
			private static WebElement sbr12throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[12]/td[5]")
			private static WebElement sbr12throw_5thColumn;
			
			
			
			// 13th
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[13]/td[2]")
			private static WebElement sbr13throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[13]/td[3]")
			private static WebElement sbr13throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[13]/td[4]")
			private static WebElement sbr13throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[13]/td[5]")
			private static WebElement sbr13throw_5thColumn;
			
			
			// 14th
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[14]/td[2]")
			private static WebElement sbr14throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[14]/td[3]")
			private static WebElement sbr14throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[14]/td[4]")
			private static WebElement sbr14throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[14]/td[5]")
			private static WebElement sbr14throw_5thColumn;
			
			
			
			
			@FindBy(xpath="//tbody[@id='LandingGridBody']/tr")
			private static List<WebElement> stockLedgerHometableRowCount;
			
			
			 // Reports Table Row List  
		  	@FindBy(xpath="//tbody[@id='tblBodyReportRender']/tr[1]/td")
			private static List<WebElement> reportsRow1List;
			
			@FindBy(xpath="//tbody[@id='tblBodyReportRender']/tr[2]/td")
			private static List<WebElement> reportsRow2List;
			
			@FindBy(xpath="//tbody[@id='tblBodyReportRender']/tr[3]/td")
			private static List<WebElement> reportsRow3List;
			
			@FindBy(xpath="//tbody[@id='tblBodyReportRender']/tr[4]/td")
			private static List<WebElement> reportsRow4List;
			
			@FindBy(xpath="//tbody[@id='tblBodyReportRender']/tr[5]/td")
			private static List<WebElement> reportsRow5List;
			
			@FindBy(xpath="//tbody[@id='tblBodyReportRender']/tr[6]/td")
			private static List<WebElement> reportsRow6List;
			
			@FindBy(xpath="//tbody[@id='tblBodyReportRender']/tr[7]/td")
			private static List<WebElement> reportsRow7List;
			
			@FindBy(xpath="//tbody[@id='tblBodyReportRender']/tr[8]/td")
			private static List<WebElement> reportsRow8List;
			
			@FindBy(xpath="//tbody[@id='tblBodyReportRender']/tr[9]/td")
			private static List<WebElement> reportsRow9List;
			
			@FindBy(xpath="//tbody[@id='tblBodyReportRender']/tr[10]/td")
			private static List<WebElement> reportsRow10List;
			
			@FindBy(xpath="//tbody[@id='tblBodyReportRender']/tr[11]/td")
			private static List<WebElement> reportsRow11List;
			
			@FindBy(xpath="//tbody[@id='tblBodyReportRender']/tr[12]/td")
			private static List<WebElement> reportsRow12List;
			
			@FindBy(xpath="//tbody[@id='tblBodyReportRender']/tr[13]/td")
			private static List<WebElement> reportsRow13List;
			
			@FindBy(xpath="//tbody[@id='tblBodyReportRender']/tr[14]/td")
			private static List<WebElement> reportsRow14List;
			
			
			
			@FindBy(xpath="//tbody[@id='tblBodyReportRender']")
			private static WebElement reportsTable;
			
			
			
			@FindBy(xpath="//select[@id='RITCombobox__1']")
			private static WebElement stockMovementDropdown;
			
			
			@FindBy(xpath="//input[@id='RITCheckbox__3']")
			private static WebElement includeAllVoucherChkBox;
			
			
			@FindBy(xpath="//span[contains(text(),'Actual consumption report')]")
			private static WebElement  actualConsumptionReport;
			
		
		
			
			
	//*******************************************************Methods Starts Here**************************************************************		
			

			
	
	
	
	
	
		public static boolean checkLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());

			LoginPage lp=new LoginPage(getDriver()); 
			
		    String unamelt= excelReader.getCellData(xlSheetName, 7, 5);
		  
		    String pawslt= excelReader.getCellData(xlSheetName, 8, 5);
		      
		    lp.enterUserName(unamelt);
		    
		    lp.enterPassword(pawslt);
		    
		    String compname=excelReader.getCellData(xlSheetName, 9, 5);

			 Select oSelect = new Select(companyDropDownList);
			 
			 List <WebElement> elementCount = oSelect.getOptions();

			 int cqSize = elementCount.size();
			 
			 System.out.println("CompanyDropdownList Count :"+cqSize);

			    int i;
			  
				for(i=0; i<elementCount.size(); i++) 
				{
				
				  elementCount.get(i).getText();
				  
			      String optionName = elementCount.get(i).getText();
			     if(optionName.toUpperCase().startsWith(compname.toUpperCase()))
			     {
				  System.out.println("Company List"+elementCount.get(i).getText());
					  elementCount.get(i).click();
				 }
			 
			    }

		    lp.clickOnSignInBtn();
		    
		    lp.reLogin(unamelt, pawslt, compname);
	
			Thread.sleep(8000);
			
			String actdashboardFooter=dashbordPageFooterFiledTxt.getText().substring(0, 26);
			String expdashboardFooter= excelReader.getCellData(xlSheetName, 10, 6);
			excelReader.setCellData(xlfile, xlSheetName, 10, 7, actdashboardFooter);
			
			
			System.out.println("ACT  :  "+actdashboardFooter+"*");
			System.out.println("Exp  :  "+expdashboardFooter+"*");
			
			if (actdashboardFooter.equalsIgnoreCase(expdashboardFooter)) 
			{
				 excelReader.setCellData(xlfile, xlSheetName, 6, 8, resPass);
				return true;
			}
			else
			{
				 excelReader.setCellData(xlfile, xlSheetName, 6, 8, resFail);
				return false;
			}
		}
		
			
	
		
		
		public static boolean checkLoginAndRestoreOptionsCompanyAndLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
    	{
    		LoginPage lp=new LoginPage(getDriver()); 
    		
    	    String unamelt="su";
    	  
    	    String pawslt="su";
    	      
    	    lp.enterUserName(unamelt);
    	    
    	    lp.enterPassword(pawslt);
    	    
    	    String compname="Automation Options Company";

   		 Select oSelect = new Select(companyDropDownList);
   		 
   		 List <WebElement> elementCount = oSelect.getOptions();

   		 int cqSize = elementCount.size();
   		 
   		 System.out.println("CompanyDropdownList Count :"+cqSize);

   		    int i;
   		  
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
    	  
    		Thread.sleep(8000);
    		
    		String actUserInfo=userNameDisplayLogo.getText();
    		
    		System.out.println("User Info  : "+actUserInfo);
    		
    		System.out.println("User Info Capture Text  :  "+userNameDisplayLogo.getText());
    		
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
    	     
    	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
    	     userNameDisplayLogo.click();
    		
    		
    	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
    	     logoutOption.click();

    	     Thread.sleep(3000);
    	     
    	     lp.enterUserName(unamelt);
    	     
    	     lp.enterPassword(pawslt);
    	     
    	     Thread.sleep(2000);
    	     
    	    
    		 System.out.println("CompanyDropdownList Count :"+cqSize);

    		    
    		  
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
    	     
    	     Thread.sleep(8000);
    	     
    	     String actUserInfo1=userNameDisplayLogo.getText();
    			
    		System.out.println("User Info  : "+actUserInfo1);
    		
    		System.out.println("User Info Capture Text  :  "+userNameDisplayLogo.getText());
    		
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
		
		
	
	
	
			
		public boolean checkOpeningStockNewSavingVoucher1WithWACOGSITEMWithWarehouseAsHyderabad() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksNewVoucher));
			openingStocksNewVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String actScreenOpened = checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
			warehouseTxt.click();
			warehouseTxt.sendKeys(Keys.SPACE);
            
			int warehouselist=warehouseHeaderComboList.size();
			
            for (int i = 0; i < warehouselist; i++) 
			{
				String warehouse=warehouseHeaderComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("HYDERABAD")) 
				{
					warehouseHeaderComboList.get(i).click();
					warehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					enter_ItemTxt.sendKeys(Keys.TAB);
					break;
				}
			}
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			enter_UnitTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("20");
			enter_Quantity.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("8.5");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys("Q,20");
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(2000);
			
			//actual
			String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
			String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
			String actRMA3=RMAPopupTableRow3Column1.getAttribute("data-value");
			String actRMA4=RMAPopupTableRow4Column1.getAttribute("data-value");
			String actRMA5=RMAPopupTableRow5Column1.getAttribute("data-value");
			String actRMA6=RMAPopupTableRow6Column1.getAttribute("data-value");
			String actRMA7=RMAPopupTableRow7Column1.getAttribute("data-value");
			String actRMA8=RMAPopupTableRow8Column1.getAttribute("data-value");
			String actRMA9=RMAPopupTableRow9Column1.getAttribute("data-value");
			String actRMA10=RMAPopupTableRow10Column1.getAttribute("data-value");
			String actRMA11=RMAPopupTableRow11Column1.getAttribute("data-value");
			String actRMA12=RMAPopupTableRow12Column1.getAttribute("data-value");
			String actRMA13=RMAPopupTableRow13Column1.getAttribute("data-value");
			String actRMA14=RMAPopupTableRow14Column1.getAttribute("data-value");
			String actRMA15=RMAPopupTableRow15Column1.getAttribute("data-value");
			String actRMA16=RMAPopupTableRow16Column1.getAttribute("data-value");
			String actRMA17=RMAPopupTableRow17Column1.getAttribute("data-value");
			String actRMA18=RMAPopupTableRow18Column1.getAttribute("data-value");
			String actRMA19=RMAPopupTableRow19Column1.getAttribute("data-value");
			String actRMA20=RMAPopupTableRow20Column1.getAttribute("data-value");
			String actDocNo=RMAPopupTableRow1Column2.getText();
			
			
			
			//expected
			String expRMA1="Q";
			String expRMA2="Q1";
			String expRMA3="Q2";
			String expRMA4="Q3";
			String expRMA5="Q4";
			String expRMA6="Q5";
			String expRMA7="Q6";
			String expRMA8="Q7";
			String expRMA9="Q8";
			String expRMA10="Q9";
			String expRMA11="Q10";
			String expRMA12="Q11";
			String expRMA13="Q12";
			String expRMA14="Q13";
			String expRMA15="Q14";
			String expRMA16="Q15";
			String expRMA17="Q16";
			String expRMA18="Q17";
			String expRMA19="Q18";
			String expRMA20="Q19";
			
			String expDocNo="NDT44:1";
			
			
			System.out.println("********* Row1        : "+actRMA1  +"  value expected  "+expRMA1);
			System.out.println("********* Row2        : "+actRMA2  +"  value expected  "+expRMA2);
			System.out.println("********* Row3        : "+actRMA3  +"  value expected  "+expRMA3);
			System.out.println("********* Row4        : "+actRMA4  +"  value expected  "+expRMA4);
			System.out.println("********* Row5        : "+actRMA5  +"  value expected  "+expRMA5);
			System.out.println("********* Row6        : "+actRMA6  +"  value expected  "+expRMA6);
			System.out.println("********* Row7        : "+actRMA7  +"  value expected  "+expRMA7);
			System.out.println("********* Row8        : "+actRMA8  +"  value expected  "+expRMA8);
			System.out.println("********* Row9        : "+actRMA9  +"  value expected  "+expRMA9);
			System.out.println("********* Row10       : "+actRMA10 +"  value expected  "+expRMA10);
			System.out.println("********* Row11       : "+actRMA11 +"  value expected  "+expRMA11);
			System.out.println("********* Row12       : "+actRMA12 +"  value expected  "+expRMA12);
			System.out.println("********* Row13       : "+actRMA13 +"  value expected  "+expRMA13);
			System.out.println("********* Row14       : "+actRMA14 +"  value expected  "+expRMA14);
			System.out.println("********* Row15       : "+actRMA15 +"  value expected  "+expRMA15);
			System.out.println("********* Row16       : "+actRMA16 +"  value expected  "+expRMA16);
			System.out.println("********* Row17       : "+actRMA17 +"  value expected  "+expRMA17);
			System.out.println("********* Row18       : "+actRMA18 +"  value expected  "+expRMA18);
			System.out.println("********* Row19       : "+actRMA19 +"  value expected  "+expRMA19);
			System.out.println("********* Row20       : "+actRMA20 +"  value expected  "+expRMA20);
			System.out.println("********* Document No : "+actDocNo +"  value expected  "+expDocNo);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
		
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			
			String actRMAColumn=enter_RMA.getAttribute("value");
			
			String expRMAColumn="Q,Q1,Q2,Q3,Q4,Q5,Q6,Q7,Q8,Q9,Q10,Q11,Q12,Q13,Q14,Q15,Q16,Q17,Q18,Q19";
			
			System.out.println("******** RMA Column: "+actRMAColumn+"  value expected  "+expRMAColumn);
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
	
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  
			
			if(actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) 
				&& actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)	
				&& actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10) && actRMA11.equalsIgnoreCase(expRMA11) && actRMA12.equalsIgnoreCase(expRMA12) 
				&& actRMA13.equalsIgnoreCase(expRMA13) && actRMA14.equalsIgnoreCase(expRMA14) && actRMA15.equalsIgnoreCase(expRMA15) && actRMA16.equalsIgnoreCase(expRMA16)	
				&& actRMA17.equalsIgnoreCase(expRMA17) && actRMA18.equalsIgnoreCase(expRMA18) && actRMA19.equalsIgnoreCase(expRMA19) && actRMA20.equalsIgnoreCase(expRMA20) 
				
				&& actRMAColumn.equalsIgnoreCase(expRMAColumn)
			    && actSaving==expSaving)
			{
				return true;
			}
			else
			{
				return false;
			}	
		}
		

		
		
		
		
		public boolean checkOpeningStocksNewPreviousButtonAndVerifyDataInSavedVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
			previousBtn.click();
			
			String expMessage="Loading Message";
			
			String actMessage=checkValidationMessage(expMessage);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			//actual
			String actDocNo=documentNumberTxt.getAttribute("value");
			String actWarehouse=warehouseTxt.getAttribute("value");
			String actItem=select1stRow_1stColumn.getText();
			String actQuantity=select1stRow_3rdColumn.getText();
			String actRate=select1stRow_4thColumn.getText();
			String actGross=select1stRow_5thColumn.getText();
			String actRMA=select1stRow_9thColumn.getText();

			
			//expected
			String expDocNo="1";
			String expWarehouse="HYDERABAD";
			String expItem="WA COGS ITEM";
			String expQuantity="20.00";
			String expRate="8.50";
			String expGross="170.00";
			String expRMA="Q,Q1,Q2,Q3,Q4,Q5,Q6,Q7,Q8,Q9,Q10,Q11,Q12,Q13,Q14,Q15,Q16,Q17,Q18,Q19";
			
			
			System.out.println("********** Voucher No  : "+actDocNo      +"  value expected  "+expDocNo);
			System.out.println("********** Warehouse   : "+actWarehouse  +"  value expected  "+expWarehouse);
			System.out.println("********** Item        : "+actItem       +"  value expected  "+expItem);
			System.out.println("********** Quantity    : "+actQuantity   +"  value expected  "+expQuantity);
			System.out.println("********** Rate        : "+actRate       +"  value expected  "+expRate);
			System.out.println("********** Gross       : "+actGross      +"  value expected  "+expGross);
			System.out.println("********** RMA Column  : "+actRMA        +"  value expected  "+expRMA);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupExpansionBtn));
			RMAPopupExpansionBtn.click();
			
			Thread.sleep(2000);
			
			//actual
			String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
			String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
			String actRMA3=RMAPopupTableRow3Column1.getAttribute("data-value");
			String actRMA4=RMAPopupTableRow4Column1.getAttribute("data-value");
			String actRMA5=RMAPopupTableRow5Column1.getAttribute("data-value");
			String actRMA6=RMAPopupTableRow6Column1.getAttribute("data-value");
			String actRMA7=RMAPopupTableRow7Column1.getAttribute("data-value");
			String actRMA8=RMAPopupTableRow8Column1.getAttribute("data-value");
			String actRMA9=RMAPopupTableRow9Column1.getAttribute("data-value");
			String actRMA10=RMAPopupTableRow10Column1.getAttribute("data-value");
			String actRMA11=RMAPopupTableRow11Column1.getAttribute("data-value");
			String actRMA12=RMAPopupTableRow12Column1.getAttribute("data-value");
			String actRMA13=RMAPopupTableRow13Column1.getAttribute("data-value");
			String actRMA14=RMAPopupTableRow14Column1.getAttribute("data-value");
			String actRMA15=RMAPopupTableRow15Column1.getAttribute("data-value");
			String actRMA16=RMAPopupTableRow16Column1.getAttribute("data-value");
			String actRMA17=RMAPopupTableRow17Column1.getAttribute("data-value");
			String actRMA18=RMAPopupTableRow18Column1.getAttribute("data-value");
			String actRMA19=RMAPopupTableRow19Column1.getAttribute("data-value");
			String actRMA20=RMAPopupTableRow20Column1.getAttribute("data-value");
			String actRMADocNo=RMAPopupTableRow1Column2.getText();
			
			
			//expected
			String expRMA1="Q";
			String expRMA2="Q1";
			String expRMA3="Q2";
			String expRMA4="Q3";
			String expRMA5="Q4";
			String expRMA6="Q5";
			String expRMA7="Q6";
			String expRMA8="Q7";
			String expRMA9="Q8";
			String expRMA10="Q9";
			String expRMA11="Q10";
			String expRMA12="Q11";
			String expRMA13="Q12";
			String expRMA14="Q13";
			String expRMA15="Q14";
			String expRMA16="Q15";
			String expRMA17="Q16";
			String expRMA18="Q17";
			String expRMA19="Q18";
			String expRMA20="Q19";
			String expRMADocNo="NDT44:1";
			
			
			System.out.println("********* Row1         : "+actRMA1     +"  value expected  "+expRMA1);
			System.out.println("********* Row2         : "+actRMA2     +"  value expected  "+expRMA2);
			System.out.println("********* Row3         : "+actRMA3     +"  value expected  "+expRMA3);
			System.out.println("********* Row4         : "+actRMA4     +"  value expected  "+expRMA4);
			System.out.println("********* Row5         : "+actRMA5     +"  value expected  "+expRMA5);
			System.out.println("********* Row6         : "+actRMA6     +"  value expected  "+expRMA6);
			System.out.println("********* Row7         : "+actRMA7     +"  value expected  "+expRMA7);
			System.out.println("********* Row8         : "+actRMA8     +"  value expected  "+expRMA8);
			System.out.println("********* Row9         : "+actRMA9     +"  value expected  "+expRMA9);
			System.out.println("********* Row10        : "+actRMA10    +"  value expected  "+expRMA10);
			System.out.println("********* Row11        : "+actRMA11    +"  value expected  "+expRMA11);
			System.out.println("********* Row12        : "+actRMA12    +"  value expected  "+expRMA12);
			System.out.println("********* Row13        : "+actRMA13    +"  value expected  "+expRMA13);
			System.out.println("********* Row14        : "+actRMA14    +"  value expected  "+expRMA14);
			System.out.println("********* Row15        : "+actRMA15    +"  value expected  "+expRMA15);
			System.out.println("********* Row16        : "+actRMA16    +"  value expected  "+expRMA16);
			System.out.println("********* Row17        : "+actRMA17    +"  value expected  "+expRMA17);
			System.out.println("********* Row18        : "+actRMA18    +"  value expected  "+expRMA18);
			System.out.println("********* Row19        : "+actRMA19    +"  value expected  "+expRMA19);
			System.out.println("********* Row20        : "+actRMA20    +"  value expected  "+expRMA20);
			System.out.println("********* Document No  : "+actRMADocNo +"  value expected  "+expRMADocNo);
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupCancelBtn));
			RMAPopupCancelBtn.click();
			
			Thread.sleep(2000);
							
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(savedVoucher1stRowVoucherNumer));
			String actVoucherHomeDocNo = savedVoucher1stRowVoucherNumer.getText();
			String expVoucherHomeDocNo = "1";
			
			System.out.println("**************************** Saved Voucher No in Homepage: "+actVoucherHomeDocNo+" Value Expected "+expVoucherHomeDocNo);
			
			if(actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) && actItem.equalsIgnoreCase(expItem) 
				&& actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross) 
				&& actRMA.equalsIgnoreCase(expRMA)
				
				&& actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) 
				&& actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)	
				&& actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10) && actRMA11.equalsIgnoreCase(expRMA11) && actRMA12.equalsIgnoreCase(expRMA12) 
				&& actRMA13.equalsIgnoreCase(expRMA13) && actRMA14.equalsIgnoreCase(expRMA14) && actRMA15.equalsIgnoreCase(expRMA15) && actRMA16.equalsIgnoreCase(expRMA16)	
				&& actRMA17.equalsIgnoreCase(expRMA17) && actRMA18.equalsIgnoreCase(expRMA18) && actRMA19.equalsIgnoreCase(expRMA19) && actRMA20.equalsIgnoreCase(expRMA20)
				
				&& actVoucherHomeDocNo.equalsIgnoreCase(expVoucherHomeDocNo))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		
		
		
		
		
		
		
		public boolean checkOpeningStockNewSavingVoucher2WithWACOGSITEMWithWarehouseAsSecunderabad() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String actScreenOpened = checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
			warehouseTxt.click();
			warehouseTxt.sendKeys(Keys.SPACE);
            
			int warehouselist=warehouseHeaderComboList.size();
			
            for (int i = 0; i < warehouselist; i++) 
			{
				String warehouse=warehouseHeaderComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("SECUNDERABAD")) 
				{
					warehouseHeaderComboList.get(i).click();
					warehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					enter_ItemTxt.sendKeys(Keys.TAB);
					break;
				}
			}
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			enter_UnitTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("10");
			enter_Quantity.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("2");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys("SEC#RMA1,10");
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(2000);
			
			//actual
			String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
			String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
			String actRMA3=RMAPopupTableRow3Column1.getAttribute("data-value");
			String actRMA4=RMAPopupTableRow4Column1.getAttribute("data-value");
			String actRMA5=RMAPopupTableRow5Column1.getAttribute("data-value");
			String actRMA6=RMAPopupTableRow6Column1.getAttribute("data-value");
			String actRMA7=RMAPopupTableRow7Column1.getAttribute("data-value");
			String actRMA8=RMAPopupTableRow8Column1.getAttribute("data-value");
			String actRMA9=RMAPopupTableRow9Column1.getAttribute("data-value");
			String actRMA10=RMAPopupTableRow10Column1.getAttribute("data-value");
			String actDocNo=RMAPopupTableRow1Column2.getText();
			
			
			
			//expected
			String expRMA1="SEC#RMA1";
			String expRMA2="SEC#RMA2";
			String expRMA3="SEC#RMA3";
			String expRMA4="SEC#RMA4";
			String expRMA5="SEC#RMA5";
			String expRMA6="SEC#RMA6";
			String expRMA7="SEC#RMA7";
			String expRMA8="SEC#RMA8";
			String expRMA9="SEC#RMA9";
			String expRMA10="SEC#RMA10";
			String expDocNo="NDT44:2";
			
			
			System.out.println("********* Row1        : "+actRMA1  +"  value expected  "+expRMA1);
			System.out.println("********* Row2        : "+actRMA2  +"  value expected  "+expRMA2);
			System.out.println("********* Row3        : "+actRMA3  +"  value expected  "+expRMA3);
			System.out.println("********* Row4        : "+actRMA4  +"  value expected  "+expRMA4);
			System.out.println("********* Row5        : "+actRMA5  +"  value expected  "+expRMA5);
			System.out.println("********* Row6        : "+actRMA6  +"  value expected  "+expRMA6);
			System.out.println("********* Row7        : "+actRMA7  +"  value expected  "+expRMA7);
			System.out.println("********* Row8        : "+actRMA8  +"  value expected  "+expRMA8);
			System.out.println("********* Row9        : "+actRMA9  +"  value expected  "+expRMA9);
			System.out.println("********* Row10       : "+actRMA10 +"  value expected  "+expRMA10);
			System.out.println("********* Document No : "+actDocNo +"  value expected  "+expDocNo);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
		
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			
			String actRMAColumn=enter_RMA.getAttribute("value");
			
			String expRMAColumn="SEC#RMA1,SEC#RMA2,SEC#RMA3,SEC#RMA4,SEC#RMA5,SEC#RMA6,SEC#RMA7,SEC#RMA8,SEC#RMA9,SEC#RMA10";
			
			System.out.println("******** RMA Column: "+actRMAColumn+"  value expected  "+expRMAColumn);
			
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
	
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
		
			if(actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) 
				&& actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)	
				&& actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10)  
				
				&& actRMAColumn.equalsIgnoreCase(expRMAColumn)
			    && actSaving==expSaving)
			{
				return true;
			}
			else
			{
				return false;
			}	
		}
		
		
		
		
		
		
		
		public boolean checkOpeningStocksNewPreviousButtonAndVerifyDataInSavedVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
			previousBtn.click();
			
			String expMessage="Loading Message";
			
			String actMessage=checkValidationMessage(expMessage);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			//actual
			String actDocNo=documentNumberTxt.getAttribute("value");
			String actWarehouse=warehouseTxt.getAttribute("value");
			String actItem=select1stRow_1stColumn.getText();
			String actQuantity=select1stRow_3rdColumn.getText();
			String actRate=select1stRow_4thColumn.getText();
			String actGross=select1stRow_5thColumn.getText();
			String actRMA=select1stRow_9thColumn.getText();

			
			//expected
			String expDocNo="2";
			String expWarehouse="SECUNDERABAD";
			String expItem="WA COGS ITEM";
			String expQuantity="10.00";
			String expRate="2.00";
			String expGross="20.00";
			String expRMA="SEC#RMA1,SEC#RMA2,SEC#RMA3,SEC#RMA4,SEC#RMA5,SEC#RMA6,SEC#RMA7,SEC#RMA8,SEC#RMA9,SEC#RMA10";
			
			
			System.out.println("********** Voucher No  : "+actDocNo      +"  value expected  "+expDocNo);
			System.out.println("********** Warehouse   : "+actWarehouse  +"  value expected  "+expWarehouse);
			System.out.println("********** Item        : "+actItem       +"  value expected  "+expItem);
			System.out.println("********** Quantity    : "+actQuantity   +"  value expected  "+expQuantity);
			System.out.println("********** Rate        : "+actRate       +"  value expected  "+expRate);
			System.out.println("********** Gross       : "+actGross      +"  value expected  "+expGross);
			System.out.println("********** RMA Column  : "+actRMA        +"  value expected  "+expRMA);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupExpansionBtn));
			RMAPopupExpansionBtn.click();
			
			Thread.sleep(2000);
			
			//actual
			String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
			String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
			String actRMA3=RMAPopupTableRow3Column1.getAttribute("data-value");
			String actRMA4=RMAPopupTableRow4Column1.getAttribute("data-value");
			String actRMA5=RMAPopupTableRow5Column1.getAttribute("data-value");
			String actRMA6=RMAPopupTableRow6Column1.getAttribute("data-value");
			String actRMA7=RMAPopupTableRow7Column1.getAttribute("data-value");
			String actRMA8=RMAPopupTableRow8Column1.getAttribute("data-value");
			String actRMA9=RMAPopupTableRow9Column1.getAttribute("data-value");
			String actRMA10=RMAPopupTableRow10Column1.getAttribute("data-value");
			String actRMADocNo=RMAPopupTableRow1Column2.getText();
			
			
			//expected
			String expRMA1="SEC#RMA1";
			String expRMA2="SEC#RMA2";
			String expRMA3="SEC#RMA3";
			String expRMA4="SEC#RMA4";
			String expRMA5="SEC#RMA5";
			String expRMA6="SEC#RMA6";
			String expRMA7="SEC#RMA7";
			String expRMA8="SEC#RMA8";
			String expRMA9="SEC#RMA9";
			String expRMA10="SEC#RMA10";
			String expRMADocNo="NDT44:2";
			
			
			System.out.println("********* Row1         : "+actRMA1     +"  value expected  "+expRMA1);
			System.out.println("********* Row2         : "+actRMA2     +"  value expected  "+expRMA2);
			System.out.println("********* Row3         : "+actRMA3     +"  value expected  "+expRMA3);
			System.out.println("********* Row4         : "+actRMA4     +"  value expected  "+expRMA4);
			System.out.println("********* Row5         : "+actRMA5     +"  value expected  "+expRMA5);
			System.out.println("********* Row6         : "+actRMA6     +"  value expected  "+expRMA6);
			System.out.println("********* Row7         : "+actRMA7     +"  value expected  "+expRMA7);
			System.out.println("********* Row8         : "+actRMA8     +"  value expected  "+expRMA8);
			System.out.println("********* Row9         : "+actRMA9     +"  value expected  "+expRMA9);
			System.out.println("********* Row10        : "+actRMA10    +"  value expected  "+expRMA10);
			System.out.println("********* Document No  : "+actRMADocNo +"  value expected  "+expRMADocNo);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupCancelBtn));
			RMAPopupCancelBtn.click();
			
			Thread.sleep(2000);
							
			if(actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) && actItem.equalsIgnoreCase(expItem) 
				&& actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross) 
				&& actRMA.equalsIgnoreCase(expRMA)
				
				&& actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) 
				&& actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)	
				&& actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		
		
		
		
		
	
		public boolean checkOpeningStockNewEditingVoucher2WithWACOGSITEMWithWarehouseAsSecunderabad() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
			select1stRow_3rdColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys(Keys.END);
			enter_Quantity.sendKeys(Keys.SHIFT, Keys.HOME);
			enter_Quantity.sendKeys("12");
			enter_Quantity.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("2");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.SPACE);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupClearBtn));
			RMAPopupClearBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys("SEC#RMA1,12");
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(2000);
			
			//actual
			String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
			String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
			String actRMA3=RMAPopupTableRow3Column1.getAttribute("data-value");
			String actRMA4=RMAPopupTableRow4Column1.getAttribute("data-value");
			String actRMA5=RMAPopupTableRow5Column1.getAttribute("data-value");
			String actRMA6=RMAPopupTableRow6Column1.getAttribute("data-value");
			String actRMA7=RMAPopupTableRow7Column1.getAttribute("data-value");
			String actRMA8=RMAPopupTableRow8Column1.getAttribute("data-value");
			String actRMA9=RMAPopupTableRow9Column1.getAttribute("data-value");
			String actRMA10=RMAPopupTableRow10Column1.getAttribute("data-value");
			String actRMA11=RMAPopupTableRow11Column1.getAttribute("data-value");
			String actRMA12=RMAPopupTableRow12Column1.getAttribute("data-value");
			String actDocNo=RMAPopupTableRow1Column2.getText();
			
			
			
			//expected
			String expRMA1="SEC#RMA1";
			String expRMA2="SEC#RMA2";
			String expRMA3="SEC#RMA3";
			String expRMA4="SEC#RMA4";
			String expRMA5="SEC#RMA5";
			String expRMA6="SEC#RMA6";
			String expRMA7="SEC#RMA7";
			String expRMA8="SEC#RMA8";
			String expRMA9="SEC#RMA9";
			String expRMA10="SEC#RMA10";
			String expRMA11="SEC#RMA11";
			String expRMA12="SEC#RMA12";
			String expDocNo="NDT44:2";
			
			
			System.out.println("********* Row1        : "+actRMA1  +"  value expected  "+expRMA1);
			System.out.println("********* Row2        : "+actRMA2  +"  value expected  "+expRMA2);
			System.out.println("********* Row3        : "+actRMA3  +"  value expected  "+expRMA3);
			System.out.println("********* Row4        : "+actRMA4  +"  value expected  "+expRMA4);
			System.out.println("********* Row5        : "+actRMA5  +"  value expected  "+expRMA5);
			System.out.println("********* Row6        : "+actRMA6  +"  value expected  "+expRMA6);
			System.out.println("********* Row7        : "+actRMA7  +"  value expected  "+expRMA7);
			System.out.println("********* Row8        : "+actRMA8  +"  value expected  "+expRMA8);
			System.out.println("********* Row9        : "+actRMA9  +"  value expected  "+expRMA9);
			System.out.println("********* Row10       : "+actRMA10 +"  value expected  "+expRMA10);
			System.out.println("********* Row11       : "+actRMA11 +"  value expected  "+expRMA11);
			System.out.println("********* Row12       : "+actRMA12 +"  value expected  "+expRMA12);
			System.out.println("********* Document No : "+actDocNo +"  value expected  "+expDocNo);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
		
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			
			String actRMAColumn=enter_RMA.getAttribute("value");
			
			String expRMAColumn="SEC#RMA1,SEC#RMA2,SEC#RMA3,SEC#RMA4,SEC#RMA5,SEC#RMA6,SEC#RMA7,SEC#RMA8,SEC#RMA9,SEC#RMA10,SEC#RMA11,SEC#RMA12";
			
			System.out.println("******** RMA Column: "+actRMAColumn+"  value expected  "+expRMAColumn);
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
	
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
		
			
			
			if(actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) 
				&& actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)	
				&& actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10) && actRMA11.equalsIgnoreCase(expRMA11) && actRMA12.equalsIgnoreCase(expRMA12) 
				
				&& actRMAColumn.equalsIgnoreCase(expRMAColumn)
			    && actSaving==expSaving)
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}	
		}
		
		
		
		
		

		
		// Sales Invoice
		
		public boolean checkSalesInvoiceVATVoucher1WithWarehouseAsHyderabadAndItemAsWACOGSITEM() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String actScreenOpened=checkValidationMessage("Screen opened");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			System.out.println("Customer Account List Size: "+customerAccountCount);
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase("Customer A"))
				{
					customerAccountComboList.get(i).click();
					
					customerAccountTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase("DUBAI"))
				{
					departmentComboList.get(i).click();
					
					departmentTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys("Abu Dhabi");
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("HYDERABAD"))
				{
					warehouseBodyComboList.get(i).click();
					
					pvWarehouseTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys("1");
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys("1");
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("19");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
			select1stRow_21stColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
			searchRMASelectAllChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA=enter_RMA.getAttribute("value");
			
			String expRMA="Q,Q1";
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
			String actNewReference=newReferenceTxt.getAttribute("value");
			
			String expNewReference="39.90";

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			
			
			System.out.println("****** RMA Column                 : "+actRMA			+"  value expected  "+expRMA);
			System.out.println("****** New Reference              : "+actNewReference	+"  value expected  "+expNewReference);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actNewReference.equalsIgnoreCase(expNewReference) 
				&& actSaving==expSaving) 
			{
				return true;
			}
			else
			{
				return false;
			}
		}
	
			
			
			
			
			
			
			
		public boolean checkClickPreviousAndCheckDataInSavedSalesInvoiceVATVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
			previousBtn.click();
			
			String expMessage="Voucher Loaded Successfully";
			
			String actMessage=checkValidationMessage(expMessage);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		
			String actDocNo=documentNumberTxt.getAttribute("value");
			String actCustomerAccount=customerAccountTxt.getAttribute("value");
			String actDepartment=departmentTxt.getAttribute("value");
			String actWarehouse=select1stRow_1stColumn.getText();
			String actItem=select1stRow_2ndColumn.getText();
			String actQuantity=select1stRow_11thColumn.getText();
			String actRate=select1stRow_14thColumn.getText();
			String actGross=select1stRow_15thColumn.getText();
			String actRMA=select1stRow_21stColumn.getText();
			
			
			String expDocNo="1";
			String expCustomerAccount="Customer A";
			String expDepartment="DUBAI";
			String expWarehouse="HYDERABAD";
			String expItem="WA COGS ITEM";
			String expQuantity="2.00";
			String expRate="19.00";
			String expGross="38.00";
			String expRMA="Q,Q1";
			
			
			System.out.println("********* Voucher No        : "+actDocNo+"  value expected  "+expDocNo);
			System.out.println("********* Customer Account  : "+actCustomerAccount+"  value expected  "+expCustomerAccount);
			System.out.println("********* Department        : "+actDepartment+"  value expected  "+expDepartment);
			System.out.println("********* Warehouse         : "+actWarehouse+"  value expected  "+expWarehouse);
			System.out.println("********* Item              : "+actItem+"  value expected  "+expItem);
			System.out.println("********* Quantity          : "+actQuantity+"  value expected  "+expQuantity);
			System.out.println("********* Rate              : "+actRate+"  value expected  "+expRate);
			System.out.println("********* Gross             : "+actGross+"  value expected  "+expGross);
			System.out.println("********* RMA Column        : "+actRMA+"  value expected  "+expRMA);
			
			if(actDocNo.equalsIgnoreCase(expDocNo) && actCustomerAccount.equalsIgnoreCase(expCustomerAccount) && actDepartment.equalsIgnoreCase(expDepartment) 
				&& actWarehouse.equalsIgnoreCase(expWarehouse) && actItem.equalsIgnoreCase(expItem) && actQuantity.equalsIgnoreCase(expQuantity) 
				&& actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross) && actRMA.equalsIgnoreCase(expRMA)) 
			{
				return true;
			}
			else
			{
				return false;
			}
		}
			
			
			
			
			
			
		public boolean checkEditingSalesInvoiceVATVoucher1QuantityTo4() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys("2");
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys("2");
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("19");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
			select1stRow_21stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.SPACE);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
			searchRMASelectAllChkBox.click();
			
			String actQty=searchRMAQuantityTxtField.getAttribute("value");
			String actNoOfItems=searchRMANoOfItemsTxt.getText();
			
			String expQty="4";
			String expNoOfItems="4";
			
			System.out.println("*********** Quantity: "+actQty+"  value expected  "+expQty);
			System.out.println("*********** Number of Items: "+actNoOfItems+"  value expected  "+expNoOfItems);
			
			if(actQty.equalsIgnoreCase(expQty) && actNoOfItems.equalsIgnoreCase(expNoOfItems)) 
			{
				System.out.println("***Test Pass: Search RMA Screen is Displaying with all options***");
				
				return true;
			}
			else
			{
				System.out.println("***Test Fail: Search RMA Screen is NOT Displaying with all options***");
			
				return false;
			}
		}
			
			
			
			
			
		public boolean checkSearchRMAPopupCheckBoxAfterChangingQuantityTo4() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			boolean actRMAChecked=true;
			for(int i=1; i < 5; i++)
			{
				actRMAChecked=getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]/input")).isSelected();
				System.out.println("**********Row "+i+">>>" +actRMAChecked);
				
				if(actRMAChecked==false)
				{
					System.out.println("*************** Row "+i+" Checkbox NOT selected");
					
					break;
				}
			}
			
			boolean expRMAChecked=true;

			
			
			boolean actRMAUnChecked=false;
			for(int i=5; i < 21; i++)
			{
				actRMAUnChecked=getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]/input")).isSelected();
				System.out.println("**********Row "+i+">>>" +actRMAUnChecked);
				
				if(actRMAUnChecked==true)
				{
					System.out.println("*************** Row "+i+" Checkbox selected");
					
					break;
				}
				
			}
			
			boolean expRMAUnChecked=false;
			
			System.out.println("******* RMA Checked   : "+actRMAChecked   +"  value expected  "+expRMAChecked);
			System.out.println("******* RMA UnChecked : "+actRMAUnChecked +"  value expected  "+expRMAUnChecked);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
		
			Thread.sleep(2000);
			
			if(actRMAChecked==expRMAChecked && actRMAUnChecked==expRMAUnChecked) 
			{
				return true;
			}
			else
			{
				return false;
			}
		}
			
			
			
			
			
			
		public boolean checkSavingSalesInvoiceVATVoucher1WithQuantity4() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
			String actNewReference=newReferenceTxt.getAttribute("value");
			
			String expNewReference="79.80";

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  
			
			System.out.println("New Reference  : "+actNewReference +"  value expected  "+expNewReference);
			
			if(actNewReference.equalsIgnoreCase(expNewReference) && 
					actSaving==expSaving) 
			{
				return true;
			}
			else
			{
				return false;
			}
		}
			
		
		
			
			
		
		public boolean checkEditingSalesInvoiceVATVoucher1QuantityAs3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
			previousBtn.click();
			
			String actMessage=checkValidationMessage("Voucher Loaded Successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys("2");
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys("1");
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("19");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
			select1stRow_21stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.SPACE);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			String actQty=searchRMAQuantityTxtField.getAttribute("value");
			String actNoOfItems=searchRMANoOfItemsTxt.getText();
			
			String expQty="3";
			String expNoOfItems="3";
			
			System.out.println("********** Quantity: "+actQty+"  value expected  "+expQty);
			System.out.println("********** Number of Items: "+actNoOfItems+"  value expected  "+expNoOfItems);
			
			if(actQty.equalsIgnoreCase(expQty) && actNoOfItems.equalsIgnoreCase(expNoOfItems)) 
			{
				return true;
			}
			else
			{
				return false;
			}
		}
			
			
			
			
			
			
		public boolean checkSearchRMAPopupCheckBoxAfterChangingQuantityTo3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			boolean actRMAChecked=true;
			for(int i=1; i < 4; i++)
			{
				actRMAChecked=getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]/input")).isSelected();
				System.out.println("**********Row "+i+">>>" +actRMAChecked);
				
				if(actRMAChecked==false)
				{
					System.out.println("*************** Row "+i+" Checkbox NOT selected");
					
					break;
				}
			}
			
			boolean expRMAChecked=true;
			
			
			boolean actRMAUnChecked=false;
			for(int i=4; i < 21; i++)
			{
				actRMAUnChecked=getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]/input")).isSelected();
				System.out.println("**********Row "+i+">>>" +actRMAUnChecked);
				
				if(actRMAUnChecked==true)
				{
					System.out.println("*************** Row "+i+" Checkbox selected");
					
					break;
				}
				
			}
			
			boolean expRMAUnChecked=false;
			
			Thread.sleep(2000);
			if(actRMAChecked==expRMAChecked && actRMAUnChecked==expRMAUnChecked) 
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
			
			
		
		
			
		
		public boolean checkSavingSalesInvoiceVATVoucher1WithQuantity3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();

			Thread.sleep(2000);
		
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
			String actNewReference=newReferenceTxt.getAttribute("value");
			
			String expNewReference="59.85";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			
			
			System.out.println("New Reference: "+actNewReference+"  value expected  "+expNewReference);
			
			if(actNewReference.equalsIgnoreCase(expNewReference) && actSaving==expSaving) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
	
			
				
				
				
				
				
				
		//To Verify Delete option In  Opening Stocks Voucher : 1
		
		public boolean checkOpeningStockNewDeletingVoucher1InEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksNewVoucher));
			openingStocksNewVoucher.click();
		
			Thread.sleep(3000);

			int count = grid_VoucherNoList.size();
			
			for (int i = 0; i < count; i++) 
			{
				String voucherNo = grid_VoucherNoList.get(i).getText();
				
				if (voucherNo.equalsIgnoreCase("1")) 
				{
					grid_CheckBoxList.get(i).click();
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
			editBtn.click();

			String actLoadingMessage=checkValidationMessage("Voucher loaded successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));		
			new_DeleteBtn.click();
			
			getWaitForAlert();
			String actAlertTxt=getAlert().getText();
			
			getAlert().accept();
			
			String expMessage="Deleting this document has has resulted in negative stock.,RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			String expAlertTxt="Are you sure that you want to delete the Selected Voucher?";
			
			
			System.out.println("********* Alert Text: "+actAlertTxt+"  value expected  "+expAlertTxt);
			System.out.println("********* Error Message: "+actMessage+"  value expected  "+expMessage);
			
			if(actAlertTxt.equalsIgnoreCase(expAlertTxt) && actMessage.equalsIgnoreCase(expMessage)) 
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		
		
		
		
		public boolean checkDeletingopeningStocksNewVoucher1InHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
			deleteBtn.click();
			
			getWaitForAlert();
			String actAlertTxt=getAlert().getText();
			
			getAlert().accept();
			
			String expMessage="VoucherNo - 1: RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			String expAlertTxt="Are you sure that you want to delete the Selected Voucher?";
			
			System.out.println("********* Alert Text: "+actAlertTxt+"  value expected  "+expAlertTxt);
			System.out.println("********* Error Message: "+actMessage+"  value expected  "+expMessage);
			
			if(actAlertTxt.equalsIgnoreCase(expAlertTxt) && actMessage.equalsIgnoreCase(expMessage)) 
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		
			
		
		
		
			
		// Verify Edit option in Opening Stock Quantity by Updating Quantity below the Consume Qty(Sales Qty as 3) 						
		
		public boolean checkUpdateQuantityAs2InOpeningStockNewVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
			editBtn.click();
			
			String LoadingMessage=checkValidationMessage("Voucher Loaded Successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
			select1stRow_3rdColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("2");
			enter_Quantity.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupExpansionBtn));
			RMAPopupExpansionBtn.click();
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			
			String actRMAColumn=select1stRow_9thColumn.getText();
			
			String expRMAColumn="Q,Q1";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			String expMessage="RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("********* RMA Column Data: "+actRMAColumn+"  value expected  "+expRMAColumn);
			
			System.out.println("*********** Error Message: "+actMessage  +"  value expected  "+expMessage);
			
			
			if(actRMAColumn.equalsIgnoreCase(expRMAColumn) && actMessage.equalsIgnoreCase(expMessage))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		
		
		
			
			
			
		public boolean checkSavingopeningStocksNewVoucher1ByChangingWarehouseFromHyderabadToSecunderabad() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			
			getWaitForAlert();
			String actAlert=getAlert().getText();
			
			getAlert().accept();
			
			Thread.sleep(2000);
			
			int count = grid_VoucherNoList.size();
			
			for (int i = 0; i < count; i++) 
			{
				String voucherNo = grid_VoucherNoList.get(i).getText();
				
				if (voucherNo.equalsIgnoreCase("1")) 
				{
					if (grid_CheckBoxList.get(i).isSelected()) 
					{
						break;
					}
					else 
					{
						grid_CheckBoxList.get(i).click();
						
						break;
					}
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
			editBtn.click();
			
			String LoadingMessage=checkValidationMessage("Voucher Loaded Successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			String actDocNo=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
			warehouseTxt.click();
			warehouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			warehouseTxt.sendKeys(Keys.SPACE);
			
			Thread.sleep(2000);
            
			int warehouselist=warehouseHeaderComboList.size();
			
			for (int i = 0; i < warehouselist; i++) 
			{
				String warehouse=warehouseHeaderComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("SECUNDERABAD")) 
				{
					warehouseHeaderComboList.get(i).click();
					
					warehouseTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			String actWarehouse=warehouseTxt.getAttribute("value");
			
			
			String expDocNo="1";
			String expWarehouse="SECUNDERABAD";
			
			
			System.out.println("********** Voucher No : "+actDocNo      +"  value expected  "+expDocNo);
			System.out.println("********** Warehouse  : "+actWarehouse  +"  value expected  "+actWarehouse);
			
			
			if(actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(actWarehouse))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
	
	
		
		
		
		public boolean checkSavingopeningStocksNewVoucher1WithWarehouseAsSecunderabad() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
			select1stRow_5thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupExpansionBtn));
			RMAPopupExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.TAB);
			
			String actRMA=select1stRow_9thColumn.getText();
			
			String expRMA="Q,Q1,Q2,Q3,Q4,Q5,Q6,Q7,Q8,Q9,Q10,Q11,Q12,Q13,Q14,Q15,Q16,Q17,Q18,Q19";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			String expMessage="VoucherNo - 1: RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("******* RMA Column: "+actRMA+"  value expected  "+expRMA);
			
			System.out.println("*********** Error Message: "+actMessage+"  value expected  "+expMessage);
		
			if(actRMA.equalsIgnoreCase(expRMA) && actMessage.equalsIgnoreCase(expMessage))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				int count = grid_VoucherNoList.size();
				
				for (int i = 0; i < count; i++) 
				{
					String voucherNo = grid_VoucherNoList.get(i).getText();
					
					if (voucherNo.equalsIgnoreCase("1")) 
					{
						if (grid_CheckBoxList.get(i).isSelected()) 
						{
							break;
						}
						else 
						{
							grid_CheckBoxList.get(i).click();
							
							break;
						}
					}
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
				editBtn.click();

				String LoadingMessage1=checkValidationMessage("Voucher Loaded Successfully");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
				warehouseTxt.click();
				warehouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
				warehouseTxt.sendKeys(Keys.SPACE);
				
				Thread.sleep(2000);
	            
				int warehouselist=warehouseHeaderComboList.size();
				
				for (int i = 0; i < warehouselist; i++) 
				{
					String warehouse=warehouseHeaderComboList.get(i).getText();
					
					System.out.println(i+"  "+warehouse);
					
					if (warehouse.equalsIgnoreCase("HYDERABAD")) 
					{
						System.out.println("Entered If Loop");
						
						warehouseHeaderComboList.get(i).click();
						
						warehouseTxt.sendKeys(Keys.TAB);
						
						break;
					}
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				String expMessage2="Voucher saved successfully";
				
				String actMessage2=checkValidationMessage(expMessage);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
			
		
		
		
		
		
		
		// Verify Edit option in Opening Stock Voucher by updating quantity to same as consumed quantity(Sales Invoice)		
		
		public boolean checkEditingopeningStocksNewVoucher1AndSavingWithQuantityAs3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksNewVoucher));
			openingStocksNewVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			
			int count = grid_VoucherNoList.size();
			
			for (int i = 0; i < count; i++) 
			{
				String voucherNo = grid_VoucherNoList.get(i).getText();
				
				if (voucherNo.equalsIgnoreCase("1")) 
				{
					getAction().doubleClick(grid_VoucherNoList.get(i)).build().perform();
					break;
				}
			}
			
			String LoadingMessage=checkValidationMessage("Voucher saved successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
			select1stRow_3rdColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("3");
			enter_Quantity.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupExpansionBtn));
			RMAPopupExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			String actRMA1=RMAPopupTableRow1Column1.getText();
			String actRMA2=RMAPopupTableRow2Column1.getText();
			String actRMA3=RMAPopupTableRow3Column1.getText();
			String actDocNo=RMAPopupTableRow1Column2.getText();
			String actQuantity=RMAPopupQuantityTxtField.getAttribute("value");
			
			String expRMA1="Q";
			String expRMA2="Q1";
			String expRMA3="Q2";
			String expDocNo="NDT44:1";
			String expQuantity="3";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			String actRMA=select1stRow_9thColumn.getText();
			
			String expRMA="Q,Q1,Q2";
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			
			System.out.println("********* Row 1: "+actRMA1+"  value expected  "+expRMA1);
			System.out.println("********* Row 2: "+actRMA2+"  value expected  "+expRMA2);
			System.out.println("********* Row 3: "+actRMA3+"  value expected  "+expRMA3);
			System.out.println("********* Document No: "+actDocNo+"  value expected  "+expDocNo);
			System.out.println("********* Quantity: "+actQuantity+"  value expected  "+expQuantity);
			System.out.println("********* RMA Column: "+actRMA+"  value expected  "+expRMA);
			
			
			if(actQuantity.equalsIgnoreCase(expQuantity) && actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) 
				&& actRMA3.equalsIgnoreCase(expRMA3) && actDocNo.equalsIgnoreCase(expDocNo) && actRMA.equalsIgnoreCase(expRMA)
				&& actSaving==expSaving)
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
	
		
		
		
		
		
		
		
		public boolean checkopeningStocksNewVoucher1SavedDataWithQuantity3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			
			int count = grid_VoucherNoList.size();
			
			for (int i = 0; i < count; i++) 
			{
				String voucherNo = grid_VoucherNoList.get(i).getText();
				
				if (voucherNo.equalsIgnoreCase("1")) 
				{
					getAction().doubleClick(grid_VoucherNoList.get(i)).build().perform();
					break;
				}
			}
			
			String LoadingMessage=checkValidationMessage("Voucher saved successfully");
			
			String actDocNo=documentNumberTxt.getAttribute("value");
			String actWarehouse=warehouseTxt.getAttribute("value");
			String actItem=select1stRow_1stColumn.getText();
			String actQuantity=select1stRow_3rdColumn.getText();
			String actRate=select1stRow_4thColumn.getText();
			String actGross=select1stRow_5thColumn.getText();
			String actRMA=select1stRow_9thColumn.getText();
			
			
			String expDocNo="1";
			String expWarehouse="Hyderabad";
			String expItem="WA COGS ITEM";
			String expQuantity="3.00";
			String expRate="8.50";
			String expGross="25.50";
			String expRMA="Q,Q1,Q2";
			
			
			System.out.println("********** Voucher No: "+actDocNo+"  value expected  "+expDocNo);
			System.out.println("********** Warehouse: "+actWarehouse+"  value expected  "+actWarehouse);
			System.out.println("********** Item: "+actItem+"  value expected  "+expItem);
			System.out.println("********** Quantity: "+actQuantity+"  value expected  "+expQuantity);
			System.out.println("********** Quantity: "+actRate+"  value expected  "+expRate);
			System.out.println("********** Gross: "+actGross+"  value expected  "+expGross);
			System.out.println("********** RMA: "+actRMA+"  value expected  "+expRMA);
			
			
			if(actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(actWarehouse) && actItem.equalsIgnoreCase(expItem) &&
				actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross) &&
				actRMA.equalsIgnoreCase(expRMA))
			
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		
		
		
		
		public boolean checkopeningStocksNewVoucher1ChangingQuantityFrom3To20() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
			select1stRow_3rdColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("20");
			enter_Quantity.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupExpansionBtn));
			RMAPopupExpansionBtn.click();
			
			Thread.sleep(2000);
					
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupClearBtn));
			RMAPopupClearBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys("Q,20");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(2000);
			
			//actual
			String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
			String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
			String actRMA3=RMAPopupTableRow3Column1.getAttribute("data-value");
			String actRMA4=RMAPopupTableRow4Column1.getAttribute("data-value");
			String actRMA5=RMAPopupTableRow5Column1.getAttribute("data-value");
			String actRMA6=RMAPopupTableRow6Column1.getAttribute("data-value");
			String actRMA7=RMAPopupTableRow7Column1.getAttribute("data-value");
			String actRMA8=RMAPopupTableRow8Column1.getAttribute("data-value");
			String actRMA9=RMAPopupTableRow9Column1.getAttribute("data-value");
			String actRMA10=RMAPopupTableRow10Column1.getAttribute("data-value");
			String actRMA11=RMAPopupTableRow11Column1.getAttribute("data-value");
			String actRMA12=RMAPopupTableRow12Column1.getAttribute("data-value");
			String actRMA13=RMAPopupTableRow13Column1.getAttribute("data-value");
			String actRMA14=RMAPopupTableRow14Column1.getAttribute("data-value");
			String actRMA15=RMAPopupTableRow15Column1.getAttribute("data-value");
			String actRMA16=RMAPopupTableRow16Column1.getAttribute("data-value");
			String actRMA17=RMAPopupTableRow17Column1.getAttribute("data-value");
			String actRMA18=RMAPopupTableRow18Column1.getAttribute("data-value");
			String actRMA19=RMAPopupTableRow19Column1.getAttribute("data-value");
			String actRMA20=RMAPopupTableRow20Column1.getAttribute("data-value");
			String actDocNo=RMAPopupTableRow1Column2.getText();
			
			
			//expected
			String expRMA1="Q";
			String expRMA2="Q1";
			String expRMA3="Q2";
			String expRMA4="Q3";
			String expRMA5="Q4";
			String expRMA6="Q5";
			String expRMA7="Q6";
			String expRMA8="Q7";
			String expRMA9="Q8";
			String expRMA10="Q9";
			String expRMA11="Q10";
			String expRMA12="Q11";
			String expRMA13="Q12";
			String expRMA14="Q13";
			String expRMA15="Q14";
			String expRMA16="Q15";
			String expRMA17="Q16";
			String expRMA18="Q17";
			String expRMA19="Q18";
			String expRMA20="Q19";
			String expDocNo="NDT44:1";
			
			
			System.out.println("********* Row1: "+actRMA1+"  value expected  "+expRMA1);
			System.out.println("********* Row2: "+actRMA2+"  value expected  "+expRMA2);
			System.out.println("********* Row3: "+actRMA3+"  value expected  "+expRMA3);
			System.out.println("********* Row4: "+actRMA4+"  value expected  "+expRMA4);
			System.out.println("********* Row5: "+actRMA5+"  value expected  "+expRMA5);
			System.out.println("********* Row6: "+actRMA6+"  value expected  "+expRMA6);
			System.out.println("********* Row7: "+actRMA7+"  value expected  "+expRMA7);
			System.out.println("********* Row8: "+actRMA8+"  value expected  "+expRMA8);
			System.out.println("********* Row9: "+actRMA9+"  value expected  "+expRMA9);
			System.out.println("********* Row10: "+actRMA10+"  value expected  "+expRMA10);
			System.out.println("********* Row11: "+actRMA11+"  value expected  "+expRMA11);
			System.out.println("********* Row12: "+actRMA12+"  value expected  "+expRMA12);
			System.out.println("********* Row13: "+actRMA13+"  value expected  "+expRMA13);
			System.out.println("********* Row14: "+actRMA14+"  value expected  "+expRMA14);
			System.out.println("********* Row15: "+actRMA15+"  value expected  "+expRMA15);
			System.out.println("********* Row16: "+actRMA16+"  value expected  "+expRMA16);
			System.out.println("********* Row17: "+actRMA17+"  value expected  "+expRMA17);
			System.out.println("********* Row18: "+actRMA18+"  value expected  "+expRMA18);
			System.out.println("********* Row19: "+actRMA19+"  value expected  "+expRMA19);
			System.out.println("********* Row20: "+actRMA20+"  value expected  "+expRMA20);
			System.out.println("********* Document No: "+actDocNo+"  value expected  "+expDocNo);
			
			if(actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) && 
			   actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)	&&
			   actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10) && actRMA11.equalsIgnoreCase(expRMA11) && actRMA12.equalsIgnoreCase(expRMA12) && 
			   actRMA13.equalsIgnoreCase(expRMA13) && actRMA14.equalsIgnoreCase(expRMA14) && actRMA15.equalsIgnoreCase(expRMA15) && actRMA16.equalsIgnoreCase(expRMA16)	&&
			   actRMA17.equalsIgnoreCase(expRMA17) && actRMA18.equalsIgnoreCase(expRMA18) && actRMA19.equalsIgnoreCase(expRMA19) && actRMA20.equalsIgnoreCase(expRMA20))
		     {
				return true;
			 }
			 else
			 {
		    	return false;
			 }
		}
		
		
		
	
		
		public boolean checkSavingopeningStocksNewVoucher1ByChangingQuantityFrom20To22() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
			select1stRow_3rdColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("22");
			enter_Quantity.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupExpansionBtn));
			RMAPopupExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys("Q20");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(2000);
			
			//actual
			String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
			String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
			String actRMA3=RMAPopupTableRow3Column1.getAttribute("data-value");
			String actRMA4=RMAPopupTableRow4Column1.getAttribute("data-value");
			String actRMA5=RMAPopupTableRow5Column1.getAttribute("data-value");
			String actRMA6=RMAPopupTableRow6Column1.getAttribute("data-value");
			String actRMA7=RMAPopupTableRow7Column1.getAttribute("data-value");
			String actRMA8=RMAPopupTableRow8Column1.getAttribute("data-value");
			String actRMA9=RMAPopupTableRow9Column1.getAttribute("data-value");
			String actRMA10=RMAPopupTableRow10Column1.getAttribute("data-value");
			String actRMA11=RMAPopupTableRow11Column1.getAttribute("data-value");
			String actRMA12=RMAPopupTableRow12Column1.getAttribute("data-value");
			String actRMA13=RMAPopupTableRow13Column1.getAttribute("data-value");
			String actRMA14=RMAPopupTableRow14Column1.getAttribute("data-value");
			String actRMA15=RMAPopupTableRow15Column1.getAttribute("data-value");
			String actRMA16=RMAPopupTableRow16Column1.getAttribute("data-value");
			String actRMA17=RMAPopupTableRow17Column1.getAttribute("data-value");
			String actRMA18=RMAPopupTableRow18Column1.getAttribute("data-value");
			String actRMA19=RMAPopupTableRow19Column1.getAttribute("data-value");
			String actRMA20=RMAPopupTableRow20Column1.getAttribute("data-value");
			String actRMA21=RMAPopupTableRow21Column1.getAttribute("data-value");
			String actRMA22=RMAPopupTableRow22Column1.getAttribute("data-value");
			String actDocNo=RMAPopupTableRow1Column2.getText();
			
			
			//expected
			String expRMA1="Q";
			String expRMA2="Q1";
			String expRMA3="Q2";
			String expRMA4="Q3";
			String expRMA5="Q4";
			String expRMA6="Q5";
			String expRMA7="Q6";
			String expRMA8="Q7";
			String expRMA9="Q8";
			String expRMA10="Q9";
			String expRMA11="Q10";
			String expRMA12="Q11";
			String expRMA13="Q12";
			String expRMA14="Q13";
			String expRMA15="Q14";
			String expRMA16="Q15";
			String expRMA17="Q16";
			String expRMA18="Q17";
			String expRMA19="Q18";
			String expRMA20="Q19";
			String expRMA21="Q20";
			String expRMA22="Q21";
			String expDocNo="NDT44:1";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			String actRMA=select1stRow_9thColumn.getText();
			
			String expRMA="Q,Q1,Q2,Q3,Q4,Q5,Q6,Q7,Q8,Q9,Q10,Q11,Q12,Q13,Q14,Q15,Q16,Q17,Q18,Q19,Q20,Q21";
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			
			
			System.out.println("********* Row1: "+actRMA1+"  value expected  "+expRMA1);
			System.out.println("********* Row2: "+actRMA2+"  value expected  "+expRMA2);
			System.out.println("********* Row3: "+actRMA3+"  value expected  "+expRMA3);
			System.out.println("********* Row4: "+actRMA4+"  value expected  "+expRMA4);
			System.out.println("********* Row5: "+actRMA5+"  value expected  "+expRMA5);
			System.out.println("********* Row6: "+actRMA6+"  value expected  "+expRMA6);
			System.out.println("********* Row7: "+actRMA7+"  value expected  "+expRMA7);
			System.out.println("********* Row8: "+actRMA8+"  value expected  "+expRMA8);
			System.out.println("********* Row9: "+actRMA9+"  value expected  "+expRMA9);
			System.out.println("********* Row10: "+actRMA10+"  value expected  "+expRMA10);
			System.out.println("********* Row11: "+actRMA11+"  value expected  "+expRMA11);
			System.out.println("********* Row12: "+actRMA12+"  value expected  "+expRMA12);
			System.out.println("********* Row13: "+actRMA13+"  value expected  "+expRMA13);
			System.out.println("********* Row14: "+actRMA14+"  value expected  "+expRMA14);
			System.out.println("********* Row15: "+actRMA15+"  value expected  "+expRMA15);
			System.out.println("********* Row16: "+actRMA16+"  value expected  "+expRMA16);
			System.out.println("********* Row17: "+actRMA17+"  value expected  "+expRMA17);
			System.out.println("********* Row18: "+actRMA18+"  value expected  "+expRMA18);
			System.out.println("********* Row19: "+actRMA19+"  value expected  "+expRMA19);
			System.out.println("********* Row20: "+actRMA20+"  value expected  "+expRMA20);
			System.out.println("********* Row21: "+actRMA21+"  value expected  "+expRMA21);
			System.out.println("********* Row20: "+actRMA22+"  value expected  "+expRMA22);
			System.out.println("********* Document No: "+actDocNo+"  value expected  "+expDocNo);
			System.out.println("********** RMA Column: "+actRMA+"  value expected  "+expRMA);
			
			if(actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) && 
			   actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)	&&
			   actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10) && actRMA11.equalsIgnoreCase(expRMA11) && actRMA12.equalsIgnoreCase(expRMA12) && 
			   actRMA13.equalsIgnoreCase(expRMA13) && actRMA14.equalsIgnoreCase(expRMA14) && actRMA15.equalsIgnoreCase(expRMA15) && actRMA16.equalsIgnoreCase(expRMA16)	&&
			   actRMA17.equalsIgnoreCase(expRMA17) && actRMA18.equalsIgnoreCase(expRMA18) && actRMA19.equalsIgnoreCase(expRMA19) && actRMA20.equalsIgnoreCase(expRMA20) &&
			   actRMA21.equalsIgnoreCase(expRMA21) && actRMA22.equalsIgnoreCase(expRMA22)
			   
			   && actRMA.equalsIgnoreCase(expRMA)
			   
			   && actSaving==expSaving)
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
	
		
		
		

		
			
		// Verify Saving Opening Stocks Voucher 
			
		public boolean checkSavingopeningStocksNewVoucher3WithWarehouseAsHyderabadAndQuantityAs2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksNewVoucher));
			openingStocksNewVoucher.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String ScreenOpenedMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
			warehouseTxt.click();
			warehouseTxt.sendKeys(Keys.SPACE);
            
			int warehouselist=warehouseHeaderComboList.size();
			
              for (int i = 0; i < warehouselist; i++) 
				{
					String warehouse=warehouseHeaderComboList.get(i).getText();
					
					System.out.println(i+"  "+warehouse);
					
					if (warehouse.equalsIgnoreCase("HYDERABAD")) 
					{
						warehouseHeaderComboList.get(i).click();
						
						warehouseTxt.sendKeys(Keys.TAB);
						
						break;
					}
				}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			enter_UnitTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("2");
			enter_Quantity.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("5.5");
			enter_Rate.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys("a");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(2000);
			
			String actQty=RMAPopupQuantityTxtField.getAttribute("value");
			String actNoOfItems=RMAPopupQuantityTxtField.getAttribute("value");
			String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
			String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
			String actDocNo=RMAPopupTableRow1Column2.getAttribute("data-value");

			
			String expQty="2";
			String expNoOfItems="2";
			String expRMA1="a";
			String expRMA2="a1";
			String expDocNo="NDT44:3";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			String actRMA=select1stRow_9thColumn.getText();
			
			String expRMA="a,a1";
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
	
			
			System.out.println("********** Quantity: "+actQty+"  value expected  "+expQty);
			System.out.println("********** Number of Items: "+actNoOfItems+"  value expected  "+expNoOfItems);
			System.out.println("********** Row 1: "+actRMA1+"  value expected  "+expRMA1);
			System.out.println("********** Row 2: "+actRMA2+"  value expected  "+expRMA2);
			System.out.println("********** Document No: "+actDocNo+"  value expected  "+expDocNo);
			System.out.println("*********** RMA Column: "+actRMA+"  value expected  "+expRMA);
			
			if(actQty.equalsIgnoreCase(expQty) && actNoOfItems.equalsIgnoreCase(expNoOfItems) && actRMA1.equalsIgnoreCase(expRMA1) && 
				actRMA2.equalsIgnoreCase(expRMA2) && actDocNo.equalsIgnoreCase(expDocNo) &&
				actRMA.equalsIgnoreCase(expRMA) &&
				actSaving==expSaving && actDocNo.equalsIgnoreCase(expDocNo))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
			
			
			
			
			
			
		// To Verify Saving Sales Invoice Voucher by consuming partial 9Qty from RMA1 and Verify Inventory reports
			
		public boolean checkSavingSalesInvoiceVATVoucher2WithQuantityAs9() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String actScreenOpened=checkValidationMessage("Screen opened");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			System.out.println("Customer Account List Size: "+customerAccountCount);
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase("Customer A"))
				{
					customerAccountComboList.get(i).click();
					
					customerAccountTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase("DUBAI"))
				{
					departmentComboList.get(i).click();
					
					departmentTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys("Abu Dhabi");
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("HYDERABAD"))
				{
					warehouseBodyComboList.get(i).click();
					
					pvWarehouseTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys("6");
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys("3");
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("6");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
			select1stRow_21stColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
			searchRMASelectAllChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA=enter_RMA.getAttribute("value");
			
			String expRMA="a,a1,Q3,Q4,Q5,Q6,Q7,Q8,Q9";
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
			String actNewReference=newReferenceTxt.getAttribute("value");
			
			String expNewReference="56.70";

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
	
			
			System.out.println("****** RMA Column                 : "+actRMA			+"  value expected  "+expRMA);
			System.out.println("****** New Reference              : "+actNewReference	+"  value expected  "+expNewReference);
			

			if(actRMA.equalsIgnoreCase(expRMA) && actNewReference.equalsIgnoreCase(expNewReference) && actSaving==expSaving)
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
			
			
			


		

		
		// Verify Purchase Voucher N
			
		public boolean checkSavingPurchaseVoucherNVoucher1WithWarehouseHyderabadAndQuantityAs100() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String ScreenOpenedMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
				
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+vendorAccount);
				
				if (vendorAccount.equalsIgnoreCase("Vendor A"))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase("DUBAI"))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();
			
			enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
			int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
			
			for (int i = 0; i < pvvGridWarehouseListCount; i++) 
			{
				String warehouse=pvvGridWarehouseList.get(i).getText();
				if (warehouse.equalsIgnoreCase("Hyderabad")) 
				{
					pvvGridWarehouseList.get(i).click();
					break;
				}
			}
	        enter_WarehouseTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
	        int pvvGridItemListCount=pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) 
			{
				String Item=pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxCode));
	        enter_PvTaxCode.sendKeys(Keys.TAB);
				
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	        enter_AccountTxt.click();
			
	        Thread.sleep(2000);
	        
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			enter_Quantity.sendKeys("100");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
			select1stRow_19thColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys("1");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String rma=enter_RMA.getAttribute("value");
			System.out.println("********** RMA Column: "+rma);
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  
			
			if(rma.equalsIgnoreCase("1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,"
				+ "51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100") && 
				actSaving==expSaving) 
			{
				return true;
			}
			else
			{
				return false;
			}
		}
			
			
			
			
			
			
			
		public boolean checkClickingOKInRMAPopUpWithRMAAsEmptyInPurchaseVoucherNVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
				
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+vendorAccount);
				
				if (vendorAccount.equalsIgnoreCase("Vendor A"))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase("DUBAI"))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();
			
			enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
			int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
			
			for (int i = 0; i < pvvGridWarehouseListCount; i++) 
			{
				String warehouse=pvvGridWarehouseList.get(i).getText();
				if (warehouse.equalsIgnoreCase("Hyderabad")) 
				{
					pvvGridWarehouseList.get(i).click();
					break;
				}
			}
	        enter_WarehouseTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
	        int pvvGridItemListCount=pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) 
			{
				String Item=pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxCode));
	        enter_PvTaxCode.sendKeys(Keys.TAB);
				
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	        enter_AccountTxt.click();
			
	        Thread.sleep(2000);
	        
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			enter_Quantity.sendKeys("1");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
			select1stRow_19thColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
	
			Thread.sleep(2000);
			
			getWaitForAlert();
			String actAlert = getAlert().getText();
			getAlert().accept();
			
			String expAlert = "RMA Number Cannot be Left Blank, ";
			String expAlert1 = "Do you want to continue with the selected quantity ?";
			
			Thread.sleep(2000);
			
			getWaitForAlert();
			String actAlert2=getAlert().getText();
			getAlert().accept();
			
			String expAlert2 = "RMA Number Cannot be Left Blank";
					
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupCancelBtn));
			RMAPopupCancelBtn.click();
			
			Thread.sleep(2000);
			
			System.out.println("************** Alert: "+actAlert  +" Value Expected "+expAlert+" "+expAlert1);
			System.out.println("************** Alert: "+actAlert2 +" Value Expected "+expAlert2);
			
			if(actAlert.startsWith(expAlert) && actAlert.endsWith(expAlert1) &&
			   actAlert2.equalsIgnoreCase(expAlert2))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
				
			
			
			
			
			
			
		public boolean checkSavingPurchaseVoucherNVoucher2WithoutRMA() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			String expMessage="Provide data for mandatory fieldsRMA";
			
			String actMessage=checkValidationMessage(expMessage);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
			new_newBtn.click();
			
			getWaitForAlert();
			String actAlert = getAlert().getText();
			
			getAlert().accept();
			
			String expAlert1 = "The contents of the voucher are changed.";
			String expAlert2 = "Are you sure that you want to ignore the changes?";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			String act2=documentNumberTxt.getAttribute("value");
			
			System.out.println("********* Error Message: "+actMessage+"  value expected  "+expMessage);
			System.out.println("********* Alert Text: "+actAlert+"  value expected  "+expAlert1+" "+expAlert2);
			
			if(actMessage.equalsIgnoreCase(expMessage) && actAlert.startsWith(expAlert1) && actAlert.endsWith(expAlert2))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
			
			
			
		
		
			
			
			
		//Excess In Stocks 
			
		public boolean checkSavingExcessesInStocksVoucher1WithWarehouseAsHyderabadAndQuantityAs3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
			invTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(excessesInStocksVoucher));
			excessesInStocksVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();

			String ScreenOpenedMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("HYDERABAD"))
				{
					System.out.println("Entered If Loop");
					
					warehouseBodyComboList.get(i).click();
					
					pvWarehouseTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					enter_ItemTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			enter_UnitTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("3");
			enter_Quantity.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys("101");
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(2000);
			
			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		    Date date=new Date();
		    System.out.println(df.format(date));
		    String expdate=df.format(date);
			
			String actQty=RMAPopupQuantityTxtField.getAttribute("value");
			System.out.println("*******************************"+actQty);
			String actNoOfItems=RMAPopupNumberOfItems.getText();
			System.out.println("*******************************"+actNoOfItems);
			
			String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
			System.out.println("*******************************"+actRMA1);
			String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
			System.out.println("*******************************"+actRMA2);
			String actRMA3=RMAPopupTableRow3Column1.getAttribute("data-value");
			System.out.println("*******************************"+actRMA3);
			String actDocNo=RMAPopupTableRow1Column2.getAttribute("data-value");
			System.out.println("*******************************"+actDocNo);
			String actDocDate=RMAPopupTableRow1Column3.getAttribute("data-value");
			System.out.println("*******************************"+actDocDate);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.TAB);
			
			String actRMA=select1stRow_11thColumn.getText();
			
			System.out.println("*********** RMA Column: "+actRMA);
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			
				
			if(actQty.equalsIgnoreCase("3") && actNoOfItems.equalsIgnoreCase("3") && actRMA1.equalsIgnoreCase("101") && actRMA2.equalsIgnoreCase("102") && 
				actRMA3.equalsIgnoreCase("103") && actDocNo.equalsIgnoreCase("ExeStk:1") && actDocDate.equalsIgnoreCase(expdate) && actRMA.equalsIgnoreCase("101,102,103") &&
				actSaving==expSaving)  
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
			
			
			
			
			
			
		
			
		
		// Shortages in Stocks 
		
		public boolean checkSavingShortageInStocksVoucher1ConsumeQuantityFromExcessesInStocksVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(shortagesInStockVoucher));
			shortagesInStockVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();

			String ScreenOpenedMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("HYDERABAD"))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					enter_ItemTxt.sendKeys(Keys.TAB);
					break;
				}
			}	
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
			select1stRow_5thColumn.click();
			enter_Quantity.sendKeys("3");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
			select1stRow_7thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASearchTxtField));
			
			for(int i=101;i<=103;i++)
			{
				searchRMASearchTxtField.clear();
				searchRMASearchTxtField.sendKeys(String.valueOf(i));
				Thread.sleep(2000);
				getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]")).click();
				Thread.sleep(2000);
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA = enter_RMA.getAttribute("value");
			String expRMA = "101,102,103";
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();

			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			
			
			System.out.println("********* RMA Column    : "+actRMA     +"  Value expected  " +expRMA);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actSaving==expSaving) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
			
			
			
			
			
		// Editing Excess Stocks Voucher by Updating Quantity below the consume quantity in Shortage in Stocks	
			
		public boolean checkEditingExcessInStocksVoucher1SavingWithQuantityBelowConsumeQuantityInShortageInStocks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
			invTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(excessesInStocksVoucher));
			excessesInStocksVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			grid_ChkBox1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
			editBtn.click();

			String LoadingMessage=checkValidationMessage("Voucher loaded Successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
			select1stRow_4thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("2");
			enter_Quantity.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupExpansionBtn));
			RMAPopupExpansionBtn.click();
		
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			
			String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
			String expRMA1="101";
			
			String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
			String expRMA2="102";
			
			String actDocNo=RMAPopupTableRow1Column2.getText();
			String expDocNo="ExeStk:1";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.TAB);
			
			String actRMA=select1stRow_11thColumn.getText();
			String expRMA="101,102";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			String expMessage="RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("*********   "+actRMA1   +" Value Expected "+expRMA1);
			System.out.println("*********   "+actRMA2   +" Value Expected "+expRMA2);
			System.out.println("*********   "+actDocNo  +" Value Expected "+expDocNo);
			System.out.println("RMA Column: "+actRMA    +" Value Expected "+expRMA);
			System.out.println("ErrorMessage"+actMessage+" Value Expected "+expMessage);
			
			if(actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actDocNo.equalsIgnoreCase(expDocNo) && actRMA.equalsIgnoreCase(expRMA) &&
				actMessage.equalsIgnoreCase(expMessage))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				getWaitForAlert();
				getAlert().accept();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				getWaitForAlert();
				getAlert().accept();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
		
		
		
		
		// Deleting  Excess in Stocks Voucher (from Voucher Home page & Voucher Entry Page) where RMA quantity is Consumed in Shortage in Stocks Voucher	
			
		public boolean checkDeletingExcessInStockVoucher1InEntryPageConsumedInShortageInStocks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
			invTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(excessesInStocksVoucher));
			excessesInStocksVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			grid_ChkBox1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
			editBtn.click();

			String LoadingMessage=checkValidationMessage("Voucher loaded Successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
			new_DeleteBtn.click();
			
			Thread.sleep(2000);
			
			getWaitForAlert();
			getAlert().accept();

			String expMessage="RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("************ Error Message: "+actMessage+"  value expected  "+expMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
			
			
			
			
			
				
			
			
		public boolean checkDeletingExcessInStockVoucher1InHomepageConsumedInShortageInStocks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
			deleteBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
			
			String expMessage="VoucherNo - 1: RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("********* Error Message: "+actMessage+"  value expected  "+expMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
				
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
			
				
			
		
		
		
		
		
		
		
		
		
		
		
		
			
		// Verify Suspending Excess in Stocks Voucher (From Voucher Home  Page and Entry Page) where RMA quantity is Consumed in Shortage in Stocks Voucher	
			
		// Should not Allow to Suspend, as they already consumed in Shortage in stocks	
		public boolean checkSuspendingExcessInStockVoucher1InEntryPageConsumedInShortageInStocks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
			invTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(excessesInStocksVoucher));
			excessesInStocksVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			grid_ChkBox1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
			editBtn.click();

			String LoadingMessage=checkValidationMessage("Voucher loaded Successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
			new_SuspendBtn.click();

			String expMessage="RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("******** Error Message: "+actMessage+"  value expected  "+expMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage)) 
			{
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
				getAction().doubleClick(grid_ChkBox1).build().perform();
				
				String LoadingMessage1=checkValidationMessage("Voucher loaded Successfully");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				String expMessage1="Voucher saved Successfully";
				
				String actMessage1=checkValidationMessage(expMessage1);
				
				return false;
			}
		}
			
			
			
			
			
			
			
			
		// Should not Allow to Suspend, as they already consumed in shortage in stocks
		public boolean checkSuspendingExcessInStockVoucher1InHomepageConsumedInShortageInStocks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			grid_ChkBox1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(suspendBtn));
			suspendBtn.click();
			
			String expMessage="VoucherNo - 1: RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("********** Error Message: "+actMessage+"  value expected  "+expMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
				getAction().doubleClick(grid_ChkBox1).build().perform();
			
				String LoadingMessage=checkValidationMessage("Voucher loaded Successfully");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				String expMessage1="Voucher saved Successfully";
				
				String actMessage1=checkValidationMessage(expMessage1);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		

			
		
		
		
		
		// Verify Suspending Shortages in Stocks Voucher and consider the Balance RMA Quantities in Inward/Outward RMA Pop-ups and in reports like Stock Ledger and Stock Balances by Products by RMA  Report Screens	
		
		public boolean checkSuspendingShortageInStocksVoucher1InEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(shortagesInStockVoucher));
			shortagesInStockVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			getAction().doubleClick(grid_ChkBox1).build().perform();

			String LoadingMessage=checkValidationMessage("Loading Message");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
			new_SuspendBtn.click();

			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
		
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_1stRowSuspendStatus));
			String actSuspendStatus=grid_1stRowSuspendStatus.getText();
			String expSuspendStatus="true";
					
			System.out.println("******** Suspend Status : "+actSuspendStatus +"  value expected  "+expSuspendStatus);
			
			
			if(actSaving==expSaving && actSuspendStatus.equalsIgnoreCase(expSuspendStatus)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
			
			
			
		
		
		
			
		
		//Stock Transfer Voucher 
		public boolean checkSavingStockTransferVoucher1WithFromWarehouseAsHyderabadAndToWarehouseAsSecunderabad() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransfersVoucher));
			stockTransfersVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
				
			String ScreenOpenedMessage=checkValidationMessage("Screen opened");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferWarehouse1));
			stockTransferWarehouse1.click();
			stockTransferWarehouse1.sendKeys(Keys.SPACE);
            
			int warehouse1Count=st_warehouse1HeaderComboList.size();
			
              for (int i = 0; i < warehouse1Count; i++) 
				{
					String warehouse=st_warehouse1HeaderComboList.get(i).getText();
					System.out.println(i+"  "+warehouse);
					
					if (warehouse.equalsIgnoreCase("HYDERABAD")) 
					{
						st_warehouse1HeaderComboList.get(i).click();
						stockTransferWarehouse1.sendKeys(Keys.TAB);
						break;
					}
				}
					
	            getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	            select1stRow_1stColumn.click();
	            getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Warehouse2Txt));
	            enter_Warehouse2Txt.sendKeys(Keys.SPACE);
	            
				int warehouse2Count=st_warehouse2BodyComboList.size();
				
				 for (int i = 0; i < warehouse2Count; i++) 
				{
					String warehouse=st_warehouse2BodyComboList.get(i).getText();
					
					System.out.println(i+"  "+warehouse);
					
					if (warehouse.equalsIgnoreCase("SECUNDERABAD")) 
					{
						st_warehouse2BodyComboList.get(i).click();
						enter_Warehouse2Txt.sendKeys(Keys.TAB);
						break;
					}
				}
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				enter_ItemTxt.sendKeys(Keys.SPACE);
				
				int itemsCount=itemComboList.size();
				
				for (int i = 0; i < itemsCount; i++) 
				{
					String item=itemComboList.get(i).getText();
					
					System.out.println(i+"  "+item);
					
					if (item.equalsIgnoreCase("WA COGS ITEM")) 
					{
						itemComboList.get(i).click();
						enter_ItemTxt.sendKeys(Keys.TAB);
						break;
					}
				}	

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
				enter_UnitTxt.sendKeys(Keys.TAB);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
				enter_Quantity.sendKeys("2");
			
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
				select1stRow_6thColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
				enter_Rate.sendKeys("3");
				enter_Rate.sendKeys(Keys.TAB);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
				enter_Gross.sendKeys(Keys.TAB);
			
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
				RMAPopupSerialNoExpansionBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASearchTxtField));
				
				for(int i=4; i < 6; i++)
				{
					searchRMASearchTxtField.clear();
					searchRMASearchTxtField.sendKeys(String.valueOf(i));
					Thread.sleep(3000);
					getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]")).click();
					Thread.sleep(2000);
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
				searchRMAOkBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
				salesInvoiceRMAPopupOkBtn.click();
				
				Thread.sleep(2000);

				String actRMAColumn=enter_RMA.getAttribute("value");
				String expRMAColumn="4,5";
				
				String docno=documentNumberTxt.getAttribute("value");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();

				boolean actSaving=checkBackgroundSavingMessage(docno);
				boolean expSaving=true;

				System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			
				
				System.out.println("********* RMA Column     : "+actRMAColumn +" Value Expected "+expRMAColumn);
				
				if(actRMAColumn.equalsIgnoreCase(expRMAColumn) && actSaving==expSaving) 
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
					new_CloseBtn.click();
					
					Thread.sleep(2000);
					
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
					voucherhomeCloseBtn.click();
					
					Thread.sleep(2000);
					
					return true;
				}
				else
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
					new_CloseBtn.click();
					
					Thread.sleep(2000);
					
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
					voucherhomeCloseBtn.click();
					
					Thread.sleep(2000);
					
					return false;
				}
			}
			
			
			
		
			
			
	
		// Suspending Stocks Transfer Voucher 	
		
		public boolean checkSuspendingStockTransferVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransfersVoucher));
			stockTransfersVoucher.click();
		
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			getAction().doubleClick(grid_ChkBox1).build().perform();
		
			String LoadingMessage=checkValidationMessage("Voucher loaded Successfully");
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
			new_SuspendBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
		
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_1stRowSuspendStatus));
			String actSuspendStatus=grid_1stRowSuspendStatus.getText();
			String expSuspendStatus="true";
			
			System.out.println("******** Suspend Statue : "+actSuspendStatus +" Value Expected "+expSuspendStatus);
			
			if(actSaving==expSaving && actSuspendStatus.equalsIgnoreCase(expSuspendStatus))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
		
		
		
		
		
		// Verify Saving the Stock Transfer Voucher without Warehouse-2 RMA
			
		public boolean checkStockTransferVoucher2WithWarehouse1IsSameAsWarehouse2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransfersVoucher));
			stockTransfersVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferWarehouse1));
			stockTransferWarehouse1.click();
			stockTransferWarehouse1.sendKeys(Keys.SPACE);
            
			int warehouse1list=st_warehouse1HeaderComboList.size();
			
              for (int i = 0; i < warehouse1list; i++) 
				{
					String warehouse=st_warehouse1HeaderComboList.get(i).getText();
					
					System.out.println(i+"  "+warehouse);
					
					if (warehouse.equalsIgnoreCase("HYDERABAD")) 
					{
						st_warehouse1HeaderComboList.get(i).click();
						stockTransferWarehouse1.sendKeys(Keys.TAB);
						break;
					}
				}

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Warehouse2Txt));
	            enter_Warehouse2Txt.sendKeys(Keys.SPACE);
	            
				int warehouse2list=st_warehouse2BodyComboList.size();
				
				
				 for (int i = 0; i < warehouse2list; i++) 
				{
					String warehouse=st_warehouse2BodyComboList.get(i).getText();
					
					System.out.println(i+"  "+warehouse);
					
					if (warehouse.equalsIgnoreCase("HYDERABAD")) 
					{
						st_warehouse2BodyComboList.get(i).click();
						enter_Warehouse2Txt.sendKeys(Keys.TAB);
						break;
					}
				}
				
				try{
				    if(errorMessage.isDisplayed())
				    {
				    	System.out.println("***Test Pass: NOT Allowing to select warehouse same, IF LOOP***");
					  
				    	return true;	
					}
				    else
				    {
				    	System.out.println("***Test Pass: Allowing to select warehouse same, ELSE LOOP***");
					   
				    	return false;
				    }
				}
				 catch(Exception e)
				{
				      System.out.println("***Test Pass: Allowing to select warehouse same, CATCH BLOCK***");
				    
				      System.out.println(e.getMessage());
				    
				      return false;
			    }
		    }
			
			
			
			
			
			
			
		public boolean checkSavingStockTransferVoucher1WithWarehouse2Blank() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	    {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Warehouse2Txt));
			enter_Warehouse2Txt.sendKeys(Keys.LEFT_SHIFT,Keys.HOME);
			enter_Warehouse2Txt.sendKeys(Keys.SPACE);
			
			Thread.sleep(2000);
			
			enter_Warehouse2Txt.sendKeys(Keys.ESCAPE);
			enter_Warehouse2Txt.sendKeys(Keys.TAB);
			
			//enter_Warehouse2Txt.sendKeys(Keys.ARROW_DOWN,Keys.ARROW_DOWN,Keys.ARROW_DOWN,Keys.ARROW_DOWN,Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					enter_ItemTxt.sendKeys(Keys.TAB);
					break;
				}
			}	
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			enter_UnitTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("2");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
			select1stRow_6thColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("5");
			
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASearchTxtField));
			
			for(int i=10; i < 12;i++)
			{
				searchRMASearchTxtField.clear();
				searchRMASearchTxtField.sendKeys(String.valueOf(i));
	
				Thread.sleep(3000);
				
				getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]")).click();
				Thread.sleep(2000);
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
		
			Thread.sleep(2000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			String expMessage="Value missing in column - Warehouse - 2";
			
			String actMessage=checkValidationMessage(expMessage);

			Thread.sleep(2000);
			
			System.out.println("********* Error Message: "+actMessage+"  value expected  "+expMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				return true;
			}
			else
			{
				return false;
			}
	    }
			
			
			
				
		
		
		
		public boolean checkSavingStockTransferVoucher2Warehouse2AsSecunderabad() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	    {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Warehouse2Txt));
            enter_Warehouse2Txt.sendKeys(Keys.SPACE);
            
			int warehouselist=st_warehouse2BodyComboList.size();
			
			for (int i = 0; i < warehouselist; i++) 
			{
				String warehouse=st_warehouse2BodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("SECUNDERABAD")) 
				{
					st_warehouse2BodyComboList.get(i).click();
					enter_Warehouse2Txt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.TAB);
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			enter_UnitTxt.sendKeys(Keys.TAB);
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("2");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
			select1stRow_6thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("3");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.SPACE);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASearchTxtField));
			
			for(int i=10; i < 12;i++)
			{
				searchRMASearchTxtField.clear();
				searchRMASearchTxtField.sendKeys(String.valueOf(i));
	
				Thread.sleep(3000);
				
				getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]")).click();
				Thread.sleep(2000);
			}
			
			Thread.sleep(2000);
			
			for(int i=4; i < 6; i++)
			{
				searchRMASearchTxtField.clear();
				searchRMASearchTxtField.sendKeys(String.valueOf(i));
				Thread.sleep(3000);
				getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]")).click();
				Thread.sleep(2000);
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			enter_RMA.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			String actRMAColumn=select1stRow_11thColumn.getText();
			String expRMAColumn="4,5";
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
		
			
			System.out.println("******* RMA Column   : "+actRMAColumn +" value expected "+expRMAColumn);
			
			
			if(actRMAColumn.equalsIgnoreCase(expRMAColumn) && actSaving==expSaving)
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
			
			
			
			
			
			
			
		// Saving Sales Invoice by consuming quantity after Saving Stock Transfer Voucher
				
		//Sales Invoice
		public boolean checkSavingSalesInvoiceVATVoucher3AfterStockTransfer() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();

			String ScreenOpenedMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			System.out.println("Customer Account List Size: "+customerAccountCount);
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase("Customer A"))
				{
					customerAccountComboList.get(i).click();
					
					customerAccountTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase("DUBAI"))
				{
					departmentComboList.get(i).click();
					
					departmentTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys("Abu Dhabi");
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("SECUNDERABAD"))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys("1");
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys("1");
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("3");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
			select1stRow_21stColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
			searchRMASelectAllChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA=enter_RMA.getAttribute("value");
			
			String expRMA="4,5";
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
			String actNewReference=newReferenceTxt.getAttribute("value");
			
			String expNewReference="6.30";

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
	
			
			System.out.println("****** RMA Column                 : "+actRMA			+"  value expected  "+expRMA);
			System.out.println("****** New Reference              : "+actNewReference	+"  value expected  "+expNewReference);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actNewReference.equalsIgnoreCase(expNewReference) 
				&& actSaving==expSaving) 
			{
				return true;
			}
			else
			{
				return false;
			}
		}
			
			
		
		
		
		
		// Suspending Sales Invoice  
			
		public boolean checkSuspendingSalesInvoiceVATVoucher3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
			previousBtn.click();
		
			String LoadingMessage=checkValidationMessage("Voucher loaded Successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
			new_SuspendBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  
			if(actSaving==expSaving) 
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		
			
		
		
		
			
		public boolean checkSavingSalesInvoiceVATVoucher4() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			System.out.println("Customer Account List Size: "+customerAccountCount);
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase("Customer A"))
				{
					customerAccountComboList.get(i).click();
					
					customerAccountTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase("DUBAI"))
				{
					departmentComboList.get(i).click();
					
					departmentTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys("Abu Dhabi");
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("SECUNDERABAD"))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys("1");
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys("1");
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("3");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
			select1stRow_21stColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
			searchRMASelectAllChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA=enter_RMA.getAttribute("value");
			
			String expRMA="4,5";
			String docno=documentNumberTxt.getAttribute("value");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
			String actNewReference=newReferenceTxt.getAttribute("value");
			
			String expNewReference="6.30";

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
	
			
			System.out.println("****** RMA Column                 : "+actRMA			+"  value expected  "+expRMA);
			System.out.println("****** New Reference              : "+actNewReference	+"  value expected  "+expNewReference);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actNewReference.equalsIgnoreCase(expNewReference) 
				&& actSaving==expSaving) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
			
			

		
			
			
			
		// Verify ReSaving the Suspended Sales Invoice Voucher:3
			
		public boolean checkResavingSalesInvoiceVATVoucher3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			int VouchersCount = grid_VoucherNoList.size();
			
			String actSuspendStatus = "";
			String expSuspendStatus = "true";
			
			for (int i = 0; i < VouchersCount; i++) 
			{
				String VoucherNo = grid_VoucherNoList.get(i).getText();
				
				if (VoucherNo.equalsIgnoreCase("3")) 
				{
					actSuspendStatus = grid_SuspendStatusList.get(i).getText();
					
					getAction().doubleClick(grid_CheckBoxList.get(i)).build().perform();
				}
			}
				
			String LoadingMessage=checkValidationMessage("Voucher loaded Successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
		
			Thread.sleep(2000);
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String expMessage="RMA is not available. It may have been consumed elsewhere.";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("********** Suspend Status "+actSuspendStatus +"  value expected  "+expSuspendStatus);
			System.out.println("********** Error Message  "+actMessage       +"  value expected  "+expMessage);
			
			if(actSuspendStatus.equalsIgnoreCase(expSuspendStatus) && actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
		
	
	
	
			
			
			
		// Verify Edit in Stock Transfer Voucher by Updating Quantity below the consume quantity (Sales Invoice VAT)	
			
		public boolean checkEditingStockTransferVoucher2WithQuantityBelowConsumedQuantityInSalesInvoiceVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransfersVoucher));
			stockTransfersVoucher.click();
		
			Thread.sleep(3000);
			
			int VouchersCount = grid_VoucherNoList.size();
			
			for (int i = 0; i < VouchersCount; i++) 
			{
				String VoucherNo = grid_VoucherNoList.get(i).getText();
				
				if (VoucherNo.equalsIgnoreCase("2")) 
				{
					getAction().doubleClick(grid_CheckBoxList.get(i)).build().perform();
				}
			}
		
			String LoadingMessage=checkValidationMessage("Voucher loaded Successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
			select1stRow_4thColumn.click();
			enter_Quantity.clear();
			enter_Quantity.sendKeys("1");
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
			select1stRow_6thColumn.click();
			enter_Rate.sendKeys(Keys.TAB);
			
			enter_Gross.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupExpansionBtn));
			RMAPopupExpansionBtn.click();
		
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			boolean actRMAChecked=true;
			boolean expRMAChecked=true;
			
			for(int i=4; i < 5; i++)
			{
				actRMAChecked=getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]/input")).isSelected();
			}
			
			System.out.println("RMAChecked : "+actRMAChecked+" Value Expected "+expRMAChecked);
			
			if(actRMAChecked==expRMAChecked) 
			{
				return true;
			}
			else
			{
				return false;
			}
		}
			
		
		
		
			
			
		public boolean checkSavingStockTransferVoucher2WithQuantity1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
			searchRMASelectAllChkBox.click();
			
			Thread.sleep(2000);
			
			searchRMASelectAllChkBox.click();
			
			for(int i=4;i<=4;i++)
			{
				searchRMASearchTxtField.sendKeys(String.valueOf(i));
				Thread.sleep(3000);
				getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]")).click();
				Thread.sleep(2000);
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA = enter_RMA.getAttribute("value");
			String expRMA = "4";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			String expMessage="RMA is not available. It may have been consumed elsewhere.";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("******** Error Message : "+actMessage  +" Value Expected "+expMessage);
			System.out.println("******** RMA Column    : "+actRMA      +" Value Expected "+expRMA);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actMessage.equalsIgnoreCase(expMessage)) 
			{
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
				previousBtn.click();
				
				String LoadingMessage1=checkValidationMessage("Voucher loaded Successfully");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
				select1stRow_4thColumn.click();
				enter_Quantity.clear();
				enter_Quantity.sendKeys("2");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
				select1stRow_6thColumn.click();
				enter_Rate.sendKeys(Keys.TAB);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
				enter_Gross.sendKeys(Keys.TAB);
			
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupExpansionBtn));
				RMAPopupExpansionBtn.click();
			
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
				RMAPopupSerialNoExpansionBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
				searchRMASelectAllChkBox.click();
				
				Thread.sleep(2000);
				
				searchRMASelectAllChkBox.click();
				
				for(int i=4;i<=5;i++)
				{
					searchRMASearchTxtField.clear();
					searchRMASearchTxtField.sendKeys(String.valueOf(i));
					Thread.sleep(3000);
					getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]")).click();
					Thread.sleep(2000);
				}
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
				searchRMAOkBtn.click();
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
				salesInvoiceRMAPopupOkBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();

				String expMessage2="Voucher saved successfully";
				
				String actMessage2=checkValidationMessage(expMessage2);
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
			
			
					
		
				
			
		// Deleting Stock Transfer Voucher (from Voucher Home page & Voucher Entry Page) where Transferred Qty is consumed in Sales Invoice Voucher	
			
		public boolean checkDeletingStockTransferVoucher2InEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
			previousBtn.click();
			
			String LoadingMessage=checkValidationMessage("Voucher loaded Successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
			new_DeleteBtn.click();
		
			getWaitForAlert();
			getAlert().accept();
		
			String expMessage="Voucher deleted Successfully";
			
			String actMessage=checkValidationMessage(expMessage);
			
			Thread.sleep(2000);
			
			System.out.println("********** Error Message: "+actMessage+"  value expected  "+expMessage);
			
			if(actMessage.contains(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
				new_newBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferWarehouse1));
				stockTransferWarehouse1.click();
				stockTransferWarehouse1.sendKeys(Keys.SPACE);
	            
				int warehouse1list=st_warehouse1HeaderComboList.size();
				System.out.println("******* Warehouse-1 List: "+warehouse1list);	
				
				for (int i = 0; i < warehouse1list; i++) 
				{
					String warehouse=st_warehouse1HeaderComboList.get(i).getText();
					
					System.out.println(i+"  "+warehouse);
					
					if (warehouse.equalsIgnoreCase("HYDERABAD")) 
					{
						System.out.println("Entered If Loop");
						
						st_warehouse1HeaderComboList.get(i).click();
						
						stockTransferWarehouse1.sendKeys(Keys.TAB);
						
						break;
					}
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Warehouse2Txt));
	            enter_Warehouse2Txt.sendKeys(Keys.SPACE);
	            
				int warehouse2list=st_warehouse2BodyComboList.size();
				
				System.out.println("******* Warehouse-2 List: "+warehouse2list);				
				
				 for (int i = 0; i < warehouse2list; i++) 
				{
					String warehouse=st_warehouse2BodyComboList.get(i).getText();
					
					System.out.println(i+"  "+warehouse);
					
					if (warehouse.equalsIgnoreCase("SECUNDERABAD")) 
					{
						System.out.println("Entered If Loop");
						
						st_warehouse2BodyComboList.get(i).click();
						
						enter_Warehouse2Txt.sendKeys(Keys.TAB);
						
						break;
					}
				}
				
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				enter_ItemTxt.sendKeys(Keys.SPACE);
				
				int itemsCount=itemComboList.size();
				
				for (int i = 0; i < itemsCount; i++) 
				{
					String item=itemComboList.get(i).getText();
					
					System.out.println(i+"  "+item);
					
					if (item.equalsIgnoreCase("WA COGS ITEM")) 
					{
						itemComboList.get(i).click();
						enter_ItemTxt.sendKeys(Keys.TAB);
						break;
					}
				}	
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
				enter_UnitTxt.sendKeys(Keys.TAB);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
				enter_Quantity.sendKeys("2");
		
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
				select1stRow_6thColumn.click();
				enter_Rate.sendKeys("3");
				enter_Rate.sendKeys(Keys.TAB);
				
				enter_Gross.sendKeys(Keys.TAB);

				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
				RMAPopupSerialNoExpansionBtn.click();
				
				Thread.sleep(2000);
				
				for(int i=4;i<=5;i++)
				{
					searchRMASearchTxtField.clear();
					searchRMASearchTxtField.sendKeys(String.valueOf(i));
					Thread.sleep(3000);
					getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]")).click();
					Thread.sleep(2000);
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
				searchRMAOkBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
				salesInvoiceRMAPopupOkBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				String expMessage2="Voucher saved Successfully";
				
				String actMessage2=checkValidationMessage(expMessage2);
				
				System.out.println("************ Error Message: "+actMessage2+"  value expected  "+expMessage2);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
			
				Thread.sleep(3000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_HeaderChkBox));
				
				return false;
			}
			else
			{
				return true;
			}
		}
			
			
			
		
			
			
		
		public boolean checkDeletingStockTransferVoucher2InHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			int count = grid_VoucherNoList.size();
			
			for (int i = 0; i < count; i++) 
			{
				String voucherNo = grid_VoucherNoList.get(i).getText();
				
				if (voucherNo.equalsIgnoreCase("2")) 
				{
					grid_CheckBoxList.get(i).click();
				}
			}
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
			deleteBtn.click();
		
			Thread.sleep(2000);
			
			getWaitForAlert();
			getAlert().accept();
		
			String expMessage="VoucherNo - 2: Voucher deleted Successfully";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("********  Error Message: "+actMessage+"  value expected  "+expMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
				newBtn.click();
				
				String ScreenOpeningMessage = checkValidationMessage("Screen opened");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferWarehouse1));
				stockTransferWarehouse1.click();
				stockTransferWarehouse1.sendKeys(Keys.SPACE);
	            
				int warehouse1list=st_warehouse1HeaderComboList.size();
				System.out.println("******* Warehouse-1 List: "+warehouse1list);	
				
               for (int i = 0; i < warehouse1list; i++) 
				{
					String warehouse=st_warehouse1HeaderComboList.get(i).getText();
					
					System.out.println(i+"  "+warehouse);
					
					if (warehouse.equalsIgnoreCase("HYDERABAD")) 
					{
						System.out.println("Entered If Loop");
						
						st_warehouse1HeaderComboList.get(i).click();
						
						stockTransferWarehouse1.sendKeys(Keys.TAB);
						
						break;
					}
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Warehouse2Txt));
	            enter_Warehouse2Txt.sendKeys(Keys.SPACE);
	            
				int warehouse2list=st_warehouse2BodyComboList.size();
				
				System.out.println("******* Warehouse-2 List: "+warehouse2list);				
				
				 for (int i = 0; i < warehouse2list; i++) 
				{
					String warehouse=st_warehouse2BodyComboList.get(i).getText();
					
					System.out.println(i+"  "+warehouse);
					
					if (warehouse.equalsIgnoreCase("SECUNDERABAD")) 
					{
						st_warehouse2BodyComboList.get(i).click();
						enter_Warehouse2Txt.sendKeys(Keys.TAB);
						break;
					}
				}
				
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				enter_ItemTxt.sendKeys(Keys.SPACE);
				
				int itemsCount=itemComboList.size();
				
				for (int i = 0; i < itemsCount; i++) 
				{
					String item=itemComboList.get(i).getText();
					
					System.out.println(i+"  "+item);
					
					if (item.equalsIgnoreCase("WA COGS ITEM")) 
					{
						itemComboList.get(i).click();
						enter_ItemTxt.sendKeys(Keys.TAB);
						break;
					}
				}	
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
				enter_UnitTxt.sendKeys(Keys.TAB);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
				enter_Quantity.sendKeys("2");
		
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
				select1stRow_6thColumn.click();
				enter_Rate.sendKeys("3");
				enter_Rate.sendKeys(Keys.TAB);
				
				enter_Gross.sendKeys(Keys.TAB);

				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
				RMAPopupSerialNoExpansionBtn.click();
				
				Thread.sleep(2000);
				
				for(int i=4;i<=5;i++)
				{
					searchRMASearchTxtField.clear();
					searchRMASearchTxtField.sendKeys(String.valueOf(i));
					Thread.sleep(2000);
					getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]")).click();
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
				searchRMAOkBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
				salesInvoiceRMAPopupOkBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				String expMessage2="Voucher saved Successfully";
				
				String actMessage2=checkValidationMessage(expMessage2);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);

				return false;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
		}
		
			
			
		
		
			
		// Suspending Stock Transfer Voucher (From Voucher Home  Page and Entry Page) where Transferred stocks consumed in Sales Invoice 	
			
		public boolean checkSuspendingStockTransferVoucher2InEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransfersVoucher));
			stockTransfersVoucher.click();
			
			Thread.sleep(3000);
			
			int VouchersCount = grid_VoucherNoList.size();
			
			for (int i = 0; i < VouchersCount; i++) 
			{
				String VoucherNo = grid_VoucherNoList.get(i).getText();
				
				if (VoucherNo.equalsIgnoreCase("2")) 
				{
					getAction().doubleClick(grid_CheckBoxList.get(i)).build().perform();
				}
			}
			
			String LoadingMessage=checkValidationMessage("Voucher loaded Successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
			new_SuspendBtn.click();
			
			String expMessage="RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("********* Error Message: "+actMessage+"  value expected  "+expMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
				previousBtn.click();
				
				String LoadingMessage1=checkValidationMessage("Voucher loaded Successfully");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				String expMessage2="Voucher saved Successfully";
				
				String actMessage2=checkValidationMessage(expMessage2);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
				
				
				
				
				
		// Should not allow to Suspend the Voucher, already consumed in sales invoice
		public boolean checkStockTransferVoucher2SuspendInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			int VouchersCount = grid_VoucherNoList.size();
			
			for (int i = 0; i < VouchersCount; i++) 
			{
				String VoucherNo = grid_VoucherNoList.get(i).getText();
				
				if (VoucherNo.equalsIgnoreCase("2")) 
				{
					if (grid_CheckBoxList.get(i).isSelected()) 
					{
						break;
					} 
					else 
					{
						grid_CheckBoxList.get(i).click();
						break;
					}
				}
			}
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(suspendBtn));
			suspendBtn.click();

			String expMessage="RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
		
			System.out.println("********* Error Message: "+actMessage+"  value expected  "+expMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("***Test Pass: NOT allowing to Suspend the voucher in Homepage***");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				
				for (int i = 0; i < VouchersCount; i++) 
				{
					String VoucherNo = grid_VoucherNoList.get(i).getText();
					
					if (VoucherNo.equalsIgnoreCase("2")) 
					{
						getAction().doubleClick(grid_CheckBoxList.get(i)).build().perform();
					}
				}
				
				String LoadingMessage1=checkValidationMessage("Voucher loaded Successfully");
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				String expMessage2="Voucher saved Successfully";
				
				String actMessage2=checkValidationMessage(expMessage2);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
	
				
				
				
				
				
				
		// Verify Transferring quantity from Multiple RMA of RMA outward to Multiple RMA of RMA Inward pop-up 
			
		public boolean checkConsumingMultipleRMAQuantityInStockTransferVoucher3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransfersVoucher));
			stockTransfersVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String LoadingMessage=checkValidationMessage("Voucher loaded Successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferWarehouse1));
			stockTransferWarehouse1.click();
			stockTransferWarehouse1.sendKeys(Keys.SPACE);
            
			int warehouse1list=st_warehouse1HeaderComboList.size();
			System.out.println("******* Warehouse-1 List: "+warehouse1list);	
			
              for (int i = 0; i < warehouse1list; i++) 
				{
					String warehouse=st_warehouse1HeaderComboList.get(i).getText();
					
					System.out.println(i+"  "+warehouse);
					
					if (warehouse.equalsIgnoreCase("HYDERABAD")) 
					{
						System.out.println("Entered If Loop");
						
						st_warehouse1HeaderComboList.get(i).click();
						
						stockTransferWarehouse1.sendKeys(Keys.TAB);
						
						break;
					}
				}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
            
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Warehouse2Txt));
            enter_Warehouse2Txt.sendKeys(Keys.SPACE);
            
			int warehouselist=st_warehouse2BodyComboList.size();
			
			System.out.println("******* Warehouse List: "+warehouselist);				
			
			 for (int i = 0; i < warehouselist; i++) 
			{
				String warehouse=st_warehouse2BodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("SECUNDERABAD")) 
				{
					System.out.println("Entered If Loop");
					
					st_warehouse2BodyComboList.get(i).click();
					
					enter_Warehouse2Txt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					System.out.println("Entered If Loop");
					
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}	
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			enter_UnitTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("100");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
			select1stRow_6thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("1");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			for(int i=4;i < 104;i++)
			{
				getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]")).click();
				int temp=i+1;
				JavascriptExecutor js = (JavascriptExecutor) getDriver();
				js.executeScript("arguments[0].scrollIntoView();",getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+temp+"]/td[1]")) );
			}
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA = enter_RMA.getAttribute("value");
			String expRMA = "6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,"
							+ "51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,"
							+ "101,102,103,Q10,Q11";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
			new_newBtn.click();
			
			getWaitForAlert();
			String actAlert=getAlert().getText();
			String expAlert1="The contents of the voucher are changed.";
			String expAlert2="Are you sure that you want to ignore the changes?";
			getAlert().accept();
			
			System.out.println(actRMA);
			System.out.println(expRMA);
			System.out.println("Alert : "+actAlert+" Value Expected "+expAlert1+" "+expAlert2);
			
			
			if(actRMA.equalsIgnoreCase(expRMA) && actAlert.startsWith(expAlert1) && actAlert.endsWith(expAlert2))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
				
		
				
				
				
					
					
		
		
				
				
		// Verify allowing to allocating Qty in RMA inward pop up without allocating Qty in RMA outward pop-up and save the voucher by keeping RMA columns blank in Entry Page
			
		public boolean checkSavingStockTransferVoucherWithoutRMA() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransfersVoucher));
			stockTransfersVoucher.click();
		
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		
			String LoadingMessage=checkValidationMessage("Voucher loaded Successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferWarehouse1));
			stockTransferWarehouse1.click();
			stockTransferWarehouse1.sendKeys(Keys.SPACE);
            
			int warehouse1list=st_warehouse1HeaderComboList.size();
			
            for (int i = 0; i < warehouse1list; i++) 
			{
				String warehouse=st_warehouse1HeaderComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("HYDERABAD")) 
				{
					st_warehouse1HeaderComboList.get(i).click();
					stockTransferWarehouse1.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Warehouse2Txt));
            enter_Warehouse2Txt.sendKeys(Keys.SPACE);
            
			int warehouse2list=st_warehouse2BodyComboList.size();
			
			for (int i = 0; i < warehouse2list; i++) 
			{
				String warehouse=st_warehouse2BodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("SECUNDERABAD")) 
				{
					st_warehouse2BodyComboList.get(i).click();
					enter_Warehouse2Txt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					enter_ItemTxt.sendKeys(Keys.TAB);
					break;
				}
			}	
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			enter_UnitTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("1");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
			select1stRow_6thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("1");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupCancelBtn));
			salesInvoiceRMAPopupCancelBtn.click();
		
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
		
			String expMessage="Provide data for mandatory fieldsRMA";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("*********** Error Message: "+actMessage+"  value expected  "+expMessage);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			
			Thread.sleep(2000);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
	
			
			
			
				
				
			
			
		// Verify Delete Row option in the Inward / Outward Voucher and Search Option in the RMA Pop-up 		
			
		public boolean checkDeletingARowInStockTransferVoucher3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
			inventoryTransactionsStocksMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransfersVoucher));
			stockTransfersVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String LoadingMessage=checkValidationMessage("Voucher loaded Successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferWarehouse1));
			stockTransferWarehouse1.click();
			stockTransferWarehouse1.sendKeys(Keys.SPACE);
            
			int warehouselist=st_warehouse1HeaderComboList.size();
			
            for (int i = 0; i < warehouselist; i++) 
			{
				String warehouse=st_warehouse1HeaderComboList.get(i).getText();
				
				if (warehouse.equalsIgnoreCase("HYDERABAD")) 
				{
					st_warehouse1HeaderComboList.get(i).click();
					stockTransferWarehouse1.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Warehouse2Txt));
            enter_Warehouse2Txt.sendKeys(Keys.SPACE);
            
			int warehouse1list=st_warehouse2BodyComboList.size();
			
			for (int i = 0; i < warehouse1list; i++) 
			{
				String warehouse=st_warehouse2BodyComboList.get(i).getText();
				
				if (warehouse.equalsIgnoreCase("SECUNDERABAD")) 
				{
					st_warehouse2BodyComboList.get(i).click();
					enter_Warehouse2Txt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					enter_ItemTxt.sendKeys(Keys.TAB);
					break;
				}
			}	
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			enter_UnitTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("5");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
			select1stRow_6thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("1");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			for(int i=4;i<=8;i++)
			{
				getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]")).click();
				int temp=i+1;
				JavascriptExecutor js = (JavascriptExecutor) getDriver();
				js.executeScript("arguments[0].scrollIntoView();",getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+temp+"]/td[1]")) );
			}
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
		
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA=enter_RMA.getAttribute("value");
			String expRMA="6,7,8,9,10";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(firstRowIndex));
			JavascriptExecutor js = (JavascriptExecutor) getDriver();
			js.executeScript("arguments[0].scrollIntoView();",firstRowIndex);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(firstRowIndex));
			firstRowIndex.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteRowBtn));
			deleteRowBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			boolean actWarehouse2=select1stRow_1stColumn.getText().isEmpty();
			boolean expWarehouse2=true;
					
			System.out.println("********* RMA Column  : "+actRMA        +" Value Expected "+expRMA);
			System.out.println("********* Warehouse-2 : "+actWarehouse2 +" Value Expected "+expWarehouse2);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			
			Thread.sleep(2000);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actWarehouse2==expWarehouse2)
			{
				return true;
			}
			else
			{
				return false;
			}
		}
			
				
				
				
				
		
		
				
		
		// Save Purchase Voucher by verifying whether RMA accepting to allocate Quantity upto their capacity or more than their capacity .		
				
		public boolean checkSavingPurchaseVATVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String LoadingMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
				
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+vendorAccount);
				
				if (vendorAccount.equalsIgnoreCase("Vendor A"))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase("DUBAI"))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();
			
			enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
			int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
			
			for (int i = 0; i < pvvGridWarehouseListCount; i++) 
			{
				String warehouse=pvvGridWarehouseList.get(i).getText();
				if (warehouse.equalsIgnoreCase("Hyderabad")) 
				{
					pvvGridWarehouseList.get(i).click();
					break;
				}
			}
	        enter_WarehouseTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
	        int pvvGridItemListCount=pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) 
			{
				String Item=pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxCode));
	        enter_PvTaxCode.sendKeys(Keys.TAB);
				
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	        enter_AccountTxt.click();
			
	        Thread.sleep(2000);
	        
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			enter_Quantity.sendKeys("50");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("25");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
			select1stRow_19thColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys("a");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA=enter_RMA.getAttribute("value");
			String expRMA= "a,a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,a22,a23,a24,a25,"
					       + "a26,a27,a28,a29,a30,a31,a32,a33,a34,a35,a36,a37,a38,a39,a40,a41,a42,a43,a44,a45,a46,a47,a48,a49";
			
			System.out.println(actRMA);
			System.out.println(expRMA);
			
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  
			if(actRMA.equalsIgnoreCase(expRMA) && actSaving==expSaving) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
				
				
				
		
		
				
		// Verify Save Option in Inward / Outward Voucher By Giving Voucher Quantity as negative Qty
				
		public boolean checkSavingPurchaseVoucherVATVoucher3WithNegativeQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String LoadingMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
				
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+vendorAccount);
				
				if (vendorAccount.equalsIgnoreCase("Vendor A"))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase("DUBAI"))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();
			
			enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
			int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
			
			for (int i = 0; i < pvvGridWarehouseListCount; i++) 
			{
				String warehouse=pvvGridWarehouseList.get(i).getText();
				if (warehouse.equalsIgnoreCase("Hyderabad")) 
				{
					pvvGridWarehouseList.get(i).click();
					break;
				}
			}
	        enter_WarehouseTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
	        int pvvGridItemListCount=pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) 
			{
				String Item=pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxCode));
	        enter_PvTaxCode.sendKeys(Keys.TAB);
				
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	        enter_AccountTxt.click();
			
	        Thread.sleep(2000);
	        
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			enter_Quantity.sendKeys("-10");
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("15");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
			select1stRow_19thColumn.click();
					
			String expMessage="Quantity Should be Greater than zero.";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("Error Message : "+actMessage+"  Value Expected  "+expMessage);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupCancelBtn));
			RMAPopupCancelBtn.click();
		
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
		
			getWaitForAlert();
			getAlert().accept();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			
			Thread.sleep(2000);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
				
				
		
		
		
		
		
		
		// Sales Invoice Voucher 

		public boolean checkSavingSalesInvoiceVATVoucher5WithNegativeQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
		
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			System.out.println("Customer Account List Size: "+customerAccountCount);
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase("Customer A"))
				{
					customerAccountComboList.get(i).click();
					
					customerAccountTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase("DUBAI"))
				{
					departmentComboList.get(i).click();
					
					departmentTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys("Abu Dhabi");
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("SECUNDERABAD"))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys("-5");
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys("-5");
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("15");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
			select1stRow_21stColumn.click();
			
			String expMessage="Quantity Should be Greater than zero.";
			
			String actMessage=checkValidationMessage(expMessage);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupCancelBtn));
			salesInvoiceRMAPopupCancelBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			
			Thread.sleep(2000);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
				
				
				
				
				
				
				
				
		public boolean checkSavingSalesInvoiceVATVoucher5() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
		
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			System.out.println("Customer Account List Size: "+customerAccountCount);
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase("Customer A"))
				{
					customerAccountComboList.get(i).click();
					
					customerAccountTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase("DUBAI"))
				{
					departmentComboList.get(i).click();
					
					departmentTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys("Abu Dhabi");
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("HYDERABAD"))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys("10");
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys("5");
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
			select1stRow_21stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
		
			Thread.sleep(2000);
			
			for(int i=4;i<=18;i++)
			{
				getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[1]")).click();
				int temp=i+1;
				JavascriptExecutor js = (JavascriptExecutor) getDriver();
				js.executeScript("arguments[0].scrollIntoView();",getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+temp+"]/td[1]")) );
			}
		
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
		
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA=enter_RMA.getAttribute("value");
			String expRMA="6,7,8,9,a,10,11,12,13,14,15,16,17,18,19";
			
			System.out.println(actRMA);
			System.out.println(expRMA);
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
			String actNewReference=newReferenceTxt.getAttribute("value");
			String expNewReference="157.50";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  

			System.out.println("********* New Reference: "+actNewReference +"  value expected  "+expNewReference);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actNewReference.equalsIgnoreCase(expNewReference) && actSaving==expSaving) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
		
		
		
		
		
		
				
		//Sales Return Voucher
		
		public boolean checkSavingSalesReturnsVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
     		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesReturnsVoucher));
			salesReturnsVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();

			String ScreenOpeningMessage=checkValidationMessage("Screen opened");	
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				if (customerAccount.equalsIgnoreCase("Customer A"))
				{
					customerAccountComboList.get(i).click();
					customerAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				if (department.equalsIgnoreCase("INDIA"))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("SECUNDERABAD"))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					enter_ItemTxt.sendKeys(Keys.TAB);
					break;
				}
			}	
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
			select1stRow_5thColumn.click();
			enter_Quantity.sendKeys("15");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
			select1stRow_7thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("1");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_13thColumn));
			select1stRow_13thColumn.click();

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys("s");
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMAColumn=enter_RMA.getAttribute("value");
			String expRMAColumn="s,s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,s11,s12,s13,s14";
			
			System.out.println(actRMAColumn);
			System.out.println(expRMAColumn);
					
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
			String actNewReference=newReferenceTxt.getAttribute("value");
			String expNewReference="15.00";
			
			System.out.println("********* New Reference: "+actNewReference+" value expected "+expNewReference);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String expMessage1="Voucher saved successfully";
			String expMessage2="1";
			
			String actMessage=checkValidationMessage(expMessage1);

			System.out.println("*********** Error Message: "+actMessage+"  value expected  "+expMessage1+" : "+expMessage2);
			
			if(actRMAColumn.equalsIgnoreCase(expRMAColumn) && actNewReference.equalsIgnoreCase(expNewReference) && actMessage.startsWith(expMessage1) 
					&& actMessage.endsWith(expMessage2)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
			
				
				
					
					

		
		
		
		
		// Verify Sales Invoice Voucher RMA Pop-up whether Returned quantity is displaying or not 
	
		public boolean checkSalesReturnRMAQuantityInSalesInvoiceVATVoucher6() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			System.out.println("Customer Account List Size: "+customerAccountCount);
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase("Customer A"))
				{
					customerAccountComboList.get(i).click();
					
					customerAccountTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase("DUBAI"))
				{
					departmentComboList.get(i).click();
					
					departmentTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys("Abu Dhabi");
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("SECUNDERABAD"))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys("1");
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys("1");
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("1");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
			select1stRow_21stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
		
			Thread.sleep(2000);
			
			ArrayList<String>  actRMA= new ArrayList<String>(); 
			for(int i=1; i < 16; i++)
			{
				String rma =getDriver().findElement(By.xpath("//*[@id='RMASearchTable']/tbody/tr["+i+"]/td[2]")).getText();
				
				actRMA.add(rma);
			}
			
			 ArrayList<String>  expRMA= new ArrayList<String>(); 
			
			 expRMA.add("s, s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14");
			
			System.out.println(actRMA);
			System.out.println(expRMA);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMACancelBtn));
			searchRMACancelBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupCloseBtn));
			RMAPopupCloseBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
		
			Thread.sleep(2000);
			
			
			if(actRMA.toString().equalsIgnoreCase(expRMA.toString())) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
			
			
		
		
		
		
		// Save Purchase Voucher forRMA and for HYD Warehouse RMA Allocate full upto their capacity
			
				
		//Purchase Voucher 
	
		public boolean checkPurchaseVoucher3EntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
				
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+vendorAccount);
				
				if (vendorAccount.equalsIgnoreCase("Vendor A"))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase("DUBAI"))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();
			
			enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
			int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
			
			for (int i = 0; i < pvvGridWarehouseListCount; i++) 
			{
				String warehouse=pvvGridWarehouseList.get(i).getText();
				if (warehouse.equalsIgnoreCase("Hyderabad")) 
				{
					pvvGridWarehouseList.get(i).click();
					break;
				}
			}
	        enter_WarehouseTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
	        int pvvGridItemListCount=pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) 
			{
				String Item=pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxCode));
	        enter_PvTaxCode.sendKeys(Keys.TAB);
				
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	        enter_AccountTxt.click();
			
	        Thread.sleep(2000);
	        
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			enter_Quantity.sendKeys("377");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
			select1stRow_19thColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys("c");
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(6000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA = enter_RMA.getAttribute("value");
			String expRMA = "c,c1,c2,c3,c4,c5,c6,c7,c8,c9,c10,c11,c12,c13,c14,c15,c16,c17,c18,c19,c20,c21,c22,c23,c24,c25,c26,c27,c28,c29,c30,c31,c32,c33,c34,c35,"
							+ "c36,c37,c38,c39,c40,c41,c42,c43,c44,c45,c46,c47,c48,c49,c50,c51,c52,c53,c54,c55,c56,c57,c58,c59,c60,c61,c62,c63,c64,c65,c66,c67,c68,c69,c70,c71,c72,c73,"
							+ "c74,c75,c76,c77,c78,c79,c80,c81,c82,c83,c84,c85,c86,c87,c88,c89,c90,c91,c92,c93,c94,c95,c96,c97,c98,c99,c100,c101,c102,c103,c104,c105,c106,c107,c108,c109,"
							+ "c110,c111,c112,c113,c114,c115,c116,c117,c118,c119,c120,c121,c122,c123,c124,c125,c126,c127,c128,c129,c130,c131,c132,c133,c134,c135,c136,c137,c138,c139,c140,"
							+ "c141,c142,c143,c144,c145,c146,c147,c148,c149,c150,c151,c152,c153,c154,c155,c156,c157,c158,c159,c160,c161,c162,c163,c164,c165,c166,c167,c168,c169,c170,c171,"
							+ "c172,c173,c174,c175,c176,c177,c178,c179,c180,c181,c182,c183,c184,c185,c186,c187,c188,c189,c190,c191,c192,c193,c194,c195,c196,c197,c198,c199,c200,c201,c202,"
							+ "c203,c204,c205,c206,c207,c208,c209,c210,c211,c212,c213,c214,c215,c216,c217,c218,c219,c220,c221,c222,c223,c224,c225,c226,c227,c228,c229,c230,c231,c232,c233,"
							+ "c234,c235,c236,c237,c238,c239,c240,c241,c242,c243,c244,c245,c246,c247,c248,c249,c250,c251,c252,c253,c254,c255,c256,c257,c258,c259,c260,c261,c262,c263,c264,"
							+ "c265,c266,c267,c268,c269,c270,c271,c272,c273,c274,c275,c276,c277,c278,c279,c280,c281,c282,c283,c284,c285,c286,c287,c288,c289,c290,c291,c292,c293,c294,c295,"
							+ "c296,c297,c298,c299,c300,c301,c302,c303,c304,c305,c306,c307,c308,c309,c310,c311,c312,c313,c314,c315,c316,c317,c318,c319,c320,c321,c322,c323,c324,c325,c326,"
							+ "c327,c328,c329,c330,c331,c332,c333,c334,c335,c336,c337,c338,c339,c340,c341,c342,c343,c344,c345,c346,c347,c348,c349,c350,c351,c352,c353,c354,c355,c356,c357,"
							+ "c358,c359,c360,c361,c362,c363,c364,c365,c366,c367,c368,c369,c370,c371,c372,c373,c374,c375,c376";
			
			System.out.println(actRMA);
			System.out.println(expRMA);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
			String actNewReference=newReferenceTxt.getAttribute("value");
			String expNewReference="3958.50";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String expMessage1="Voucher saved successfully";
			String expMessage2="3";
			
			String actMessage=checkValidationMessage(expMessage1);
			
			System.out.println("************ New Reference: "+actNewReference +"  value expected  "+expNewReference);
			System.out.println("************ Error Message: "+actMessage      +"  value expected  "+expMessage1+" : "+expMessage2);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actNewReference.equalsIgnoreCase(expNewReference) && actMessage.startsWith(expMessage1) 
				&& actMessage.endsWith(expMessage2)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
		
		
		
		
		
		
				
				
		//Purchase Returns
				
		public boolean checkSavingPurchaseReturnsVoucher1ConsumingAllAvailableQuantityInWarehouseHyderabad() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseReturnsVoucher));
			purchaseReturnsVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();

			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
			
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				if (vendorAccount.equalsIgnoreCase("Vendor A"))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase("DUBAI"))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				if (warehouse.equalsIgnoreCase("HYDERABAD"))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					enter_ItemTxt.sendKeys(Keys.TAB);
					break;
				}
			}	
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
			select1stRow_7thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("525");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_10thColumn));
			select1stRow_10thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_15thColumn));
			select1stRow_15thColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
			searchRMASelectAllChkBox.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA = enter_RMA.getAttribute("value");
			String expRMA = "1,2,3,c,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,"
							+ "56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,"
							+ "94,95,96,97,98,99,a1,a2,a3,a4,a5,a6,a7,a8,a9,c1,c2,c3,c4,c5,c6,c7,c8,c9,100,101,102,103,a10,a11,a12,a13,a14,a15,a16,"
							+ "a17,a18,a19,a20,a21,a22,a23,a24,a25,a26,a27,a28,a29,a30,a31,a32,a33,a34,a35,a36,a37,a38,a39,a40,a41,a42,a43,a44,a45,"
							+ "a46,a47,a48,a49,c10,c11,c12,c13,c14,c15,c16,c17,c18,c19,c20,c21,c22,c23,c24,c25,c26,c27,c28,c29,c30,c31,c32,c33,c34,"
							+ "c35,c36,c37,c38,c39,c40,c41,c42,c43,c44,c45,c46,c47,c48,c49,c50,c51,c52,c53,c54,c55,c56,c57,c58,c59,c60,c61,c62,c63,"
							+ "c64,c65,c66,c67,c68,c69,c70,c71,c72,c73,c74,c75,c76,c77,c78,c79,c80,c81,c82,c83,c84,c85,c86,c87,c88,c89,c90,c91,c92,"
							+ "c93,c94,c95,c96,c97,c98,c99,Q10,Q11,Q12,Q13,Q14,Q15,Q16,Q17,Q18,Q19,Q20,Q21,c100,c101,c102,c103,c104,c105,c106,c107,"
							+ "c108,c109,c110,c111,c112,c113,c114,c115,c116,c117,c118,c119,c120,c121,c122,c123,c124,c125,c126,c127,c128,c129,c130,"
							+ "c131,c132,c133,c134,c135,c136,c137,c138,c139,c140,c141,c142,c143,c144,c145,c146,c147,c148,c149,c150,c151,c152,c153,"
							+ "c154,c155,c156,c157,c158,c159,c160,c161,c162,c163,c164,c165,c166,c167,c168,c169,c170,c171,c172,c173,c174,c175,c176,"
							+ "c177,c178,c179,c180,c181,c182,c183,c184,c185,c186,c187,c188,c189,c190,c191,c192,c193,c194,c195,c196,c197,c198,c199,"
							+ "c200,c201,c202,c203,c204,c205,c206,c207,c208,c209,c210,c211,c212,c213,c214,c215,c216,c217,c218,c219,c220,c221,c222,"
							+ "c223,c224,c225,c226,c227,c228,c229,c230,c231,c232,c233,c234,c235,c236,c237,c238,c239,c240,c241,c242,c243,c244,c245,"
							+ "c246,c247,c248,c249,c250,c251,c252,c253,c254,c255,c256,c257,c258,c259,c260,c261,c262,c263,c264,c265,c266,c267,c268,"
							+ "c269,c270,c271,c272,c273,c274,c275,c276,c277,c278,c279,c280,c281,c282,c283,c284,c285,c286,c287,c288,c289,c290,c291,"
							+ "c292,c293,c294,c295,c296,c297,c298,c299,c300,c301,c302,c303,c304,c305,c306,c307,c308,c309,c310,c311,c312,c313,c314,"
							+ "c315,c316,c317,c318,c319,c320,c321,c322,c323,c324,c325,c326,c327,c328,c329,c330,c331,c332,c333,c334,c335,c336,c337,"
							+ "c338,c339,c340,c341,c342,c343,c344,c345,c346,c347,c348,c349,c350,c351,c352,c353,c354,c355,c356,c357,c358,c359,c360,"
							+ "c361,c362,c363,c364,c365,c366,c367,c368,c369,c370,c371,c372,c373,c374,c375,c376";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
			String actNewReference=newReferenceTxt.getAttribute("value");
			String expNewReference="6163.50";

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			Thread.sleep(10000);
			
			String expMessage1="Voucher saved successfully";
			String expMessage2="1";
			
			String actMessage=checkValidationMessage(expMessage1);
			
			System.out.println("New Reference: "+actNewReference +"  value expected  "+expNewReference);
			System.out.println("Error Message: "+actMessage      +"  value expected  "+expMessage1+" : "+expMessage2);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actNewReference.equalsIgnoreCase(expNewReference) && actMessage.startsWith(expMessage1) 
				&& actMessage.endsWith(expMessage2) ) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
		
		
		
		
		
		
					
		//Sales Invoice Voucher 
		
		public boolean checkSavingSalesInvoiceVoucher6ConsumingAllAvailableQuantityInWarehouseSecunderabad() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
		
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			System.out.println("Customer Account List Size: "+customerAccountCount);
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase("Customer A"))
				{
					customerAccountComboList.get(i).click();
					
					customerAccountTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase("DUBAI"))
				{
					departmentComboList.get(i).click();
					
					departmentTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys("Abu Dhabi");
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase("SECUNDERABAD"))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase("WA COGS ITEM")) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys("12");
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys("15");
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
			select1stRow_21stColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
			searchRMASelectAllChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA=enter_RMA.getAttribute("value");
			String expRMA="s,s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,s11,s12,s13,s14,SEC#RMA1,SEC#RMA2,SEC#RMA3,SEC#RMA4,SEC#RMA5,SEC#RMA6,"
						  + "SEC#RMA7,SEC#RMA8,SEC#RMA9,SEC#RMA10,SEC#RMA11,SEC#RMA12";
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
			String actNewReference=newReferenceTxt.getAttribute("value");
			
			String expNewReference="283.50";

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  
			System.out.println("****** RMA Column                 : "+actRMA			+"  value expected  "+expRMA);
			System.out.println("****** New Reference              : "+actNewReference	+"  value expected  "+expNewReference);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actNewReference.equalsIgnoreCase(expNewReference) 
				&& actSaving==expSaving) 
			{		
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
				
				
				
			
		
		
		
		
		//Sales Return Voucher
		public boolean checkEditingSalesReturnsSavingWithBelowConsumedQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesReturnsVoucher));
			salesReturnsVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			getAction().doubleClick(grid_ChkBox1).build().perform();
			
			String LoadingMessage=checkValidationMessage("Loading Message");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
			select1stRow_5thColumn.click();
			enter_Quantity.clear();
			enter_Quantity.sendKeys("2");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
			select1stRow_12thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupExpansionBtn));
			RMAPopupExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA=enter_RMA.getAttribute("value");
			String expRMA="s,s1";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String expMessage="RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("RMA Column    : "+actRMA     +"  value expected  "+expRMA);
			System.out.println("Error Message : "+actMessage +"  value expected  "+expMessage);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				getWaitForAlert();
				getAlert().accept();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				getWaitForAlert();
				getAlert().accept();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
				
				
				
				
				
				
				
				
		// Deleting Sales Return (from Voucher Home page & Voucher Entry Page) where RMA quantity is Consumed in Sales Invoice
				
		public boolean checkDeletingSalesReturnVoucher1InEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesReturnsVoucher));
			salesReturnsVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			getAction().doubleClick(grid_ChkBox1).build().perform();
			
			String LoadingMessage=checkValidationMessage("Loading Message");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
			new_DeleteBtn.click();
			
			getWaitForAlert();
			getAlert().accept();

			String expMessage1="Deleting this document has has resulted in negative stock.,";
			String expMessage2="RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage1);
			
			System.out.println("Error Message: "+actMessage+"  value expected  "+expMessage1+" "+expMessage2);
			
			if(actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
				
				
				
				
				
				
			
		public boolean checkDeletingSalesReturnVoucher1InHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			grid_ChkBox1.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesReturnsVoucherHomeDeleteBtn));
			salesReturnsVoucherHomeDeleteBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
			
			String expMessage="VoucherNo - 1: RMA has already been used in another voucher";
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("Error Message : "+actMessage+"  value expected  "+expMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))  
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
				
				
				
				
		
		
		// Suspending Sales Return (from Voucher Home page & Voucher Entry Page) where RMA quantity is Consumed in Sales Invoice
				
		public boolean checkSuspendingSalesReturnVoucher1InEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesReturnsVoucher));
			salesReturnsVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			getAction().doubleClick(grid_ChkBox1).build().perform();
			
			String LoadingMessage=checkValidationMessage("Loading Message");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
			new_SuspendBtn.click();
			
			String expMessage="RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);	
			
			System.out.println("Error Message : "+actMessage+"  value expected  "+expMessage);
			//Voucher saved successfully : 1
			
			if(actMessage.startsWith(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
				previousBtn.click();
				
				String LoadingMessage1=checkValidationMessage("Loading Message");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
				billRefPickIcon.click();
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
				billRefOkBtn.click();
				
				String expMessage2="Voucher saved successfully";
				
				String actMessage2=checkValidationMessage(expMessage2);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
				
		
		
		
				
				
		// Should NOT allow to Suspend the Voucher
		public boolean checkSuspendingSalesReturnVoucher1InHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			if (grid_ChkBox1.isSelected()==false) 
			{
				grid_ChkBox1.click();
			}
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(suspendBtn));
			suspendBtn.click();
			
			String expMessage="VoucherNo - 1: RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("Error Message : "+actMessage+"  value expected  "+expMessage);
			
			if(actMessage.contains(expMessage)) 
			{
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
				getAction().doubleClick(grid_ChkBox1).build().perform();
				
				String LoadingMessage1=checkValidationMessage("Loading Message");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
				billRefPickIcon.click();
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
				billRefOkBtn.click();
				
				String expMessage2="Voucher saved successfully";
				
				String actMessage2=checkValidationMessage(expMessage2);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();

				Thread.sleep(2000);
				
				return false;
			}
		}
		
				
				
				
				
				
				
				
		// Verify Edit option in Purchase Voucher and Update the Quantity to Below the Consumed quantity in Purchase Return 
				
		public boolean checkEditingPurchaseVoucherVATVoucher3SavingWithQuantityBelowPurchaseReturn() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			
			int count = grid_VoucherNoList.size();
			
			for (int i = 0; i < count; i++) 
			{
				String VoucherNo = grid_VoucherNoList.get(i).getText();
				if (VoucherNo.equalsIgnoreCase("3")) 
				{
					getAction().doubleClick(grid_CheckBoxList.get(i)).build().perform();
				}
			}
			
			String LoadingMessage=checkValidationMessage("Loading Message");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			enter_Quantity.sendKeys("350");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
			select1stRow_19thColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupExpansionBtn));
			RMAPopupExpansionBtn.click();
		
			Thread.sleep(2000);
					
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
		
			Thread.sleep(2000);
								
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
		
			String expMessage="RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				getWaitForAlert();
				getAlert().accept();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();

				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				getWaitForAlert();
				getAlert().accept();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();

				Thread.sleep(2000);
				
				return false;
			}
		}
				
				
				
		
		
		
		
		
		
				
		// Deleting Purchase Voucher (from Voucher Home page & Voucher Entry Page) where RMA quantity is Consumed in Purchase Return 
				
		public boolean checkDeletingPurchaseVoucherVATVoucher3InEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			
			int count = grid_VoucherNoList.size();
			
			for (int i = 0; i < count; i++) 
			{
				String VoucherNo = grid_VoucherNoList.get(i).getText();
				if (VoucherNo.equalsIgnoreCase("3")) 
				{
					getAction().doubleClick(grid_CheckBoxList.get(i)).build().perform();
				}
			}
			
			String LoadingMessage=checkValidationMessage("Loading Message");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
			new_DeleteBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
			
			String expMessage1="Deleting this document has has resulted in negative stock.,";
			String expMessage2="RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage1);
			
			System.out.println("Error Message: "+actMessage+"  value expected  "+expMessage1+" "+expMessage2);
			
			if(actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				return false;
			}
		}
				
		
		
		
		
		
		public boolean checkDeletingPurchaseVoucherVATVoucher3InHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			
			int count = grid_VoucherNoList.size();
			
			for (int i = 0; i < count; i++) 
			{
				String VoucherNo = grid_VoucherNoList.get(i).getText();
				
				if (VoucherNo.equalsIgnoreCase("3")) 
				{
					if (grid_CheckBoxList.get(i).isSelected()==false) 
					{
						grid_CheckBoxList.get(i).click();
					}
				}
			}
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
			deleteBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
			
			String expMessage="VoucherNo - 3: RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("Error Message : "+actMessage+"  value expected  "+expMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
				
				
				
				
				
				
				
		// Suspending Purchase Voucher (from Voucher Home page & Voucher Entry Page) where RMA quantity is Consumed in Purchase Return  	
				
		public boolean checkSuspendingPurchaseVoucherVATVoucher3InEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			
			int count = grid_VoucherNoList.size();
			
			for (int i = 0; i < count; i++) 
			{
				String VoucherNo = grid_VoucherNoList.get(i).getText();
				if (VoucherNo.equalsIgnoreCase("3")) 
				{
					getAction().doubleClick(grid_CheckBoxList.get(i)).build().perform();
				}
			}
			
			String LoadingMessage=checkValidationMessage("Loading Message");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
					
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
			new_SuspendBtn.click();
			
			String expMessage="RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("Error Message : "+actMessage+"  value expected  "+expMessage);
			
			Thread.sleep(2000);
			
			if(actMessage.contains(expMessage))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
				
				for (int i = 0; i < count; i++) 
				{
					String VoucherNo = grid_VoucherNoList.get(i).getText();
					if (VoucherNo.equalsIgnoreCase("3")) 
					{
						getAction().doubleClick(grid_CheckBoxList.get(i)).build().perform();
					}
				}
				
				String LoadingMessage1=checkValidationMessage("Loading Message");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
				billRefPickIcon.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
				billRefOkBtn.click();
				
				String expMessage2="Voucher saved successfully";
				
				String actMessage2=checkValidationMessage(expMessage2);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
				
			
			
				
		// Should not allow to Suspend the Voucher
		public boolean checkSuspendingPurchaseVoucherVATVoucher3InHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			
			int count = grid_VoucherNoList.size();
			
			for (int i = 0; i < count; i++) 
			{
				String VoucherNo = grid_VoucherNoList.get(i).getText();
				
				if (VoucherNo.equalsIgnoreCase("3")) 
				{
					if (grid_CheckBoxList.get(i).isSelected()==false) 
					{
						grid_CheckBoxList.get(i).click();
					}
				}
			}
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(suspendBtn));
			suspendBtn.click();
			
			String expMessage="VoucherNo - 3: RMA has already been used in another voucher";
			
			String actMessage=checkValidationMessage(expMessage);
			
			Thread.sleep(2000);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				return true;
			}
			else
			{
				
				for (int i = 0; i < count; i++) 
				{
					String VoucherNo = grid_VoucherNoList.get(i).getText();
					if (VoucherNo.equalsIgnoreCase("3")) 
					{
						getAction().doubleClick(grid_CheckBoxList.get(i)).build().perform();
					}
				}
				
				String LoadingMessage=checkValidationMessage("Loading Message");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
				billRefPickIcon.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
				billRefOkBtn.click();
				
				String expMessage2="Voucher saved successfully";
				
				String actMessage2=checkValidationMessage(expMessage2);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
		
		
		
		
		
		public boolean checkDeletingSalesInvoiceVATVoucher6() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			
			int count = grid_VoucherNoList.size();
			
			for (int i = 0; i < count; i++) 
			{
				String VoucherNo = grid_VoucherNoList.get(i).getText();
				if (VoucherNo.equalsIgnoreCase("6")) 
				{
					getAction().doubleClick(grid_CheckBoxList.get(i)).build().perform();
				}
			}
			
			String LoadingMessage=checkValidationMessage("Loading Message");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
			new_DeleteBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
			
			String expMessage="Voucher deleted Successfully";
			
			String actMessage=checkValidationMessage(expMessage);
			
			Thread.sleep(2000);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
			
		
		//********************************************************** RMA Reports Starts Here ******************************************************
		
		
		
		// Inventory Reports Starts from Here

		public boolean checkItemQueryOfWACOGSITEM() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemQuery));
			itemQuery.click();
			
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(iq_ItemTxt));
			iq_ItemTxt.click();
			iq_ItemTxt.sendKeys("WA COGS ITEM");
			Thread.sleep(2000);
			iq_ItemTxt.sendKeys(Keys.TAB);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(iq_getStockBtn));
			iq_getStockBtn.click();
			
			Thread.sleep(5000);
			
			String actiq_OpeningStocks				=iq_OpeningStocks.getText();
			String actiq_PendingPurchaseOrders		=iq_PendingPurchaseOrders.getText();
			String actiq_CurrentStock				=iq_CurrentStock.getText();
			String actiq_PendingSalesOrders			=iq_PendingSalesOrders.getText();
			String actiq_AvgStockRate				=iq_AvgStockRate.getText();
			String actiq_QtyToBeOrdered				=iq_QtyToBeOrdered.getText();
			String actiq_Value						=iq_Value.getText();
			String actiq_BaseUnit     				=iq_BaseUnit.getText();
			
			String expiq_OpeningStocks				="36.00";
			String expiq_PendingPurchaseOrders		="0.00";
			String expiq_CurrentStock				="27.00";
			String expiq_PendingSalesOrders			="0.00";
			String expiq_AvgStockRate				="11.74";
			String expiq_QtyToBeOrdered				="0.00";
			String expiq_Value						="316.94";
			String expiq_BaseUnit     				="Pcs";
			
			// Recent Stock Transaction -Purchases
			
			String actpurVoucherNoR1	=pur1strow_2ndColumn.getText();
			String actpurVendorR1		=pur1strow_3rdColumn.getText();
			String actpurQtyR1			=pur1strow_4thColumn.getText();
			String actpurUnitR1			=pur1strow_5thColumn.getText();
			String actpurRateR1			=pur1strow_6thColumn.getText();
			String actpurCurrencyR1		=pur1strow_7thColumn.getText();
			String actpurQtyInBaseR1	=pur1strow_8thColumn.getText();
			
			String actpurVoucherNoR2	=pur2ndrow_2ndColumn.getText();
			String actpurVendorR2		=pur2ndrow_3rdColumn.getText();
			String actpurQtyR2			=pur2ndrow_4thColumn.getText();
			String actpurUnitR2			=pur2ndrow_5thColumn.getText();
			String actpurRateR2			=pur2ndrow_6thColumn.getText();
			String actpurCurrencyR2		=pur2ndrow_7thColumn.getText();
			String actpurQtyInBaseR2	=pur2ndrow_8thColumn.getText();
			
			String actpurVoucherNoR3	=pur3rdrow_2ndColumn.getText();
			String actpurVendorR3		=pur3rdrow_3rdColumn.getText();
			String actpurQtyR3			=pur3rdrow_4thColumn.getText();
			String actpurUnitR3			=pur3rdrow_5thColumn.getText();
			String actpurRateR3			=pur3rdrow_6thColumn.getText();
			String actpurCurrencyR3		=pur3rdrow_7thColumn.getText();
			String actpurQtyInBaseR3	=pur3rdrow_8thColumn.getText();
			
			String actpurTotalWa		=pur4throw_2ndColumn.getText();
			
			
			// Expected
			
			String exppurVoucherNoR1	="NDT52:1";
			String exppurVendorR1		="Vendor A";
			String exppurQtyR1			="100.00";
			String exppurUnitR1			="Pcs";
			String exppurRateR1			="10.00";
			String exppurCurrencyR1		="Indian Rupees";
			String exppurQtyInBaseR1	="100.0000000000";
			
			String exppurVoucherNoR2	="NDT52:2";
			String exppurVendorR2		="Vendor A";
			String exppurQtyR2			="50.00";
			String exppurUnitR2			="Pcs";
			String exppurRateR2			="25.00";
			String exppurCurrencyR2		="Indian Rupees";
			String exppurQtyInBaseR2	="50.0000000000";
			
			String exppurVoucherNoR3	="NDT52:3";
			String exppurVendorR3		="Vendor A";
			String exppurQtyR3		    ="377.00";
			String exppurUnitR3			="Pcs";
			String exppurRateR3			="10.00";
			String exppurCurrencyR3		="Indian Rupees";
			String exppurQtyInBaseR3	="377.0000000000";
			
			String exppurTotalWa		="527.00";
			
			getAction().sendKeys(Keys.CONTROL,Keys.END).build().perform();
			
			
			// Recent Stock Transaction -Sales
			
			String actsalVoucherNoR1	=sal1strow_2ndColumn.getText();
			String actsalCustomerR1		=sal1strow_3rdColumn.getText();
			String actsalQtyR1			=sal1strow_4thColumn.getText();
			String actsalUnitR1			=sal1strow_5thColumn.getText();
			String actsalRateR1			=sal1strow_6thColumn.getText();
			String actsalCurrencyR1		=sal1strow_7thColumn.getText();
			String actsalQtyInBaseR1	=sal1strow_8thColumn.getText();
			
			String actsalVoucherNoR2	=sal2ndrow_2ndColumn.getText();
			String actsalCustomerR2		=sal2ndrow_3rdColumn.getText();
			String actsalQtyR2			=sal2ndrow_4thColumn.getText();
			String actsalUnitR2			=sal2ndrow_5thColumn.getText();
			String actsalRateR2			=sal2ndrow_6thColumn.getText();
			String actsalCurrencyR2		=sal2ndrow_7thColumn.getText();
			String actsalQtyInBaseR2	=sal2ndrow_8thColumn.getText();
			
			String actsalVoucherNoR3	=sal3rdrow_2ndColumn.getText();
			String actsalCustomerR3		=sal3rdrow_3rdColumn.getText();
			String actsalQtyR3			=sal3rdrow_4thColumn.getText();
			String actsalUnitR3			=sal3rdrow_5thColumn.getText();
			String actsalRateR3			=sal3rdrow_6thColumn.getText();
			String actsalCurrencyR3		=sal3rdrow_7thColumn.getText();
			String actsalQtyInBaseR3	=sal3rdrow_8thColumn.getText();
			
			String actsalVoucherNoR4	=sal4throw_2ndColumn.getText();
			String actsalCustomerR4		=sal4throw_3rdColumn.getText();
			String actsalQtyR4			=sal4throw_4thColumn.getText();
			String actsalUnitR4			=sal4throw_5thColumn.getText();
			String actsalRateR4			=sal4throw_6thColumn.getText();
			String actsalCurrencyR4		=sal4throw_7thColumn.getText();
			String actsalQtyInBaseR4	=sal4throw_8thColumn.getText();
			
			String actsalTotalWa		=sal5throw_2ndColumn.getText();

			// Expected
			String expsalVoucherNoR1	="NDT55:1";
			String expsalCustomerR1		="Customer A";
			String expsalQtyR1			="3.00";
			String expsalUnitR1			="Pcs";
			String expsalRateR1			="19.00";
			String expsalCurrencyR1		="Indian Rupees";
			String expsalQtyInBaseR1	="3.0000000000";
			
			String expsalVoucherNoR2	="NDT55:2";
			String expsalCustomerR2 	="Customer A";
			String expsalQtyR2			="9.00";
			String expsalUnitR2			="Pcs";
			String expsalRateR2			="6.00";
			String expsalCurrencyR2		="Indian Rupees";
			String expsalQtyInBaseR2	="9.0000000000";
			
			String expsalVoucherNoR3	="NDT55:4";
			String expsalCustomerR3		="Customer A";
			String expsalQtyR3			="2.00";
			String expsalUnitR3			="Pcs";
			String expsalRateR3			="3.00";
			String expsalCurrencyR3		="Indian Rupees";
			String expsalQtyInBaseR3	="2.0000000000";
			
			String expsalVoucherNoR4	="NDT55:5";
			String expsalCustomerR4 	="Customer A";
			String expsalQtyR4			="15.00";
			String expsalUnitR4			="Pcs";
			String expsalRateR4			="10.00";
			String expsalCurrencyR4		="Indian Rupees";
			String expsalQtyInBaseR4	="15.0000000000";
			
			String expsalTotalWa		="29.00";
			

			
			// WareHouse Table
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(BalancebyWareHouseExpandBtn));
			BalancebyWareHouseExpandBtn.click();
			
			Thread.sleep(3000);
			
			getAction().sendKeys(Keys.CONTROL,Keys.END).build().perform();
			
			String actwareHouseR1		=wh1strow_2ndColumn.getText();
			String actstockR1			=wh1strow_3rdColumn.getText();
			String actrateR1			=wh1strow_4thColumn.getText();
			String actvalueR1			=wh1strow_5thColumn.getText();
			
			String acttotalStock		=whStockTotal.getText();
			String acttotalValue		=whValueTotal.getText();
			
			
			String expwareHouseR1		="SECUNDERABAD";
			String expstockR1			="27.00";
			String exprateR1			="9.46";
			String expvalueR1			="255.49";
			
			String exptotalStock		="27.00";
			String exptotalValue		="255.49";
			

			// Seller And Buying rates Table
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sellingAndBuyingRatesExpandBtn));
			sellingAndBuyingRatesExpandBtn.click();
			
			Thread.sleep(2000);
			
			getAction().sendKeys(Keys.CONTROL,Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			String actsellingRateR1		=sbr1strow_2ndColumn.getText();
			String actsellingValueR1	=sbr1strow_3rdColumn.getText();
			String actbuyingRateR1		=sbr1strow_4thColumn.getText();
			String actbuyingValueR1		=sbr1strow_5thColumn.getText();
			
			String actsellingRateR2		=sbr2ndrow_2ndColumn.getText();
			String actsellingValueR2	=sbr2ndrow_3rdColumn.getText();
			String actbuyingRateR2		=sbr2ndrow_4thColumn.getText();
			String actbuyingValueR2		=sbr2ndrow_5thColumn.getText();
			
			String actsellingRateR3		=sbr3rdrow_2ndColumn.getText();
			String actsellingValueR3	=sbr3rdrow_3rdColumn.getText();
			String actbuyingRateR3		=sbr3rdrow_4thColumn.getText();
			String actbuyingValueR3		=sbr3rdrow_5thColumn.getText();
			
			String actsellingRateR4		=sbr4throw_2ndColumn.getText();
			String actsellingValueR4	=sbr4throw_3rdColumn.getText();
			String actbuyingRateR4		=sbr4throw_4thColumn.getText();
			String actbuyingValueR4		=sbr4throw_5thColumn.getText();
			
			String actsellingRateR5		=sbr5throw_2ndColumn.getText();
			String actsellingValueR5	=sbr5throw_3rdColumn.getText();
			String actbuyingRateR5		=sbr5throw_4thColumn.getText();
			String actbuyingValueR5		=sbr5throw_5thColumn.getText();
			
			String actsellingRateR6		=sbr6throw_2ndColumn.getText();
			String actsellingValueR6	=sbr6throw_3rdColumn.getText();
			String actbuyingRateR6		=sbr6throw_4thColumn.getText();
			String actbuyingValueR6		=sbr6throw_5thColumn.getText();
			
			String actsellingRateR7		=sbr7throw_2ndColumn.getText();
			String actsellingValueR7	=sbr7throw_3rdColumn.getText();
			String actbuyingRateR7		=sbr7throw_4thColumn.getText();
			String actbuyingValueR7		=sbr7throw_5thColumn.getText();
			
			String actsellingRateR8		=sbr8throw_2ndColumn.getText();
			String actsellingValueR8	=sbr8throw_3rdColumn.getText();
			String actbuyingRateR8		=sbr8throw_4thColumn.getText();
			String actbuyingValueR8		=sbr8throw_5thColumn.getText();
			
			String actsellingRateR9		=sbr9throw_2ndColumn.getText();
			String actsellingValueR9	=sbr9throw_3rdColumn.getText();
			String actbuyingRateR9		=sbr9throw_4thColumn.getText();
			String actbuyingValueR9		=sbr9throw_5thColumn.getText();
			
			String actsellingRateR10	=sbr10throw_2ndColumn.getText();
			String actsellingValueR10	=sbr10throw_3rdColumn.getText();
			String actbuyingRateR10		=sbr10throw_4thColumn.getText();
			String actbuyingValueR10	=sbr10throw_5thColumn.getText();
			
			String actsellingRateR11	=sbr11throw_2ndColumn.getText();
			String actsellingValueR11	=sbr11throw_3rdColumn.getText();
			String actbuyingRateR11		=sbr11throw_4thColumn.getText();
			String actbuyingValueR11	=sbr11throw_5thColumn.getText();
			
			String actsellingRateR12	=sbr12throw_2ndColumn.getText();
			String actsellingValueR12	=sbr12throw_3rdColumn.getText();
			String actbuyingRateR12 	=sbr12throw_4thColumn.getText();
			String actbuyingValueR12	=sbr12throw_5thColumn.getText();
			
			String actsellingRateR13	=sbr13throw_2ndColumn.getText();
			String actsellingValueR13	=sbr13throw_3rdColumn.getText();
			String actbuyingRateR13 	=sbr13throw_4thColumn.getText();
			String actbuyingValueR13    =sbr13throw_5thColumn.getText();
			
			String actsellingRateR14	=sbr14throw_2ndColumn.getText();
			String actsellingValueR14	=sbr14throw_3rdColumn.getText();
			String actbuyingRateR14 	=sbr14throw_4thColumn.getText();
			String actbuyingValueR14    =sbr14throw_5thColumn.getText();
			
			
			// Expected
			
			String expsellingRateR1		="So";
			String expsellingValueR1	="0.00";
			String expbuyingRateR1		="Po";
			String expbuyingValueR1		="0.00";

			String expsellingRateR2		="St";
			String expsellingValueR2	="0.00";
			String expbuyingRateR2		="Pt";
			String expbuyingValueR2		="0.00";
			
			String expsellingRateR3		="Val 3";
			String expsellingValueR3	="0.00";
			String expbuyingRateR3		="Val 3";
			String expbuyingValueR3		="0.00";
			
			String expsellingRateR4		="Val 4";
			String expsellingValueR4	="0.00";
			String expbuyingRateR4		="Val 4";
			String expbuyingValueR4		="0.00";
			
			String expsellingRateR5		="Val 5";
			String expsellingValueR5	="0.00";
			String expbuyingRateR5		="Val 5";
			String expbuyingValueR5		="0.00";
			
			String expsellingRateR6		="Val 6";
			String expsellingValueR6	="0.00";
			String expbuyingRateR6		="Val 6";
			String expbuyingValueR6		="0.00";
			
			String expsellingRateR7		="Val 7";
			String expsellingValueR7	="0.00";
			String expbuyingRateR7		="Val 7";
			String expbuyingValueR7		="0.00";
			
			String expsellingRateR8		="Val 8";
			String expsellingValueR8	="0.00";
			String expbuyingRateR8		="Val 8";
			String expbuyingValueR8		="0.00";
			
			String expsellingRateR9		="Val 9";
			String expsellingValueR9	="0.00";
			String expbuyingRateR9		="Val 9";
			String expbuyingValueR9		="0.00";
			
			String expsellingRateR10	="Val 10";
			String expsellingValueR10	="0.00";
			String expbuyingRateR10		="Val 10";
			String expbuyingValueR10	="0.00";
			
			String expsellingRateR11	="Val 11";
			String expsellingValueR11	="0.00";
			String expbuyingRateR11		="Val 11";
			String expbuyingValueR11	="0.00";
			
			String expsellingRateR12	="Val 12";
			String expsellingValueR12	="0.00";
			String expbuyingRateR12 	="Val 12";
			String expbuyingValueR12	="0.00";
			
			String expsellingRateR13	="Val 13";
			String expsellingValueR13	="0.00";
			String expbuyingRateR13 	="Val 13";
			String expbuyingValueR13    ="0.00";
			
			String expsellingRateR14	="Val 14";
			String expsellingValueR14	="0.00";
			String expbuyingRateR14 	="Val 14";
			String expbuyingValueR14    ="0.00";
			
			System.out.println("************************************checkItemQueryOfBatchWAItem*************************************");
			
			System.out.println("iq_OpeningStocks     		 :   "+actiq_OpeningStocks					+" Value Expected : "+expiq_OpeningStocks);
			System.out.println("iq_PendingPurchaseOrders     :   "+actiq_PendingPurchaseOrders			+" Value Expected : "+expiq_PendingPurchaseOrders);
			System.out.println("iq_CurrentStock     		 :   "+actiq_CurrentStock					+" Value Expected : "+expiq_CurrentStock);
			System.out.println("iq_PendingSalesOrders     	 :   "+actiq_PendingSalesOrders				+" Value Expected : "+expiq_PendingSalesOrders);
			System.out.println("iq_AvgStockRate     		 :   "+actiq_AvgStockRate					+" Value Expected : "+expiq_AvgStockRate);
			System.out.println("iq_QtyToBeOrdered     		 :   "+actiq_QtyToBeOrdered					+" Value Expected : "+expiq_QtyToBeOrdered);
			System.out.println("iq_Value     				 :   "+actiq_Value							+" Value Expected : "+expiq_Value);
			System.out.println("iq_BaseUnit     			 :   "+actiq_BaseUnit						+" Value Expected : "+expiq_BaseUnit);
			
			System.out.println("********************* Purchase ******************");
			
			System.out.println("purVoucherNoR1     			 :   "+actpurVoucherNoR1					+" Value Expected : "+exppurVoucherNoR1);
			System.out.println("purVendorR1     			 :   "+actpurVendorR1						+" Value Expected : "+exppurVendorR1);
			System.out.println("purQtyR1     				 :   "+actpurQtyR1							+" Value Expected : "+exppurQtyR1);
			System.out.println("purUnitR1     				 :   "+actpurUnitR1							+" Value Expected : "+exppurUnitR1);
			System.out.println("purRateR1     				 :   "+actpurRateR1							+" Value Expected : "+exppurRateR1);
			System.out.println("purCurrencyR1     			 :   "+actpurCurrencyR1						+" Value Expected : "+exppurCurrencyR1);
			System.out.println("purQtyInBaseR1     			 :   "+actpurQtyInBaseR1					+" Value Expected : "+exppurQtyInBaseR1);
			
			System.out.println("purVoucherNoR2     			 :   "+actpurVoucherNoR2					+" Value Expected : "+exppurVoucherNoR2);
			System.out.println("purVendorR2     			 :   "+actpurVendorR2						+" Value Expected : "+exppurVendorR2);
			System.out.println("purQtyR2     				 :   "+actpurQtyR2							+" Value Expected : "+exppurQtyR2);
			System.out.println("purUnitR2     				 :   "+actpurUnitR2							+" Value Expected : "+exppurUnitR2);
			System.out.println("purRateR2     				 :   "+actpurRateR2							+" Value Expected : "+exppurRateR2);
			System.out.println("purCurrencyR2     			 :   "+actpurCurrencyR2						+" Value Expected : "+exppurCurrencyR2);
			System.out.println("purQtyInBaseR2     			 :   "+actpurQtyInBaseR2					+" Value Expected : "+exppurQtyInBaseR2);
			
			System.out.println("purVoucherNoR3     			 :   "+actpurVoucherNoR3					+" Value Expected : "+exppurVoucherNoR3);
			System.out.println("purVendorR3     			 :   "+actpurVendorR3						+" Value Expected : "+exppurVendorR3);
			System.out.println("purQtyR3     				 :   "+actpurQtyR3							+" Value Expected : "+exppurQtyR3);
			System.out.println("purUnitR3     				 :   "+actpurUnitR3							+" Value Expected : "+exppurUnitR3);
			System.out.println("purRateR3     				 :   "+actpurRateR3							+" Value Expected : "+exppurRateR3);
			System.out.println("purCurrencyR3     			 :   "+actpurCurrencyR3						+" Value Expected : "+exppurCurrencyR3);
			System.out.println("purQtyInBaseR3     			 :   "+actpurQtyInBaseR3					+" Value Expected : "+exppurQtyInBaseR3);
			
			System.out.println("purTotalWa                   :   "+actpurTotalWa                        +" Value expected : "+exppurTotalWa);
			
			System.out.println("********************* Sales ******************");
			
			System.out.println("salVoucherNoR1     			 :   "+actsalVoucherNoR1					+" Value Expected : "+expsalVoucherNoR1);
			System.out.println("salCustomerR1     			 :   "+actsalCustomerR1						+" Value Expected : "+expsalCustomerR1);
			System.out.println("salQtyR1     				 :   "+actsalQtyR1							+" Value Expected : "+expsalQtyR1);
			System.out.println("salUnitR1     				 :   "+actsalUnitR1							+" Value Expected : "+expsalUnitR1);
			System.out.println("salRateR1     				 :   "+actsalRateR1							+" Value Expected : "+expsalRateR1);
			System.out.println("salCurrencyR1     			 :   "+actsalCurrencyR1						+" Value Expected : "+expsalCurrencyR1);
			System.out.println("salQtyInBaseR1     			 :   "+actsalQtyInBaseR1					+" Value Expected : "+expsalQtyInBaseR1);
			
			System.out.println("salVoucherNoR2     			 :   "+actsalVoucherNoR2					+" Value Expected : "+expsalVoucherNoR2);
			System.out.println("salCustomerR2     			 :   "+actsalCustomerR2						+" Value Expected : "+expsalCustomerR2);
			System.out.println("salQtyR2     				 :   "+actsalQtyR2							+" Value Expected : "+expsalQtyR2);
			System.out.println("salUnitR2     				 :   "+actsalUnitR2							+" Value Expected : "+expsalUnitR2);
			System.out.println("salRateR2     				 :   "+actsalRateR2							+" Value Expected : "+expsalRateR2);
			System.out.println("salCurrencyR2     			 :   "+actsalCurrencyR2						+" Value Expected : "+expsalCurrencyR2);
			System.out.println("salQtyInBaseR2     			 :   "+actsalQtyInBaseR2					+" Value Expected : "+expsalQtyInBaseR2);
			
			System.out.println("salVoucherNoR3     			 :   "+actsalVoucherNoR3					+" Value Expected : "+expsalVoucherNoR3);
			System.out.println("salCustomerR3     			 :   "+actsalCustomerR3						+" Value Expected : "+expsalCustomerR3);
			System.out.println("salQtyR3     				 :   "+actsalQtyR3							+" Value Expected : "+expsalQtyR3);
			System.out.println("salUnitR3     				 :   "+actsalUnitR3							+" Value Expected : "+expsalUnitR3);
			System.out.println("salRateR3     				 :   "+actsalRateR3							+" Value Expected : "+expsalRateR3);
			System.out.println("salCurrencyR3     			 :   "+actsalCurrencyR3						+" Value Expected : "+expsalCurrencyR3);
			System.out.println("salQtyInBaseR3     			 :   "+actsalQtyInBaseR3					+" Value Expected : "+expsalQtyInBaseR3);
			
			System.out.println("salVoucherNoR4     			 :   "+actsalVoucherNoR4					+" Value Expected : "+expsalVoucherNoR4);
			System.out.println("salCustomerR4     			 :   "+actsalCustomerR4						+" Value Expected : "+expsalCustomerR4);
			System.out.println("salQtyR4     				 :   "+actsalQtyR4							+" Value Expected : "+expsalQtyR4);
			System.out.println("salUnitR4     				 :   "+actsalUnitR4							+" Value Expected : "+expsalUnitR4);
			System.out.println("salRateR4     				 :   "+actsalRateR4							+" Value Expected : "+expsalRateR4);
			System.out.println("salCurrencyR4     			 :   "+actsalCurrencyR4						+" Value Expected : "+expsalCurrencyR4);
			System.out.println("salQtyInBaseR4     			 :   "+actsalQtyInBaseR4					+" Value Expected : "+expsalQtyInBaseR4);
			
			System.out.println("********************* WareHouse ******************");
			
			System.out.println("wareHouseR1     	 :   "+actwareHouseR1				+" Value Expected : "+expwareHouseR1);
			System.out.println("stockR1          	 :   "+actstockR1					+" Value Expected : "+expstockR1);
			System.out.println("rateR1	          	 :   "+actrateR1					+" Value Expected : "+exprateR1);
			System.out.println("valueR1	          	 :   "+actvalueR1					+" Value Expected : "+expvalueR1);
			System.out.println("totalStock	         :   "+acttotalStock				+" Value Expected : "+exptotalStock);
			System.out.println("totalValue	         :   "+acttotalValue				+" Value Expected : "+exptotalValue);
			
			System.out.println("*************************Selling and Buying rates********************");
			
			System.out.println("sellingRateR1     			 :   "+actsellingRateR1						+" Value Expected : "+expsellingRateR1);
			System.out.println("sellingValueR1  			 :   "+actsellingValueR1					+" Value Expected : "+expsellingValueR1);
			System.out.println("buyingRateR1     			 :   "+actbuyingRateR1						+" Value Expected : "+expbuyingRateR1);
			System.out.println("buyingValueR1     			 :   "+actbuyingValueR1						+" Value Expected : "+expbuyingValueR1);
			
			System.out.println("sellingRateR2     			 :   "+actsellingRateR2						+" Value Expected : "+expsellingRateR2);
			System.out.println("sellingValueR2  			 :   "+actsellingValueR2					+" Value Expected : "+expsellingValueR2);
			System.out.println("buyingRateR2     			 :   "+actbuyingRateR2						+" Value Expected : "+expbuyingRateR2);
			System.out.println("buyingValueR2     			 :   "+actbuyingValueR2						+" Value Expected : "+expbuyingValueR2);
			
			System.out.println("sellingRateR3     			 :   "+actsellingRateR3						+" Value Expected : "+expsellingRateR3);
			System.out.println("sellingValueR3  			 :   "+actsellingValueR3					+" Value Expected : "+expsellingValueR3);
			System.out.println("buyingRateR3     			 :   "+actbuyingRateR3						+" Value Expected : "+expbuyingRateR3);
			System.out.println("buyingValueR3     			 :   "+actbuyingValueR3						+" Value Expected : "+expbuyingValueR3);
			
			System.out.println("sellingRateR4     			 :   "+actsellingRateR4						+" Value Expected : "+expsellingRateR4);
			System.out.println("sellingValueR4  			 :   "+actsellingValueR4					+" Value Expected : "+expsellingValueR4);
			System.out.println("buyingRateR4     			 :   "+actbuyingRateR4						+" Value Expected : "+expbuyingRateR4);
			System.out.println("buyingValueR4     			 :   "+actbuyingValueR4						+" Value Expected : "+expbuyingValueR4);
			
			System.out.println("sellingRateR5     			 :   "+actsellingRateR5						+" Value Expected : "+expsellingRateR5);
			System.out.println("sellingValueR5  			 :   "+actsellingValueR5					+" Value Expected : "+expsellingValueR5);
			System.out.println("buyingRateR5     			 :   "+actbuyingRateR5						+" Value Expected : "+expbuyingRateR5);
			System.out.println("buyingValueR5     			 :   "+actbuyingValueR5						+" Value Expected : "+expbuyingValueR5);
			
			System.out.println("sellingRateR6     			 :   "+actsellingRateR6						+" Value Expected : "+expsellingRateR6);
			System.out.println("sellingValueR6  			 :   "+actsellingValueR6					+" Value Expected : "+expsellingValueR6);
			System.out.println("buyingRateR6     			 :   "+actbuyingRateR6						+" Value Expected : "+expbuyingRateR6);
			System.out.println("buyingValueR6     			 :   "+actbuyingValueR6						+" Value Expected : "+expbuyingValueR6);
			
			System.out.println("sellingRateR7     			 :   "+actsellingRateR7						+" Value Expected : "+expsellingRateR7);
			System.out.println("sellingValueR7  			 :   "+actsellingValueR7					+" Value Expected : "+expsellingValueR7);
			System.out.println("buyingRateR7     			 :   "+actbuyingRateR7						+" Value Expected : "+expbuyingRateR7);
			System.out.println("buyingValueR7     			 :   "+actbuyingValueR7						+" Value Expected : "+expbuyingValueR7);
			
			System.out.println("sellingRateR8     			 :   "+actsellingRateR8						+" Value Expected : "+expsellingRateR8);
			System.out.println("sellingValueR8  			 :   "+actsellingValueR8					+" Value Expected : "+expsellingValueR8);
			System.out.println("buyingRateR8     			 :   "+actbuyingRateR8						+" Value Expected : "+expbuyingRateR8);
			System.out.println("buyingValueR8     			 :   "+actbuyingValueR8						+" Value Expected : "+expbuyingValueR8);
			
			System.out.println("sellingRateR9     			 :   "+actsellingRateR9						+" Value Expected : "+expsellingRateR9);
			System.out.println("sellingValueR9  			 :   "+actsellingValueR9					+" Value Expected : "+expsellingValueR9);
			System.out.println("buyingRateR9     			 :   "+actbuyingRateR9						+" Value Expected : "+expbuyingRateR9);
			System.out.println("buyingValueR9     			 :   "+actbuyingValueR9						+" Value Expected : "+expbuyingValueR9);
			
			System.out.println("sellingRateR10     			 :   "+actsellingRateR10					+" Value Expected : "+expsellingRateR10);
			System.out.println("sellingValueR10  			 :   "+actsellingValueR10					+" Value Expected : "+expsellingValueR10);
			System.out.println("buyingRateR10     			 :   "+actbuyingRateR10						+" Value Expected : "+expbuyingRateR10);
			System.out.println("buyingValueR10     			 :   "+actbuyingValueR10					+" Value Expected : "+expbuyingValueR10);
			
			System.out.println("sellingRateR11     			 :   "+actsellingRateR11					+" Value Expected : "+expsellingRateR11);
			System.out.println("sellingValueR11  			 :   "+actsellingValueR11					+" Value Expected : "+expsellingValueR11);
			System.out.println("buyingRateR11     			 :   "+actbuyingRateR11						+" Value Expected : "+expbuyingRateR11);
			System.out.println("buyingValueR11     			 :   "+actbuyingValueR11					+" Value Expected : "+expbuyingValueR11);
			
			System.out.println("sellingRateR12     			 :   "+actsellingRateR12					+" Value Expected : "+expsellingRateR12);
			System.out.println("sellingValueR12  			 :   "+actsellingValueR12					+" Value Expected : "+expsellingValueR12);
			System.out.println("buyingRateR12     			 :   "+actbuyingRateR12						+" Value Expected : "+expbuyingRateR12);
			System.out.println("buyingValueR12     			 :   "+actbuyingValueR12					+" Value Expected : "+expbuyingValueR12);
			
			System.out.println("sellingRateR13     			 :   "+actsellingRateR13					+" Value Expected : "+expsellingRateR13);
			System.out.println("sellingValueR13  			 :   "+actsellingValueR13					+" Value Expected : "+expsellingValueR13);
			System.out.println("buyingRateR13     			 :   "+actbuyingRateR13						+" Value Expected : "+expbuyingRateR13);
			System.out.println("buyingValueR13     			 :   "+actbuyingValueR13					+" Value Expected : "+expbuyingValueR13);

			System.out.println("sellingRateR14     			 :   "+actsellingRateR14					+" Value Expected : "+expsellingRateR14);
			System.out.println("sellingValueR14  			 :   "+actsellingValueR14					+" Value Expected : "+expsellingValueR14);
			System.out.println("buyingRateR14     			 :   "+actbuyingRateR14						+" Value Expected : "+expbuyingRateR14);
			System.out.println("buyingValueR14     			 :   "+actbuyingValueR14					+" Value Expected : "+expbuyingValueR14);
			
			if(actiq_OpeningStocks.equalsIgnoreCase(expiq_OpeningStocks) && actiq_PendingPurchaseOrders.equalsIgnoreCase(expiq_PendingPurchaseOrders)
					&& actiq_CurrentStock.equalsIgnoreCase(expiq_CurrentStock) && actiq_PendingSalesOrders.equalsIgnoreCase(expiq_PendingSalesOrders)
					&& actiq_AvgStockRate.equalsIgnoreCase(expiq_AvgStockRate) && actiq_QtyToBeOrdered.equalsIgnoreCase(expiq_QtyToBeOrdered)
					&& actiq_Value.equalsIgnoreCase(expiq_Value) && actiq_BaseUnit.equalsIgnoreCase(expiq_BaseUnit)
					
					&& actpurVoucherNoR1.equalsIgnoreCase(exppurVoucherNoR1) && actpurVendorR1.equalsIgnoreCase(exppurVendorR1) && actpurQtyR1.equalsIgnoreCase(exppurQtyR1)
					&& actpurUnitR1.equalsIgnoreCase(exppurUnitR1) && actpurRateR1.equalsIgnoreCase(exppurRateR1) && actpurCurrencyR1.equalsIgnoreCase(exppurCurrencyR1)
					&& actpurQtyInBaseR1.equalsIgnoreCase(exppurQtyInBaseR1)
					
					&& actpurVoucherNoR2.equalsIgnoreCase(exppurVoucherNoR2) && actpurVendorR2.equalsIgnoreCase(exppurVendorR2) && actpurQtyR2.equalsIgnoreCase(exppurQtyR2)
					&& actpurUnitR2.equalsIgnoreCase(exppurUnitR2) && actpurRateR2.equalsIgnoreCase(exppurRateR2) && actpurCurrencyR2.equalsIgnoreCase(exppurCurrencyR2)
					&& actpurQtyInBaseR2.equalsIgnoreCase(exppurQtyInBaseR2)
					
					&& actpurVoucherNoR3.equalsIgnoreCase(exppurVoucherNoR3) && actpurVendorR3.equalsIgnoreCase(exppurVendorR3) && actpurQtyR3.equalsIgnoreCase(exppurQtyR3)
					&& actpurUnitR3.equalsIgnoreCase(exppurUnitR3) && actpurRateR3.equalsIgnoreCase(exppurRateR3) && actpurCurrencyR3.equalsIgnoreCase(exppurCurrencyR3)
					&& actpurQtyInBaseR3.equalsIgnoreCase(exppurQtyInBaseR3)
					
					&& actpurTotalWa.equalsIgnoreCase(exppurTotalWa)
					
					
					&& actsalVoucherNoR1.equalsIgnoreCase(expsalVoucherNoR1) && actsalCustomerR1.equalsIgnoreCase(expsalCustomerR1) && actsalQtyR1.equalsIgnoreCase(expsalQtyR1)
					&& actsalUnitR1.equalsIgnoreCase(expsalUnitR1) && actsalRateR1.equalsIgnoreCase(expsalRateR1) && actsalCurrencyR1.equalsIgnoreCase(expsalCurrencyR1)
					&& actsalQtyInBaseR1.equalsIgnoreCase(expsalQtyInBaseR1)
					
					&& actsalVoucherNoR2.equalsIgnoreCase(expsalVoucherNoR2) && actsalCustomerR2.equalsIgnoreCase(expsalCustomerR2) && actsalQtyR2.equalsIgnoreCase(expsalQtyR2)
					&& actsalUnitR2.equalsIgnoreCase(expsalUnitR2) && actsalRateR2.equalsIgnoreCase(expsalRateR2) && actsalCurrencyR2.equalsIgnoreCase(expsalCurrencyR2)
					&& actsalQtyInBaseR2.equalsIgnoreCase(expsalQtyInBaseR2)
					
					&& actsalVoucherNoR3.equalsIgnoreCase(expsalVoucherNoR3) && actsalCustomerR3.equalsIgnoreCase(expsalCustomerR3) && actsalQtyR3.equalsIgnoreCase(expsalQtyR3)
					&& actsalUnitR3.equalsIgnoreCase(expsalUnitR3) && actsalRateR3.equalsIgnoreCase(expsalRateR3) && actsalCurrencyR3.equalsIgnoreCase(expsalCurrencyR3)
					&& actsalQtyInBaseR3.equalsIgnoreCase(expsalQtyInBaseR3)
					
					&& actsalVoucherNoR4.equalsIgnoreCase(expsalVoucherNoR4) && actsalCustomerR4.equalsIgnoreCase(expsalCustomerR4) && actsalQtyR4.equalsIgnoreCase(expsalQtyR4)
					&& actsalUnitR4.equalsIgnoreCase(expsalUnitR4) && actsalRateR4.equalsIgnoreCase(expsalRateR4) && actsalCurrencyR4.equalsIgnoreCase(expsalCurrencyR4)
					&& actsalQtyInBaseR4.equalsIgnoreCase(expsalQtyInBaseR4)
					
					&& actsalTotalWa.equalsIgnoreCase(expsalTotalWa)
					
					&& actwareHouseR1.equalsIgnoreCase(expwareHouseR1) && actstockR1.equalsIgnoreCase(expstockR1) && actrateR1.equalsIgnoreCase(exprateR1)
					&& actvalueR1.equalsIgnoreCase(expvalueR1) && acttotalStock.equalsIgnoreCase(exptotalStock) && acttotalValue.equalsIgnoreCase(exptotalValue)
					
					&& actsellingRateR1.equalsIgnoreCase(expsellingRateR1) && actsellingValueR1.equalsIgnoreCase(expsellingValueR1)
					&& actbuyingRateR1.equalsIgnoreCase(expbuyingRateR1) && actbuyingValueR1.equalsIgnoreCase(expbuyingValueR1)
					
					&& actsellingRateR2.equalsIgnoreCase(expsellingRateR2) && actsellingValueR2.equalsIgnoreCase(expsellingValueR2)
					&& actbuyingRateR2.equalsIgnoreCase(expbuyingRateR2) && actbuyingValueR2.equalsIgnoreCase(expbuyingValueR2)
					
					&& actsellingRateR3.equalsIgnoreCase(expsellingRateR3) && actsellingValueR3.equalsIgnoreCase(expsellingValueR3)
					&& actbuyingRateR3.equalsIgnoreCase(expbuyingRateR3) && actbuyingValueR3.equalsIgnoreCase(expbuyingValueR3)
					
					&& actsellingRateR4.equalsIgnoreCase(expsellingRateR4) && actsellingValueR4.equalsIgnoreCase(expsellingValueR4)
					&& actbuyingRateR4.equalsIgnoreCase(expbuyingRateR4) && actbuyingValueR4.equalsIgnoreCase(expbuyingValueR4)
					
					&& actsellingRateR5.equalsIgnoreCase(expsellingRateR5) && actsellingValueR4.equalsIgnoreCase(expsellingValueR5)
					&& actbuyingRateR5.equalsIgnoreCase(expbuyingRateR5) && actbuyingValueR5.equalsIgnoreCase(expbuyingValueR5)
					
					&& actsellingRateR6.equalsIgnoreCase(expsellingRateR6) && actsellingValueR6.equalsIgnoreCase(expsellingValueR6)
					&& actbuyingRateR6.equalsIgnoreCase(expbuyingRateR6) && actbuyingValueR6.equalsIgnoreCase(expbuyingValueR6)
					
					&& actsellingRateR7.equalsIgnoreCase(expsellingRateR7) && actsellingValueR7.equalsIgnoreCase(expsellingValueR7)
					&& actbuyingRateR7.equalsIgnoreCase(expbuyingRateR7) && actbuyingValueR7.equalsIgnoreCase(expbuyingValueR7)
					
					&& actsellingRateR8.equalsIgnoreCase(expsellingRateR8) && actsellingValueR8.equalsIgnoreCase(expsellingValueR8)
					&& actbuyingRateR8.equalsIgnoreCase(expbuyingRateR8) && actbuyingValueR8.equalsIgnoreCase(expbuyingValueR8)
					
					&& actsellingRateR9.equalsIgnoreCase(expsellingRateR9) && actsellingValueR9.equalsIgnoreCase(expsellingValueR9)
					&& actbuyingRateR9.equalsIgnoreCase(expbuyingRateR9) && actbuyingValueR9.equalsIgnoreCase(expbuyingValueR9)
					
					&& actsellingRateR10.equalsIgnoreCase(expsellingRateR10) && actsellingValueR10.equalsIgnoreCase(expsellingValueR10)
					&& actbuyingRateR10.equalsIgnoreCase(expbuyingRateR10) && actbuyingValueR10.equalsIgnoreCase(expbuyingValueR10)
					
					&& actsellingRateR11.equalsIgnoreCase(expsellingRateR11) && actsellingValueR11.equalsIgnoreCase(expsellingValueR11)
					&& actbuyingRateR11.equalsIgnoreCase(expbuyingRateR11) && actbuyingValueR11.equalsIgnoreCase(expbuyingValueR11)
					
					&& actsellingRateR12.equalsIgnoreCase(expsellingRateR12) && actsellingValueR12.equalsIgnoreCase(expsellingValueR12)
					&& actbuyingRateR12.equalsIgnoreCase(expbuyingRateR12) && actbuyingValueR12.equalsIgnoreCase(expbuyingValueR12)
					
					&& actsellingRateR13.equalsIgnoreCase(expsellingRateR13) && actsellingValueR13.equalsIgnoreCase(expsellingValueR13)
					&& actbuyingRateR13.equalsIgnoreCase(expbuyingRateR13) && actbuyingValueR13.equalsIgnoreCase(expbuyingValueR13)
					
					&& actsellingRateR14.equalsIgnoreCase(expsellingRateR14) && actsellingValueR14.equalsIgnoreCase(expsellingValueR14)
					&& actbuyingRateR14.equalsIgnoreCase(expbuyingRateR14) && actbuyingValueR14.equalsIgnoreCase(expbuyingValueR14))
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(iq_closeBtn));
				iq_closeBtn.click();
				
				Thread.sleep(2000);

				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(iq_closeBtn));
				iq_closeBtn.click();
				
				Thread.sleep(2000);

				return false;
			}
		}
		
		
				
		
		
		
		
		
		
		// Stock Ledger

		public boolean checkRMAStockLedgerReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedger));
			stockLedger.click();
			
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			Thread.sleep(2000);
			
			int rowcount=stockLedgerHometableRowCount.size();
			
			System.out.println(rowcount);
			
			for (int i = 1; i <= rowcount; i++) 
			{
				WebElement name=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[12]"));
				
				String actname=name.getText();
				
				System.out.println(actname);
				
				if(actname.equalsIgnoreCase("WA COGS ITEM"))
				{
					WebElement index=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[8]/div/label/input"));
					index.click();
					Thread.sleep(2000);
					break;
				}
			}
			
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			int reportsRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
			String actRow1List = reportsRow1ListArray.toString();
			String expRow1List = "[WA COGS ITEM WA COGS ITEM, , , , , , , , , , , , , ]";
			
			

			int reportsRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsRow2ListArray.add(data);
			}
			String actRow2List = reportsRow2ListArray.toString();
			String expRow2List = "[Opening Balance, 36.00, , , , 36.00, , , 222.00, 6.17, , , ]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[NDT55:1, , , 3.00, 6.06, 33.00, , 18.19, 203.50, 6.17, , , ]";
			
			
			
			int reportsRow4ListCount = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow4ListCount;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
			String actRow4List = reportsRow4ListArray.toString();
			String expRow4List = "[NDT55:2, , , 9.00, 6.17, 24.00, , 55.50, 148.00, 6.17, , , ]";
			
			
			
			int reportsRow5ListCount = reportsRow5List.size();
			ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow5ListCount;i++)
			{
				String data = reportsRow5List.get(i).getText();
				reportsRow5ListArray.add(data);
			}
			String actRow5List = reportsRow5ListArray.toString();
			String expRow5List = "[NDT52:1, 100.00, 10.50, , , 124.00, 1,050.00, , 1,198.00, 9.66, , , ]";
			
			
			
			int reportsRow6ListCount = reportsRow6List.size();
			ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow6ListCount;i++)
			{
				String data = reportsRow6List.get(i).getText();
				reportsRow6ListArray.add(data);
			}
			String actRow6List = reportsRow6ListArray.toString();
			String expRow6List = "[ExeStk:1, 3.00, 10.00, , , 127.00, 30.00, , 1,228.00, 9.67, , , ]";
			
			
			
			int reportsRow7ListCount = reportsRow7List.size();
			ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow7ListCount;i++)
			{
				String data = reportsRow7List.get(i).getText();
				reportsRow7ListArray.add(data);
			}
			String actRow7List = reportsRow7ListArray.toString();
			String expRow7List = "[NDT55:4, , , 2.00, 9.67, 125.00, , 19.34, 1,208.66, 9.67, , , ]";
			
			
			
			int reportsRow8ListCount = reportsRow8List.size();
			ArrayList<String> reportsRow8ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow8ListCount;i++)
			{
				String data = reportsRow8List.get(i).getText();
				reportsRow8ListArray.add(data);
			}
			String actRow8List = reportsRow8ListArray.toString();
			String expRow8List = "[StkTrf:2, , , 2.00, 9.67, 123.00, , 19.34, 1,189.32, 9.67, , , ]";
			
			
			
			int reportsRow9ListCount = reportsRow9List.size();
			ArrayList<String> reportsRow9ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow9ListCount;i++)
			{
				String data = reportsRow9List.get(i).getText();
				reportsRow9ListArray.add(data);
			}
			String actRow9List = reportsRow9ListArray.toString();
			String expRow9List = "[StkTrf:2, 2.00, 9.67, , , 125.00, 19.34, , 1,208.66, 9.67, , , ]";
			
			
			
			int reportsRow10ListCount = reportsRow10List.size();
			ArrayList<String> reportsRow10ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow10ListCount;i++)
			{
				String data = reportsRow10List.get(i).getText();
				reportsRow10ListArray.add(data);
			}
			String actRow10List = reportsRow10ListArray.toString();
			String expRow10List = "[NDT52:2, 50.00, 26.25, , , 175.00, 1,312.50, , 2,521.16, 14.41, , , ]";
			
			
			
			int reportsRow11ListCount = reportsRow11List.size();
			ArrayList<String> reportsRow11ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow11ListCount;i++)
			{
				String data = reportsRow11List.get(i).getText();
				reportsRow11ListArray.add(data);
			}
			String actRow11List = reportsRow11ListArray.toString();
			String expRow11List = "[NDT55:5, , , 15.00, 14.41, 160.00, , 216.10, 2,305.06, 14.41, , , ]";
			
			
			
			int reportsRow12ListCount = reportsRow12List.size();
			ArrayList<String> reportsRow12ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow12ListCount;i++)
			{
				String data = reportsRow12List.get(i).getText();
				reportsRow12ListArray.add(data);
			}
			String actRow12List = reportsRow12ListArray.toString();
			String expRow12List = "[SalRet:1, 15.00, 14.41, , , 175.00, 216.15, , 2,521.21, 14.41, , , ]";
			
			
			
			int reportsRow13ListCount = reportsRow13List.size();
			ArrayList<String> reportsRow13ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow13ListCount;i++)
			{
				String data = reportsRow13List.get(i).getText();
				reportsRow13ListArray.add(data);
			}
			String actRow13List = reportsRow13ListArray.toString();
			String expRow13List = "[NDT52:3, 377.00, 10.50, , , 552.00, 3,958.50, , 6,479.71, 11.74, , , ]";
			
			
			
			int reportsRow14ListCount = reportsRow14List.size();
			ArrayList<String> reportsRow14ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow14ListCount;i++)
			{
				String data = reportsRow14List.get(i).getText();
				reportsRow14ListArray.add(data);
			}
			String actRow14List = reportsRow14ListArray.toString();
			String expRow14List = "[PurRet:1, , , 525.00, 11.74, 27.00, , 6,163.50, 316.94, 11.74, , , ]";
			
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_NextBtn));
			report_NextBtn.click();
			
			Thread.sleep(2000);
			
			int reportsP2Row1ListCount = reportsRow1List.size();
			ArrayList<String> reportsP2Row1ListArray = new ArrayList<String>();
			for(int i=2;i<reportsP2Row1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsP2Row1ListArray.add(data);
			}
			String actP2Row1List = reportsP2Row1ListArray.toString();
			String expP2Row1List = "[, 583.00, 81.33, 556.00, 57.71, 27.00, 6,586.49, 6,491.96, 316.94, 133.54, , , ]";
			
			System.out.println("************************************checkWACOGSITEMStockLedgerReport********************************************");
		
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			System.out.println(actRow4List);
			System.out.println(expRow4List);
			
			System.out.println(actRow5List);
			System.out.println(expRow5List);
			
			System.out.println(actRow6List);
			System.out.println(expRow6List);
			
			System.out.println(actRow7List);
			System.out.println(expRow7List);
			
			System.out.println(actRow8List);
			System.out.println(expRow8List);
			
			System.out.println(actRow9List);
			System.out.println(expRow9List);
			
			System.out.println(actRow10List);
			System.out.println(expRow10List);
			
			System.out.println(actRow11List);
			System.out.println(expRow11List);
			
			System.out.println(actRow12List);
			System.out.println(expRow12List);
			
			System.out.println(actRow13List);
			System.out.println(expRow13List);
			
			System.out.println(actRow14List);
			System.out.println(expRow14List);
			
			System.out.println(actP2Row1List);
			System.out.println(expP2Row1List);
			
			Thread.sleep(2000);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List) 
				&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List) 
				&& actRow10List.equalsIgnoreCase(expRow10List) && actRow11List.equalsIgnoreCase(expRow11List) && actRow12List.equalsIgnoreCase(expRow12List) 
				&& actRow13List.equalsIgnoreCase(expRow13List) && actRow14List.equalsIgnoreCase(expRow14List) && actP2Row1List.equalsIgnoreCase(expP2Row1List))
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
			
		
		
		
		
		
		
		
		
		public boolean checkOpeningStockRegisterReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksRegister));
			openingStocksRegister.click();
			
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
			
			int reportsRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
			String actRow1List = reportsRow1ListArray.toString();
			String expRow1List = "[WA COGS ITEM, 22.00, 8.50, 187.00]";
			
			

			int reportsRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsRow2ListArray.add(data);
			}
			String actRow2List = reportsRow2ListArray.toString();
			String expRow2List = "[WA COGS ITEM, 12.00, 2.00, 24.00]";
		

			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[WA COGS ITEM, 2.00, 5.50, 11.00]";
			
			
			
			int reportsRow4ListCount = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow4ListCount;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
			String actRow4List = reportsRow4ListArray.toString();
			String expRow4List = "[Grand Total, , 36.00, 16.00, 222.00]";
			
			
			System.out.println("***************************checkOpenigStockRegisterReport*********************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			System.out.println(actRow4List);
			System.out.println(expRow4List);
			
			Thread.sleep(2000);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
					&& actRow4List.equalsIgnoreCase(expRow4List))
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				
				return false;
			}
		}

		
		
		
		
		
		
		
		
		
		
		public boolean checkStockStatementReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockStatement));
			stockStatement.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			int rowcount=stockLedgerHometableRowCount.size();
			
			System.out.println(rowcount);
			
			for (int i = 1; i <= rowcount; i++) 
			{
				WebElement name=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[12]"));
				
				String actname=name.getText();
				
				System.out.println(actname);
				
				if(actname.equalsIgnoreCase("WA COGS ITEM") )
				{
					WebElement index=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[8]/div/label/input"));
					index.click();
					
					break;
				}

			}
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));

			int reportsRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
			String actRow1List = reportsRow1ListArray.toString();
			String expRow1List = "[WA COGS ITEM WA COGS ITEM, , , , , ]";
			

			
			int reportsRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsRow2ListArray.add(data);
			}
			String actRow2List = reportsRow2ListArray.toString();
			String expRow2List = "[36.00, 547.00, 556.00, 27.00, 316.94]";
		

			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[36.00, 547.00, 556.00, 27.00, 316.94]";
			
			
			System.out.println("***************************checkStockStatementReport*********************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			Thread.sleep(2000);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List))
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}

		
		
		
		
		
		
		public boolean checkStockMovementReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockMovement));
			stockMovement.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			   
			int rowcount=stockLedgerHometableRowCount.size();
			
			System.out.println(rowcount);
			
			for (int i = 1; i <= rowcount; i++) 
			{
				WebElement name=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[12]"));
				
				String actname=name.getText();
				
				System.out.println(actname);
				
				if(actname.equalsIgnoreCase("WA COGS ITEM") )
				{
					WebElement index=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[8]/div/label/input"));
					index.click();
					
					break;
				}
			}
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));

			int reportsRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
			String actDefaultRow1List = reportsRow1ListArray.toString();
			String expDefaultRow1List = "[WA COGS ITEM, 36.00, 222.00, 547.00, 6,586.49, 556.00, 6,491.55, 27.00, 316.94, 11.74]";
			
			

			int reportsRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsRow2ListArray.add(data);
			}
			String actDefaultRow2List = reportsRow2ListArray.toString();
			String expDefaultRow2List = "[36.00, 222.00, 547.00, 6,586.49, 556.00, 6,491.55, 27.00, 316.94, 11.74]";
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			
	        Thread.sleep(2000);
	        
	        getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockMovementDropdown));
			
			Select sm=new Select(stockMovementDropdown);
			sm.selectByVisibleText("by Warehouse");	
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
			
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsByWarehouseRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[HYDERABAD, , , , , , , , , ]";
			
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsByWarehouseRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[WA COGS ITEM, 24.00, 198.00, 530.00, 6,351.00, 554.00, 6,472.63, , , ]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[Sub Total, 24.00, 198.00, 530.00, 6,351.00, 554.00, 6,472.63, , , ]";
			
			
			
			int reportsRow4ListCount = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow4ListCount;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
			String actRow4List = reportsRow4ListArray.toString();
			String expRow4List = "[SECUNDERABAD, , , , , , , , , ]";
			
			
			
			int reportsRow5ListCount = reportsRow5List.size();
			ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow5ListCount;i++)
			{
				String data = reportsRow5List.get(i).getText();
				reportsRow5ListArray.add(data);
			}
			String actRow5List = reportsRow5ListArray.toString();
			String expRow5List = "[WA COGS ITEM, 12.00, 24.00, 17.00, 235.49, 2.00, 4.00, 27.00, 255.49, 9.46]";
			
			
			
			int reportsRow6ListCount = reportsRow6List.size();
			ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow6ListCount;i++)
			{
				String data = reportsRow6List.get(i).getText();
				reportsRow6ListArray.add(data);
			}
			String actRow6List = reportsRow6ListArray.toString();
			String expRow6List = "[Sub Total, 12.00, 24.00, 17.00, 235.49, 2.00, 4.00, 27.00, 255.49, 9.46]";
			
			
			
			int reportsRow7ListCount = reportsRow7List.size();
			ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow7ListCount;i++)
			{
				String data = reportsRow7List.get(i).getText();
				reportsRow7ListArray.add(data);
			}
			String actRow7List = reportsRow7ListArray.toString();
			String expRow7List = "[Grand Total, 36.00, 222.00, 547.00, 6,586.49, 556.00, 6,468.63, 27.00, 255.49, 9.46]";
			
			
			System.out.println("*********************************checkStockMovementReport*****************************************");
			System.out.println(actDefaultRow1List);
			System.out.println(expDefaultRow1List);
			System.out.println(actDefaultRow2List);
			System.out.println(expDefaultRow2List);
			
			System.out.println("********************************checkStockMovementReportOnSelectingStockMoventAsByWareHouse***********************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			System.out.println(actRow4List);
			System.out.println(expRow4List);
			
			System.out.println(actRow5List);
			System.out.println(expRow5List);
			
			System.out.println(actRow6List);
			System.out.println(expRow6List);
			
			System.out.println(actRow7List);
			System.out.println(expRow7List);
			
			Thread.sleep(2000);	
			
			if(actDefaultRow1List.equalsIgnoreCase(expDefaultRow1List) && actDefaultRow2List.equalsIgnoreCase(expDefaultRow2List)
				&& actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List) 
				&& actRow7List.equalsIgnoreCase(expRow7List))
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
		
		
		
		
		
		
		
		// Multi Level Stock Movement
		
		public boolean checkMultiLevelStockMovementReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(multiLevelStockMovement));
			multiLevelStockMovement.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			   
			int rowcount=stockLedgerHometableRowCount.size();
			
			System.out.println(rowcount);
			
			for (int i = 1; i <= rowcount; i++) 
			{
				WebElement name=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[12]"));

				String actname=name.getText();
				
				if(actname.equalsIgnoreCase("WA COGS ITEM") )
				{
					WebElement index=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[8]/div/label/input"));
					index.click();
					break;
				}
			}
			
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
			
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsByWarehouseRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[WA COGS ITEM, 36.00, 222.00, 547.00, 6,586.49, 556.00, 6,491.96, 27.00, 316.94, 11.74]";
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsByWarehouseRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[Grand Total, 36.00, 222.00, 547.00, 6,586.49, 556.00, 6,491.96, 27.00, 316.94, 11.74]";
			
			System.out.println("*********************************checkMultiLevelStockMovementReport*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			Thread.sleep(2000);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List))
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		// Virtual Stock Analysis
		
		public boolean checkVirtualStockAnalysisReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(virtualStockAnalysis));
			virtualStockAnalysis.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			   
			int rowcount=stockLedgerHometableRowCount.size();
			
			System.out.println(rowcount);
			
			for (int i = 1; i <= rowcount; i++) 
			{
				WebElement name=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[12]"));

				String actname=name.getText();
				
				if(actname.equalsIgnoreCase("WA COGS ITEM"))
				{
					WebElement index=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[8]/div/label/input"));
					index.click();
					
					break;
				}
			}
			
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
			
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsByWarehouseRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[WA COGS ITEM, 27.00, , , , 27.00]";
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsByWarehouseRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[Grand Total, 27.00, , , , 27.00]";
			
			System.out.println("*********************************checkVirtualStockAnalysisReport*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			Thread.sleep(2000);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List))
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
		
		
		
		
		
		
		
		// Stock Valuation 
		
		public boolean checkStockValuationReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockValuation));
			stockValuation.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			int rowcount=stockLedgerHometableRowCount.size();
			
			System.out.println(rowcount);
			
			for (int i = 1; i <= rowcount; i++) 
			{
				WebElement name=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[12]"));
				String actname=name.getText();

				if(actname.equalsIgnoreCase("WA COGS ITEM"))
				{
					WebElement index=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[8]/div/label/input"));
					index.click();
					
					break;
				}
			}
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
			
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsByWarehouseRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[WA COGS ITEM, WA COGS ITEM, WA COGS ITEM, 27.00, 316.94, 11.74, ]";
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsByWarehouseRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[Total, , , 27.00, 316.94, 11.74, ]";
			
			System.out.println("*********************************checkStockValuationReport*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			Thread.sleep(2000);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List))
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);				
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);		
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		
		// Stock report By Tag
		
		public boolean checkStockReportbyTagReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockReportByTag));
			stockReportByTag.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			int rowcount=stockLedgerHometableRowCount.size();
			
			for (int i = 1; i <= rowcount; i++) 
			{
				WebElement name=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[12]"));
				
				String actname=name.getText();
				
				if(actname.equalsIgnoreCase("WA COGS ITEM") )
				{
					WebElement index=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[8]/div/label/input"));
					index.click();
					
					break;
				}
			}
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsByWarehouseRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[WA COGS ITEM, WA COGS ITEM, 27.00, 11.74, 316.94, 27.00, 9.46, 255.49, , , , , , , , , , , , , , , ]";
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsByWarehouseRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[Grand Total, , 27.00, 11.74, 316.94, 27.00, 9.46, 255.49, , , , , , , , , , , , , , , ]";
			
			System.out.println("*********************************checkStockReportByTagReport*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			Thread.sleep(2000);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List))
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		

		
		
		
		
		// ABC Analysis
		
		public boolean checkABCAnalysisReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(abcAnalysis));
			abcAnalysis.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";

			int rowcount=stockLedgerHometableRowCount.size();
			
			System.out.println(rowcount);
			
			for (int i = 1; i <= rowcount; i++) 
			{
				WebElement name=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[12]"));
				String actname=name.getText();
				
				if(actname.equalsIgnoreCase("WA COGS ITEM") )
				{
					WebElement index=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[8]/div/label/input"));
					index.click();
					
					break;
				}
			}
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
			
			int reportsByWarehouseRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsByWarehouseRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsByWarehouseRow1ListArray.add(data);
			}
			String actRow1List = reportsByWarehouseRow1ListArray.toString();
			String expRow1List = "[A Product, , ]";
			

			int reportsByWarehouseRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsByWarehouseRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsByWarehouseRow2ListArray.add(data);
			}
			String actRow2List = reportsByWarehouseRow2ListArray.toString();
			String expRow2List = "[WA COGS ITEM, 27.00, 316.94]";
			
			
			int reportsByWarehouseRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsByWarehouseRow3ListArray = new ArrayList<String>();
			for(int i=1;i<reportsByWarehouseRow3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsByWarehouseRow3ListArray.add(data);
			}
			String actRow3List = reportsByWarehouseRow3ListArray.toString();
			String expRow3List = "[Total, 27.00, 316.94]";
			
			System.out.println("*********************************checkAbcAnalysisReport*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			Thread.sleep(2000);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List))
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
			
		}
		
		
		
		
		
		
		
		
		// Stock Analysis By Batch 

		public boolean checkStockAnalysisByBatchReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAnalysisByBatch));
			stockAnalysisByBatch.click();
			
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";

			int rowcount=stockLedgerHometableRowCount.size();
			
			System.out.println(rowcount);
			
			for (int i = 1; i <= rowcount; i++) 
			{
				WebElement name=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[12]"));
				
				String actname=name.getText();
				
				if(actname.equalsIgnoreCase("WA COGS ITEM") )
				{
					WebElement index=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[8]/div/label/input"));
					index.click();
					
					break;
				}

			}
			
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			boolean actReportTable = reportsTable.getText().isEmpty();
			boolean expReportTable = true;
			
			System.out.println("*********************************checkStockAnalysisByBatchReport*****************************************");
			
			System.out.println("Report Table isEmpty   : "+actReportTable    +" Value Expected : "+expReportTable);
			
			if(actReportTable==expReportTable)
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		
		// Stock Analysis By RMA
		
		public boolean checkStockAnalysisByRmaReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAnalysisByRMA));
			stockAnalysisByRMA.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			int rowcount=stockLedgerHometableRowCount.size();
			
			System.out.println(rowcount);
			
			for (int i = 1; i <= rowcount; i++) 
			{
				WebElement name=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[12]"));
				
				String actname=name.getText();
				
				System.out.println(actname);
				
				if(actname.equalsIgnoreCase("WA COGS ITEM"))
				{
					WebElement index=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[8]/div/label/input"));
					index.click();
					Thread.sleep(2000);
					break;
				}
			}
			
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			int reportsRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
			String actRow1List = reportsRow1ListArray.toString();
			String expRow1List = "[s, 1.00, 14.41, 14.41]";
			
			

			int reportsRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsRow2ListArray.add(data);
			}
			String actRow2List = reportsRow2ListArray.toString();
			String expRow2List = "[s1, 1.00, 14.41, 14.41]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[s2, 1.00, 14.41, 14.41]";
			
			
			
			int reportsRow4ListCount = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow4ListCount;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
			String actRow4List = reportsRow4ListArray.toString();
			String expRow4List = "[s3, 1.00, 14.41, 14.41]";
			
			
			
			int reportsRow5ListCount = reportsRow5List.size();
			ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow5ListCount;i++)
			{
				String data = reportsRow5List.get(i).getText();
				reportsRow5ListArray.add(data);
			}
			String actRow5List = reportsRow5ListArray.toString();
			String expRow5List = "[s4, 1.00, 14.41, 14.41]";
			
			
			
			int reportsRow6ListCount = reportsRow6List.size();
			ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow6ListCount;i++)
			{
				String data = reportsRow6List.get(i).getText();
				reportsRow6ListArray.add(data);
			}
			String actRow6List = reportsRow6ListArray.toString();
			String expRow6List = "[s5, 1.00, 14.41, 14.41]";
			
			
			
			int reportsRow7ListCount = reportsRow7List.size();
			ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow7ListCount;i++)
			{
				String data = reportsRow7List.get(i).getText();
				reportsRow7ListArray.add(data);
			}
			String actRow7List = reportsRow7ListArray.toString();
			String expRow7List = "[s6, 1.00, 14.41, 14.41]";
			
			
			
			int reportsRow8ListCount = reportsRow8List.size();
			ArrayList<String> reportsRow8ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow8ListCount;i++)
			{
				String data = reportsRow8List.get(i).getText();
				reportsRow8ListArray.add(data);
			}
			String actRow8List = reportsRow8ListArray.toString();
			String expRow8List = "[s7, 1.00, 14.41, 14.41]";
			
			
			
			int reportsRow9ListCount = reportsRow9List.size();
			ArrayList<String> reportsRow9ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow9ListCount;i++)
			{
				String data = reportsRow9List.get(i).getText();
				reportsRow9ListArray.add(data);
			}
			String actRow9List = reportsRow9ListArray.toString();
			String expRow9List = "[s8, 1.00, 14.41, 14.41]";
			
			
			
			int reportsRow10ListCount = reportsRow10List.size();
			ArrayList<String> reportsRow10ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow10ListCount;i++)
			{
				String data = reportsRow10List.get(i).getText();
				reportsRow10ListArray.add(data);
			}
			String actRow10List = reportsRow10ListArray.toString();
			String expRow10List = "[s9, 1.00, 14.41, 14.41]";
			
			
			
			int reportsRow11ListCount = reportsRow11List.size();
			ArrayList<String> reportsRow11ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow11ListCount;i++)
			{
				String data = reportsRow11List.get(i).getText();
				reportsRow11ListArray.add(data);
			}
			String actRow11List = reportsRow11ListArray.toString();
			String expRow11List = "[s10, 1.00, 14.41, 14.41]";
			
			
			
			int reportsRow12ListCount = reportsRow12List.size();
			ArrayList<String> reportsRow12ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow12ListCount;i++)
			{
				String data = reportsRow12List.get(i).getText();
				reportsRow12ListArray.add(data);
			}
			String actRow12List = reportsRow12ListArray.toString();
			String expRow12List = "[s11, 1.00, 14.41, 14.41]";
			
			
			
			int reportsRow13ListCount = reportsRow13List.size();
			ArrayList<String> reportsRow13ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow13ListCount;i++)
			{
				String data = reportsRow13List.get(i).getText();
				reportsRow13ListArray.add(data);
			}
			String actRow13List = reportsRow13ListArray.toString();
			String expRow13List = "[s12, 1.00, 14.41, 14.41]";
			
			
			
			int reportsRow14ListCount = reportsRow14List.size();
			ArrayList<String> reportsRow14ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow14ListCount;i++)
			{
				String data = reportsRow14List.get(i).getText();
				reportsRow14ListArray.add(data);
			}
			String actRow14List = reportsRow14ListArray.toString();
			String expRow14List = "[s13, 1.00, 14.41, 14.41]";
			
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_NextBtn));
			report_NextBtn.click();
			
			Thread.sleep(2000);

			int reportsP2Row1ListCount = reportsRow1List.size();
			ArrayList<String> reportsP2Row1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsP2Row1ListArray.add(data);
			}
			String actP2Row1List = reportsP2Row1ListArray.toString();
			String expP2Row1List = "[s14, 1.00, 14.41, 14.41]";
			
			

			int reportsP2Row2ListCount = reportsRow2List.size();
			ArrayList<String> reportsP2Row2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsP2Row2ListArray.add(data);
			}
			String actP2Row2List = reportsP2Row2ListArray.toString();
			String expP2Row2List = "[SEC#RMA1, 1.00, 2.00, 2.00]";
			
			
			
			int reportsP2Row3ListCount = reportsRow3List.size();
			ArrayList<String> reportsP2Row3ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsP2Row3ListArray.add(data);
			}
			String actP2Row3List = reportsP2Row3ListArray.toString();
			String expP2Row3List = "[SEC#RMA2, 1.00, 2.00, 2.00]";
			
			
			
			int reportsP2Row4ListCount = reportsRow4List.size();
			ArrayList<String> reportsP2Row4ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row4ListCount;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsP2Row4ListArray.add(data);
			}
			String actP2Row4List = reportsP2Row4ListArray.toString();
			String expP2Row4List = "[SEC#RMA3, 1.00, 2.00, 2.00]";
			
			
			
			int reportsP2Row5ListCount = reportsRow5List.size();
			ArrayList<String> reportsP2Row5ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row5ListCount;i++)
			{
				String data = reportsRow5List.get(i).getText();
				reportsP2Row5ListArray.add(data);
			}
			String actP2Row5List = reportsP2Row5ListArray.toString();
			String expP2Row5List = "[SEC#RMA4, 1.00, 2.00, 2.00]";
			
			
			
			int reportsP2Row6ListCount = reportsRow6List.size();
			ArrayList<String> reportsP2Row6ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row6ListCount;i++)
			{
				String data = reportsRow6List.get(i).getText();
				reportsP2Row6ListArray.add(data);
			}
			String actP2Row6List = reportsP2Row6ListArray.toString();
			String expP2Row6List = "[SEC#RMA5, 1.00, 2.00, 2.00]";
			
			
			
			int reportsP2Row7ListCount = reportsRow7List.size();
			ArrayList<String> reportsP2Row7ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row7ListCount;i++)
			{
				String data = reportsRow7List.get(i).getText();
				reportsP2Row7ListArray.add(data);
			}
			String actP2Row7List = reportsP2Row7ListArray.toString();
			String expP2Row7List = "[SEC#RMA6, 1.00, 2.00, 2.00]";
			
			
			
			int reportsP2Row8ListCount = reportsRow8List.size();
			ArrayList<String> reportsP2Row8ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row8ListCount;i++)
			{
				String data = reportsRow8List.get(i).getText();
				reportsP2Row8ListArray.add(data);
			}
			String actP2Row8List = reportsP2Row8ListArray.toString();
			String expP2Row8List = "[SEC#RMA7, 1.00, 2.00, 2.00]";
			
			
			
			int reportsP2Row9ListCount = reportsRow9List.size();
			ArrayList<String> reportsP2Row9ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row9ListCount;i++)
			{
				String data = reportsRow9List.get(i).getText();
				reportsP2Row9ListArray.add(data);
			}
			String actP2Row9List = reportsP2Row9ListArray.toString();
			String expP2Row9List = "[SEC#RMA8, 1.00, 2.00, 2.00]";
			
			
			
			int reportsP2Row10ListCount = reportsRow10List.size();
			ArrayList<String> reportsP2Row10ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row10ListCount;i++)
			{
				String data = reportsRow10List.get(i).getText();
				reportsP2Row10ListArray.add(data);
			}
			String actP2Row10List = reportsP2Row10ListArray.toString();
			String expP2Row10List = "[SEC#RMA9, 1.00, 2.00, 2.00]";
			
			
			
			int reportsP2Row11ListCount = reportsRow11List.size();
			ArrayList<String> reportsP2Row11ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row11ListCount;i++)
			{
				String data = reportsRow11List.get(i).getText();
				reportsP2Row11ListArray.add(data);
			}
			String actP2Row11List = reportsRow11ListArray.toString();
			String expP2Row11List = "[s10, 1.00, 14.41, 14.41]";
			
			
			
			int reportsP2Row12ListCount = reportsRow12List.size();
			ArrayList<String> reportsP2Row12ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row12ListCount;i++)
			{
				String data = reportsRow12List.get(i).getText();
				reportsP2Row12ListArray.add(data);
			}
			String actP2Row12List = reportsP2Row12ListArray.toString();
			String expP2Row12List = "[SEC#RMA11, 1.00, 2.00, 2.00]";
			
			
			
			int reportsP2Row13ListCount = reportsRow13List.size();
			ArrayList<String> reportsP2Row13ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row13ListCount;i++)
			{
				String data = reportsRow13List.get(i).getText();
				reportsP2Row13ListArray.add(data);
			}
			String actP2Row13List = reportsP2Row13ListArray.toString();
			String expP2Row13List = "[SEC#RMA12, 1.00, 2.00, 2.00]";
			
			
			
			int reportsP2Row14ListCount = reportsRow14List.size();
			ArrayList<String> reportsP2Row14ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row14ListCount;i++)
			{
				String data = reportsRow14List.get(i).getText();
				reportsP2Row14ListArray.add(data);
			}
			String actP2Row14List = reportsP2Row14ListArray.toString();
			String expP2Row14List = "[Grand Total, 27.00, 240.15, 240.15]";
			
			System.out.println("*********************************checkStockAnalysisByRmaReport*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			System.out.println(actRow4List);
			System.out.println(expRow4List);
			
			System.out.println(actRow5List);
			System.out.println(expRow5List);
			
			System.out.println(actRow6List);
			System.out.println(expRow6List);
			
			System.out.println(actRow7List);
			System.out.println(expRow7List);
			
			System.out.println(actRow8List);
			System.out.println(expRow8List);
			
			System.out.println(actRow9List);
			System.out.println(expRow9List);
			
			System.out.println(actRow10List);
			System.out.println(expRow10List);
			
			System.out.println(actRow11List);
			System.out.println(expRow11List);
			
			System.out.println(actRow12List);
			System.out.println(expRow12List);
			
			System.out.println(actRow13List);
			System.out.println(expRow13List);
			
			System.out.println(actRow14List);
			System.out.println(expRow14List);
			
			System.out.println(actP2Row1List);
			System.out.println(expP2Row1List);
			
			System.out.println(actP2Row2List);
			System.out.println(expP2Row2List);
			
			System.out.println(actP2Row3List);
			System.out.println(expP2Row3List);
			
			System.out.println(actP2Row4List);
			System.out.println(expP2Row4List);
			
			System.out.println(actP2Row5List);
			System.out.println(expP2Row5List);
			
			System.out.println(actP2Row6List);
			System.out.println(expP2Row6List);
			
			System.out.println(actP2Row7List);
			System.out.println(expP2Row7List);
			
			System.out.println(actP2Row8List);
			System.out.println(expP2Row8List);
			
			System.out.println(actP2Row9List);
			System.out.println(expP2Row9List);
			
			System.out.println(actP2Row10List);
			System.out.println(expP2Row10List);
			
			System.out.println(actP2Row11List);
			System.out.println(expP2Row11List);
			
			System.out.println(actP2Row12List);
			System.out.println(expP2Row12List);
			
			System.out.println(actP2Row13List);
			System.out.println(expP2Row13List);
			
			System.out.println(actP2Row14List);
			System.out.println(expP2Row14List);
			
			Thread.sleep(2000);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List) 
				&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List) 
				&& actRow10List.equalsIgnoreCase(expRow10List) && actRow11List.equalsIgnoreCase(expRow11List) && actRow12List.equalsIgnoreCase(expRow12List) 
				&& actRow13List.equalsIgnoreCase(expRow13List) && actRow14List.equalsIgnoreCase(expRow14List) 
				&& actP2Row1List.equalsIgnoreCase(expP2Row1List) && actP2Row2List.equalsIgnoreCase(expP2Row2List) && actP2Row3List.equalsIgnoreCase(expP2Row3List) 
				&& actP2Row4List.equalsIgnoreCase(expP2Row4List) && actP2Row5List.equalsIgnoreCase(expP2Row5List) && actP2Row6List.equalsIgnoreCase(expP2Row6List) 
				&& actP2Row7List.equalsIgnoreCase(expP2Row7List) && actP2Row8List.equalsIgnoreCase(expP2Row8List) && actP2Row9List.equalsIgnoreCase(expP2Row9List) 
				&& actP2Row10List.equalsIgnoreCase(expP2Row10List) && actP2Row11List.equalsIgnoreCase(expP2Row11List) && actP2Row12List.equalsIgnoreCase(expP2Row12List) 
				&& actP2Row13List.equalsIgnoreCase(expP2Row13List) && actP2Row14List.equalsIgnoreCase(expP2Row14List))
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
			
		}
		
		
		
		
		
		
		
		
		
		
		// Stock Details By Bins
		
		public boolean checkStockDetailsByBinsReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(binsReportMenu));
			binsReportMenu.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockDetailsByBins));
			stockDetailsByBins.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			boolean actReportsTable = reportsTable.getText().isEmpty();
			boolean expReportsTable = true;
			
			System.out.println("*********************************checkStockDetailsByBinsReport*****************************************");
			
			System.out.println("Reports Table isEmpty   : "+actReportsTable   +" Value Expected : "+expReportsTable);
			
			if(actReportsTable==expReportsTable)
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		
		
		// Stock Balance By Bins

		public boolean checkStockBalanaceByBinsReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(binsReportMenu));
			binsReportMenu.click();
			
			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockBalancesByBins));
			stockBalancesByBins.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			boolean actReportsTable = reportsTable.getText().isEmpty();
			boolean expReportsTable = true;
			

			System.out.println("*********************************checkStockBalanceByBinsReport*****************************************");
			
			System.out.println("Reports Table isEmpty    : "+actReportsTable   +" Value Expected : "+expReportsTable);
			
			if(actReportsTable==expReportsTable)
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		//Stock Balance By Item By Bins
		
		public boolean checkStockBalancesByItemByBinsReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(binsReportMenu));
			binsReportMenu.click();
			
			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockBalancesByItemByBins));
			stockBalancesByItemByBins.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			boolean actReportsTable = reportsTable.getText().isEmpty();
			boolean expReportsTable = true;
			
			System.out.println("*********************************checkStockBalancesByItemByBinsReport*****************************************");
			
			System.out.println("Reports Table isEmpty  : "+actReportsTable+" Value Expected: "+expReportsTable);
			
			
			if(actReportsTable==expReportsTable)
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		
		// Expired Stock By Bins

		public boolean checkExpiredStockByBinsReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(binsReportMenu));
			binsReportMenu.click();
			
			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(expiredStockByBins));
			expiredStockByBins.click();
			
			Thread.sleep(3000);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			boolean actReportsTable = reportsTable.getText().isEmpty();
			boolean expReportsTable = true;

			System.out.println("report_Body : "+actReportsTable+" Value Expected : "+expReportsTable);
			
			if(actReportsTable==expReportsTable)
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		
		// Stock Balance By WareHouse

		public boolean checkStockBalanceByWarehouseReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockBalanceByWarehouse));
			stockBalanceByWarehouse.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			int reportsRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
			String actRow1List = reportsRow1ListArray.toString();
			String expRow1List = "[SECUNDERABAD SECUNDERABAD, , , , , ]";
			
			

			int reportsRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsRow2ListArray.add(data);
			}
			String actRow2List = reportsRow2ListArray.toString();
			String expRow2List = "[WA COGS ITEM, , , 27.00, , 27.00]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[Grand Total, , , 27.00, , 27.00]";

			System.out.println("*********************************checkStockBalanceByWarehouseReport*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
	
	
		
		
		
		
		
		
		
		public boolean checkStockAgeingAnalysisReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAgeingAnalysisMenu));
			stockAgeingAnalysisMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ageingAnalysis));
			ageingAnalysis.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
			
			int reportsRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
			String actRow1List = reportsRow1ListArray.toString();
			String expRow1List = "[WA COGS ITEM, 27.00, 316.94, 27.00, 316.94, , , , , , , , , , , , , , , ]";
			
			

			int reportsRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsRow2ListArray.add(data);
			}
			String actRow2List = reportsRow2ListArray.toString();
			String expRow2List = "[Total, 27.00, 316.94, 27.00, 316.94, , , , , , , , , , , , , , , ]";
			
			
			System.out.println("*****************************************checkStockAgeingAnalysisReports***************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		// Ageing Analysis By Batch
		
		public boolean checkAgeingAnalysisByBatchReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();

			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAgeingAnalysisMenu));
			stockAgeingAnalysisMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ageingAnalysisByBatch));
			ageingAnalysisByBatch.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			boolean actReportsTable = reportsTable.getText().isEmpty();
			boolean expReportsTable = true;
			
			System.out.println("*****************************************checkAgeingAnalysisByBatchReports***************************************");
			
			System.out.println("Reports Table isEmpty : "+actReportsTable+"  Value Expected  "+expReportsTable);
			
			if(actReportsTable==expReportsTable)
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		
		
		// Ageing Analysis By Rma
		
		public boolean checkAgeingAnalysisByRmaReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();

			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAgeingAnalysisMenu));
			stockAgeingAnalysisMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ageingAnalysisByRMA));
			ageingAnalysisByRMA.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
			
			int reportsRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
			String actRow1List = reportsRow1ListArray.toString();
			String expRow1List = "[WA COGS ITEM WA COGS ITEM, , , , , , , , , , , , , , , , , , ]";
			
			

			int reportsRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsRow2ListArray.add(data);
			}
			String actRow2List = reportsRow2ListArray.toString();
			String expRow2List = "[s, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[s1, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			
			
			int reportsRow4ListCount = reportsRow4List.size();
			ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow4ListCount;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
			String actRow4List = reportsRow4ListArray.toString();
			String expRow4List = "[s10, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			
			
			int reportsRow5ListCount = reportsRow5List.size();
			ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow5ListCount;i++)
			{
				String data = reportsRow5List.get(i).getText();
				reportsRow5ListArray.add(data);
			}
			String actRow5List = reportsRow5ListArray.toString();
			String expRow5List = "[s11, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			
			
			int reportsRow6ListCount = reportsRow6List.size();
			ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow6ListCount;i++)
			{
				String data = reportsRow6List.get(i).getText();
				reportsRow6ListArray.add(data);
			}
			String actRow6List = reportsRow6ListArray.toString();
			String expRow6List = "[s12, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			
			
			int reportsRow7ListCount = reportsRow7List.size();
			ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow7ListCount;i++)
			{
				String data = reportsRow7List.get(i).getText();
				reportsRow7ListArray.add(data);
			}
			String actRow7List = reportsRow7ListArray.toString();
			String expRow7List = "[s13, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			
			
			int reportsRow8ListCount = reportsRow8List.size();
			ArrayList<String> reportsRow8ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow8ListCount;i++)
			{
				String data = reportsRow8List.get(i).getText();
				reportsRow8ListArray.add(data);
			}
			String actRow8List = reportsRow8ListArray.toString();
			String expRow8List = "[s14, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			
			
			int reportsRow9ListCount = reportsRow9List.size();
			ArrayList<String> reportsRow9ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow9ListCount;i++)
			{
				String data = reportsRow9List.get(i).getText();
				reportsRow9ListArray.add(data);
			}
			String actRow9List = reportsRow9ListArray.toString();
			String expRow9List = "[s2, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			
			
			int reportsRow10ListCount = reportsRow10List.size();
			ArrayList<String> reportsRow10ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow10ListCount;i++)
			{
				String data = reportsRow10List.get(i).getText();
				reportsRow10ListArray.add(data);
			}
			String actRow10List = reportsRow10ListArray.toString();
			String expRow10List = "[s3, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			
			
			int reportsRow11ListCount = reportsRow11List.size();
			ArrayList<String> reportsRow11ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow11ListCount;i++)
			{
				String data = reportsRow11List.get(i).getText();
				reportsRow11ListArray.add(data);
			}
			String actRow11List = reportsRow11ListArray.toString();
			String expRow11List = "[s4, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			
			
			int reportsRow12ListCount = reportsRow12List.size();
			ArrayList<String> reportsRow12ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow12ListCount;i++)
			{
				String data = reportsRow12List.get(i).getText();
				reportsRow12ListArray.add(data);
			}
			String actRow12List = reportsRow12ListArray.toString();
			String expRow12List = "[s5, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			
			
			int reportsRow13ListCount = reportsRow13List.size();
			ArrayList<String> reportsRow13ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow13ListCount;i++)
			{
				String data = reportsRow13List.get(i).getText();
				reportsRow13ListArray.add(data);
			}
			String actRow13List = reportsRow13ListArray.toString();
			String expRow13List = "[s6, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			
			
			int reportsRow14ListCount = reportsRow14List.size();
			ArrayList<String> reportsRow14ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow14ListCount;i++)
			{
				String data = reportsRow14List.get(i).getText();
				reportsRow14ListArray.add(data);
			}
			String actRow14List = reportsRow14ListArray.toString();
			String expRow14List = "[s7, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_NextBtn));
			report_NextBtn.click();
			
			Thread.sleep(2000);

			int reportsP2Row1ListCount = reportsRow1List.size();
			ArrayList<String> reportsP2Row1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsP2Row1ListArray.add(data);
			}
			String actP2Row1List = reportsP2Row1ListArray.toString();
			String expP2Row1List = "[s8, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			

			int reportsP2Row2ListCount = reportsRow2List.size();
			ArrayList<String> reportsP2Row2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsP2Row2ListArray.add(data);
			}
			String actP2Row2List = reportsP2Row2ListArray.toString();
			String expP2Row2List = "[s9, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			
			
			int reportsP2Row3ListCount = reportsRow3List.size();
			ArrayList<String> reportsP2Row3ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsP2Row3ListArray.add(data);
			}
			String actP2Row3List = reportsP2Row3ListArray.toString();
			String expP2Row3List = "[SEC#RMA1, 1.00, 11.74, , , , , , , , , , , , , 1.00, 11.74, , ]";
			
			
			
			int reportsP2Row4ListCount = reportsRow4List.size();
			ArrayList<String> reportsP2Row4ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row4ListCount;i++)
			{
				String data = reportsRow4List.get(i).getText();
				reportsP2Row4ListArray.add(data);
			}
			String actP2Row4List = reportsP2Row4ListArray.toString();
			String expP2Row4List = "[SEC#RMA10, 1.00, 11.74, , , , , , , , , , , , , 1.00, 11.74, , ]";
			
			
			
			int reportsP2Row5ListCount = reportsRow5List.size();
			ArrayList<String> reportsP2Row5ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row5ListCount;i++)
			{
				String data = reportsRow5List.get(i).getText();
				reportsP2Row5ListArray.add(data);
			}
			String actP2Row5List = reportsP2Row5ListArray.toString();
			String expP2Row5List = "[SEC#RMA11, 1.00, 11.74, , , , , , , , , , , , , 1.00, 11.74, , ]";
			
			
			
			int reportsP2Row6ListCount = reportsRow6List.size();
			ArrayList<String> reportsP2Row6ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row6ListCount;i++)
			{
				String data = reportsRow6List.get(i).getText();
				reportsP2Row6ListArray.add(data);
			}
			String actP2Row6List = reportsP2Row6ListArray.toString();
			String expP2Row6List = "[SEC#RMA12, 1.00, 11.74, , , , , , , , , , , , , 1.00, 11.74, , ]";
			
			
			
			int reportsP2Row7ListCount = reportsRow7List.size();
			ArrayList<String> reportsP2Row7ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row7ListCount;i++)
			{
				String data = reportsRow7List.get(i).getText();
				reportsP2Row7ListArray.add(data);
			}
			String actP2Row7List = reportsP2Row7ListArray.toString();
			String expP2Row7List = "[SEC#RMA2, 1.00, 11.74, , , , , , , , , , , , , 1.00, 11.74, , ]";
			
			
			
			int reportsP2Row8ListCount = reportsRow8List.size();
			ArrayList<String> reportsP2Row8ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row8ListCount;i++)
			{
				String data = reportsRow8List.get(i).getText();
				reportsP2Row8ListArray.add(data);
			}
			String actP2Row8List = reportsP2Row8ListArray.toString();
			String expP2Row8List = "[SEC#RMA3, 1.00, 11.74, , , , , , , , , , , , , 1.00, 11.74, , ]";
			
			
			
			int reportsP2Row9ListCount = reportsRow9List.size();
			ArrayList<String> reportsP2Row9ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row9ListCount;i++)
			{
				String data = reportsRow9List.get(i).getText();
				reportsP2Row9ListArray.add(data);
			}
			String actP2Row9List = reportsP2Row9ListArray.toString();
			String expP2Row9List = "[SEC#RMA4, 1.00, 11.74, , , , , , , , , , , , , 1.00, 11.74, , ]";
			
			
			
			int reportsP2Row10ListCount = reportsRow10List.size();
			ArrayList<String> reportsP2Row10ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row10ListCount;i++)
			{
				String data = reportsRow10List.get(i).getText();
				reportsP2Row10ListArray.add(data);
			}
			String actP2Row10List = reportsP2Row10ListArray.toString();
			String expP2Row10List = "[SEC#RMA5, 1.00, 11.74, , , , , , , , , , , , , 1.00, 11.74, , ]";
			
			
			
			int reportsP2Row11ListCount = reportsRow11List.size();
			ArrayList<String> reportsP2Row11ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row11ListCount;i++)
			{
				String data = reportsRow11List.get(i).getText();
				reportsP2Row11ListArray.add(data);
			}
			String actP2Row11List = reportsRow11ListArray.toString();
			String expP2Row11List = "[s4, 1.00, 11.74, 1.00, 11.74, , , , , , , , , , , , , , ]";
			
			
			
			int reportsP2Row12ListCount = reportsRow12List.size();
			ArrayList<String> reportsP2Row12ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row12ListCount;i++)
			{
				String data = reportsRow12List.get(i).getText();
				reportsP2Row12ListArray.add(data);
			}
			String actP2Row12List = reportsP2Row12ListArray.toString();
			String expP2Row12List = "[SEC#RMA7, 1.00, 11.74, , , , , , , , , , , , , 1.00, 11.74, , ]";
			
			
			
			int reportsP2Row13ListCount = reportsRow13List.size();
			ArrayList<String> reportsP2Row13ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row13ListCount;i++)
			{
				String data = reportsRow13List.get(i).getText();
				reportsP2Row13ListArray.add(data);
			}
			String actP2Row13List = reportsP2Row13ListArray.toString();
			String expP2Row13List = "[SEC#RMA8, 1.00, 11.74, , , , , , , , , , , , , 1.00, 11.74, , ]";
			
			
			
			int reportsP2Row14ListCount = reportsRow14List.size();
			ArrayList<String> reportsP2Row14ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP2Row14ListCount;i++)
			{
				String data = reportsRow14List.get(i).getText();
				reportsP2Row14ListArray.add(data);
			}
			String actP2Row14List = reportsP2Row14ListArray.toString();
			String expP2Row14List = "[SEC#RMA9, 1.00, 11.74, , , , , , , , , , , , , 1.00, 11.74, , ]";
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_NextBtn));
			report_NextBtn.click();
			
			Thread.sleep(3000);

			int reportsP3Row1ListCount = reportsRow1List.size();
			ArrayList<String> reportsP3Row1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsP3Row1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsP3Row1ListArray.add(data);
			}
			String actP3Row1List = reportsP3Row1ListArray.toString();
			String expP3Row1List = "[Grand Total, 27.00, 316.94, 15.00, 176.08, , , , , , , , , , , 12.00, 140.86, , ]";
			
			
			System.out.println("*****************************************checkAgeingAnalysisByRMAReports***************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			System.out.println(actRow4List);
			System.out.println(expRow4List);
			
			System.out.println(actRow5List);
			System.out.println(expRow5List);
			
			System.out.println(actRow6List);
			System.out.println(expRow6List);
			
			System.out.println(actRow7List);
			System.out.println(expRow7List);
			
			System.out.println(actRow8List);
			System.out.println(expRow8List);
			
			System.out.println(actRow9List);
			System.out.println(expRow9List);
			
			System.out.println(actRow10List);
			System.out.println(expRow10List);
			
			System.out.println(actRow11List);
			System.out.println(expRow11List);
			
			System.out.println(actRow12List);
			System.out.println(expRow12List);
			
			System.out.println(actRow13List);
			System.out.println(expRow13List);
			
			System.out.println(actRow14List);
			System.out.println(expRow14List);
			
			System.out.println(actP2Row1List);
			System.out.println(expP2Row1List);
			
			System.out.println(actP2Row2List);
			System.out.println(expP2Row2List);
			
			System.out.println(actP2Row3List);
			System.out.println(expP2Row3List);
			
			System.out.println(actP2Row4List);
			System.out.println(expP2Row4List);
			
			System.out.println(actP2Row5List);
			System.out.println(expP2Row5List);
			
			System.out.println(actP2Row6List);
			System.out.println(expP2Row6List);
			
			System.out.println(actP2Row7List);
			System.out.println(expP2Row7List);
			
			System.out.println(actP2Row8List);
			System.out.println(expP2Row8List);
			
			System.out.println(actP2Row9List);
			System.out.println(expP2Row9List);
			
			System.out.println(actP2Row10List);
			System.out.println(expP2Row10List);
			
			System.out.println(actP2Row11List);
			System.out.println(expP2Row11List);
			
			System.out.println(actP2Row12List);
			System.out.println(expP2Row12List);
			
			System.out.println(actP2Row13List);
			System.out.println(expP2Row13List);
			
			System.out.println(actP2Row14List);
			System.out.println(expP2Row14List);
			
			System.out.println(actP3Row1List);
			System.out.println(expP3Row1List);
			
			Thread.sleep(2000);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List) 
				&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List) 
				&& actRow10List.equalsIgnoreCase(expRow10List) && actRow11List.equalsIgnoreCase(expRow11List) && actRow12List.equalsIgnoreCase(expRow12List) 
				&& actRow13List.equalsIgnoreCase(expRow13List) && actRow14List.equalsIgnoreCase(expRow14List) 
				&& actP2Row1List.equalsIgnoreCase(expP2Row1List) && actP2Row2List.equalsIgnoreCase(expP2Row2List) && actP2Row3List.equalsIgnoreCase(expP2Row3List) 
				&& actP2Row4List.equalsIgnoreCase(expP2Row4List) && actP2Row5List.equalsIgnoreCase(expP2Row5List) && actP2Row6List.equalsIgnoreCase(expP2Row6List) 
				&& actP2Row7List.equalsIgnoreCase(expP2Row7List) && actP2Row8List.equalsIgnoreCase(expP2Row8List) && actP2Row9List.equalsIgnoreCase(expP2Row9List) 
				&& actP2Row10List.equalsIgnoreCase(expP2Row10List) && actP2Row11List.equalsIgnoreCase(expP2Row11List) && actP2Row12List.equalsIgnoreCase(expP2Row12List) 
				&& actP2Row13List.equalsIgnoreCase(expP2Row13List) && actP2Row14List.equalsIgnoreCase(expP2Row14List)
				&& actP3Row1List.equalsIgnoreCase(expP3Row1List))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}

		
		
		
		
		
		
		
		// Ageing Stock ByBins

		public boolean checkAgeingStockByBinsReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();

			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAgeingAnalysisMenu));
			stockAgeingAnalysisMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ageingStockByBins));
			ageingStockByBins.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			boolean actReportsTable = reportsTable.getText().isEmpty();
			boolean expReportsTable = true;
			
			System.out.println("*****************************************checkAgeingStockByBinsReports***************************************");
			
			System.out.println("Reports Table isEmpty : "+actReportsTable+"  Value Expected  "+expReportsTable);
				
			if(actReportsTable==expReportsTable)
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		// Ageing By Item By Bins

		public boolean checkAgeingByItemByBinsReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();

			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAgeingAnalysisMenu));
			stockAgeingAnalysisMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ageingByItemByBins));
			ageingByItemByBins.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			boolean actReportsTable = reportsTable.getText().isEmpty();
			boolean expReportsTable = true;
			
			System.out.println("*****************************************checkAgeingByItemByBinsReports***************************************");
			
			System.out.println("Reports Table isEmpty : "+actReportsTable+"  Value Expected  "+expReportsTable);
			
			if(actReportsTable==expReportsTable)
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		
		
		// Reorder Report

		public boolean checkReorderReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			
			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reorderReport));
			reorderReport.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			boolean actReportsTable = reportsTable.getText().isEmpty();
			boolean expReportsTable = true;
			
			System.out.println("*********************************checkReorderReportReport*****************************************");
			
			System.out.println("Reports Table isEmpty : "+actReportsTable+"  Value Expected  "+expReportsTable);
			
			if(actReportsTable==expReportsTable)
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		// Fast Moving Item (XYZ Analysis)

		public boolean checkFastMovingItemsReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();

			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fastMovingItem));
			fastMovingItem.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().equals(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
			
			int reportsRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
			String actRow1List = reportsRow1ListArray.toString();
			String expRow1List = "[WA COGS ITEM, WA COGS ITEM, WA COGS ITEM, 27.00, 255.49, 9.46, SECUNDERABAD]";
			
			

			int reportsRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsRow2ListArray.add(data);
			}
			String actRow2List = reportsRow2ListArray.toString();
			String expRow2List = "[Grand Total, , , 27.00, 255.49, 9.46, ]";
			
			System.out.println("*********************************checkFastMovingItemsReport*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		
		// Slow Moving Item

		public boolean checkSlowMovingItemsReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();

			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(slowMovingItem));
			slowMovingItem.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			int reportsRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
			String actRow1List = reportsRow1ListArray.toString();
			String expRow1List = "[WA COGS ITEM, WA COGS ITEM, WA COGS ITEM, 27.00, 255.49, 9.46, SECUNDERABAD]";
			
			

			int reportsRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsRow2ListArray.add(data);
			}
			String actRow2List = reportsRow2ListArray.toString();
			String expRow2List = "[Grand Total, , , 27.00, 255.49, 9.46, ]";
			
			
			System.out.println("*********************************checkSlowMovingItemsReport*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		// PeakAndLowBalances
		
		public boolean checkPeakAndLowBalancesReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();

			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(peakORLowBalances));
			peakORLowBalances.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().equals(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
			String actItem = sl_1stRow1stCol.getText();
			String expItem = "WA COGS ITEM"; 
			
			boolean actNegativeBalanceDate = sl_1stRow2ndCol.getText().isEmpty();
			boolean expNegativeBalanceDate = true;
			
			boolean actNegativeBalance = sl_1stRow3rdCol.getText().isEmpty();
			boolean expNegativeBalance = true;

			String actLowestPositiveBalance = sl_1stRow5thCol.getText();
			String expLowestPositiveBalance = "22.00";
			
			String actHighestBalance = sl_1stRow7thCol.getText();
			String expHighestBalance = "552.00";
			
			boolean actTotalNegativeBalanceDate = sl_2ndRow2ndCol.getText().isEmpty();
			boolean expTotalNegativeBalanceDate = true;
			
			boolean actTotalNegativeBalance = sl_2ndRow3rdCol.getText().isEmpty();
			boolean expTotalNegativeBalance = true;
			
			boolean actTotalLowestPositiveBalanceDate = sl_2ndRow4thCol.getText().isEmpty();
			boolean expTotalLowestPositiveBalanceDate = true;
			
			String actTotalLowestPositiveBalance = sl_2ndRow5thCol.getText();
			String expTotalLowestPositiveBalance = "22.00";
			
			boolean actTotalHighestBalanceDate = sl_2ndRow6thCol.getText().isEmpty();
			boolean expTotalHighestBalanceDate = true;
			
			String actTotalHighestBalance = sl_2ndRow7thCol.getText();
			String expTotalHighestBalance = "552.00";
			
			
			boolean actResult = actNegativeBalanceDate==expNegativeBalanceDate && actNegativeBalance==expNegativeBalance 
								&& actTotalNegativeBalanceDate==expTotalNegativeBalanceDate && actTotalNegativeBalance==expTotalNegativeBalance
								&& actTotalLowestPositiveBalanceDate==expTotalLowestPositiveBalanceDate && actTotalHighestBalanceDate==expTotalHighestBalanceDate;
			
			boolean expResult = true;
			
			System.out.println("*********************************checkPeakAndLowBalancesReport*****************************************");
			
			System.out.println("Item : "+actItem+"  Value Expected  "+expItem);
			System.out.println("LowestPositiveBalance : "+actLowestPositiveBalance+"  Value Expected  "+expLowestPositiveBalance);
			System.out.println("HighestBalance : "+actHighestBalance+"  Value Expected  "+expHighestBalance);
			System.out.println("TotalLowestPositiveBalance : "+actTotalLowestPositiveBalance+"  Value Expected  "+expTotalLowestPositiveBalance);
			System.out.println("TotalHighestBalance : "+actTotalHighestBalance+"  Value Expected  "+expTotalHighestBalance);
			System.out.println("Empty cells : "+actResult+"  Value Expected  "+expResult);
			if(actItem.equalsIgnoreCase(expItem) && actLowestPositiveBalance.equalsIgnoreCase(expLowestPositiveBalance)
				&& actHighestBalance.equalsIgnoreCase(expHighestBalance) && actTotalLowestPositiveBalance.equalsIgnoreCase(expTotalLowestPositiveBalance)
				&& actTotalHighestBalance.equalsIgnoreCase(expTotalHighestBalance) && actResult==expResult)
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		
		public boolean checkBestSellingItemReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();

			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(bestSellingItem));
			bestSellingItem.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(includeAllVoucherChkBox));
			includeAllVoucherChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
			
			int reportsRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
			String actRow1List = reportsRow1ListArray.toString();
			String expRow1List = "[WA COGS ITEM, 541.00, 144.15]";
			
			

			int reportsRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsRow2ListArray.add(data);
			}
			String actRow2List = reportsRow2ListArray.toString();
			String expRow2List = "[Grand Total, 541.00, 144.15]";
			
			System.out.println("*********************************checkBestSellingItemReport*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		
		// Stock Transfer Report

		public boolean checkStockTransferRegisterReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();

			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferReport));
			stockTransferReport.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().equals(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
			
			
			int reportsRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
			String actRow1List = reportsRow1ListArray.toString();
			String expRow1List = "[WA COGS ITEM WA COGS ITEM, , , , , , , , , ]";
			
			

			int reportsRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsRow2ListArray.add(data);
			}
			String actRow2List = reportsRow2ListArray.toString();
			String expRow2List = "[, HYDERABAD, 2.00, , , , , , ]";
			
			
			
			int reportsRow3ListCount = reportsRow3List.size();
			ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
			for(int i=2;i<reportsRow3ListCount;i++)
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
			String actRow3List = reportsRow3ListArray.toString();
			String expRow3List = "[, , 2.00, , , , , , ]";
			
			
			System.out.println("*********************************checkStockTransferRegisterReport*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			System.out.println(actRow3List);
			System.out.println(expRow3List);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		
		// Stock reservation register

		public boolean checkStockReservationReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();

			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockReservationReport));
			stockReservationReport.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			int rowcount=stockLedgerHometableRowCount.size();
			
			System.out.println(rowcount);
			
			for (int i = 1; i <= rowcount; i++) 
			{
				WebElement name=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[12]"));
				
				String actname=name.getText();
				
				if(actname.equalsIgnoreCase("WA COGS ITEM"))
				{
					WebElement index=getDriver().findElement(By.xpath("//tbody[@id='LandingGridBody']/tr["+i+"]/td[8]/div/label/input"));
					index.click();
					
					break;
				}
			}
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			boolean actReportsTable = reportsTable.getText().isEmpty();
			boolean expReportsTable = true;
			
			System.out.println("*********************************checkStockReservationReportReport*****************************************");

			System.out.println("Reports Table isEmpty : "+actReportsTable+"  Value Expected  "+expReportsTable);
			
			if(actReportsTable==expReportsTable)
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		
		
		
		
		
		
		
		
		// Transaction Type wise Stock report
		
		public boolean checkTransTypeWiseStockReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();

			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionTypeWiseStockReport));
			transactionTypeWiseStockReport.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().equals(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
			
			int reportsRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
			String actRow1List = reportsRow1ListArray.toString();
			String expRow1List = "[WA COGS ITEM, 27.00, 316.94, 11.74, , , 36.00, 222.00, , , 527.00, 6,321.00, , , , , , , , , , , , , , , 15.00, "
								    + "216.15, , , 3.00, 30.00, , , , , , , , , , , , , , , 29.00, 309.13, , , , , 525.00, 6,163.50]";
			
			int reportsRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsRow2ListArray.add(data);
			}
			String actRow2List = reportsRow2ListArray.toString();
			String expRow2List = "[Grand Total, 27.00, 316.94, 11.74, , , 36.00, 222.00, , , 527.00, 6,321.00, , , , , , , , , , , , , , , 15.00, "
									+ "216.15, , , 3.00, 30.00, , , , , , , , , , , , , , , 29.00, 309.13, , , , , 525.00, 6,163.50]";
			
			
			System.out.println("*********************************checkTransTypeWiseStockReports*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		

		
		
		
		
		
		// Actual Consumption Report

		public boolean checkActualConsumptionReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();

			getAction().sendKeys(Keys.END).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(actualConsumptionReport));
			actualConsumptionReport.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByVisibleText(" As on date ");
			
			String actsl_DateOptionDropdown    =s.getFirstSelectedOption().getText();
			String expsl_DateOptionDropdown    =" As on date ";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
			sl_SelectAllItemsChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().equals(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
			
			int reportsRow1ListCount = reportsRow1List.size();
			ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow1ListCount;i++)
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
			String actRow1List = reportsRow1ListArray.toString();
			String expRow1List = "[WA COGS ITEM, 36.00, 222.00, , , 527.00, 6,321.00, 3.00, 30.00, 15.00, 216.15, , , 29.00, "
									+ "8,964.64, , , 525.00, 32,35,837.50, , , 27.00, 1,19,926.41, 32,38,012.99]";
			
			int reportsRow2ListCount = reportsRow2List.size();
			ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
			for(int i=1;i<reportsRow2ListCount;i++)
			{
				String data = reportsRow2List.get(i).getText();
				reportsRow2ListArray.add(data);
			}
			String actRow2List = reportsRow2ListArray.toString();
			String expRow2List = "[Grand Total, 36.00, 222.00, , , 527.00, 6,321.00, 3.00, 30.00, 15.00, 216.15, , , 29.00, "
									+ "8,964.64, , , 525.00, 32,35,837.50, , , 27.00, 1,19,926.41, 32,38,012.99]";
			
			System.out.println("*********************************checkActualConsumptionReports*****************************************");
			
			System.out.println(actRow1List);
			System.out.println(expRow1List);
			
			System.out.println(actRow2List);
			System.out.println(expRow2List);
			
			if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
				sl_CloseBtn.click();
				Thread.sleep(2000);
				return false;
			}
		}
		

		


		
		
		
		// Erase all Transactions after Transactions flow
		
		public boolean checkEraseAllTransactionsAfterCompletingTransactionFlow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
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
			
			String expValidationMsg="Data deleted successfully.";
			
			String actValidationMsg=checkValidationMessage(expValidationMsg);
			
			System.out.println("*************************************checkEraseAllTransactionsAfterCompletingTransactionFlow*********************************");
			System.out.println("Message  :  "+actValidationMsg +" Value Expected : "+expValidationMsg);
			
			if (actValidationMsg.equalsIgnoreCase(expValidationMsg))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		
		
		
		
		
		//********************************************************** RMA Options Starts Here ******************************************************
		
		
		
				
		// Check RMA Option from Preferences :  Two Different Items Cannot Place in Same RMA and Verify Saving Inward Document.
				
		// RMA numbers are always numeric
		
		public boolean checkConfigureSelectAndUpdateRMAAlwaysNumericChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsmenuBtn));
			SettingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigureTransactionBtn));
			ConfigureTransactionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMABtn));
			RMABtn.click();
			
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMANumberAlwaysnumericChkBox));
			
			if (RMANumberAlwaysnumericChkBox.isSelected()==false) 
			{
				RMANumberAlwaysnumericChkBox.click();
			}
			
			Thread.sleep(2000);
			
			String actRMAOptionIsSelected = Boolean.toString(RMANumberAlwaysnumericChkBoxIsSelected.isSelected());
			String expRMAOptionIsSelected =    excelReader.getCellData(xlSheetName, 12, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 12, 7, actRMAOptionIsSelected);

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			
			getWaitForAlert();
			String actAlert=getAlert().getText();
			
			String expAlert= excelReader.getCellData(xlSheetName, 13, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 13, 7, actAlert);
			
			getAlert().accept();
			
			String expMessage=  excelReader.getCellData(xlSheetName, 14, 6);
			
			String actMessage=checkValidationMessage(expMessage);
			excelReader.setCellData(xlfile, xlSheetName, 14, 7, actMessage);
			
			
			System.out.println("RMAOptionIsSelected : "+actRMAOptionIsSelected+"  value expected  "+expRMAOptionIsSelected);
			System.out.println("Alert               : "+actAlert              +"  value expected  "+expAlert);
			System.out.println("Error Message       : "+actMessage            +"  value expected  "+expMessage);
					
			if(actRMAOptionIsSelected.equalsIgnoreCase(expRMAOptionIsSelected) && actAlert.equalsIgnoreCase(expAlert) && actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				excelReader.setCellData(xlfile, xlSheetName, 11, 8, resPass);
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				excelReader.setCellData(xlfile, xlSheetName, 11, 8, resFail);
				
				return false;
			}
		}
			
			
			
				
		
			
			
		// Login and Logout After updating RMA Always Numeric
		
		public boolean checkLogoutAndLoginAfterUpdatingRMAAlwaysNumeric() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
  
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
			  userNameDisplayLogo.click();
			 
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			  logoutOption.click();
			  Thread.sleep(2000);
			  
			  System.out.println("Logout  Successfully");
	
			  
		     String  act=Boolean.toString(checkLogin());
		     String exp= excelReader.getCellData(xlSheetName, 16, 6);
              
            excelReader.setCellData(xlfile, xlSheetName, 16, 7, act);
		      
		      System.out.println(" Login status act:"+act);
		      System.out.println(" Login status exp:"+exp);
		     
		      if(act.equalsIgnoreCase(exp))
		      {
		    	 

		    	  System.out.println("***Test Pass: Login Successfull***");
			
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 15, 8, resPass);
				
				return true;
			  }
		      else
		      {
		    	 
				System.out.println("***Test Fail: Login Not Successfull***");
			
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 15, 8, resFail);

				
				return false;
			  }
		   }
		
			
			
		
		
		
		
		
		// Purchase Voucher
		
		public boolean checkPurchaseVoucherVATVoucher1RMAAsAlphabets() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
				
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+vendorAccount);
				
				if (vendorAccount.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 18, 5)))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 19, 5)))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();
			
			enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
			int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
			
			for (int i = 0; i < pvvGridWarehouseListCount; i++) 
			{
				String warehouse=pvvGridWarehouseList.get(i).getText();
				if (warehouse.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 20, 5))) 
				{
					pvvGridWarehouseList.get(i).click();
					break;
				}
			}
	        enter_WarehouseTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
	        int pvvGridItemListCount=pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) 
			{
				String Item=pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 21, 5))) 
				{
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxCode));
	        enter_PvTaxCode.sendKeys(Keys.TAB);
				
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	        enter_AccountTxt.click();
			
	        Thread.sleep(2000);
	        
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 22, 5));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 23, 5));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
			select1stRow_19thColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 24, 5));
			
			Thread.sleep(2000);
			
			boolean actSerialNo=RMAPopupSerialNumberTxtField.getText().isEmpty();
			
			
			System.out.println("********************************** Serial No Text field: "+actSerialNo);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			
	        if(actSerialNo==true) 
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
				
				
			
		
		
			
			
		public boolean checkPurchaseVoucherVATVoucher1RMAAsNumeric() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			 
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 26, 5));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMAColumn=enter_RMA.getAttribute("value");
			String expRMAColumn=  excelReader.getCellData(xlSheetName, 27, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 27, 7, actRMAColumn);

			
			String docno=documentNumberTxt.getAttribute("value");
			Thread.sleep(1000);
			
					
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String  actSaving= Boolean.toString(checkBackgroundSavingMessage(docno));
			String expSaving= excelReader.getCellData(xlSheetName, 28, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 28, 7, actSaving);

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  

			
			
			System.out.println("RMA Column     : "+actRMAColumn +"  value expected  "+expRMAColumn);
			
			if(actRMAColumn.equalsIgnoreCase(expRMAColumn) &&actSaving.equalsIgnoreCase(expSaving)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 25, 8, resPass);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 25, 8, resFail);
				
				return false;
			}
		}
		
			
			
				
				
			
				
				
		// Sales Invoice
		
		public boolean checkSavingSalesInvoiceVATVoucher1ConsumingAllAvailableQuantityInPurchaseVoucherVAT1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
		
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			System.out.println("Customer Account List Size: "+customerAccountCount);
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 30, 5)))
				{
					customerAccountComboList.get(i).click();
					
					customerAccountTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 31, 5)))
				{
					departmentComboList.get(i).click();
					
					departmentTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 32, 5));
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 33, 5)))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 34, 5))) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys(excelReader.getCellData(xlSheetName, 35, 5));
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys(excelReader.getCellData(xlSheetName, 36, 5));
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 37, 5));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
			select1stRow_21stColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
			searchRMASelectAllChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA=enter_RMA.getAttribute("value");
			String expRMA= excelReader.getCellData(xlSheetName, 38, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 38, 7, actRMA);

			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String actSaving= Boolean.toString(checkBackgroundSavingMessage(docno));
			String expSaving= excelReader.getCellData(xlSheetName, 39, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 39, 7, actSaving);
			

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  
			System.out.println("****** RMA Column                 : "+actRMA			+"  value expected  "+expRMA);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actSaving.equalsIgnoreCase(expSaving) )
			{		
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				 excelReader.setCellData(xlfile, xlSheetName, 29, 8, resPass);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 29, 8, resFail);
				
				return false;
			}
		}
		
				
			
				
			
		
		
		// Erase all Transactions after RMA Always Numeric
		
		public boolean checkEraseAllTransactionsAfterCompletingRMAAlwaysNumeric() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			 
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
				
				String actalert=getAlert().getText();
				System.out.println(" ACT ALert Displayed : "+actalert);
				
				getAlert().accept();
			}
			
			String expValidationMsg= excelReader.getCellData(xlSheetName, 41, 6);
			
			String actValidationMsg=checkValidationMessage(expValidationMsg);
			  excelReader.setCellData(xlfile, xlSheetName, 41, 7, actValidationMsg);
			
			System.out.println("*************************************checkEraseAllTransactionsAfterCompletingTransactionFlow*********************************");
			System.out.println("Message  :  "+actValidationMsg +" Value Expected : "+expValidationMsg);
			
			if (actValidationMsg.equalsIgnoreCase(expValidationMsg))
			{
				excelReader.setCellData(xlfile, xlSheetName, 40, 8, resPass);
				return true;
			}
			else
			{
				excelReader.setCellData(xlfile, xlSheetName, 40, 8, resFail);
				return false;
			}
		}
		
		
			
				
				
				
			
		// RMA numbers are Unique across all items
			
		public boolean checkConfigureRMAUniqueAcrossAllItemsChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsmenuBtn));
			SettingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigureTransactionBtn));
			ConfigureTransactionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMABtn));
			RMABtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMANumberAlwaysnumericChkBox));
			
			if (RMANumberAlwaysnumericChkBoxIsSelected.isSelected()) 
			{
				RMANumberAlwaysnumericChkBox.click();
			}
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMANumberUniqueAcrossAllItemsChkBox));
			
			if (RMANumberUniqueAcrossAllItemsChkBox.isSelected()==false) 
			{
				RMANumberUniqueAcrossAllItemsChkBox.click();
			}
			
			Thread.sleep(2000);
			
			String actRMAOptionIsSelected = Boolean.toString(RMANumberUniqueAcrossAllItemsChkBoxIsSelected.isSelected());
			String expRMAOptionIsSelected = excelReader.getCellData(xlSheetName, 43, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 43, 7, actRMAOptionIsSelected);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			
			getWaitForAlert();
			String actAlert=getAlert().getText();
			
			String expAlert= excelReader.getCellData(xlSheetName, 44, 6);
			excelReader.setCellData(xlfile, xlSheetName, 44, 7, actAlert);
			
			getAlert().accept();
			
			String expMessage= excelReader.getCellData(xlSheetName, 45, 6);
			
			String actMessage=checkValidationMessage(expMessage);
			excelReader.setCellData(xlfile, xlSheetName, 45, 7, actMessage);
			
			
			System.out.println("RMAOptionIsSelected : "+actRMAOptionIsSelected+"  value expected  "+expRMAOptionIsSelected);
			System.out.println("Alert               : "+actAlert              +"  value expected  "+expAlert);
			System.out.println("Error Message       : "+actMessage            +"  value expected  "+expMessage);
					
			if(actRMAOptionIsSelected.equalsIgnoreCase(expRMAOptionIsSelected) && actAlert.equalsIgnoreCase(expAlert) && actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 42, 8, resPass);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 42, 8, resFail);
				
				return false;
			}
		}
					
			
			
			
			
			
			

		// Logout and Login After Selecting RMA numbers are unique across all items
			
		public boolean checkLogoutAndLoginAfterRMAUniqueAcrossAllItems() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
			userNameDisplayLogo.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			logoutOption.click();

			Thread.sleep(2000);

			System.out.println("Logout  Successfully");

			String act= Boolean.toString(checkLogin());
			String exp= excelReader.getCellData(xlSheetName, 47, 6);
			excelReader.setCellData(xlfile, xlSheetName, 47, 7, act);

			System.out.println(" Login Comany ACT"+act);
			System.out.println(" Login Comany EXP"+exp);

			if(act.equalsIgnoreCase(exp) ) 
			{
				 excelReader.setCellData(xlfile, xlSheetName, 46, 8, resPass);
				return true;
			}
			else
			{
				 excelReader.setCellData(xlfile, xlSheetName, 46, 8, resFail);
				return false;
			}
		}
	
		
		
	
		
	
	
		// Purchase Voucher 
	
		public boolean checkSavingPurchaseVoucherVATVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
				
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+vendorAccount);
				
				if (vendorAccount.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 49, 5)))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 50, 5)))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();
			
			enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
			int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
			
			for (int i = 0; i < pvvGridWarehouseListCount; i++) 
			{
				String warehouse=pvvGridWarehouseList.get(i).getText();
				if (warehouse.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 51, 5))) 
				{
					pvvGridWarehouseList.get(i).click();
					break;
				}
			}
	        enter_WarehouseTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
	        int pvvGridItemListCount=pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) 
			{
				String Item=pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 52, 5))) 
				{
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxCode));
	        enter_PvTaxCode.sendKeys(Keys.TAB);
				
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	        enter_AccountTxt.click();
			
	        Thread.sleep(2000);
	        
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			enter_Quantity.sendKeys( excelReader.getCellData(xlSheetName, 53, 5));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys( excelReader.getCellData(xlSheetName, 54, 5));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
			select1stRow_19thColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys( excelReader.getCellData(xlSheetName, 55, 5));
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMAColumn=enter_RMA.getAttribute("value");
			String expRMAColumn=  excelReader.getCellData(xlSheetName, 56, 6);
			excelReader.setCellData(xlfile, xlSheetName, 56, 7, actRMAColumn);
					
			String docno=documentNumberTxt.getAttribute("value");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String actSaving= Boolean.toString(checkBackgroundSavingMessage(docno));
			String expSaving= excelReader.getCellData(xlSheetName, 57, 6);
			excelReader.setCellData(xlfile, xlSheetName, 57, 7, actSaving);

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  
			Thread.sleep(2000);
			
			System.out.println("RMA Column     : "+actRMAColumn +"  value expected  "+expRMAColumn);
			
			if(actRMAColumn.equalsIgnoreCase(expRMAColumn) && actSaving.equalsIgnoreCase(expSaving)) 
			{
				 excelReader.setCellData(xlfile, xlSheetName, 48, 8, resPass);
				return true;
			}
			else
			{
				 excelReader.setCellData(xlfile, xlSheetName, 48, 8, resFail);
				return false;
			}
		}
			
		
		
		@FindBy(xpath="//*[@id='idGlobalError']/div/div[2]")
			
		public static WebElement rmaAlertTxt;
		
		
		public boolean checkSavingPurchaseVoucherVATVoucher2WithRMAUnique() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
				
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+vendorAccount);
				
				if (vendorAccount.equalsIgnoreCase(  excelReader.getCellData(xlSheetName, 59, 5)))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase(  excelReader.getCellData(xlSheetName, 60, 5)))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();
			
			enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
			int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
			
			for (int i = 0; i < pvvGridWarehouseListCount; i++) 
			{
				String warehouse=pvvGridWarehouseList.get(i).getText();
				if (warehouse.equalsIgnoreCase(  excelReader.getCellData(xlSheetName, 61, 5))) 
				{
					pvvGridWarehouseList.get(i).click();
					break;
				}
			}
	        enter_WarehouseTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
	        int pvvGridItemListCount=pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) 
			{
				String Item=pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase(  excelReader.getCellData(xlSheetName, 62, 5))) 
				{
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxCode));
	        enter_PvTaxCode.sendKeys(Keys.TAB);
				
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	        enter_AccountTxt.click();
			
	        Thread.sleep(2000);
	        
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			enter_Quantity.sendKeys(  excelReader.getCellData(xlSheetName, 63, 5));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys(  excelReader.getCellData(xlSheetName, 64, 5));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
			select1stRow_19thColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys(  excelReader.getCellData(xlSheetName, 65, 5));
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			Thread.sleep(2000);
			
			//getWaitForAlert();
			//String actAlert= getAlert().getText();
			String actAlert = rmaAlertTxt.getText();
			
			String expAlert=   excelReader.getCellData(xlSheetName, 66, 6);
			  excelReader.setCellData(xlfile, xlSheetName, 66, 7, actAlert);
			
			
			System.out.println("****************************** Alert Text: "+actAlert);
			
			//getAlert().accept();
		
			if(actAlert.equalsIgnoreCase(expAlert) )
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				getWaitForAlert();
				getAlert().accept();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 58, 8, resPass);

				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				getWaitForAlert();
				getAlert().accept();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 58, 8, resFail);

				
				return false;
			}
		}
			
		
		
			
			
			
			
		
		// Sales Invoice Voucher 
	
		public boolean checkSavingSalesInvoiceVATVoucher1ConsumingAllAvailableQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
		
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			System.out.println("Customer Account List Size: "+customerAccountCount);
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 68, 5)))
				{
					customerAccountComboList.get(i).click();
					
					customerAccountTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 69, 5)))
				{
					departmentComboList.get(i).click();
					
					departmentTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 70, 5));
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 71, 5)))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 72, 5))) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys(excelReader.getCellData(xlSheetName, 73, 5));
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys(excelReader.getCellData(xlSheetName, 74, 5));
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 75, 5));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
			select1stRow_21stColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
			searchRMASelectAllChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA= enter_RMA.getAttribute("value");
			String expRMA= excelReader.getCellData(xlSheetName, 76, 6);
			excelReader.setCellData(xlfile, xlSheetName, 76, 7, actRMA);
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String actSaving= Boolean.toString(checkBackgroundSavingMessage(docno));
			String expSaving= excelReader.getCellData(xlSheetName, 77, 6);
			excelReader.setCellData(xlfile, xlSheetName, 77, 7, actRMA);

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  
			System.out.println("****** RMA Column                 : "+actRMA			+"  value expected  "+expRMA);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actSaving.equalsIgnoreCase(expSaving)) 
			{		
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				  excelReader.setCellData(xlfile, xlSheetName, 67, 8, resPass);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				  excelReader.setCellData(xlfile, xlSheetName, 67, 8, resFail);
				
				return false;
			}
		}
		
		
			
			
		
		
		// Erase all Transactions after RMA numbers are unique across all items
		
		public boolean checkEraseAllTransactionsAfterCompletingRMANumbersAreUniqueAcrossAllItems() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());

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
			
			String expValidationMsg=  excelReader.getCellData(xlSheetName, 79, 6);
			
			String actValidationMsg=checkValidationMessage(expValidationMsg);
			excelReader.setCellData(xlfile, xlSheetName, 79, 7, actValidationMsg);
			
			
			System.out.println("Message  :  "+actValidationMsg +" Value Expected : "+expValidationMsg);
			
			if (actValidationMsg.equalsIgnoreCase(expValidationMsg))
			{
				 excelReader.setCellData(xlfile, xlSheetName, 78, 8, resPass);
				return true;
			}
			else
			{
				 excelReader.setCellData(xlfile, xlSheetName, 78, 8, resFail);
				return false;
			}
		}
		
		
		
			
			
			
			
		// Input RMA Numbers before Quantity
			
		public boolean checkInputRMABeforeQuantityChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			  excelReader=new ExcelReader(POJOUtility.getExcelPath());

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsmenuBtn));
			SettingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigureTransactionBtn));
			ConfigureTransactionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMABtn));
			RMABtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMANumberUniqueAcrossAllItemsChkBox));
			
			if (RMANumberUniqueAcrossAllItemsChkBoxIsSelected.isSelected()) 
			{
				RMANumberUniqueAcrossAllItemsChkBox.click();
			}
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InputRMANoBeforeQuantityChkBox));
			
			if (InputRMANoBeforeQuantityChkBoxIsSelected.isSelected()==false) 
			{
				InputRMANoBeforeQuantityChkBox.click();
			}
			
			Thread.sleep(2000);
			
			String  actRMAOptionIsSelected = Boolean.toString(InputRMANoBeforeQuantityChkBoxIsSelected.isSelected());
			String  expRMAOptionIsSelected =   excelReader.getCellData(xlSheetName, 81, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 81, 7, actRMAOptionIsSelected);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			
			getWaitForAlert();
			String actAlert=getAlert().getText();
			
			String expAlert= excelReader.getCellData(xlSheetName, 82, 6);
			excelReader.setCellData(xlfile, xlSheetName, 82, 7, actAlert);
			
			getAlert().accept();
			
			String expMessage= excelReader.getCellData(xlSheetName, 83, 6);
			
			
			String actMessage=checkValidationMessage(expMessage);
			excelReader.setCellData(xlfile, xlSheetName, 83, 7, actMessage);
			
			
			System.out.println("RMAOptionIsSelected : "+actRMAOptionIsSelected+"  value expected  "+expRMAOptionIsSelected);
			System.out.println("Alert               : "+actAlert              +"  value expected  "+expAlert);
			System.out.println("Error Message       : "+actMessage            +"  value expected  "+expMessage);
					
			if(actRMAOptionIsSelected.equalsIgnoreCase(expRMAOptionIsSelected) && actAlert.equalsIgnoreCase(expAlert) && actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				excelReader.setCellData(xlfile, xlSheetName, 80, 8, resPass);

				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				excelReader.setCellData(xlfile, xlSheetName, 80, 8, resFail);

				
				return false;
			}
		}
		
			
			

		// Logout and Login After Selecting Input RMA number before quantity
			
		public boolean checkLogoutAndLoginAfterRMABeforeQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
			userNameDisplayLogo.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			logoutOption.click();

			Thread.sleep(2000);

			System.out.println("Logout  Successfully");

			String act= Boolean.toString(checkLogin());
			String exp= excelReader.getCellData(xlSheetName, 85, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 85, 7, act);


			System.out.println(" Login Comany ACT"+act);
			System.out.println(" Login Comany EXP"+exp);

			if(act.equalsIgnoreCase(exp))  
			{
				excelReader.setCellData(xlfile, xlSheetName, 84, 8, resPass);

				return true;
			}
			else
			{
				excelReader.setCellData(xlfile, xlSheetName, 84, 8, resFail);

				return false;
			}
		}
		
			
		// Purchase Voucher
		
		public boolean checkSavingPurchaseVoucherVATVoucher1RMABeforeQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String ScreenOpenedMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
				
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+vendorAccount);
				
				if (vendorAccount.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 87, 5)))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 88, 5)))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();
			
			enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
			int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
			
			for (int i = 0; i < pvvGridWarehouseListCount; i++) 
			{
				String warehouse=pvvGridWarehouseList.get(i).getText();
				if (warehouse.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 89, 5))) 
				{
					pvvGridWarehouseList.get(i).click();
					break;
				}
			}
	        enter_WarehouseTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
	        int pvvGridItemListCount=pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) 
			{
				String Item=pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 90, 5))) 
				{
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxCode));
	        enter_PvTaxCode.sendKeys(Keys.TAB);
				
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	        enter_AccountTxt.click();
			
	        Thread.sleep(2000);
		
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
	        select1stRow_9thColumn.click();
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 91, 5));
			RMAPopupSerialNumberTxtField.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupQuantityTxtField));
			RMAPopupQuantityTxtField.sendKeys(excelReader.getCellData(xlSheetName, 92, 5));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMAColumn= enter_RMA.getAttribute("value");
			String expRMAColumn= excelReader.getCellData(xlSheetName, 93, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 93, 7, actRMAColumn);
			
			enter_RMA.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_10thColumn));
			String actQty= select1stRow_10thColumn.getText();
			String expQty= excelReader.getCellData(xlSheetName, 94, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 94, 7, actQty);
					
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 95, 5));
			Thread.sleep(2000);
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String actSaving= Boolean.toString(checkBackgroundSavingMessage(docno));
			String expSaving=excelReader.getCellData(xlSheetName, 96, 6);
			excelReader.setCellData(xlfile, xlSheetName, 96, 7, actSaving);

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			
			System.out.println("RMA          : "+actRMAColumn +"  value expected  "+expRMAColumn);
			System.out.println("Quantity     : "+actQty       +"  value expected  "+expQty);
			
			if(actRMAColumn.equalsIgnoreCase(expRMAColumn) && actQty.equalsIgnoreCase(expQty) && actSaving.equalsIgnoreCase(expSaving) )
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				 excelReader.setCellData(xlfile, xlSheetName, 86, 8, resPass);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				 excelReader.setCellData(xlfile, xlSheetName, 86, 8, resFail);
				return false;
			}
		}
				
			
			
			
		
		
		
		
		// Sales Invoice Voucher 
				
		public boolean checkSavingSalesInvoiceVATVoucher1ConsumingAllAvailableQuantityWithRMABeforeQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());

			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
		
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			System.out.println("Customer Account List Size: "+customerAccountCount);
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 98, 5)))
				{
					customerAccountComboList.get(i).click();
					
					customerAccountTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 99, 5)))
				{
					departmentComboList.get(i).click();
					
					departmentTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 100, 5));
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 101, 5)))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 102, 5))) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys(excelReader.getCellData(xlSheetName, 103, 5));
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys(excelReader.getCellData(xlSheetName, 104, 5));
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RD));
			enter_RD.sendKeys(Keys.TAB);
					
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
			searchRMASelectAllChkBox.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
					
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA  = enter_RMA.getAttribute("value");
			String expRMA  =  excelReader.getCellData(xlSheetName, 105, 6);
			excelReader.setCellData(xlfile, xlSheetName, 105, 7, actRMA);
			enter_RMA.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
			String actQty  = select1stRow_12thColumn.getText();
			String expQty  =  excelReader.getCellData(xlSheetName, 106, 6);
			excelReader.setCellData(xlfile, xlSheetName, 106, 7, actQty);
					
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_15thColumn));
			select1stRow_15thColumn.click();
			enter_Rate.sendKeys( excelReader.getCellData(xlSheetName, 107, 6));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(2000);
					
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String actSaving  =  Boolean.toString(checkBackgroundSavingMessage(docno));
			String expSaving  =  excelReader.getCellData(xlSheetName, 108, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 108, 7, actSaving);


			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  

			System.out.println("RMA Column   : "+actRMA			+"  value expected  "+expRMA);
			System.out.println("Quantity     : "+actQty			+"  value expected  "+expQty);
			
			
			if(actRMA.equalsIgnoreCase(expRMA) && actQty.equalsIgnoreCase(expQty) && actSaving.equalsIgnoreCase(expSaving)) 
				
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 97, 8, resPass);

				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 97, 8, resFail);

				
				return false;
			}
		}
				
		
		
		
		
		
		
		// Erase all Transactions after RMA number Before Quantity
		
		public boolean checkEraseAllTransactionsAfterCompletingRMANumberBeforeQuantity() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());

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
			
			String expValidationMsg =   excelReader.getCellData(xlSheetName, 110, 6);
			
			String actValidationMsg=checkValidationMessage(expValidationMsg);
			 excelReader.setCellData(xlfile, xlSheetName, 110, 7, actValidationMsg);
			
			System.out.println("Message  :  "+actValidationMsg +" Value Expected : "+expValidationMsg);
			
			if (actValidationMsg.equalsIgnoreCase(expValidationMsg))
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
		
		
		
					
					
					
			
				
				
			
		 // RMA number Cannot be reuse
		
		public boolean checkInputRMANoCannotBeReuseChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			  excelReader=new ExcelReader(POJOUtility.getExcelPath());
			  
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsmenuBtn));
			SettingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigureTransactionBtn));
			ConfigureTransactionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMABtn));
			RMABtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InputRMANoBeforeQuantityChkBox));
			
			if (InputRMANoBeforeQuantityChkBoxIsSelected.isSelected()) 
			{
				InputRMANoBeforeQuantityChkBox.click();
			}
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMANumberCantBeUsedChkBox));
			
			if (RMANumberCantBeUsedChkBoxIsSelected.isSelected()==false) 
			{
				RMANumberCantBeUsedChkBox.click();
			}
			
			Thread.sleep(2000);
			
			String  actRMAOptionIsSelected = Boolean.toString(RMANumberCantBeUsedChkBoxIsSelected.isSelected());
			String  expRMAOptionIsSelected = excelReader.getCellData(xlSheetName, 112, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 112, 7, actRMAOptionIsSelected);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			
			getWaitForAlert();
			
			String actAlert=getAlert().getText();
			
			String expAlert = excelReader.getCellData(xlSheetName, 113, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 113, 7, actAlert);
			
			getAlert().accept();
			
			String expMessage = excelReader.getCellData(xlSheetName, 114, 6);
			
			String actMessage=checkValidationMessage(expMessage);
			 excelReader.setCellData(xlfile, xlSheetName, 114, 7, actMessage);
			
			
			
			System.out.println("RMAOptionIsSelected : "+actRMAOptionIsSelected+"  value expected  "+expRMAOptionIsSelected);
			System.out.println("Alert               : "+actAlert              +"  value expected  "+expAlert);
			System.out.println("Error Message       : "+actMessage            +"  value expected  "+expMessage);
					
			if(actRMAOptionIsSelected.equalsIgnoreCase(expRMAOptionIsSelected) && actAlert.equalsIgnoreCase(expAlert) && actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 111, 8, resPass);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 111, 8, resFail);
				
				return false;
			}
		}
		
				
		
		
		
		// Login and Logout	
				
		public boolean checkLogoutAndLoginAfterRMACantReuse() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
			userNameDisplayLogo.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			logoutOption.click();

			Thread.sleep(2000);

			System.out.println("Logout  Successfully");

			String act=Boolean.toString(checkLogin());
			String  exp =   excelReader.getCellData(xlSheetName, 116, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 116, 7, act);


			System.out.println(" Login Comany ACT"+act);
			System.out.println(" Login Comany EXP"+exp);

			if(act.equalsIgnoreCase(exp)  )
				
			{
				 excelReader.setCellData(xlfile, xlSheetName, 115, 8, resPass);
				return true;
			}
			else
			{
				 excelReader.setCellData(xlfile, xlSheetName, 115, 8, resFail);
				return false;
			}
		}

		
		
		
		
		
		
		
		// Purchase Voucher 
		
		public boolean checkSavingPurchaseVoucherVATVoucher1WithRMACantReuse() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			
			   excelReader=new ExcelReader(POJOUtility.getExcelPath());
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
				
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+vendorAccount);
				
				if (vendorAccount.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 118, 5)))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 119, 5)))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();
			
			enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
			int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
			
			for (int i = 0; i < pvvGridWarehouseListCount; i++) 
			{
				String warehouse=pvvGridWarehouseList.get(i).getText();
				if (warehouse.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 120, 5))) 
				{
					pvvGridWarehouseList.get(i).click();
					break;
				}
			}
	        enter_WarehouseTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
	        int pvvGridItemListCount=pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) 
			{
				String Item=pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 121, 5))) 
				{
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxCode));
	        enter_PvTaxCode.sendKeys(Keys.TAB);
				
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	        enter_AccountTxt.click();
			
	        Thread.sleep(2000);
	        
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			enter_Quantity.sendKeys( excelReader.getCellData(xlSheetName, 122, 5));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys( excelReader.getCellData(xlSheetName, 123, 5));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
			select1stRow_19thColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys( excelReader.getCellData(xlSheetName, 124, 5));
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMAColumn=enter_RMA.getAttribute("value");
			String expRMAColumn=   excelReader.getCellData(xlSheetName, 125, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 125, 7, actRMAColumn);
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String actSaving = Boolean.toString(checkBackgroundSavingMessage(docno));
			String  expSaving=  excelReader.getCellData(xlSheetName, 126, 6);

			 excelReader.setCellData(xlfile, xlSheetName, 126, 7, actSaving);
			 
			 
			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			
			
			Thread.sleep(2000);
			
			System.out.println("RMA Column     : "+actRMAColumn +"  value expected  "+expRMAColumn);
			
			if(actRMAColumn.equalsIgnoreCase(expRMAColumn) && actSaving.equalsIgnoreCase(expSaving)) 
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
			
		
		
				
				
				
			
		// Sales Invoice Voucher 
	
		public boolean checkSavingSalesInvoiceVATVoucher1ConsumingAllAvailableQuantityWithRMACantReuse() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			
			  excelReader=new ExcelReader(POJOUtility.getExcelPath());
			  
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
		
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			System.out.println("Customer Account List Size: "+customerAccountCount);
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 128, 5)))
				{
					customerAccountComboList.get(i).click();
					
					customerAccountTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 129, 5)))
				{
					departmentComboList.get(i).click();
					
					departmentTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys( excelReader.getCellData(xlSheetName, 130, 5));
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 131, 5)))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 132, 5))) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys( excelReader.getCellData(xlSheetName, 133, 5));
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys( excelReader.getCellData(xlSheetName, 134, 5));
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys( excelReader.getCellData(xlSheetName, 135, 5));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
			select1stRow_21stColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
			searchRMASelectAllChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA = enter_RMA.getAttribute("value");
			String expRMA =   excelReader.getCellData(xlSheetName, 136, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 136, 7, actRMA);
			
			String docno=documentNumberTxt.getAttribute("value");
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String actSaving = Boolean.toString(checkBackgroundSavingMessage(docno));
			String  expSaving= excelReader.getCellData(xlSheetName, 137, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 137, 7, actSaving);

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			
			
			System.out.println("****** RMA Column                 : "+actRMA			+"  value expected  "+expRMA);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actSaving.equalsIgnoreCase(expSaving) )
			{		
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				excelReader.setCellData(xlfile, xlSheetName, 127, 8, resPass);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				excelReader.setCellData(xlfile, xlSheetName, 127, 8, resFail);
				
				return false;
			}
		}
			
		
		
		
		
		
		
		public boolean checkSavingPurchaseVoucherVATVoucher2WithAlreadyUsedRMA() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());

			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
				
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+vendorAccount);
				
				if (vendorAccount.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 139, 5)))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 140, 5)))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();
			
			enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
			int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
			
			for (int i = 0; i < pvvGridWarehouseListCount; i++) 
			{
				String warehouse=pvvGridWarehouseList.get(i).getText();
				if (warehouse.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 141, 5))) 
				{
					pvvGridWarehouseList.get(i).click();
					break;
				}
			}
	        enter_WarehouseTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
	        int pvvGridItemListCount=pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) 
			{
				String Item=pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 142, 5))) 
				{
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxCode));
	        enter_PvTaxCode.sendKeys(Keys.TAB);
				
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	        enter_AccountTxt.click();
			
	        Thread.sleep(2000);
	        
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 143, 5));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 144, 5));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
			select1stRow_19thColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 145, 5));
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			Thread.sleep(2000);
			
			String actAlert = rmaAlertTxt.getText();
			System.err.println(actAlert);
			
			String expAlert= excelReader.getCellData(xlSheetName, 146, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 146, 7, actAlert);
			 
			/*getWaitForAlert();
			String actAlert=getAlert().getText();
			String expAlert= excelReader.getCellData(xlSheetName, 146, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 146, 7, actAlert);
			
			System.out.println("****************************** Alert Text: "+actAlert);
			
			getAlert().accept();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupCancelBtn));
			RMAPopupCancelBtn.click();
			
			Thread.sleep(2000);
			*/
			if(actAlert.equalsIgnoreCase(expAlert))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				getWaitForAlert();
				getAlert().accept();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 138, 8, resPass);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				getWaitForAlert();
				getAlert().accept();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 138, 8, resFail);
				
				return false;
			}
		}
			
			
			
			
			
				
		
		
		// Erase all Transactions after RMA Cant Reuse
		
		public boolean checkEraseAllTransactionsAfterCompletingRMACantReuse() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
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
			
			String expValidationMsg= excelReader.getCellData(xlSheetName, 148, 6);
			
			String actValidationMsg=checkValidationMessage(expValidationMsg);
			excelReader.setCellData(xlfile, xlSheetName, 148, 7, actValidationMsg);
			
			System.out.println("Message  :  "+actValidationMsg +" Value Expected : "+expValidationMsg);
			
			if (actValidationMsg.equalsIgnoreCase(expValidationMsg))
			{
				 excelReader.setCellData(xlfile, xlSheetName, 147, 8, resPass);

				return true;
			}
			else
			{
				 excelReader.setCellData(xlfile, xlSheetName, 147, 8, resFail);

				return false;
			}
		}
		
		
		
		
		
			
		// Prefix voucher number to the serial number in inwards
		
		public boolean checkPrefixVoucherNumberToTheSerialNunberInInwardsChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsmenuBtn));
			SettingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigureTransactionBtn));
			ConfigureTransactionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMABtn));
			RMABtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMANumberCantBeUsedChkBox));
			
			if (RMANumberCantBeUsedChkBoxIsSelected.isSelected()) 
			{
				RMANumberCantBeUsedChkBox.click();
			}
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PrefixVoucherNoChkBox));
			
			if (PrefixVoucherNoChkBoxIsSelected.isSelected()==false) 
			{
				PrefixVoucherNoChkBox.click();
			}
			
			Thread.sleep(2000);
			
			String actRMAOptionIsSelected = Boolean.toString(PrefixVoucherNoChkBoxIsSelected.isSelected());
			String  expRMAOptionIsSelected =   excelReader.getCellData(xlSheetName, 150, 6);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			
			getWaitForAlert();
			String actAlert=getAlert().getText();
			
			String expAlert=   excelReader.getCellData(xlSheetName, 151, 6);
			   excelReader.setCellData(xlfile, xlSheetName, 151, 7, actAlert);
			getAlert().accept();
			
			String expMessage=   excelReader.getCellData(xlSheetName, 152, 6);
			
			String actMessage=checkValidationMessage(expMessage);
			   excelReader.setCellData(xlfile, xlSheetName, 152, 7, actMessage);
			
			System.out.println("RMAOptionIsSelected : "+actRMAOptionIsSelected+"  value expected  "+expRMAOptionIsSelected);
			System.out.println("Alert               : "+actAlert              +"  value expected  "+expAlert);
			System.out.println("Error Message       : "+actMessage            +"  value expected  "+expMessage);
					
			if(actRMAOptionIsSelected.equalsIgnoreCase(expRMAOptionIsSelected) && actAlert.equalsIgnoreCase(expAlert) && actMessage.equalsIgnoreCase(expMessage)) 
			{
				 excelReader.setCellData(xlfile, xlSheetName, 149, 8, resPass);
				return true;
			}
			else
			{
				 excelReader.setCellData(xlfile, xlSheetName, 149, 8, resFail);
				return false;
			}
		}
			
					
					
				
				
		
				
		
		// Logout and Login After Prefix voucher number to the serial number in inwards
			
		public boolean checkLogoutAndLoginAfterRMAPrefixVoucherNo() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());

		 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
			userNameDisplayLogo.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			logoutOption.click();

			Thread.sleep(2000);

			System.out.println("Logout  Successfully");

			String  act= Boolean.toString(checkLogin());
			String  exp=  excelReader.getCellData(xlSheetName, 154, 6);
            excelReader.setCellData(xlfile, xlSheetName, 154, 7, act);


			System.out.println(" Login Comany ACT"+act);
			System.out.println(" Login Comany EXP"+exp);

			if(act.equalsIgnoreCase(exp))  
			{
				excelReader.setCellData(xlfile, xlSheetName, 153, 8, resPass);
				return true;
			}
			else
			{
				excelReader.setCellData(xlfile, xlSheetName, 153, 8, resFail);
				return false;
			}
		}
				
					
			
		
		
		
		
		
		
		// Purchase Voucher 
		
		public boolean checkSavingPurchaseVoucherVATVoucher1RMAPrefixVoucherNo() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
				
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+vendorAccount);
				
				if (vendorAccount.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 156, 5)))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 157, 5)))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();
			
			enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
			int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
			
			for (int i = 0; i < pvvGridWarehouseListCount; i++) 
			{
				String warehouse=pvvGridWarehouseList.get(i).getText();
				if (warehouse.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 158, 5))) 
				{
					pvvGridWarehouseList.get(i).click();
					break;
				}
			}
	        enter_WarehouseTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
	        int pvvGridItemListCount=pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) 
			{
				String Item=pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 159, 5))) 
				{
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxCode));
	        enter_PvTaxCode.sendKeys(Keys.TAB);
				
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	        enter_AccountTxt.click();
			
	        Thread.sleep(2000);
	        
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			enter_Quantity.sendKeys( excelReader.getCellData(xlSheetName, 160, 5));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys( excelReader.getCellData(xlSheetName, 161, 5));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
			select1stRow_19thColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys( excelReader.getCellData(xlSheetName, 162, 5));
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			ArrayList<String> actRMA= new ArrayList<>();
	        for (int i = 1; i < 11; i++) 
	        {
	        	WebElement a=getDriver().findElement(By.xpath("//tbody[@id='RMA_Table_body']/tr["+i+"]/td[2]"));
	        	actRMA.add(a.getAttribute("data-value"));
	        }
			
	        ArrayList<String>  expRMA= new ArrayList<String>(); 
	         
			expRMA.add( excelReader.getCellData(xlSheetName, 163, 5));
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMAColumn=enter_RMA.getAttribute("value");
			String expRMAColumn=  excelReader.getCellData(xlSheetName, 164, 6);
			excelReader.setCellData(xlfile, xlSheetName, 164, 7, actRMAColumn);
					
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String actSaving= Boolean.toString(checkBackgroundSavingMessage(docno));
			String  expSaving=  excelReader.getCellData(xlSheetName, 165, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 165, 7, actSaving);

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			
			Thread.sleep(2000);
			
			System.out.println(actRMA);
			System.out.println(expRMA);
			System.out.println("RMA Column     : "+actRMAColumn +"  value expected  "+expRMAColumn);
			
			if(actRMA.toString().equalsIgnoreCase(expRMA.toString()) && actSaving.equalsIgnoreCase(expSaving)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				Thread.sleep(2000);
				excelReader.setCellData(xlfile, xlSheetName, 155, 8, resPass);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				Thread.sleep(2000);
				excelReader.setCellData(xlfile, xlSheetName, 155, 8, resFail);
				return false;
			}
		}
			
			
			
				
			
		
				
		// Sales Invoice

		public boolean checkSavingSalesInvoiceVATVoucher1ConsumeAllAvailableQuantityRMAPrefixVoucherNo() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
		
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			System.out.println("Customer Account List Size: "+customerAccountCount);
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 167, 5)))
				{
					customerAccountComboList.get(i).click();
					
					customerAccountTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 168, 5)))
				{
					departmentComboList.get(i).click();
					
					departmentTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys(excelReader.getCellData(xlSheetName, 169, 5));
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 170, 5)))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 171, 5))) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys(excelReader.getCellData(xlSheetName, 172, 5));
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys(excelReader.getCellData(xlSheetName, 173, 5));
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 174, 5));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
			select1stRow_21stColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
			searchRMASelectAllChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA=enter_RMA.getAttribute("value");
			String expRMA=excelReader.getCellData(xlSheetName, 175, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 175, 7, actRMA);
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String actSaving = Boolean.toString(checkBackgroundSavingMessage(docno));
			String  expSaving = excelReader.getCellData(xlSheetName, 176, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 176, 7, actSaving);

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			
			System.out.println("****** RMA Column                 : "+actRMA			+"  value expected  "+expRMA);
			
			if(actRMA.equalsIgnoreCase(expRMA) &&actSaving.equalsIgnoreCase(expSaving) )
			{		
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 166, 8, resPass);

				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 166, 8, resFail);

				
				return false;
			}
		}
				
					
					
		
		
		
		
		
		// Erase all Transactions after Prefix voucher number to the serial number in inwards
		
		public boolean checkEraseAllTransactionsAfterCompletingPrefixVoucherNumberToTheSerialNumberInInwards() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			 
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
			
			String expValidationMsg= excelReader.getCellData(xlSheetName, 178, 6);
			
			String actValidationMsg=checkValidationMessage(expValidationMsg);
			
			 excelReader.setCellData(xlfile, xlSheetName, 178, 7, actValidationMsg);
			
			System.out.println("Message  :  "+actValidationMsg +" Value Expected : "+expValidationMsg);
			
			if (actValidationMsg.equalsIgnoreCase(expValidationMsg))
			{
				 excelReader.setCellData(xlfile, xlSheetName, 177, 8, resPass);
				return true;
			}
			else
			{
				 excelReader.setCellData(xlfile, xlSheetName, 177, 8, resFail);
				return false;
			}
		}
		
		
				
				
				
				
		// RMA based alternative quantity
			
		public boolean checkRMABasedAlternateQuantityChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsmenuBtn));
			SettingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigureTransactionBtn));
			ConfigureTransactionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMABtn));
			RMABtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PrefixVoucherNoChkBox));
			
			if (PrefixVoucherNoChkBoxIsSelected.isSelected()) 
			{
				PrefixVoucherNoChkBox.click();
			}
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAbasedalternateQuantityChkBox));
			
			if (RMAbasedalternateQuantityChkBoxIsSelected.isSelected()==false) 
			{
				RMAbasedalternateQuantityChkBox.click();
			}
			
			Thread.sleep(2000);
			
			String actRMAOptionIsSelected = Boolean.toString(RMAbasedalternateQuantityChkBoxIsSelected.isSelected());
			String expRMAOptionIsSelected = excelReader.getCellData(xlSheetName, 180, 6);
            excelReader.setCellData(xlfile, xlSheetName, 180, 7, actRMAOptionIsSelected);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			
			getWaitForAlert();
			String actAlert=getAlert().getText();
			
			String expAlert= excelReader.getCellData(xlSheetName, 181, 6);
            excelReader.setCellData(xlfile, xlSheetName, 181, 7, actAlert);
			
            getAlert().accept();
			
			String expMessage= excelReader.getCellData(xlSheetName, 182, 6);
			
			String actMessage=checkValidationMessage(expMessage);
			  excelReader.setCellData(xlfile, xlSheetName, 182, 7, actMessage);
			
			System.out.println("RMAOptionIsSelected : "+actRMAOptionIsSelected+"  value expected  "+expRMAOptionIsSelected);
			System.out.println("Alert               : "+actAlert              +"  value expected  "+expAlert);
			System.out.println("Error Message       : "+actMessage            +"  value expected  "+expMessage);
					
			if(actRMAOptionIsSelected.equalsIgnoreCase(expRMAOptionIsSelected) && actAlert.equalsIgnoreCase(expAlert) && actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 179, 8, resPass);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 189, 8, resFail);
				
				return false;
			}
		}
			
			
			
			
			
		
		
		// RMA based alternative quantity
		
		public boolean checkSelectingMaintainInventoryInAlternateUnitCheckBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			  excelReader=new ExcelReader(POJOUtility.getExcelPath());


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsmenuBtn));
			SettingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigureTransactionBtn));
			ConfigureTransactionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InventoryBtn));
			InventoryBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MaintainInventoryInAlternateUnitChkBox));
			
			if (MaintainInventoryInAlternateUnitChkBoxIsSelected.isSelected()==false) 
			{
				MaintainInventoryInAlternateUnitChkBox.click();
			}
			
			Thread.sleep(2000);
			
			String actMaintainInventoryInAlternateUnitChkBox = Boolean.toString(MaintainInventoryInAlternateUnitChkBoxIsSelected.isSelected());
			String  expMaintainInventoryInAlternateUnitChkBox =  excelReader.getCellData(xlSheetName, 184, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 184, 7, actMaintainInventoryInAlternateUnitChkBox);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			
			getWaitForAlert();
			String actAlert=getAlert().getText();
			String expAlert= excelReader.getCellData(xlSheetName, 185, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 185, 7, actAlert);
			getAlert().accept();
			
			String expMessage= excelReader.getCellData(xlSheetName, 186, 6);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, xlSheetName, 186, 7, actMessage);
			
			System.out.println("MaintainInventoryInAlternateUnitChkBox : "+actMaintainInventoryInAlternateUnitChkBox+"  value expected  "+expMaintainInventoryInAlternateUnitChkBox);
			System.out.println("Alert               : "+actAlert              +"  value expected  "+expAlert);
			System.out.println("Error Message       : "+actMessage            +"  value expected  "+expMessage);
					
			if(actMaintainInventoryInAlternateUnitChkBox.equalsIgnoreCase(expMaintainInventoryInAlternateUnitChkBox) && actAlert.equalsIgnoreCase(expAlert) 
					&& actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				excelReader.setCellData(xlfile, xlSheetName, 183, 8, resPass);

				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				excelReader.setCellData(xlfile, xlSheetName, 183, 8, resFail);

				return false;
			}
		}
		
		
		
		
				
				
		
		// Logout and Login 
				
		public boolean checkLogoutAndLoginAfterRMABasedAlternateQty() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
			userNameDisplayLogo.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			logoutOption.click();

			Thread.sleep(2000);

			System.out.println("Logout  Successfully");

			String act= Boolean.toString(checkLogin());
			String exp=  excelReader.getCellData(xlSheetName, 188, 6);
            
          excelReader.setCellData(xlfile, xlSheetName, 188, 7, act);


			System.out.println(" Login Comany ACT"+act);
			System.out.println(" Login Comany EXP"+exp);

			if(act.equalsIgnoreCase(exp))  
			{
				  excelReader.setCellData(xlfile, xlSheetName, 187, 8, resPass);
				return true;
			}
			else
			{
				  excelReader.setCellData(xlfile, xlSheetName, 187, 8, resPass);
				return false;
			}
		}
				
			
		
		// Purchase Voucher
		
		public boolean checkSavingPurchaseVoucherVATVoucher1WithRMAAlternateQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
				
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+vendorAccount);
				
				if (vendorAccount.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 190, 5)))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 191, 5)))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();
			
			enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
			int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
			
			for (int i = 0; i < pvvGridWarehouseListCount; i++) 
			{
				String warehouse=pvvGridWarehouseList.get(i).getText();
				if (warehouse.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 192, 5))) 
				{
					pvvGridWarehouseList.get(i).click();
					break;
				}
			}
	        enter_WarehouseTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
	        int pvvGridItemListCount=pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) 
			{
				String Item=pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 193, 5))) 
				{
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxCode));
	        enter_PvTaxCode.sendKeys(Keys.TAB);
				
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	        enter_AccountTxt.click();
			
	        Thread.sleep(2000);
	        
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			enter_Quantity.sendKeys( excelReader.getCellData(xlSheetName, 194, 5));
			enter_Quantity.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AlternateQuantity));
			enter_AlternateQuantity.sendKeys( excelReader.getCellData(xlSheetName, 195, 5));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
			select1stRow_12thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 196, 5));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_20thColumn));
			select1stRow_20thColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys( excelReader.getCellData(xlSheetName, 197, 5));
			
			Thread.sleep(2000);
			
			String actSerialNo=RMAPopupSerialNumberTxtField.getAttribute("value");
			System.out.println("********************************** Serial No Text field: "+actSerialNo);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMAColumn=enter_RMA.getAttribute("value");
			String expRMAColumn=  excelReader.getCellData(xlSheetName, 198, 6);
            
          excelReader.setCellData(xlfile, xlSheetName, 198, 7, actRMAColumn);

				
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String actSaving= Boolean.toString(checkBackgroundSavingMessage(docno));
			String expSaving=  excelReader.getCellData(xlSheetName, 199, 6);
            
          excelReader.setCellData(xlfile, xlSheetName, 199, 7, actSaving);


			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  
			
			System.out.println(actRMAColumn);
			System.out.println(expRMAColumn);
			
			if(actRMAColumn.equalsIgnoreCase(expRMAColumn) && actSaving.equalsIgnoreCase(expSaving) )
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 189, 8, resPass);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 189, 8, resFail);
				return false;
			}
		}
		
			
			
				
				
			
				
				
		// Sales Invoice
		
		public boolean checkSavingSalesInvoiceVATVoucher1WithAlternateQuantityConsuming12Quantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			  excelReader=new ExcelReader(POJOUtility.getExcelPath());

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
		
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			System.out.println("Customer Account List Size: "+customerAccountCount);
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 201, 5)))
				{
					customerAccountComboList.get(i).click();
					
					customerAccountTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 202, 5)))
				{
					departmentComboList.get(i).click();
					
					departmentTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys( excelReader.getCellData(xlSheetName, 203, 5));
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 204, 5)))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 205, 5))) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys( excelReader.getCellData(xlSheetName, 206, 5));
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys( excelReader.getCellData(xlSheetName, 207, 5));
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RD));
			enter_RD.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AlternateQuantity));
			enter_AlternateQuantity.sendKeys( excelReader.getCellData(xlSheetName, 208, 5));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_15thColumn));
			select1stRow_15thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys( excelReader.getCellData(xlSheetName, 209, 5));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_22ndColumn));
			select1stRow_22ndColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
			searchRMASelectAllChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA=enter_RMA.getAttribute("value");
			String expRMA=  excelReader.getCellData(xlSheetName, 210, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 210, 7, actRMA);

			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String actSaving= Boolean.toString(checkBackgroundSavingMessage(docno));
			String  expSaving=  excelReader.getCellData(xlSheetName, 211, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 211, 7, actSaving);


			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  

			System.out.println("****** RMA Column                 : "+actRMA			+"  value expected  "+expRMA);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actSaving.equalsIgnoreCase(expSaving) )
			{		
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 200, 8, resPass);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 200, 8, resFail);
				return false;
			}
		}
		
				
		
		
		
		
		
		public boolean checkSavingSalesInvoiceVATVoucher2WithAlternateQuantityConsuming12Quantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
		
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			System.out.println("Customer Account List Size: "+customerAccountCount);
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 213, 5)))
				{
					customerAccountComboList.get(i).click();
					
					customerAccountTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			System.out.println("Department List Size: "+departmentCount);
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				System.out.println(i+"  "+department);
				
				if (department.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 214, 5)))
				{
					departmentComboList.get(i).click();
					
					departmentTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys( excelReader.getCellData(xlSheetName, 215, 5));
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			System.out.println("Warehouse List Size: "+warehouseCount);
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				System.out.println(i+"  "+warehouse);
				
				if (warehouse.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 216, 5)))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				System.out.println(i+"  "+item);
				
				if (item.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 217, 5))) 
				{
					itemComboList.get(i).click();
					
					enter_ItemTxt.sendKeys(Keys.TAB);
					
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys( excelReader.getCellData(xlSheetName, 218, 5));
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys( excelReader.getCellData(xlSheetName, 219, 5));
			enter_FQ.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RD));
			enter_RD.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AlternateQuantity));
			enter_AlternateQuantity.sendKeys( excelReader.getCellData(xlSheetName, 220, 5));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_15thColumn));
			select1stRow_15thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys( excelReader.getCellData(xlSheetName, 221, 5));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_22ndColumn));
			select1stRow_22ndColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
			RMAPopupSerialNoExpansionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMASelectAllChkBox));
			searchRMASelectAllChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
			searchRMAOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceRMAPopupOkBtn));
			salesInvoiceRMAPopupOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			String actRMA = enter_RMA.getAttribute("value");
			String expRMA =  excelReader.getCellData(xlSheetName, 222, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 222, 7, actRMA);

			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String actSaving= Boolean.toString(checkBackgroundSavingMessage(docno));
			String expSaving= excelReader.getCellData(xlSheetName, 223, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 223, 7, actSaving);

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  
			System.out.println("****** RMA Column                 : "+actRMA			+"  value expected  "+expRMA);
			
			if(actRMA.equalsIgnoreCase(expRMA) && actSaving.equalsIgnoreCase(expSaving)) 
			{		
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 212, 8, resPass);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 212, 8, resFail);
				return false;
			}
		}
		
				
		
		
		
		
		public boolean checkUnCheckRMABasedAlternateQuantityChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			  excelReader=new ExcelReader(POJOUtility.getExcelPath());
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsmenuBtn));
			SettingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigureTransactionBtn));
			ConfigureTransactionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMABtn));
			RMABtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAbasedalternateQuantityChkBox));
			
			if (RMAbasedalternateQuantityChkBoxIsSelected.isSelected()) 
			{
				RMAbasedalternateQuantityChkBox.click();
			}
			
			Thread.sleep(2000);
			
			String actRMAOptionIsSelected = Boolean.toString(RMAbasedalternateQuantityChkBoxIsSelected.isSelected());
			String expRMAOptionIsSelected =  excelReader.getCellData(xlSheetName, 225, 6);
            
          excelReader.setCellData(xlfile, xlSheetName, 225, 7, actRMAOptionIsSelected);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			
			getWaitForAlert();
			String actAlert=getAlert().getText();
			String expAlert= excelReader.getCellData(xlSheetName, 226, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 226, 7, actAlert);
			
			getAlert().accept();
			
			String expMessage= excelReader.getCellData(xlSheetName, 227, 6);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, xlSheetName, 227, 7, actMessage);
			
			System.out.println("RMAOptionIsSelected : "+actRMAOptionIsSelected+"  value expected  "+expRMAOptionIsSelected);
			System.out.println("Alert               : "+actAlert              +"  value expected  "+expAlert);
			System.out.println("Error Message       : "+actMessage            +"  value expected  "+expMessage);
					
			if(actRMAOptionIsSelected.equalsIgnoreCase(expRMAOptionIsSelected) && actAlert.equalsIgnoreCase(expAlert) && actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 224, 8, resPass);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 224, 8, resFail);
				
				return false;
			}
		}
			
			
			
			
			
		
		
		// RMA based alternative quantity
		
		public boolean checkUnCheckMaintainInventoryInAlternateUnitCheckBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			   excelReader=new ExcelReader(POJOUtility.getExcelPath());

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsmenuBtn));
			SettingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigureTransactionBtn));
			ConfigureTransactionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InventoryBtn));
			InventoryBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MaintainInventoryInAlternateUnitChkBox));
			
			if (MaintainInventoryInAlternateUnitChkBoxIsSelected.isSelected()) 
			{
				MaintainInventoryInAlternateUnitChkBox.click();
			}
			
			Thread.sleep(2000);
			
			String actMaintainInventoryInAlternateUnitChkBox = Boolean.toString(MaintainInventoryInAlternateUnitChkBoxIsSelected.isSelected());
			String expMaintainInventoryInAlternateUnitChkBox = excelReader.getCellData(xlSheetName, 229, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 229, 7, actMaintainInventoryInAlternateUnitChkBox);

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			
			getWaitForAlert();
			String actAlert=getAlert().getText();
			String expAlert= excelReader.getCellData(xlSheetName, 230, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 230, 7, actAlert);

			
			getAlert().accept();
			
			String expMessage= excelReader.getCellData(xlSheetName, 231, 6);
			
			String actMessage=checkValidationMessage(expMessage);
			excelReader.setCellData(xlfile, xlSheetName, 231, 7, actMessage);
			
			System.out.println("MaintainInventoryInAlternateUnitChkBox : "+actMaintainInventoryInAlternateUnitChkBox+"  value expected  "+expMaintainInventoryInAlternateUnitChkBox);
			System.out.println("Alert               : "+actAlert              +"  value expected  "+expAlert);
			System.out.println("Error Message       : "+actMessage            +"  value expected  "+expMessage);
					
			if(actMaintainInventoryInAlternateUnitChkBox.equalsIgnoreCase(expMaintainInventoryInAlternateUnitChkBox) && actAlert.equalsIgnoreCase(expAlert) 
					&& actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				excelReader.setCellData(xlfile, xlSheetName, 228, 8, resPass);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				excelReader.setCellData(xlfile, xlSheetName, 228, 8, resFail);
				return false;
			}
		}
		
			
				
			
		
		
		// Erase all Transactions after RMA Based Alternate Quantity
		
		public boolean checkEraseAllTransactionsAfterCompletingRMABasedAlternateQuantity() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
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
			
			String expValidationMsg=  excelReader.getCellData(xlSheetName, 233, 6);
			
			String actValidationMsg=checkValidationMessage(expValidationMsg);
			
			excelReader.setCellData(xlfile, xlSheetName, 233, 7, actValidationMsg);
			
			System.out.println("*************************************checkEraseAllTransactionsAfterCompletingTransactionFlow*********************************");
			System.out.println("Message  :  "+actValidationMsg +" Value Expected : "+expValidationMsg);
			
			
			Thread.sleep(2000);
			
			userNameDisplayLogo.click();
			
			
			Thread.sleep(2000);
			logoutOption.click();
			
			
			
			if (actValidationMsg.equalsIgnoreCase(expValidationMsg))
			{
				 excelReader.setCellData(xlfile, xlSheetName, 232, 8, resPass);


				return true;
			}
			else
			{
				 excelReader.setCellData(xlfile, xlSheetName, 232, 8, resPass);

				return false;
			}
		}
		
		
		
		public static boolean checkBackgroundSavingNegativeMessage(String docno) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			try
			{
				HashSet<String> actMessage = new HashSet<String>();

				for (int i = 0; i < 2; i++) 
				{
					String data = checkValidationMessage("");
					actMessage.add(data);
				}

				HashSet<String> expMessage = new HashSet<String>();

				expMessage.add("Voucher saved successfully : "+docno);
				
				expMessage.add("This Transaction will make the Stock Negative");

				System.out.println("Actual Message    : "+actMessage);
				System.out.println("Expected Message  : "+expMessage);

			    if(actMessage.equals(expMessage))
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
			System.err.println("Error Message NOT Found or NOT Clickable");
				System.err.println(e.getMessage());
		       return false;
				
			} 
		}	

				
		
		
		public boolean checkLoginAndEnableReservationIsMandatoryChkbox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
            excelReader=new ExcelReader(POJOUtility.getExcelPath());
			checkLogin();
			
			
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
			
			String expValidationMsg= excelReader.getCellData(xlSheetName, 235, 6);
			
			String actValidationMsg=checkValidationMessage(expValidationMsg);
			
			excelReader.setCellData(xlfile, xlSheetName, 235, 7, actValidationMsg);

			
			System.out.println("*************************************checkLoginAndEnableReservationIsMandatoryChkbox*********************************");
			System.out.println("Message  :  "+actValidationMsg +" Value Expected : "+expValidationMsg);
		
			
			Thread.sleep(1999);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseAllOkBtn));
			SettingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigureTransactionBtn));
			ConfigureTransactionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InventoryBtn));
			InventoryBtn.click();
			
			Thread.sleep(1999);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingInven_ReservationIsMandatoryChkbox));
			if (settingInven_ReservationIsMandatoryChkboxIsSelected.isSelected()==false)
			{
				settingInven_ReservationIsMandatoryChkbox.click();
			}
			
			
           Thread.sleep(2000);
			
			String actsettingInven_ReservationIsMandatoryChkbox = Boolean.toString(settingInven_ReservationIsMandatoryChkboxIsSelected.isSelected());
			String expsettingInven_ReservationIsMandatoryChkbox =  excelReader.getCellData(xlSheetName, 236, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 236, 7, actsettingInven_ReservationIsMandatoryChkbox);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			
			getWaitForAlert();
			String actAlert=getAlert().getText();
			
			String expAlert= excelReader.getCellData(xlSheetName, 237, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 237, 7, actAlert);
			
			getAlert().accept();
			
			String expMessage= excelReader.getCellData(xlSheetName, 238, 6);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, xlSheetName, 238, 7, actMessage);
			
			System.out.println("settingInven_ReservationIsMandatoryChkbox : "+actsettingInven_ReservationIsMandatoryChkbox+"  value expected  "+expsettingInven_ReservationIsMandatoryChkbox);
			System.out.println("Alert               : "+actAlert              +"  value expected  "+expAlert);
			System.out.println("Error Message       : "+actMessage            +"  value expected  "+expMessage);
					
			if(actsettingInven_ReservationIsMandatoryChkbox.equalsIgnoreCase(expsettingInven_ReservationIsMandatoryChkbox)
					&& actAlert.equalsIgnoreCase(expAlert) && actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 234, 8, resPass);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				 excelReader.setCellData(xlfile, xlSheetName, 234, 8, resFail);
				
				return false;
			}
		}
			
		@FindBy(xpath="//input[@id='ReservationIsMandatory']")
		private static WebElement settingInven_ReservationIsMandatoryChkboxIsSelected;
		
		
		
		@FindBy(xpath="//input[@id='ReservationIsMandatory']/following-sibling::span")
		private static WebElement settingInven_ReservationIsMandatoryChkbox;
		
		
		
		
		@FindBy(xpath="//*[@id='StockSettings']/button")
		private static WebElement inven_StockSettingExpandBtn;
		
		
		
		@FindBy(xpath="//input[@id='CheckNegativeStockGroup_1']")
		private static WebElement settingInven_StopRadioBtn;
		
		
		@FindBy(xpath="//input[@id='DonotAllowPartialReservations']")
		private static WebElement settingInven_DonotAllowPartialReservationsIsSelected;
		
		
		@FindBy(xpath="//input[@id='DonotAllowPartialReservations']/following-sibling::span")
		private static WebElement settingInven_DonotAllowPartialReservations;
		
		
		
		
		

		public boolean checkSavingPurchaseVoucherVATWithBRItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String ScreenOpenedMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
				
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				if (vendorAccount.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 240, 5)))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				if (department.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 241, 5)))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();
			
			enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
			int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
			
			for (int i = 0; i < pvvGridWarehouseListCount; i++) 
			{
				String warehouse=pvvGridWarehouseList.get(i).getText();
				if (warehouse.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 242, 5))) 
				{
					pvvGridWarehouseList.get(i).click();
					break;
				}
			}
	        enter_WarehouseTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
	        int pvvGridItemListCount=pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) 
			{
				String Item=pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 243, 5))) 
				{
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxCode));
	        enter_PvTaxCode.sendKeys(Keys.TAB);
				
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	        enter_AccountTxt.sendKeys(Keys.TAB);
			
	        Thread.sleep(2000);
	        
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 244, 5));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 245, 5));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_16thColumn));
			select1stRow_16thColumn.click();
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
			enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 246, 5));
			enter_Batch.sendKeys(Keys.TAB);
			
			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		    Date date=new Date();
	  	   
			Calendar cal=Calendar.getInstance();
			SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
			String presentDate = currentDate.format(cal.getTime());
			System.out.println("--------------Date Before Adding:: " + presentDate);
			
			cal.add(Calendar.MONTH, 1); 
			
			String nextMonthDate=df.format(cal.getTime());
			
			System.out.println("--------------Date After Adding:: " + nextMonthDate);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
			enter_Expirydate.click();
			enter_Expirydate.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			Thread.sleep(2000);
			enter_Expirydate.sendKeys(nextMonthDate,Keys.TAB);
			
			String actBatchNo=select1stRow_16thColumn.getText();
			String expBatchNo=  excelReader.getCellData(xlSheetName, 247, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 247, 7, actBatchNo);
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			String actSaving= Boolean.toString(checkBackgroundSavingMessage(docno));
			String expSaving=  excelReader.getCellData(xlSheetName, 248, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 248, 7, actSaving);

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			    
			
			System.out.println("******** Batch No      : "+actBatchNo +"  value expected  "+expBatchNo);
			
			if(actBatchNo.equalsIgnoreCase(expBatchNo) && actSaving.equalsIgnoreCase(expSaving)) 
			{
				 excelReader.setCellData(xlfile, xlSheetName, 239, 8, resPass);
				return true;
			}
			else
			{
				
				 excelReader.setCellData(xlfile, xlSheetName, 239, 8, resFail);
				return false;
			}
		}
		
		
		
		public boolean checkSavingRMAItemInPVVAT() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.sendKeys(Keys.SPACE);
				
			int vendorAccountCount=vendorAccountComboList.size();
			
			for (int i = 0; i < vendorAccountCount; i++) 
			{
				String vendorAccount=vendorAccountComboList.get(i).getText();
				
				if (vendorAccount.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 250, 5)))
				{
					vendorAccountComboList.get(i).click();
					vendorAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				if (department.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 251, 5)))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();
			
			enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
			int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
			
			for (int i = 0; i < pvvGridWarehouseListCount; i++) 
			{
				String warehouse=pvvGridWarehouseList.get(i).getText();
				if (warehouse.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 252, 5))) 
				{
					pvvGridWarehouseList.get(i).click();
					break;
				}
			}
	        enter_WarehouseTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
	        int pvvGridItemListCount=pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) 
			{
				String Item=pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 253, 5))) 
				{
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);	
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxCode));
	        enter_PvTaxCode.sendKeys(Keys.TAB);
				
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
	        enter_AccountTxt.sendKeys(Keys.TAB);
			
	        Thread.sleep(2000);
	        
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			enter_Quantity.sendKeys( excelReader.getCellData(xlSheetName, 254, 5));
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys( excelReader.getCellData(xlSheetName, 255, 5));
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PVDiscount));
			enter_PVDiscount.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_18thColumn));
			select1stRow_18thColumn.click();
		
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
			RMAPopupSerialNumberTxtField.sendKeys( excelReader.getCellData(xlSheetName, 256, 5));
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
			rmaAddBtn.click();
			
			Thread.sleep(2000);
			
			//actual
			String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
			String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
			String actRMA3=RMAPopupTableRow3Column1.getAttribute("data-value");
			String actRMA4=RMAPopupTableRow4Column1.getAttribute("data-value");
			String actRMA5=RMAPopupTableRow5Column1.getAttribute("data-value");
			String actRMA6=RMAPopupTableRow6Column1.getAttribute("data-value");
			String actRMA7=RMAPopupTableRow7Column1.getAttribute("data-value");
			String actRMA8=RMAPopupTableRow8Column1.getAttribute("data-value");
			String actRMA9=RMAPopupTableRow9Column1.getAttribute("data-value");
			String actRMA10=RMAPopupTableRow10Column1.getAttribute("data-value");
			
			
			
			//expected
			String expRMA1=  excelReader.getCellData(xlSheetName, 257, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 257, 7, expRMA1);

			String expRMA2= excelReader.getCellData(xlSheetName, 258, 6);
            excelReader.setCellData(xlfile, xlSheetName, 258, 7, expRMA2);
			
			String expRMA3= excelReader.getCellData(xlSheetName, 259, 6);
            excelReader.setCellData(xlfile, xlSheetName, 259, 7, expRMA3);
			
			String expRMA4= excelReader.getCellData(xlSheetName, 260, 6);
            excelReader.setCellData(xlfile, xlSheetName, 260, 7, expRMA4);
			
            String expRMA5= excelReader.getCellData(xlSheetName, 261, 6);
            excelReader.setCellData(xlfile, xlSheetName, 261, 7, expRMA5);
			
			String expRMA6= excelReader.getCellData(xlSheetName, 262, 6);
            excelReader.setCellData(xlfile, xlSheetName, 262, 7, expRMA6);
			
			String expRMA7=  excelReader.getCellData(xlSheetName, 263, 6);
            excelReader.setCellData(xlfile, xlSheetName, 263, 7, expRMA7);
			
            String expRMA8=  excelReader.getCellData(xlSheetName, 264, 6);
            excelReader.setCellData(xlfile, xlSheetName, 264, 7, expRMA8);
			
			String expRMA9= excelReader.getCellData(xlSheetName, 265, 6);
            excelReader.setCellData(xlfile, xlSheetName, 265, 7, expRMA9);
			
			String expRMA10= excelReader.getCellData(xlSheetName, 266, 6);
            excelReader.setCellData(xlfile, xlSheetName, 266, 7, expRMA10);
			
			
			System.out.println("********* Row1        : "+actRMA1  +"  value expected  "+expRMA1);
			System.out.println("********* Row2        : "+actRMA2  +"  value expected  "+expRMA2);
			System.out.println("********* Row3        : "+actRMA3  +"  value expected  "+expRMA3);
			System.out.println("********* Row4        : "+actRMA4  +"  value expected  "+expRMA4);
			System.out.println("********* Row5        : "+actRMA5  +"  value expected  "+expRMA5);
			System.out.println("********* Row6        : "+actRMA6  +"  value expected  "+expRMA6);
			System.out.println("********* Row7        : "+actRMA7  +"  value expected  "+expRMA7);
			System.out.println("********* Row8        : "+actRMA8  +"  value expected  "+expRMA8);
			System.out.println("********* Row9        : "+actRMA9  +"  value expected  "+expRMA9);
			System.out.println("********* Row10       : "+actRMA10 +"  value expected  "+expRMA10);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
			RMAPopupOkBtn.click();
		
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_RMA));
			
			String actRMAColumn=enter_RMA.getAttribute("value");
			
			String expRMAColumn=  excelReader.getCellData(xlSheetName, 267, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 267, 7, actRMAColumn);
			
			System.out.println("******** RMA Column: "+actRMAColumn+"  value expected  "+expRMAColumn);
			
	
			String docno=documentNumberTxt.getAttribute("value");
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
	
	        Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			
			String actSaving= Boolean.toString(checkBackgroundSavingMessage(docno));
			String  expSaving= excelReader.getCellData(xlSheetName, 268, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 268, 7, actSaving);
			

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  
			if(actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) 
				&& actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actRMA7.equalsIgnoreCase(expRMA7) && actRMA8.equalsIgnoreCase(expRMA8)	
				&& actRMA9.equalsIgnoreCase(expRMA9) && actRMA10.equalsIgnoreCase(expRMA10)  
				
				&& actRMAColumn.equalsIgnoreCase(expRMAColumn)
			    && actSaving.equalsIgnoreCase(expSaving))
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				excelReader.setCellData(xlfile, xlSheetName, 249, 8, resPass);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				excelReader.setCellData(xlfile, xlSheetName, 249, 8, resFail);
				return false;
			}	
		}
		
		
		
		@FindBy(xpath="//*[@id='140']/span")
		private static WebElement  inventoryTransactionsSalesMenu; 
			
		
		
		   @FindBy(xpath="//tbody[@id='id_header_4_table_body']/tr/td[2]")
			private static List<WebElement> customerAccountListCount; 
			
		
		   
		   public void logout() throws InterruptedException
		   {
			   Thread.sleep(2000);
			   getWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
			   userNameDisplayLogo.click();
			   
			   Thread.sleep(2000);
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			   logoutOption.click();
		   }
		   
		   
		public boolean checkClickOnSaveButtonInsalesaOrderWithoutreservaion() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			
			logout();
			
			Thread.sleep(1000);
			checkLogin();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsSalesMenu));
			inventoryTransactionsSalesMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesOrdersVoucher));
			salesOrdersVoucher.click();
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		  	
		  
			checkValidationMessage("Screen opened");
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		  		
		  	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		  	customerAccountTxt.click();
		  	customerAccountTxt.sendKeys(Keys.SPACE);
			 
		  	
		  	customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 270, 5));
		  	
		  	Thread.sleep(2000);
					
		  	customerAccountTxt.sendKeys(Keys.TAB);
				
		  	Thread.sleep(2000);
			 
		  	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		  	select1stRow_1stColumn.click();
		  	
		  	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		  	pvWareHouseTxt.click();
		  	pvWareHouseTxt.sendKeys(Keys.SPACE);
		  	
		  	int warehousecount=pvwareHouseListCount.size();
		  	
		  	System.err.println(warehousecount);
		  	
		  	for(int i=0 ; i < warehousecount ;i++)
		  	{
		  		String data=pvwareHouseListCount.get(i).getAttribute("data-name");
		  		
		  		if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 271, 5)))
		  		{
		  			pvwareHouseListCount.get(i).click();
		  			
		  			break;
		  		}	
		  	}
					
		  	pvWareHouseTxt.sendKeys(Keys.TAB);
		  	
		  	
		  	
		  	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		  	enter_ItemTxt.sendKeys(Keys.END);
		  	enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		  	enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		  	enter_ItemTxt.sendKeys(Keys.SPACE);
		  	
		  	int itemcount1=itemListCount.size();
		  	
		  	System.err.println(itemcount1);
		  	
		  	for(int i=0 ; i < itemcount1 ;i++)
		  	{
		  		String data=itemListCount.get(i).getAttribute("data-sname");
		  		
		  		if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 272, 5)))
		  		{
		  			itemListCount.get(i).click();
		  			
		  			break;
		  		}
		  	}
					
		  	enter_ItemTxt.sendKeys(Keys.TAB);
			 
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			enter_UnitTxt.sendKeys(Keys.TAB);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
			select1stRow_5thColumn.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
			so_enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 273, 5));
			so_enter_AQTxt.sendKeys(Keys.TAB);
		
			Thread.sleep(2000);
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			String ExpMessage= excelReader.getCellData(xlSheetName, 274, 6);
			String actMesage=checkValidationMessage(ExpMessage);
			excelReader.setCellData(xlfile, xlSheetName, 274, 7, actMesage);

			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			
			getWaitForAlert();
			
			getAlert().accept();
			
			if (actMesage.equalsIgnoreCase(ExpMessage)) 
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				 excelReader.setCellData(xlfile, xlSheetName, 269, 8, resPass);
				return true;
			}
			else
			{
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				 excelReader.setCellData(xlfile, xlSheetName, 269, 8, resFail);
				return false;
			}
		}
		
		
		public boolean checkMandatoryoptionWithWACogsItem() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			  excelReader=new ExcelReader(POJOUtility.getExcelPath());

			Thread.sleep(1000);
			
			logout();
			
			Thread.sleep(1000);
			
			checkLogin();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsSalesMenu));
			inventoryTransactionsSalesMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesOrdersVoucher));
			salesOrdersVoucher.click();
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		  	
		  
			checkValidationMessage("Screen opened");
		  		
		  	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		  	customerAccountTxt.click();
		  	customerAccountTxt.sendKeys(Keys.SPACE);
		  	customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 276, 5));
			 
		
		  	Thread.sleep(2000);
					
		  	customerAccountTxt.sendKeys(Keys.TAB);
				
		  	Thread.sleep(2000);
			 
		  	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		  	select1stRow_1stColumn.click();
		  	
		  	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		  	pvWareHouseTxt.click();
		  	pvWareHouseTxt.sendKeys(Keys.SPACE);
		  	
		  	pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 277, 5));
		  	Thread.sleep(2000);
					
		  	pvWareHouseTxt.sendKeys(Keys.TAB);
		  	
		  	
		  	
		  	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		  	enter_ItemTxt.sendKeys(Keys.END);
		  	enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		  	enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		  	enter_ItemTxt.sendKeys(Keys.SPACE);
		  	
		  	enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 278, 5));
		  	
		  	Thread.sleep(2000);
					
		  	enter_ItemTxt.sendKeys(Keys.TAB);
			 
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			enter_UnitTxt.sendKeys(Keys.TAB);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
			select1stRow_5thColumn.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
			so_enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 279, 5));
			so_enter_AQTxt.sendKeys(Keys.TAB);
		
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			String ExpMessage= excelReader.getCellData(xlSheetName, 280, 6);
			String actMesage=checkValidationMessage(ExpMessage);
			excelReader.setCellData(xlfile, xlSheetName, 280, 7, actMesage);

			
			
			if (actMesage.equalsIgnoreCase(ExpMessage)) 
				
			{

				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				if (getIsAlertPresent()) 
				{
					getAlert().accept();
				}
				 excelReader.setCellData(xlfile, xlSheetName, 275, 8, resPass);
				return true;
				
			}
			else
			{

				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				if (getIsAlertPresent()) 
				{
					getAlert().accept();
				}
				 excelReader.setCellData(xlfile, xlSheetName, 275, 8, resFail);
				return false;

			}
		
			
		}

		
		public boolean checkEnableDoNotAllowPartialreservation() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			  excelReader=new ExcelReader(POJOUtility.getExcelPath());
			
			Thread.sleep(1999);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsmenuBtn));
			SettingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigureTransactionBtn));
			ConfigureTransactionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InventoryBtn));
			InventoryBtn.click();
			
			Thread.sleep(1999);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingInven_ReservationIsMandatoryChkbox));
			if (settingInven_ReservationIsMandatoryChkboxIsSelected.isSelected()==true)
			{
				settingInven_ReservationIsMandatoryChkbox.click();
			}
			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingInven_DonotAllowPartialReservations));
			if (settingInven_DonotAllowPartialReservationsIsSelected.isSelected()==false)
			{
				settingInven_DonotAllowPartialReservations.click();
			}
			
           Thread.sleep(2000);
			
			String actsettingInven_DonotAllowPartialReservations = Boolean.toString(settingInven_DonotAllowPartialReservationsIsSelected.isSelected());
			String expsettingInven_DonotAllowPartialReservations =  excelReader.getCellData(xlSheetName, 282, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 282, 7, actsettingInven_DonotAllowPartialReservations);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			
			getWaitForAlert();
			String actAlert=getAlert().getText();
			
			String expAlert = excelReader.getCellData(xlSheetName, 283, 6);
            
			excelReader.setCellData(xlfile, xlSheetName, 283, 7, actAlert);
			
			getAlert().accept();
			
			String expMessage= excelReader.getCellData(xlSheetName, 284, 6);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, xlSheetName, 284, 7, actMessage);
			
			System.out.println("settingInven_DonotAllowPartialReservations : "+actsettingInven_DonotAllowPartialReservations+"  value expected  "+expsettingInven_DonotAllowPartialReservations);
			System.out.println("Alert               : "+actAlert              +"  value expected  "+expAlert);
			System.out.println("Error Message       : "+actMessage            +"  value expected  "+expMessage);
					
			if(actsettingInven_DonotAllowPartialReservations.equalsIgnoreCase(expsettingInven_DonotAllowPartialReservations)
					&& actAlert.equalsIgnoreCase(expAlert) && actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				excelReader.setCellData(xlfile, xlSheetName, 281, 8, resPass);
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				excelReader.setCellData(xlfile, xlSheetName, 281, 8, resFail);
				
				return false;
			}
			
		}
		
	
		
		public boolean checkInputMoreQtyInReservationpopupLessThanQtyInAQ() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{

			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			
			logout();
			
			Thread.sleep(1000);
			checkLogin();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsSalesMenu));
			inventoryTransactionsSalesMenu.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesOrdersVoucher));
			salesOrdersVoucher.click();
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
		  	
		  
			checkValidationMessage("Screen opened");
		  		
		  	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		  	customerAccountTxt.click();
		  	customerAccountTxt.sendKeys(Keys.SPACE);
			 
		  	customerAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 286, 5));
		  	Thread.sleep(2000);
					
		  	customerAccountTxt.sendKeys(Keys.TAB);
				
		  	Thread.sleep(2000);
			 
		  	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		  	select1stRow_1stColumn.click();
		  	
		  	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		  	pvWareHouseTxt.click();
		  	pvWareHouseTxt.sendKeys(Keys.SPACE);
		  	
		  	pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 287, 5));
		  	
		  	Thread.sleep(2000);
		  	
		  	pvWareHouseTxt.sendKeys(Keys.TAB);
		  	
		  	
		  	
		  	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		  	enter_ItemTxt.sendKeys(Keys.END);
		  	enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		  	enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		  	enter_ItemTxt.sendKeys(Keys.SPACE);
		  	
		  	enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 288, 5));
		  	
		  	Thread.sleep(2000);
		
					
		  	enter_ItemTxt.sendKeys(Keys.TAB);
			 
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
			enter_UnitTxt.sendKeys(Keys.TAB);
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
			select1stRow_5thColumn.click();
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_AQTxt));
			so_enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 289, 5));
			so_enter_AQTxt.sendKeys(Keys.TAB);
		

			Thread.sleep(2000);
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(so_enter_FQTxt));
			so_enter_FQTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(qtyToReserve));
			qtyToReserve.click();
			
			Thread.sleep(2000);
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(enterQtyToReserve));
			enterQtyToReserve.click();
			enterQtyToReserve.sendKeys(excelReader.getCellData(xlSheetName, 290, 5));
			
			
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(reservePopUpOkBtn));
			reservePopUpOkBtn.click();
			
			
			String actQTY=select1stRow_7thColumn.getText();
			String expQTY= excelReader.getCellData(xlSheetName, 291, 6);
			excelReader.setCellData(xlfile, xlSheetName, 291, 7, actQTY);
			
			System.out.println(" QTY : "+actQTY +" Value Exp: "+expQTY);
			
			if (actQTY.equalsIgnoreCase(expQTY)) 
			{
				System.out.println(" Test Pass: Qty is Updated ");
				
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				if (getIsAlertPresent()) 
				{
					getAlert().accept();
				}
				excelReader.setCellData(xlfile, xlSheetName, 285, 8, resPass);
				return true;
				
			}
			else
			{
				System.out.println(" Test Fail: Qty is Updated ");
				getWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
                new_CloseBtn.click();
				
				if (getIsAlertPresent()) 
				{
					getAlert().accept();
				}
				excelReader.setCellData(xlfile, xlSheetName, 285, 8, resFail);
				return false;

			}
			
		}
		
		//////////////////
		
		@FindBy(xpath="//div[contains(text(),'Ok')]")
		private static WebElement reservePopUpOkBtn;
		
		@FindBy(xpath="//td[@id='Id_BatchReservation_FGrid_col_1-2']")
		private static WebElement qtyToReserve;
		

		@FindBy(xpath="//input[@id='txtQtyToReserve']")
		private static WebElement enterQtyToReserve;

		@FindBy(xpath="//input[@id='id_body_33554460']")
		private static WebElement so_enter_AQTxt;
		
		@FindBy(xpath="//input[@id='id_body_33554461']")
		private static WebElement so_enter_FQTxt;

		
		public boolean checkSavingSalesInvoiceVATWithBatch() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
			financialsTransactionsSalesMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
			salesInvoiceVATVoucher.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			
			String ScreenOpeningMessage=checkValidationMessage("Screen opened");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			customerAccountTxt.sendKeys(Keys.SPACE);
			
			int customerAccountCount=customerAccountComboList.size();
			
			for (int i = 0; i < customerAccountCount; i++) 
			{
				String customerAccount=customerAccountComboList.get(i).getText();
				
				System.out.println(i+"  "+customerAccount);
				
				if (customerAccount.equalsIgnoreCase("Customer A"))
				{
					customerAccountComboList.get(i).click();
					customerAccountTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);
			
			int departmentCount=departmentComboList.size();
			
			for (int i = 0; i < departmentCount; i++) 
			{
				String department=departmentComboList.get(i).getText();
				
				if (department.equalsIgnoreCase("DUBAI"))
				{
					departmentComboList.get(i).click();
					departmentTxt.sendKeys(Keys.TAB);
					break;
				}
			}

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
			salesInvoiceVATPlaceOFSupply.click();
			salesInvoiceVATPlaceOFSupply.sendKeys("Abu Dhabi");
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWarehouseTxt));
			pvWarehouseTxt.sendKeys(Keys.SPACE);
			
			int warehouseCount=warehouseBodyComboList.size();
			
			for (int i = 0; i < warehouseCount; i++) 
			{
				String warehouse=warehouseBodyComboList.get(i).getText();
				
				if (warehouse.equalsIgnoreCase("HYDERABAD"))
				{
					warehouseBodyComboList.get(i).click();
					pvWarehouseTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.sendKeys(Keys.SPACE);
			
			int itemsCount=itemComboList.size();
			
			for (int i = 0; i < itemsCount; i++) 
			{
				String item=itemComboList.get(i).getText();
				
				if (item.equalsIgnoreCase("BR COGS ITEM")) 
				{
					itemComboList.get(i).click();
					enter_ItemTxt.sendKeys(Keys.TAB);
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccount));
			enter_SalesAccount.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQ));
			enter_AQ.sendKeys("6");
			enter_AQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQ));
			enter_FQ.sendKeys("6");
			enter_FQ.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("1.5");
			enter_Rate.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
			enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
			select1stRow_19thColumn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(batchItemName));
			
			int BatchNoCount = batchPopupBatchNoList.size();
			
			for (int j = 0; j < BatchNoCount; j++) 
			{
				String actBatchNo = batchPopupBatchNoList.get(j).getText();
				
				if (actBatchNo.equalsIgnoreCase("Batch#PVAT1")) 
				{
					getAction().doubleClick(batchPopupRowNosList.get(j)).build().perform();
				}
			}
			
	         Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
			batchOkIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
			String actBatchNo=select1stRow_19thColumn.getText();
			String expBatchNo="Batch#PVAT1";
			
			String docno=documentNumberTxt.getAttribute("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
			billRefPickIcon.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
			billRefOkBtn.click();
			
			
			boolean transactionNegativeMessage = true;
			
			if (errorMessage.getText().equalsIgnoreCase("This Transaction will make the Stock Negative")) 
			{
				System.err.println(errorMessage.getText());
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
				errorMessageCloseBtn.click();
				//transactionNegativeMessage = false;
			}
			
			Thread.sleep(1000);
			
			boolean actSaving=checkBackgroundSavingMessage(docno);
			boolean expSaving=true;

			System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
			  
			
			System.out.println("************* Batch No         : "+actBatchNo          +"  value expected  "+expBatchNo);
			
			if (actBatchNo.equalsIgnoreCase(expBatchNo) && actSaving==expSaving 
					&& transactionNegativeMessage==true) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
		}
			
			
			
		public boolean checkChangingOptionAsStopInPreferenceUnderInventory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			getDriver().navigate().refresh();
			
			Thread.sleep(1999);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsmenuBtn));
			SettingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigureTransactionBtn));
			ConfigureTransactionBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InventoryBtn));
			InventoryBtn.click();
			
			Thread.sleep(1999);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inven_StockSettingExpandBtn));
			inven_StockSettingExpandBtn.click();
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingInven_StopRadioBtn));
			if (settingInven_StopRadioBtn.isSelected()==false)
			{
				settingInven_StopRadioBtn.click();
			}
			
			
           Thread.sleep(2000);
			
			boolean actsettingInven_StopRadioBtn = settingInven_StopRadioBtn.isSelected();
			boolean expsettingInven_StopRadioBtn = true;
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			
			getWaitForAlert();
			String actAlert=getAlert().getText();
			String expAlert="Do you want to save the changes?";
			getAlert().accept();
			
			String expMessage="Data saved Successfully";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("settingInven_StopRadioBtn : "+actsettingInven_StopRadioBtn+"  value expected  "+expsettingInven_StopRadioBtn);
			System.out.println("Alert               : "+actAlert              +"  value expected  "+expAlert);
			System.out.println("Error Message       : "+actMessage            +"  value expected  "+expMessage);
					
			if(actsettingInven_StopRadioBtn==expsettingInven_StopRadioBtn
					&& actAlert.equalsIgnoreCase(expAlert) && actMessage.equalsIgnoreCase(expMessage)) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				
				return true;
			}
			else
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CloseBtn));
				CloseBtn.click();
				
				Thread.sleep(2000);
				
				return false;
			}
	
		}
				
				
	
	public boolean checkResavingSalesInvoiceVATAfterChangigInSettingWithLogoutAndLogin() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();


		Thread.sleep(2000);
		
		checkLogin();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsSalesMenu));
		financialsTransactionsSalesMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
		salesInvoiceVATVoucher.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		
		String ScreenOpeningMessage=checkValidationMessage("Screen opened");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();

		
		String loading=checkValidationMessage("Voucher Loaded Successfully");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
		String actBatchNo=select1stRow_19thColumn.getText();
		String expBatchNo="Batch#PVAT1";
		
		String docno=documentNumberTxt.getAttribute("value");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		
		boolean transactionNegativeMessage = true;
		
		if (errorMessage.getText().equalsIgnoreCase("This Transaction will make the Stock negative. It cannot be saved")) 
		{
			System.err.println(errorMessage.getText());
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
			errorMessageCloseBtn.click();
		}
		
		Thread.sleep(1000);
		
		boolean actSaving=checkBackgroundSavingMessage(docno);
		boolean expSaving=true;

		System.out.println(" Final Mesage Displayed  : "+actSaving +" Value Exp : "+expSaving);
		  
		
		System.out.println("************* Batch No         : "+actBatchNo          +"  value expected  "+expBatchNo);
		if (actBatchNo.equalsIgnoreCase(expBatchNo) && actSaving==expSaving 
				&& transactionNegativeMessage==true) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			
			Thread.sleep(2000);
			
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();
			
			Thread.sleep(2000);
			
			return false;
		}
		
	}
	
				
	
	
	public boolean checkEraseAllDataAndSavingPurchaseVoucherAfterEnableReservationMadatory() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
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
		
		String expValidationMsg="Data deleted successfully.";
		
		String actValidationMsg=checkValidationMessage(expValidationMsg);
		
		System.out.println("Message  :  "+actValidationMsg +" Value Expected : "+expValidationMsg);
	
		
		boolean actcheckSavingPurchaseVoucherVATWithBRItem=checkSavingPurchaseVoucherVATWithBRItem();
		boolean expcheckSavingPurchaseVoucherVATWithBRItem=true;
		
		System.err.println(" Actual  :  "+actcheckSavingPurchaseVoucherVATWithBRItem);
		System.err.println(" Expected : "+expcheckSavingPurchaseVoucherVATWithBRItem);
		
		if (actcheckSavingPurchaseVoucherVATWithBRItem==expcheckSavingPurchaseVoucherVATWithBRItem
				 && actValidationMsg.equalsIgnoreCase(expValidationMsg)) 
		{
			return true;
		} else 
		{

			return false;
		}
		
	}
				
		
	public boolean checkSavingSalesInvoiceDirectlyAfterEnableReserVationManadatory() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		Thread.sleep(2000);
		
		if (checkSavingSalesInvoiceVATWithBatch()==true) 
		{
		return true;	
		} 
		else 
		{
			return false;
		}
		
	}
		
	
	
	   public static boolean checkBackgroundSavingMessage(String docno) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
							

				
				try
				{
					getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(errorMessage));


					String actErrorMessage=errorMessage.getText();
					String expErrorMessage="Voucher saved successfully";
					String expErrorMessage1=docno;


					if (actErrorMessage.startsWith(expErrorMessage) && 
							actErrorMessage.endsWith(expErrorMessage1)) 
					{
						getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
						errorMessageCloseBtn.click();

						System.out.println("ValidationMessage  :  "+actErrorMessage +" Starts With  : "+expErrorMessage+" Ends With "+expErrorMessage1);

						return true;
					}

					else
					{
					return false;
					}

				}
				catch(Exception e)
				{
					System.err.println("Error Message NOT Found or NOT Clickable");
					System.err.println(e.getMessage());
					return false;
				}

			}	

	   
			
			
			
			 public RMAWithAllOptionsPage(WebDriver driver)
			 {
			    	
			    	PageFactory.initElements(driver, this);
			    	
			 }
		}
