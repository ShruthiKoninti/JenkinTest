package com.focus.Pages;

import java.util.List;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;

public class TDS_TCS_GST 
{


	@FindBy(xpath="//div[text()='Financials']")
	private static WebElement financialsMenu;

	@FindBy(xpath="//span[text()='Statutory']")
	private static WebElement statutoryMenu;

	@FindBy(xpath="//a[@id='4502']/span[text()='TDS/TCS']")
	private static WebElement TDS_TCSMenu;




	@FindBy(xpath="//a[@id='4602']/span[text()='GST']")
	private static WebElement GSTMenu;

	// Sub-Menus Under GST Menu

	@FindBy(xpath="//a[@id='4603']/span[text()='Preferences']")
	private static WebElement GSTPreferencesMenu;

	@FindBy(xpath="//a[@id='4604']/span[text()='GSTR1']")
	private static WebElement GSTR1Menu;

	// Sub-Menus Under GSTR1 Menu

	@FindBy(xpath="//a[@id='4605']/span[text()='Online Filing']")
	private static WebElement GSTR1OnlineFilingMenu;

	@FindBy(xpath="//a[@id='4606']/span[text()='Offline Filing']")
	private static WebElement GSTR1OfflineFilingMenu;

	@FindBy(xpath="//a[@id='4607']/span[text()='GSTR2A']")
	private static WebElement GSTR2AMenu;

	@FindBy(xpath="//a[@id='4610']/span[text()='GSTR3B']")
	private static WebElement GSTR3BMenu;

	@FindBy(xpath="//a[@id='4611']/span[text()='GSTR4 Offline']")
	private static WebElement GSTR4OfflineMenu;

	@FindBy(xpath="//a[@id='4612']/span[text()='GSTR9']")
	private static WebElement GSTR9Menu;

	@FindBy(xpath="//a[@id='4614']/span[text()='Bulk EWay Bill']")
	private static WebElement BulkEWayBillMenu;

	@FindBy(xpath="//a[@id='4615']/span[text()='Update Masters']")
	private static WebElement updateMastersMenu;

	@FindBy(xpath="//a[@id='4617']/span[text()='About GST']")
	private static WebElement aboutGSTMenu;



	// GST - Preferences Screen

	@FindBy(xpath="//ul[@id='report-tab']/li[1]/a")
	private static WebElement GSTPreferences_GSTTab;

	@FindBy(xpath="//ul[@id='report-tab']/li[2]/a")
	private static WebElement GSTPreferences_MiscellaneousTab;

	@FindBy(xpath="//ul[@id='report-tab']/li[3]/a")
	private static WebElement GSTPreferences_VoucherMappingTab;

	@FindBy(xpath="//ul[@id='report-tab']/li[4]/a")
	private static WebElement GSTPreferences_ProvisionalAssessmentTab;

	@FindBy(xpath="//ul[@id='report-tab']/li[5]/a")
	private static WebElement GSTPreferences_InformationTab;

	@FindBy(xpath="//ul[@id='report-tab']/li[6]/a")
	private static WebElement GSTPreferences_EWayBillTab;

	@FindBy(xpath="//ul[@id='report-tab']/li[7]/a")
	private static WebElement GSTPreferences_EInvoiceTab;
	
	
	// Elements Under GST Tab of GST Preferences
	
	// Company Details
	
	@FindBy(xpath="//div[@id='CompanyDetails']/h4/span[2]/i")
	private static WebElement GSTTab_CompanyDetailsExpandOrCollapseBtn;
	
	@FindBy(xpath="//input[@id='sCompanyName']")
	private static WebElement GSTTab_CompanyNameTxt;
	
	@FindBy(xpath="//input[@id='sLegalName']")
	private static WebElement GSTTab_LegalNameTxt;
	
	@FindBy(xpath="//input[@id='CompGSTIN']")
	private static WebElement GSTTab_GSTINTxt;
	
	@FindBy(xpath="//input[@id='CompGSTINRegDate']")
	private static WebElement GSTTab_GSTINRegistrationDateTxt;
	
	@FindBy(xpath="//input[@id='txtCompOldTIN']")
	private static WebElement GSTTab_OldTINTxt;
	
	@FindBy(xpath="//select[@id='ddCompStateName']")
	private static WebElement GSTTab_StateNameDrpDwn;
	
	@FindBy(xpath="//input[@id='txtPrefPIN']")
	private static WebElement GSTTab_PINTxt;
	
	@FindBy(xpath="//select[@id='ddTaxPayersType']")
	private static WebElement GSTTab_TypeOfTaxpayersDrpDwn;
	
	@FindBy(xpath="//input[@id='txtCompPAN']")
	private static WebElement GSTTab_PANTxt;
	
	@FindBy(xpath="//input[@id='txtCompLYGTurnover']")
	private static WebElement GSTTab_AggregateTurnoverLastFYTxt;
	
	@FindBy(xpath="//input[@id='AggregateTurnover']")
	private static WebElement GSTTab_AggregateTurnoverAprJun2017Txt;
	
	
	// Bank Details
	
	@FindBy(xpath="//input[@id='txtCompBankName']")
	private static WebElement GSTTab_NameOfTheBankTxt;
	
	@FindBy(xpath="//input[@id='txtCompBankAcctNo']")
	private static WebElement GSTTab_AccountNoTxt;
	
	@FindBy(xpath="//input[@id='txtCompBankBranch']")
	private static WebElement GSTTab_BranchTxt;
	
	@FindBy(xpath="//input[@id='txtBranchIFSC']")
	private static WebElement GSTTab_BranchIFSCTxt;
	
	@FindBy(xpath="//input[@id='txtUPIID']")
	private static WebElement GSTTab_UPIIDTxt;
	
	
	// Shipping Details
	
	@FindBy(xpath="//input[@id='rbShipAddrAccountsMaster']")
	private static WebElement GSTTab_ShippingAddressIsSameAsBillingAddressInAccountMasterRadio;
	
	@FindBy(xpath="//input[@id='rbShipAddrSeparateMaster']")
	private static WebElement GSTTab_ShippingAddressIsSeperateMasterRadio;
	
	@FindBy(xpath="//input[@id='rbShipAddrVoucherLevel']")
	private static WebElement GSTTab_ShippingAddressInVoucherLevelRadio;
	
	
	// Type of Signature (EVC)
	
	@FindBy(xpath="//input[@id='txtPAN']")
	private static WebElement GSTTab_TypeOfSign_PANNoTxt;
	
	
	
	
	// Elements Under Miscellaneous Tab of GST Preferences
	
	// Ecom_TIN
	
	@FindBy(xpath="//input[@id='rbECommSeparateMaster']")
	private static WebElement MiscellaneousTab_ECommerceOperatorAsASeparateMasterRadio;
	
	@FindBy(xpath="//input[@id='rbeTINExtraField']")
	private static WebElement MiscellaneousTab_eTINAsExtraFieldRadio;
	
	@FindBy(xpath="//select[@id='ddShipAddrSeparateMaster1']")
	private static WebElement MiscellaneousTab_ECommerceOperatorAsASeparateMasterDrpDwn;
	
	@FindBy(xpath="//input[@id='txteTINExtraField']")
	private static WebElement MiscellaneousTab_eTINAsExtraFieldTxt;
	
	
	// B2C Data in GSTR1
	
	@FindBy(xpath="//input[@id='rbB2CMonth']")
	private static WebElement MiscellaneousTab_MonthWiseRadio;
	
	@FindBy(xpath="//input[@id='rbB2CDate']")
	private static WebElement MiscellaneousTab_DateWiseRadio;
	
	
	// RCM
	
	@FindBy(xpath="//input[@id='chkrcmcalculate']")
	private static WebElement MiscellaneousTab_EnableSelfInvoicePostingForSection9_4_DailyChkbox;
	
	@FindBy(xpath="//input[@id='chkrcmmonthlycalculate']")
	private static WebElement MiscellaneousTab_EnableSelfInvoicePostingForSection9_3_MonthlyChkbox;
	
	@FindBy(xpath="//select[@id='dtRCMMonthlyAppDate']")
	private static WebElement MiscellaneousTab_ApplicationFormDrpDwn;
	
	@FindBy(xpath="//button[@id='btnRCMMonthlyAppDate']")
	private static WebElement MiscellaneousTab_RCMDataProcessButton;
	
	
	// Report UOM
	
	@FindBy(xpath="//input[@id='chkReportUOM']")
	private static WebElement MiscellaneousTab_ReportQuantityInBaseUnitChkbox;
	
	
	// Triggers
	
	@FindBy(xpath="//input[@id='chkInvoicevalidate']")
	private static WebElement MiscellaneousTab_DoNotValidateInvoiceNoInCNOrDNChkbox;
	
	@FindBy(xpath="//input[@id='chkeInvHSNWise']")
	private static WebElement MiscellaneousTab_GenerateEInvoiceHSNOrSACWiseChkbox;
	
	@FindBy(xpath="//input[@id='chkvalidateTrigger']")
	private static WebElement MiscellaneousTab_EnableVoucherValidationsInMappedGSTVoucherChkbox;
	
	@FindBy(xpath="//input[@id='chkvalidategstin']")
	private static WebElement MiscellaneousTab_ValidateGSTINInAccountMasterChkbox;
	
	@FindBy(xpath="//select[@id='dtCNDNDelinkAppDate']")
	private static WebElement MiscellaneousTab_CNDNDelinkingDateApplicationFromDrpDwn;
	
	@FindBy(xpath="//input[@id='chkbQRCodeforB2C']")
	private static WebElement MiscellaneousTab_GenerateDynamicQRCodeForB2CInvoicesChkbox;
	
	@FindBy(xpath="//input[@id='chkbItemtypeHSNcode']")
	private static WebElement MiscellaneousTab_ItemTypeBasedOnHSNOrSACCodeChkbox;
	
	
	// Voucher Mapping
	
	// GST Field Mapping
	
	@FindBy(xpath="//li[@id='Voucher_new']/span/i")
	private static WebElement VoucherMappingTab_NewBtn;
	
	
	// Table Edit and Delete Button Elements
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[1]/td[1]/button[1]")
	private static WebElement VoucherMappingTab_TableRow1EditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[1]/td[1]/button[2]")
	private static WebElement VoucherMappingTab_TableRow1DeleteBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[2]/td[1]/button[1]")
	private static WebElement VoucherMappingTab_TableRow2EditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[2]/td[1]/button[2]")
	private static WebElement VoucherMappingTab_TableRow2DeleteBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[3]/td[1]/button[1]")
	private static WebElement VoucherMappingTab_TableRow3EditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[3]/td[1]/button[2]")
	private static WebElement VoucherMappingTab_TableRow3DeleteBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[4]/td[1]/button[1]")
	private static WebElement VoucherMappingTab_TableRow4EditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[4]/td[1]/button[2]")
	private static WebElement VoucherMappingTab_TableRow4DeleteBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[5]/td[1]/button[1]")
	private static WebElement VoucherMappingTab_TableRow5EditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[5]/td[1]/button[2]")
	private static WebElement VoucherMappingTab_TableRow5DeleteBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[6]/td[1]/button[1]")
	private static WebElement VoucherMappingTab_TableRow6EditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[6]/td[1]/button[2]")
	private static WebElement VoucherMappingTab_TableRow6DeleteBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[7]/td[1]/button[1]")
	private static WebElement VoucherMappingTab_TableRow7EditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[7]/td[1]/button[2]")
	private static WebElement VoucherMappingTab_TableRow7DeleteBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[8]/td[1]/button[1]")
	private static WebElement VoucherMappingTab_TableRow8EditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[8]/td[1]/button[2]")
	private static WebElement VoucherMappingTab_TableRow8DeleteBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[9]/td[1]/button[1]")
	private static WebElement VoucherMappingTab_TableRow9EditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[9]/td[1]/button[2]")
	private static WebElement VoucherMappingTab_TableRow9DeleteBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[10]/td[1]/button[1]")
	private static WebElement VoucherMappingTab_TableRow10EditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[10]/td[1]/button[2]")
	private static WebElement VoucherMappingTab_TableRow10DeleteBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[11]/td[1]/button[1]")
	private static WebElement VoucherMappingTab_TableRow11EditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[11]/td[1]/button[2]")
	private static WebElement VoucherMappingTab_TableRow11DeleteBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[12]/td[1]/button[1]")
	private static WebElement VoucherMappingTab_TableRow12EditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[12]/td[1]/button[2]")
	private static WebElement VoucherMappingTab_TableRow12DeleteBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[13]/td[1]/button[1]")
	private static WebElement VoucherMappingTab_TableRow13EditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[13]/td[1]/button[2]")
	private static WebElement VoucherMappingTab_TableRow13DeleteBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[14]/td[1]/button[1]")
	private static WebElement VoucherMappingTab_TableRow14EditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[14]/td[1]/button[2]")
	private static WebElement VoucherMappingTab_TableRow14DeleteBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[15]/td[1]/button[1]")
	private static WebElement VoucherMappingTab_TableRow15EditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[15]/td[1]/button[2]")
	private static WebElement VoucherMappingTab_TableRow15DeleteBtn;
	

	// Voucher Type
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[1]/td[4]")
	private static WebElement VoucherMappingTab_TableRow1VoucherTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[2]/td[4]")
	private static WebElement VoucherMappingTab_TableRow2VoucherTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[3]/td[4]")
	private static WebElement VoucherMappingTab_TableRow3VoucherTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[4]/td[4]")
	private static WebElement VoucherMappingTab_TableRow4VoucherTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[5]/td[4]")
	private static WebElement VoucherMappingTab_TableRow5VoucherTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[6]/td[4]")
	private static WebElement VoucherMappingTab_TableRow6VoucherTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[7]/td[4]")
	private static WebElement VoucherMappingTab_TableRow7VoucherTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[8]/td[4]")
	private static WebElement VoucherMappingTab_TableRow8VoucherTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[9]/td[4]")
	private static WebElement VoucherMappingTab_TableRow9VoucherTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[10]/td[4]")
	private static WebElement VoucherMappingTab_TableRow10VoucherTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[11]/td[4]")
	private static WebElement VoucherMappingTab_TableRow11VoucherTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[12]/td[4]")
	private static WebElement VoucherMappingTab_TableRow12VoucherTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[13]/td[4]")
	private static WebElement VoucherMappingTab_TableRow13VoucherTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[14]/td[4]")
	private static WebElement VoucherMappingTab_TableRow14VoucherTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[15]/td[4]")
	private static WebElement VoucherMappingTab_TableRow15VoucherTypesTxt;
	
	
	
	// Business Types
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[1]/td[5]")
	private static WebElement VoucherMappingTab_TableRow1BusinessTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[2]/td[5]")
	private static WebElement VoucherMappingTab_TableRow2BusinessTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[3]/td[5]")
	private static WebElement VoucherMappingTab_TableRow3BusinessTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[4]/td[5]")
	private static WebElement VoucherMappingTab_TableRow4BusinessTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[5]/td[5]")
	private static WebElement VoucherMappingTab_TableRow5BusinessTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[6]/td[5]")
	private static WebElement VoucherMappingTab_TableRow6BusinessTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[7]/td[5]")
	private static WebElement VoucherMappingTab_TableRow7BusinessTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[8]/td[5]")
	private static WebElement VoucherMappingTab_TableRow8BusinessTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[9]/td[5]")
	private static WebElement VoucherMappingTab_TableRow9BusinessTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[10]/td[5]")
	private static WebElement VoucherMappingTab_TableRow10BusinessTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[11]/td[5]")
	private static WebElement VoucherMappingTab_TableRow11BusinessTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[12]/td[5]")
	private static WebElement VoucherMappingTab_TableRow12BusinessTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[13]/td[5]")
	private static WebElement VoucherMappingTab_TableRow13BusinessTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[14]/td[5]")
	private static WebElement VoucherMappingTab_TableRow14BusinessTypesTxt;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[15]/td[5]")
	private static WebElement VoucherMappingTab_TableRow15BusinessTypesTxt;
	
	
	
	// Nature of Document
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[1]/td[6]")
	private static WebElement VoucherMappingTab_TableRow1NatureOfDocument;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[2]/td[6]")
	private static WebElement VoucherMappingTab_TableRow2NatureOfDocument;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[3]/td[6]")
	private static WebElement VoucherMappingTab_TableRow3NatureOfDocument;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[4]/td[6]")
	private static WebElement VoucherMappingTab_TableRow4NatureOfDocument;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[5]/td[6]")
	private static WebElement VoucherMappingTab_TableRow5NatureOfDocument;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[6]/td[6]")
	private static WebElement VoucherMappingTab_TableRow6NatureOfDocument;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[7]/td[6]")
	private static WebElement VoucherMappingTab_TableRow7NatureOfDocument;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[8]/td[6]")
	private static WebElement VoucherMappingTab_TableRow8NatureOfDocument;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[9]/td[6]")
	private static WebElement VoucherMappingTab_TableRow9NatureOfDocument;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[10]/td[6]")
	private static WebElement VoucherMappingTab_TableRow10NatureOfDocument;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[11]/td[6]")
	private static WebElement VoucherMappingTab_TableRow11NatureOfDocument;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[12]/td[6]")
	private static WebElement VoucherMappingTab_TableRow12NatureOfDocument;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[13]/td[6]")
	private static WebElement VoucherMappingTab_TableRow13NatureOfDocument;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[14]/td[6]")
	private static WebElement VoucherMappingTab_TableRow14NatureOfDocument;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr[15]/td[6]")
	private static WebElement VoucherMappingTab_TableRow15NatureOfDocument;
	
	
	// Table List 
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr/td[1]/button[1]")
	private static List<WebElement> VoucherMappingTab_TableListEditBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr/td[1]/button[2]")
	private static List<WebElement> VoucherMappingTab_TableListDeleteBtn;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr/td[4]")
	private static List<WebElement> VoucherMappingTab_TableListVoucherTypes;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr/td[5]")
	private static List<WebElement> VoucherMappingTab_TableListBusinessTypes;
	
	@FindBy(xpath="//tbody[@id='tblbodyMapping_data']/tr/td[6]")
	private static List<WebElement> VoucherMappingTab_TableListNatureOfDocument;
	
	
	// GST Field Mapping
	
	@FindBy(xpath="//select[@id='ddVoucherType']")
	private static WebElement VoucherMappingTab_VoucherTypeDrpDwn;
	
	@FindBy(xpath="//select[@id='ddbusinessType']")
	private static WebElement VoucherMappingTab_BusinessTypeDrpDwn;
	
	@FindBy(xpath="//select[@id='ddnaturedoc']")
	private static WebElement VoucherMappingTab_NatureOfDocumentDrpDwn;
	
	@FindBy(xpath="//div[@id='btnvoucherdisplay']/div/span")
	private static WebElement VoucherMappingTab_GSTFieldMappingOkBtn;
	
	@FindBy(xpath="//li[@id='VocherbtnAuto']/span[text()='Create Fields ']")
	private static WebElement VoucherMappingTab_CreateFieldsBtn;
	
	@FindBy(xpath="//li[@id='Vocherbtnsave']/span[text()='Save']")
	private static WebElement VoucherMappingTab_GSTFieldMappingSaveBtn;
	
	@FindBy(xpath="//li[@id='VocherCancel']/span[text()='Cancel']")
	private static WebElement VoucherMappingTab_GSTFieldMappingCancelBtn;
	
	@FindBy(xpath="//input[@id='chkextrafield']")
	private static WebElement VoucherMappingTab_CreateExtraFieldsInDifferentTabChkbox;
	
	@FindBy(xpath="//input[@id='chkEwaybill']")
	private static WebElement VoucherMappingTab_EWayBillChkbox;
	
	
	// Business Type From
	
	@FindBy(xpath="//input[@id='rbBusinesstype']")
	private static WebElement VoucherMappingTab_BusinessTypeFromVoucherFieldsRadio;
	
	@FindBy(xpath="//input[@id='rbAccountMster']")
	private static WebElement VoucherMappingTab_BusinessTypeFromCustomerOrVendorMasterRadio;
	
	
	// GST Field Mappings Table 
	// GST Fields
	
	@FindBy(xpath="//input[@id='GstFields0']")
	private static WebElement VoucherMappingTab_TableRow1GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields1']")
	private static WebElement VoucherMappingTab_TableRow2GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields2']")
	private static WebElement VoucherMappingTab_TableRow3GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields3']")
	private static WebElement VoucherMappingTab_TableRow4GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields4']")
	private static WebElement VoucherMappingTab_TableRow5GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields5']")
	private static WebElement VoucherMappingTab_TableRow6GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields6']")
	private static WebElement VoucherMappingTab_TableRow7GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields7']")
	private static WebElement VoucherMappingTab_TableRow8GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields8']")
	private static WebElement VoucherMappingTab_TableRow9GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields9']")
	private static WebElement VoucherMappingTab_TableRow10GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields10']")
	private static WebElement VoucherMappingTab_TableRow11GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields11']")
	private static WebElement VoucherMappingTab_TableRow12GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields12']")
	private static WebElement VoucherMappingTab_TableRow13GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields13']")
	private static WebElement VoucherMappingTab_TableRow14GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields14']")
	private static WebElement VoucherMappingTab_TableRow15GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields15']")
	private static WebElement VoucherMappingTab_TableRow16GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields16']")
	private static WebElement VoucherMappingTab_TableRow17GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields17']")
	private static WebElement VoucherMappingTab_TableRow18GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields18']")
	private static WebElement VoucherMappingTab_TableRow19GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields19']")
	private static WebElement VoucherMappingTab_TableRow20GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields20']")
	private static WebElement VoucherMappingTab_TableRow21GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields21']")
	private static WebElement VoucherMappingTab_TableRow22GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields22']")
	private static WebElement VoucherMappingTab_TableRow23GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields23']")
	private static WebElement VoucherMappingTab_TableRow24GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields24']")
	private static WebElement VoucherMappingTab_TableRow25GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields25']")
	private static WebElement VoucherMappingTab_TableRow26GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields26']")
	private static WebElement VoucherMappingTab_TableRow27GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields27']")
	private static WebElement VoucherMappingTab_TableRow28GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields28']")
	private static WebElement VoucherMappingTab_TableRow29GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields29']")
	private static WebElement VoucherMappingTab_TableRow30GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields30']")
	private static WebElement VoucherMappingTab_TableRow31GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields31']")
	private static WebElement VoucherMappingTab_TableRow32GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields32']")
	private static WebElement VoucherMappingTab_TableRow33GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields33']")
	private static WebElement VoucherMappingTab_TableRow34GSTFieldsTxt;
	
	@FindBy(xpath="//input[@id='GstFields34']")
	private static WebElement VoucherMappingTab_TableRow35GSTFieldsTxt;
	
	
	
	// Focus Fields
	
	@FindBy(xpath="//select[@id='ddFocusFields0']")
	private static WebElement VoucherMappingTab_TableRow1FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields1']")
	private static WebElement VoucherMappingTab_TableRow2FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields2']")
	private static WebElement VoucherMappingTab_TableRow3FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields3']")
	private static WebElement VoucherMappingTab_TableRow4FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields4']")
	private static WebElement VoucherMappingTab_TableRow5FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields5']")
	private static WebElement VoucherMappingTab_TableRow6FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields6']")
	private static WebElement VoucherMappingTab_TableRow7FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields7']")
	private static WebElement VoucherMappingTab_TableRow8FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields8']")
	private static WebElement VoucherMappingTab_TableRow9FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields9']")
	private static WebElement VoucherMappingTab_TableRow10FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields10']")
	private static WebElement VoucherMappingTab_TableRow11FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields11']")
	private static WebElement VoucherMappingTab_TableRow12FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields12']")
	private static WebElement VoucherMappingTab_TableRow13FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields13']")
	private static WebElement VoucherMappingTab_TableRow14FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields14']")
	private static WebElement VoucherMappingTab_TableRow15FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields15']")
	private static WebElement VoucherMappingTab_TableRow16FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields16']")
	private static WebElement VoucherMappingTab_TableRow17FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields17']")
	private static WebElement VoucherMappingTab_TableRow18FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields18']")
	private static WebElement VoucherMappingTab_TableRow19FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields19']")
	private static WebElement VoucherMappingTab_TableRow20FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields20']")
	private static WebElement VoucherMappingTab_TableRow21FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields21']")
	private static WebElement VoucherMappingTab_TableRow22FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields22']")
	private static WebElement VoucherMappingTab_TableRow23FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields23']")
	private static WebElement VoucherMappingTab_TableRow24FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields24']")
	private static WebElement VoucherMappingTab_TableRow25FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields25']")
	private static WebElement VoucherMappingTab_TableRow26FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields26']")
	private static WebElement VoucherMappingTab_TableRow27FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields27']")
	private static WebElement VoucherMappingTab_TableRow28FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields28']")
	private static WebElement VoucherMappingTab_TableRow29FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields29']")
	private static WebElement VoucherMappingTab_TableRow30FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields30']")
	private static WebElement VoucherMappingTab_TableRow31FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields31']")
	private static WebElement VoucherMappingTab_TableRow32FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields32']")
	private static WebElement VoucherMappingTab_TableRow33FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields33']")
	private static WebElement VoucherMappingTab_TableRow34FocusFieldsTxt;
	
	@FindBy(xpath="//select[@id='ddFocusFields34']")
	private static WebElement VoucherMappingTab_TableRow35FocusFieldsTxt;
	
	
	
	// AccountToPost
	
	@FindBy(xpath="//select[@id='ddDebitAccont0']")
	private static WebElement VoucherMappingTab_TableRow1AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont1']")
	private static WebElement VoucherMappingTab_TableRow2AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont2']")
	private static WebElement VoucherMappingTab_TableRow3AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont3']")
	private static WebElement VoucherMappingTab_TableRow4AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont4']")
	private static WebElement VoucherMappingTab_TableRow5AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont5']")
	private static WebElement VoucherMappingTab_TableRow6AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont6']")
	private static WebElement VoucherMappingTab_TableRow7AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont7']")
	private static WebElement VoucherMappingTab_TableRow8AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont8']")
	private static WebElement VoucherMappingTab_TableRow9AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont9']")
	private static WebElement VoucherMappingTab_TableRow10AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont10']")
	private static WebElement VoucherMappingTab_TableRow11AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont11']")
	private static WebElement VoucherMappingTab_TableRow12AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont12']")
	private static WebElement VoucherMappingTab_TableRow13AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont13']")
	private static WebElement VoucherMappingTab_TableRow14AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont14']")
	private static WebElement VoucherMappingTab_TableRow15AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont15']")
	private static WebElement VoucherMappingTab_TableRow16AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont16']")
	private static WebElement VoucherMappingTab_TableRow17AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont17']")
	private static WebElement VoucherMappingTab_TableRow18AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont18']")
	private static WebElement VoucherMappingTab_TableRow19AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont19']")
	private static WebElement VoucherMappingTab_TableRow20AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont20']")
	private static WebElement VoucherMappingTab_TableRow21AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont21']")
	private static WebElement VoucherMappingTab_TableRow22AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont22']")
	private static WebElement VoucherMappingTab_TableRow23AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont23']")
	private static WebElement VoucherMappingTab_TableRow24AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont24']")
	private static WebElement VoucherMappingTab_TableRow25AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont25']")
	private static WebElement VoucherMappingTab_TableRow26AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont26']")
	private static WebElement VoucherMappingTab_TableRow27AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont27']")
	private static WebElement VoucherMappingTab_TableRow28AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont28']")
	private static WebElement VoucherMappingTab_TableRow29AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont29']")
	private static WebElement VoucherMappingTab_TableRow30AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont30']")
	private static WebElement VoucherMappingTab_TableRow31AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont31']")
	private static WebElement VoucherMappingTab_TableRow32AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont32']")
	private static WebElement VoucherMappingTab_TableRow33AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont33']")
	private static WebElement VoucherMappingTab_TableRow34AccountToPostTxt;
	
	@FindBy(xpath="//select[@id='ddDebitAccont34']")
	private static WebElement VoucherMappingTab_TableRow35AccountToPostTxt;
	
	
	// GST Field Mapping Table List
	
	@FindBy(xpath="//tbody[@id='savevoucher']/tr/td[3]/input")
	private static List<WebElement> VoucherMappingTab_TableGSTFieldsTxtList;
	
	@FindBy(xpath="//tbody[@id='savevoucher']/tr/td[4]/select")
	private static List<WebElement> VoucherMappingTab_TableFocusFieldsDrpDwnList;
	
	@FindBy(xpath="//tbody[@id='savevoucher']/tr/td[5]/select")
	private static List<WebElement> VoucherMappingTab_TableAccountToPostDrpDwnList;
	
	
	
	
	// Elements Under Information Tab of GST Preferences
	
	@FindBy(xpath="//input[@id='sContactname']")
	private static WebElement InformationTab_ContactPersonNameTxt;
	
	@FindBy(xpath="//input[@id='Mbl_No']")
	private static WebElement InformationTab_MobileNumberTxt;
	
	@FindBy(xpath="//input[@id='Phn_No']")
	private static WebElement InformationTab_PhoneNumberTxt;
	
	@FindBy(xpath="//textarea[@id='TextArea1']")
	private static WebElement InformationTab_Address1Txt;
	
	@FindBy(xpath="//textarea[@id='sAddress2']")
	private static WebElement InformationTab_Address2Txt;
	
	@FindBy(xpath="//select[@id='ddCity']")
	private static WebElement InformationTab_CityDrpDwn;
	
	@FindBy(xpath="//input[@id='txt_district']")
	private static WebElement InformationTab_DistrictTxt;
	
	@FindBy(xpath="//input[@id='txt_email']")
	private static WebElement InformationTab_EmailTxt;
	
	
	
	// Elements Under E-Way Bill Tab of GST Preferences
	
	@FindBy(xpath="//input[@id='chkShowContent']")
	private static WebElement EWayBillTab_EnableEWayBillChkbox;
	
	@FindBy(xpath="//input[@id='chkDisplyEwaybill']")
	private static WebElement EWayBillTab_DisplayEWayBillInVoucherChkbox;
	
	@FindBy(xpath="//input[@id='chkExpImpPointPIN']")
	private static WebElement EWayBillTab_CreateExpOrImpPointAsMasterChkbox;
	
	@FindBy(xpath="//input[@id='rbNone']")
	private static WebElement EWayBillTab_NoneRadio;
	
	@FindBy(xpath="//input[@id='chkSendSMS']")
	private static WebElement EWayBillTab_SendSMSForEWayBillRadio;
	
	@FindBy(xpath="//input[@id='rbonline']")
	private static WebElement EWayBillTab_OnlineRadio;
	
	
	
	
	// Elements Under E-Invoice Tab of GST Preferences
	
	@FindBy(xpath="//input[@id='chkenableEinvoice']")
	private static WebElement EInvoiceTab_EnableEInvoiceChkbox;
	
	@FindBy(xpath="//li[@id='Voucher_neweinv']/span/i")
	private static WebElement EInvoiceTab_NewBtn;
	
	@FindBy(xpath="//select[@id='ddVoucherTypeeinv']")
	private static WebElement EInvoiceTab_VoucherTypeDrpDwn;
	
	@FindBy(xpath="//div[@id='Vocherbtndisplay']/div/span")
	private static WebElement EInvoiceTab_DisplayBtn;
	
	@FindBy(xpath="//li[@id='btnSaveEInvMapping']/span/i")
	private static WebElement EInvoiceTab_SaveBtn;
	
	@FindBy(xpath="//li[@id='VocherCanceleinv']/span/i")
	private static WebElement EInvoiceTab_CancelBtn;
	
	
	
	
	// GSTR1 Online
	
	@FindBy(xpath="//input[@id='ChkQuarterly']")
	private static WebElement GSTR1Online_QuarterlyChkbox;
	
	@FindBy(xpath="//select[@id='ddDatemonth']")
	private static WebElement GSTR1Online_ForTheQuarterDrpDwn;
	
	// Sales Metrics
	
	@FindBy(xpath="//div[@id='getdata']/div[1]/div/div[1]/h4/span[2]/i")
	private static WebElement GSTR1Online_SalesMetricsExpandBtn;
	
	@FindBy(xpath="//input[@id='UnSyncInvoices']")
	private static WebElement GSTR1Online_UnSyncedInvoicesTxt;
	
	@FindBy(xpath="//input[@id='SyncInvoices']")
	private static WebElement GSTR1Online_SyncedinvoiceTxt;
	
	@FindBy(xpath="//input[@id='Reversechargeinvoice']")
	private static WebElement GSTR1Online_ReverseChargeInvoiceTxt;
	
	
	
	
	
	// GSTR2A
	
	
	
	
	
















}
