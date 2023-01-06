package com.focus.Pages;

import java.awt.AWTException;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashSet;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.omg.PortableServer.THREAD_POLICY_ID;
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
import com.testautomationguru.utility.PDFUtil;

import freemarker.template.SimpleDate;

public class RulesPage extends BaseEngine 
{
	
	@FindBy(xpath="//*[@id='id_search_menu']/input")
	private static WebElement searchMenuTxt;
	
	@FindBy(xpath="//*[@id='flexSwitchCheckDefault']")
	private static WebElement searchMenuBtn;
	
	
	//@FindBy(xpath="//div[contains(text(),'Settings')]")
	@FindBy(xpath="//*[@id='navigation_menu']/li[8]")
	private static WebElement  settingsmenuBtn;

	@FindBy(xpath="//span[contains(text(),'Configure Transactions')]")
	private static WebElement  configureTransactionBtn;

	@FindBy(xpath="//*[@id='tbPreference']")
	private static WebElement  preferencesBtn;

	//@FindBy(xpath="//div[@id='docCustomization']")
	//@FindBy(xpath="//*[@id='configureTransactions']/div[2]/div[1]/div/div/ul/li[2]/a")
	@FindBy(xpath="//*[@id='configureTransactions']/div/div[1]/div/div/ul/li[2]/a/i")
	private static WebElement  documentCustomization45Btn;

	@FindBy(xpath="//*[@id='doc_Searchbox']")
	private static WebElement  searchBox;

	@FindBy(xpath="//span[@class='icon-general icon-font6 theme_button_color']")
	private static WebElement  createVoucherBtn;

	@FindBy(xpath="//span[@id='spanDeleteVoucher']")
	private static WebElement  deletevoucherBtn;

	@FindBy(xpath="//*[@id='4609']")
	private static WebElement  recepitsVATVoucher;

	@FindBy(xpath="//a[@id='4865']")
	private static WebElement  paymentsVATVoucher;


	@FindBy(xpath="//span[contains(text(),'Receipts VAT')]")
	private static WebElement  recepitsVATMenu;


	@FindBy(xpath="//a[@id='4864']")
	private static WebElement  paymentsBtn;

	@FindBy(xpath="//a[@id='5120']")
	private static WebElement  pettyCashBtn;

	@FindBy(xpath="//a[@id='5888']")
	private static WebElement  postDatedrecepitsVATVoucher;

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

	@FindBy(xpath="//*[@id='updateButton']")
	private static WebElement  updateBtn;

	@FindBy(xpath="//a[@id='btnCustomizeClose']")
	private static WebElement  closeBtn;

	@FindBy(xpath="//*[@id='idGlobalError']/div/div[2]")
	public static WebElement errorMessage;

	@FindBy(xpath="//*[@id='idGlobalError']/div/div[1]/button")
	public static WebElement errorMessageCloseBtn;


	@FindBy(xpath="//input[@id='donotshow']")
	private static WebElement DontShowChkBox;

	@FindBy(xpath="//span[@class='pull-right']")
	private static WebElement PopUpcloseBtn;





	//Dash board Page
	@FindBy(xpath="//*[@id='dashName']")
	private static WebElement labelDashboard ;

	@FindBy(xpath="//*[@id='Select_dash']")
	private static WebElement selectDashboard ;

	@FindBy(xpath="//*[@id='Dashboard_AddDash']")
	private static WebElement newAddDashBoard;

	@FindBy(xpath="//*[@id='Dashboard_Dash_Config']")
	private static WebElement dashboardCustomizationSettings;



	@FindBy(xpath="//input[@id='donotshow']")
	private static WebElement doNotShowCheckbox;

	@FindBy(xpath="//span[@class='pull-right']")
	private static WebElement closeBtnInDemoPopupScreen;




	public static void checkPopUpWindow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		Thread.sleep(5000);

		try
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(doNotShowCheckbox));
			doNotShowCheckbox.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtnInDemoPopupScreen));
			closeBtnInDemoPopupScreen.click(); 

			System.err.println("POP UP DISPLAYED AND CLOSED SUCCESSFULLY");

		} 
		catch (Exception e)
		{
			System.err.println("NO POP UP DISPLAYED");
		}


		Thread.sleep(4000);
	}




	public static String checkValidationMessage(String ExpMessage) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		try
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
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

	//@FindBy(xpath="//input[@id='chkNewRecord_Rule_DocCustRules']")
	@FindBy(xpath="//*[@id='chkListApplyOn_DocCustRules']/div/div[2]/div/label/span")
	private static WebElement rulesNewRecordChekbox ;
	
	@FindBy(xpath="//*[@id='chkListApplyOn_DocCustRules']/div/div[2]/div/label/input")
	private static WebElement rulesNewRecordChekboxSelected ;

	//@FindBy(xpath="//input[@id='chkEdit_Rule_DocCustRules']")
	@FindBy(xpath="//*[@id='chkEdit_Rule_DocCustRules']/following-sibling::span")
	private static WebElement rulesEditChekbox ;
	
	@FindBy(xpath="//*[@id='chkListApplyOn_DocCustRules']/div/div[3]/div/div/label//input[1]")
	private static WebElement rulesEditChekboxxSelected ;

	//@FindBy(xpath="//input[@id='chkLoad_Rule_DocCustRules']")
	@FindBy(xpath="//*[@id='chkListEvaluateOn_DocCustRules']/div/div[1]/div/label/span")
	private static WebElement rulesLoadChekbox ;
	
	@FindBy(xpath="//*[@id='chkListEvaluateOn_DocCustRules']/div/div[1]/div/label/input")
	private static WebElement rulesLoadChekboxSelected ;

	//@FindBy(xpath="//input[@id='chkOnLeave_Rule_DocCustRules']")
	@FindBy(xpath="//*[@id='chkListEvaluateOn_DocCustRules']/div/div[2]/div/label/span")
	private static WebElement rulesOnLeaveChekbox ;

	@FindBy(xpath="//input[@id='chkBeforeDelete_Rule_DocCustRules']/following-sibling::span")
	private static WebElement rulesBeforeDeleteChekbox ;

	@FindBy(xpath="//*[@id='chkOnEnter_Rule_DocCustRules']/following-sibling::span")
	private static WebElement rulesOnEnterChekbox ;

	@FindBy(xpath="//*[@id='chkListStatus_DocCustRules']/div/div/div/label/span")
	private static WebElement rulesActiveChekbox ;
	
	@FindBy(xpath="//*[@id='chkListStatus_DocCustRules']/div/div/div/label/input")
	private static WebElement rulesActiveChekboxSelected ;


	@FindBy(xpath="//*[@id='chkNoRuleCondition_DocCustRules']/following-sibling::span")
	private static WebElement rulesNoConditionChekbox;


	//

	 @FindBy(xpath="//*[@id='tab_Rule_IF_DocCustRules']/a")
 	 private static WebElement rulesIfTab ;
     
     @FindBy(xpath=" //*[@id='Conditions-Context_Menu_DocCustRules']/li[2]/a")
 	 private static WebElement rulesElseIfTab ;
     
  


	@FindBy(xpath="//*[@id='tab_Rule_Else_DocCustRules']/a")
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

	@FindBy(xpath="//*[@id='tab_Rule_IF_Messaging_DocCustRules']/a")
	private static WebElement rulesIFMessageTab ;

	@FindBy(xpath="//input[@id='txtGeneralMsg_RuleIfMsg_DocCustRules']")
	private static WebElement rulesIFMessagesGenralMessageTxt;

	@FindBy(xpath="//*[@id='lstRdoMessageType_RuleIF_DocCustRules']/div[1]/div/input")
	private static WebElement rulesIFMessagesInformationRadio;

	@FindBy(xpath="//*[@id='lstRdoMessageType_RuleIF_DocCustRules']/div[2]/div/input")
	private static WebElement rulesIFMessagesWarnAndProceddRadio;

	@FindBy(xpath="//*[@id='lstRdoMessageType_RuleIF_DocCustRules']/div[3]/div/input")
	private static WebElement rulesIFMessagesWarnAndStopRadio;


	///// IF Tab ****Alerts-----Grid Elements

	@FindBy(xpath="//*[@id='tab_Rule_IF_Alerts_DocCustRules']/a")
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
	private static WebElement restricttheitemFromBeingRepeatedInEnChkBox;

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



	////////////////////////////	






	//**--------------------------------------------------------------------------------------------------------------


	//LOGOUT_LOGIN 


	@FindBy(xpath="//*[@id='txtUsername']")
	private static WebElement username;

	@FindBy(id="txtPassword")
	private static WebElement password;

	@FindBy(id="btnSignin")
	private static WebElement signIn;

	@FindBy(id="ddlCompany")
	private static WebElement companyDropDownList;

	@FindBy(xpath="//*[@id='id_mainlayoutmenu']//*[contains(text(),'SU')]")
	private static WebElement userNameDisplay;


	/* @FindBy(xpath="//*[@id='mainHeader_MainLayout']/nav/div/ul/li[7]/a/span")
			private static WebElement userNameDisplay;*/

	@FindBy(xpath="//*[@id='companyLogo']")
	private static WebElement companyLogo;

	@FindBy(xpath="//*[@id='ulCompanyDetails_HomePage']/li[1]")
	private static WebElement companyName;

	@FindBy(xpath="//*[@id='id_user_profile_icons']/li[2]/a")
	private static WebElement logoutOption;


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


		Thread.sleep(4000);
	}

	public static boolean checkBackgroundSavingMessage(String docno) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
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
			expMessage.add("Saving in background.");

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

	private static String xlfile = getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlSheetName = "RulesPage";
	

	public boolean CheckNavigateToRulesTAB() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		/*
		    LoginPage lp=new LoginPage(getDriver()); 

			String unamelt="su";

			String pawslt="su";

			lp.enterUserName(unamelt);

			lp.enterPassword(pawslt);

			lp.clickOnSignInBtn();

		    Thread.sleep(8000);*/
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTxt));
		searchMenuTxt.click();
		searchMenuTxt.sendKeys("Configure Transactions");
		searchMenuTxt.sendKeys(Keys.ENTER);
		
	
		Thread.sleep(3000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionBtn));
		configureTransactionBtn.click();
*/
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));
		documentCustomization45Btn.click();*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATVoucher));
		recepitsVATVoucher.click();

		Thread.sleep(1999);
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentsTab));
		documentsTab.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toogleExpandBtn));
		toogleExpandBtn.click();

		Thread.sleep(1999);
*/		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesBtn));
		rulesBtn.click();

		Thread.sleep(1999);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickhereToAddRuleBtn));
		clickhereToAddRuleBtn.click();

		Thread.sleep(1999);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleNameTxt));
		String actruleNameTxt=Boolean.toString(ruleNameTxt.isDisplayed());
		String expruleNameTxt=excelReader.getCellData(xlSheetName, 8, 6);
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actruleNameTxt);

		System.out.println("ruleNameTxt : "+actruleNameTxt +" Value Exp :"+expruleNameTxt);

		if(actruleNameTxt.equalsIgnoreCase(expruleNameTxt))
		{
			System.out.println("Test Pass : Rules Screen Displayed ");
			excelReader.setCellData(xlfile, xlSheetName, 8, 8, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail : Rules Screen Displayed ");
			excelReader.setCellData(xlfile, xlSheetName, 8, 8, resFail);
			return false;
		}
	}






	public void NavigateToRulesScreen() throws InterruptedException
	{
		getDriver().navigate().refresh();
		
		Thread.sleep(2000);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionBtn));
		configureTransactionBtn.click();*/
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTxt));
		searchMenuTxt.click();
		searchMenuTxt.sendKeys("Configure Transactions");
		searchMenuTxt.sendKeys(Keys.ENTER);
		
	
		Thread.sleep(2000);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));
		documentCustomization45Btn.click();*/


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATVoucher));
		recepitsVATVoucher.click();

		/*Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentsTab));
		documentsTab.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toogleExpandBtn));
		toogleExpandBtn.click();
*/
		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesBtn));
		rulesBtn.click();
		
		

		Thread.sleep(1999);
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickhereToAddRuleBtn));
		clickhereToAddRuleBtn.click();*/


		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleNameTxt));
		boolean actruleNameTxt=ruleNameTxt.isDisplayed();
		boolean expruleNameTxt=true;

		System.out.println(" ruleNameTxt :"+actruleNameTxt +" Value Exp :"+expruleNameTxt);
	}

	@FindBy(xpath = "//div[@id='id_focus_msgbox_main']")
	public static WebElement ServerErrorPopup;

	@FindBy(xpath = "//div[@id='id_focus_msgbox_title']/div[2]/span")
	public static WebElement ServerErrorPopupCloseBtn;

	@FindBy(xpath = "//*[@id='id_focus_msgbox_detail']/span/h2/i")
	public static WebElement ServerErrorPopupTxt;
	

	public boolean checkSavingRuleWithOutInput() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveRuleBtn));
		saveRuleBtn.click();
		//Thread.sleep(2000);
	    try {
			
				if(errorMessage.isDisplayed()) 
				{
					String actMessage=errorMessage.getText();
					System.out.println("Meeage Displayed : "+actMessage);
					
					errorMessageCloseBtn.click();
				}
				
				/*if(ServerErrorPopupCloseBtn.isDisplayed())
				{
					
					String actServerErrorPopupTxt=ServerErrorPopupTxt.getText();
					System.out.println("ServerErrorPopupTxt Displayed : "+actServerErrorPopupTxt);
					
					ServerErrorPopupCloseBtn.click();
					
				}*/
				
				
			} catch (Exception e) 
			{
				System.out.println(e.getMessage());
			}

		String ExpMessage=excelReader.getCellData(xlSheetName, 9, 6);
		String actMesage=checkValidationMessage(ExpMessage);
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, actMesage);

		if (actMesage.equalsIgnoreCase(ExpMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 9, 8, resPass);
			return true;

		} else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 9, 8, resFail);
			return false;

		}
	}
	
	
	@FindBy(xpath = "//input[@id='txtMasterRuleName_DocCustRules']")
	private static WebElement rulesTabRuleNameTxt;



	@FindBy(xpath = "//*[@id='Conditions-Context_Menu_DocCustRules']/li[2]/a")
	private static WebElement rulesTabElseTab;


	@FindBy(xpath = "//*[@id='Conditions-Context_Menu_DocCustRules']/li[1]/a")
	private static WebElement rulesTabIFTab;

	@FindBy(xpath = "//input[@id='txtField_RC_1_1_DocCustRules']")
	private static WebElement rulesTabIFSelectTabTxt;

	@FindBy(xpath = "//select[@id='ddlMasterFields_RULES_DocCustRules']")
	private static WebElement rulesTabIFSelectTabDrpdwn;

	@FindBy(xpath = "//select[@id='ddlExtraFields_1_1_DocCustRules']")
	private static WebElement rulesTabExtraFieldDrpdwn;

	@FindBy(xpath = "//select[@id='ddlExtraFields_2_1_DocCustRules']")
	private static WebElement rulesTabRow2ExtraFieldDrpdwn;

	@FindBy(xpath = "//select[@id='ddlExtraFields_1_2_DocCustRules']")
	private static WebElement rulesTabExtraFieldDrpdwn2;


	@FindBy(xpath = "//select[@id='ddlOperators_RC_1_1_DocCustRules']")
	private static WebElement rulesTabOperatorsDrpdwn;

	@FindBy(xpath = "//select[@id='ddlOperators_RC_2_1_DocCustRules']")
	private static WebElement rulesTabRow2OperatorsDrpdwn;

	@FindBy(xpath = "//select[@id='ddlCompareWith_RC_2_1_DocCustRules']")
	private static WebElement rulesTabRow2CompareWithDrpdwn;

	@FindBy(xpath = "//select[@id='ddlCompareWith_RC_1_1_DocCustRules']")
	private static WebElement rulesTabCompareWithDrpdwn;

	@FindBy(xpath = "//input[@id='txtCompareType_RC_1_1_DocCustRules']")
	private static WebElement rulesTabValueTxt;

	@FindBy(xpath = "//input[@id='txtCompareType_RC_2_1_DocCustRules']")
	private static WebElement rulesTabRow2ValueTxt;

	@FindBy(xpath = "//input[@id='txtCompareType_RC_1_1_DocCustRules']")
	private static WebElement rulesTabConditionTxttoEnter;

	@FindBy(xpath = "//input[@id='optControl_RulesCondition_DocCustRules']	")
	private static WebElement rulesTabConditionTxt;



	@FindBy(xpath = "//select[@id='ddlOperators_RC_1_2_DocCustRules']")
	private static WebElement rulesTabOperatorsDrpdwn2;

	@FindBy(xpath = "//select[@id='ddlCompareWith_RC_1_2_DocCustRules']")
	private static WebElement rulesTabCompareWithDrpdwn2;

	@FindBy(xpath = "//input[@id='txtCompareType_RC_1_2_DocCustRules']")
	private static WebElement rulesTabValueTxt2;


	@FindBy(xpath = "//input[@id='txtCompareType_RC_1_2_DocCustRules']")
	private static WebElement rulesTabConditionTxttoEnter2;




	@FindBy(xpath = "//input[@id='optControl_RulesCondition_DocCustRules']")
	private static WebElement rulesTabConditioDrpdwn;




	@FindBy(xpath = "//select[@id='ddlConjunction_RC_1_1_DocCustRules']")
	private static WebElement rulesTabConjunctionDrpdwn;

	@FindBy(xpath = "//input[@id='chkNoRuleCondition_DocCustRules']")
	private static WebElement rulesTabNoConditionChkbbox;

	@FindBy(xpath = "//i[@id='iconRuleSettings_DocCustRules']")
	private static WebElement rulesTabSettingsAddIcon;

	@FindBy(xpath = "//span[@class='icon-close icon-font7']")
	private static WebElement rulesTabIfConditionCloseBtn;

	@FindBy(xpath = "//*[@id='Rule_IF_DocCustRules']/div[2]/div[1]/ul/li[1]/a/span")
	private static WebElement rulesTab_IFFormattingTab;

	@FindBy(xpath = "//*[@id='Rule_IF_DocCustRules']/div[2]/div[1]/ul/li[2]/a/span")
	private static WebElement rulesTab_IFMesssageTab;

	@FindBy(xpath = "//*[@id='Rule_IF_DocCustRules']/div[2]/div[1]/ul/li[3]/a/span")
	private static WebElement rulesTab_IFAlertsTab;

	@FindBy(xpath = "//*[@id='Rule_Else_DocCustRules']/div/div[1]/ul/li[1]/a/span")
	private static WebElement rulesTab_ElseFormattingTab;

	@FindBy(xpath = "//*[@id='Rule_Else_DocCustRules']/div/div[1]/ul/li[2]/a/span")
	private static WebElement rulesTab_ElseMessageTab;

	@FindBy(xpath = "//*[@id='Rule_Else_DocCustRules']/div/div[1]/ul/li[3]/a/span")
	private static WebElement rulesTab_ElseAlertsTab;

	@FindBy(xpath = "//button[@id='btnAddRule']")
	private static WebElement rulesTab_SaveRuleBtn;

	@FindBy(xpath = "//button[contains(text(),'Delete Rule')]")
	private static WebElement rulesTab_DeleteRuleBtn;

	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-1']")
	private static WebElement rulesGrid1stRow_1stcol ;


	@FindBy(xpath="//td[@id='tblFormating_ELSEIF_Rule_DocCustRules_2_col_1-1']")
	private static WebElement ElseIFrulesGrid1stRow_1stcol ;


	@FindBy(xpath="//td[@id='tblFormating_ELSEIF_Rule_DocCustRules_2_col_1-2']")
	private static WebElement ElseIFrulesGrid1stRow_2ndcol; 

	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-2']")
	private static WebElement rulesGrid1stRow_2ndcol ;


	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-4']")
	private static WebElement rulesGrid1stRow_4thcol ;

	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-5']")
	private static WebElement rulesGrid1stRow_5thcol ;
	
	
	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-10']")
	private static WebElement rulesGrid1stRow_10thcol ;

	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_2-1']")
	private static WebElement rulesGrid2ndRow_1stcol ;

	@FindBy(xpath=" //input[@id='txtField_RC_2_1_DocCustRules']")
	private static WebElement  condition2ndRow_1stcol ;


	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_2-4']")
	private static WebElement rulesGrid2ndRow_4thcol ;

	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_2-5']")
	private static WebElement rulesGrid2ndRow_5thcol ;	    

	@FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-1']")
	private static WebElement rulesGridElse1stRow_1stcol ;

	@FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_2-1']")
	private static WebElement rulesGridElse2ndRow_1stcol ;

	@FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_2-4']")
	private static WebElement rulesGridElse2ndRow_4thcol ;

	@FindBy(xpath="//td[@id='tblFormating_Else_Rule_DocCustRules_col_2-5']")
	private static WebElement rulesGridElse2ndRow_5thcol ;

	@FindBy(xpath="//select[@id='ddlFieldsDiv_Rule_IF_DocCustRules']")
	private static WebElement rulesIFEnterFiled ;

	@FindBy(xpath="//select[@id='ddlFieldsDiv_Rule_ELSE_DocCustRules']")
	private static WebElement rulesElseEnterFiled ;



	@FindBy(xpath="//select[@id='ddlFieldsDiv_Rule_ELSEIF_DocCustRules_2']")
	private static WebElement ElseIFrulesIFEnterFiled ;

	@FindBy(xpath="//select[@id='ddlFieldsDiv_Rule_ELSEIF_DocCustRules']")
	private static WebElement ElseIFrulesElseEnterFiled ;



	@FindBy(xpath="//select[@id='ddlChangeValueFormat_ELSEIF_DocCustRules_2']")
	private static WebElement ElseIFrulesIFEnterChangeValueDrpdwn;

	@FindBy(xpath="//select[@id='ddlChangeValueFormat_IF_DocCustRules']")
	private static WebElement rulesIFEnterChangeValueDrpdwn;

	@FindBy(xpath="//select[@id='ddlChangeValueFormat_ELSE_DocCustRules']")
	private static WebElement rulesElseEnterChangeValueDrpdwn;


	@FindBy(xpath="//input[@id='txtCaptionFormat_IF_DocCustRules']")
	private static WebElement rulesIFEnterCaptionTXT;


	@FindBy(xpath="//input[@id='txtTextField']")
	private static WebElement rulesIFEnterValue;

	@FindBy(xpath="//input[@id='ValuesDivFormat_IF_DocCustRules']")
	private static WebElement rulesIFEnterTaxcodeValue;


	@FindBy(xpath="//input[@id='ValuesDivFormat_ELSEIF_DocCustRules_2']")
	private static WebElement rulesElseIFEnterTaxcodeValue;

	@FindBy(xpath="//input[@id='txtCaptionFormat_ELSEIF_DocCustRules_2']")
	private static WebElement rulesElseIFEnterCAPTIONValue;


	public boolean checkSavingRulesWithOnLoadOptionEnableInVoucherLevel() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleNameTxt));
		ruleNameTxt.click();

		rulesTabRuleNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);

		rulesTabRuleNameTxt.sendKeys(Keys.BACK_SPACE);		

		Thread.sleep(1000);
		
		rulesTabRuleNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		
		Thread.sleep(1999);

		ruleNameTxt.sendKeys(Keys.TAB);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesNewRecordChekbox));
		rulesNewRecordChekbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesLoadChekbox));
		rulesLoadChekbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesActiveChekbox));
		rulesActiveChekbox.click();

		rulesTabIFSelectTabTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectTabDrpdwn));
		Select account=new Select(rulesTabIFSelectTabDrpdwn);
		account.selectByVisibleText(excelReader.getCellData(xlSheetName, 11, 5));
		rulesTabIFSelectTabDrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		rulesTabExtraFieldDrpdwn.click();

		rulesTabExtraFieldDrpdwn.sendKeys(Keys.ARROW_DOWN);
		rulesTabExtraFieldDrpdwn.sendKeys(Keys.TAB);

		rulesTabOperatorsDrpdwn.click();

		rulesTabOperatorsDrpdwn.sendKeys(Keys.ARROW_DOWN);

		rulesTabOperatorsDrpdwn.sendKeys(Keys.TAB);

		rulesTabValueTxt.click();

		Thread.sleep(1000);
		Select Value=new Select(rulesTabCompareWithDrpdwn);
		Value.selectByValue(excelReader.getCellData(xlSheetName, 12, 5));
		rulesTabCompareWithDrpdwn.sendKeys(Keys.TAB);
		rulesTabCompareWithDrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		rulesTabConditionTxt.sendKeys(Keys.SPACE);
		rulesTabConditionTxt.sendKeys(Keys.BACK_SPACE);
		Thread.sleep(2000);

		rulesTabConditionTxt.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));

		Thread.sleep(4000);

		rulesTabConditionTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid1stRow_1stcol));
		rulesGrid1stRow_1stcol.click();

		rulesIFEnterFiled.sendKeys(excelReader.getCellData(xlSheetName, 14, 5));
		rulesIFEnterFiled.sendKeys(Keys.TAB);
		rulesIFEnterChangeValueDrpdwn.sendKeys(excelReader.getCellData(xlSheetName, 15, 5));

		rulesIFEnterChangeValueDrpdwn.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFEnterTaxcodeValue));

		rulesIFEnterTaxcodeValue.sendKeys(Keys.TAB);

		Thread.sleep(1999);
		
		rulesIFEnterCaptionTXT.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabGrid1st_10thCol));
		rulesTabGrid1st_10thCol.click();

		rulesTabGridMandatoryChkbox.click();

		Thread.sleep(1999);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessageTab));
		rulesIFMessageTab.click();

		Thread.sleep(1999);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessagesGenralMessageTxt));
		rulesIFMessagesGenralMessageTxt.click();
		rulesIFMessagesGenralMessageTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		rulesIFMessagesGenralMessageTxt.sendKeys(excelReader.getCellData(xlSheetName, 17, 5));

		if (rulesIFMessagesInformationRadio.isSelected()==false)
		{
			rulesIFMessagesInformationRadio.click();
		}

		Thread.sleep(1999);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesElseTab));
		getAction().moveToElement(rulesElseTab).build().perform();
		rulesElseTab.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGridElse1stRow_1stcol));
		rulesGridElse1stRow_1stcol.click();

		rulesElseEnterFiled.sendKeys(excelReader.getCellData(xlSheetName, 18, 5));
		rulesElseEnterFiled.sendKeys(Keys.TAB);
		
		rulesElseEnterChangeValueDrpdwn.sendKeys(excelReader.getCellData(xlSheetName, 19, 5));
		rulesElseEnterChangeValueDrpdwn.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveRuleBtn));
		saveRuleBtn.click();

		String ExpMessage=excelReader.getCellData(xlSheetName, 10, 6);
		String actMessage=checkValidationMessage(ExpMessage);
		excelReader.setCellData(xlfile, xlSheetName, 10, 7, actMessage);
		
		//boolean SavedRule=checkSavedRule();

		if (actMessage.equalsIgnoreCase(ExpMessage) /*&& SavedRule==true*/)
		{
			excelReader.setCellData(xlfile, xlSheetName, 10, 8, resPass);
			return true;

		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 10, 8, resFail);
			return false;
		}
	}
	
	
	
	
	@FindBy(xpath="//*[@id='ddlExtraFields_1_1_DocCustRules']")
	private static WebElement ruleConditionRow1Of2ndColbox;

	@FindBy(xpath="//select[@id='ddlOperators_RC_1_1_DocCustRules']")
	private static WebElement ruleConditionRow1Of3rdColbox;

	@FindBy(xpath="//*[@id='ddlCompareWith_RC_1_1_DocCustRules']")
	private static WebElement ruleConditionRow1Of4thColbox;

	@FindBy(xpath="//input[@id='txtCompareType_RC_1_1_DocCustRules']")
	private static WebElement ruleConditionRow1Of5thColbox;


	public boolean checkSavedRuleWithLoadOptionEnable() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleNameTxt));
		ruleNameTxt.click();

		rulesTabRuleNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);

		rulesTabRuleNameTxt.sendKeys(Keys.BACK_SPACE);		

		Thread.sleep(1000);
		
		rulesTabRuleNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 20, 5));
		
		Thread.sleep(1999);

		ruleNameTxt.sendKeys(Keys.TAB);

		Thread.sleep(1999);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesNewRecordChekbox));
		String actrulesNewRecordChekbox=Boolean.toString(rulesNewRecordChekboxSelected.isSelected());
		String exprulesNewRecordChekbox=excelReader.getCellData(xlSheetName, 20, 6);
		excelReader.setCellData(xlfile, xlSheetName, 20, 7, actrulesNewRecordChekbox);

		System.out.println(" rulesNewRecordChekbox: "+actrulesNewRecordChekbox +" Value "+exprulesNewRecordChekbox);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesLoadChekbox));
		String actrulesLoadChekbox=Boolean.toString(rulesLoadChekboxSelected.isSelected());
		String exprulesLoadChekbox=excelReader.getCellData(xlSheetName, 21, 6);
		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actrulesLoadChekbox);
		
		System.out.println(" rulesLoadChekbox: "+actrulesLoadChekbox +" Value "+exprulesLoadChekbox);

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesActiveChekbox));
		String actrulesActiveChekbox=Boolean.toString(rulesActiveChekboxSelected.isSelected());
		String exprulesActiveChekbox=excelReader.getCellData(xlSheetName, 22, 6);
		excelReader.setCellData(xlfile, xlSheetName, 22, 7, actrulesActiveChekbox);

		System.out.println(" rulesActiveChekbox: "+actrulesActiveChekbox +" Value "+exprulesActiveChekbox);

		Thread.sleep(1999);
		
		rulesTabIFSelectTabTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectTabDrpdwn));
		Select account=new Select(rulesTabIFSelectTabDrpdwn);

		String  actrulesTabIFSelectTabDrpdwn=account.getFirstSelectedOption().getText();
		String exprulesTabIFSelectTabDrpdwn=excelReader.getCellData(xlSheetName, 23, 6);
		excelReader.setCellData(xlfile, xlSheetName, 23, 7, actrulesTabIFSelectTabDrpdwn);

		System.out.println("rulesTabIFSelectTabDrpdwn : "+actrulesTabIFSelectTabDrpdwn +" Value : "+exprulesTabIFSelectTabDrpdwn);

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleConditionRow1Of2ndColbox));
		Select con1=new Select(ruleConditionRow1Of2ndColbox);

		String  actruleConditionRow1Of2ndColbox=con1.getFirstSelectedOption().getText();
		String expruleConditionRow1Of2ndColbox=excelReader.getCellData(xlSheetName, 24, 6);
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actruleConditionRow1Of2ndColbox);
		
		System.out.println("ruleConditionRow1Of2ndColbox : "+actruleConditionRow1Of2ndColbox +" Value : "+expruleConditionRow1Of2ndColbox);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleConditionRow1Of3rdColbox));
		Select con2=new Select(ruleConditionRow1Of3rdColbox);

		String  actruleConditionRow1Of3rdColbox=con2.getFirstSelectedOption().getText();
		String expruleConditionRow1Of3rdColbox=excelReader.getCellData(xlSheetName, 25, 6);
		excelReader.setCellData(xlfile, xlSheetName, 25, 7, actruleConditionRow1Of3rdColbox);

		System.out.println("ruleConditionRow1Of3rdColbox : "+actruleConditionRow1Of3rdColbox +" Value : "+expruleConditionRow1Of3rdColbox);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleConditionRow1Of4thColbox));
		Select con3=new Select(ruleConditionRow1Of4thColbox);

		String  actruleConditionRow1Of4thColbox=con3.getFirstSelectedOption().getText();
		String expruleConditionRow1Of4thColbox=excelReader.getCellData(xlSheetName, 26, 6);
		excelReader.setCellData(xlfile, xlSheetName, 26, 7, actruleConditionRow1Of4thColbox);

		System.out.println("ruleConditionRow1Of4thColbox : "+actruleConditionRow1Of4thColbox +" Value : "+expruleConditionRow1Of4thColbox);

		Thread.sleep(1999);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleConditionRow1Of5thColbox));

		String  actruleConditionRow1Of5thColbox=ruleConditionRow1Of5thColbox.getAttribute("value");
		String expruleConditionRow1Of5thColbox=excelReader.getCellData(xlSheetName, 27, 6);
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actruleConditionRow1Of5thColbox);

		System.out.println("ruleConditionRow1Of5thColbox : "+actruleConditionRow1Of5thColbox +" Value : "+expruleConditionRow1Of5thColbox);

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid1stRow_1stcol));
		getAction().moveToElement(rulesGrid1stRow_1stcol).build().perform();
		String actrulesGrid1stRow_1stcol=rulesGrid1stRow_1stcol.getText();
		String exprulesGrid1stRow_1stcol=excelReader.getCellData(xlSheetName, 28, 6);
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actrulesGrid1stRow_1stcol);

		System.out.println(" rulesGrid1stRow_1stcol : "+actrulesGrid1stRow_1stcol +" Value : "+exprulesGrid1stRow_1stcol);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid1stRow_2ndcol));
		String actrulesGrid1stRow_2ndcol=rulesGrid1stRow_2ndcol.getText();
		String exprulesGrid1stRow_2ndcol=excelReader.getCellData(xlSheetName, 29, 6);
		excelReader.setCellData(xlfile, xlSheetName, 29, 7, actrulesGrid1stRow_2ndcol);

		System.out.println(" rulesGrid1stRow_2ndcol : "+actrulesGrid1stRow_2ndcol +" Value : "+exprulesGrid1stRow_2ndcol);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid1stRow_5thcol));
		String actrulesGrid1stRow_4thcol=rulesGrid1stRow_5thcol.getText();
		String exprulesGrid1stRow_4thcol=excelReader.getCellData(xlSheetName, 30, 6);
		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actrulesGrid1stRow_4thcol);

		System.out.println(" rulesGrid1stRow_4thcol : "+actrulesGrid1stRow_4thcol +" Value : "+exprulesGrid1stRow_4thcol);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid1stRow_10thcol));
		rulesGrid1stRow_10thcol.click();
		rulesTabGridMandatoryChkbox.click();
		
		String actrulesGrid1stRow_10thcol=rulesGrid1stRow_10thcol.getText();
		if(actrulesGrid1stRow_10thcol.equalsIgnoreCase("false"))
		{
			rulesTabGridMandatoryChkbox.click();
			
		}
		String exprulesGrid1stRow_10thcol=excelReader.getCellData(xlSheetName, 30, 6);
		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actrulesGrid1stRow_4thcol);

		System.out.println(" rulesGrid1stRow_4thcol : "+actrulesGrid1stRow_4thcol +" Value : "+exprulesGrid1stRow_4thcol);


		Thread.sleep(1999);	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveRuleBtn));
		saveRuleBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		Thread.sleep(2000);
		
		getDriver().navigate().refresh();

		if (actrulesNewRecordChekbox.equalsIgnoreCase(exprulesNewRecordChekbox) && actrulesLoadChekbox.equalsIgnoreCase(exprulesLoadChekbox) &&
				actrulesActiveChekbox.equalsIgnoreCase(exprulesActiveChekbox) && actrulesTabIFSelectTabDrpdwn.equalsIgnoreCase(exprulesTabIFSelectTabDrpdwn) &&
				actrulesGrid1stRow_1stcol.equalsIgnoreCase(exprulesGrid1stRow_1stcol) && 
				actrulesGrid1stRow_4thcol.equalsIgnoreCase(exprulesGrid1stRow_4thcol) && 

				actruleConditionRow1Of2ndColbox.equalsIgnoreCase(expruleConditionRow1Of2ndColbox) &&
				actruleConditionRow1Of3rdColbox.equalsIgnoreCase(expruleConditionRow1Of3rdColbox) &&
				actruleConditionRow1Of4thColbox.equalsIgnoreCase(expruleConditionRow1Of4thColbox) &&
				actruleConditionRow1Of5thColbox.equalsIgnoreCase(expruleConditionRow1Of5thColbox) ) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 20, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 20, 8, resFail);
			return false;
		}
	}
	
	
	
	
	

	@FindBy(xpath="//a[@id='btnAddNewCondition']")
	private static WebElement ruleAddCndBtn;

	@FindBy(xpath="//a[@id='btnRemoveCondition']")
	private static WebElement ruleRemoveCndBtn;

	//@FindBy(xpath="//*[@id='RulesDefinition_DocCustRules']/div[3]/ul[1]/li")
	//@FindBy(xpath="(//*[@id='RulesDefinition_DocCustRules']/div[2]//div[2])[4]//ul//li//a")
	@FindBy(xpath="//*[@id='RulesDefinition_DocCustRules']/div[3]//ul//li//a")
	private static List<WebElement> ruleIfElseList;

	@FindBy(xpath="//*[@id='id_header_4_table_body']/tr/td[2]")
	private static List<WebElement>  cashAndBAnkAccountList;



	@FindBy(xpath="//*[@id='id_body_16777305']")
	private static WebElement  enterReceiptsVATTaxCode;

	@FindBy(xpath="//input[@id='id_header_268435470']")
	private static WebElement  placeofSupplyTxt;

	@FindBy(xpath="//tbody[@id='id_header_268435470_table_body']/tr/td[2]")
	private static List<WebElement>  placeofSupplyList;


	@FindBy(xpath="//tbody[@id='id_header_268435471_table_body']/tr/td[2]")
	private static List<WebElement>  jurisdictionList;

	@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr/td[3]")
	private static List<WebElement>  billwiseAdjustBillsDocList;

	@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr/td[13]")
	private static List<WebElement>  billwiseAdjustBillsAdjustColumn;

	@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr/td[2]/input")
	private static List<WebElement>  billwiseAdjustBillsChkBoxList;

	@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr[1]/td[2]/input")
	private static WebElement  billrefAdjuBills1stChkbox;

	@FindBy(xpath="//*[@id='id_body_12_table_body']/tr/td[2]")
	private static List<WebElement> bodyAccountListInGrid;

	@FindBy(xpath="//*[@id='id_body_39_table_body']/tr/td[2]")
	private static List<WebElement> bodyCreditAccountListInGrid;

	@FindBy(xpath="//input[@id='id_body_16']")
	private static WebElement  enter_Amount;



	@FindBy(xpath="//input[@id='id_body_16777307']")
	private static WebElement  enterpayVATTaxCode;

	@FindBy(xpath="//*[@id='id_transactionentry_save']")
	private static WebElement  voucherSaveBtn;

	@FindBy(xpath="//div[@id='dvHomeTransClose']")  
	private static WebElement  voucherhomeCloseBtn;



	@FindBy(xpath="//input[@id='id_body_12']")
	private static WebElement  enter_AccountTxt;

	@FindBy(xpath="//*[@id='id_body_12_table_body']/tr")
	private static List<WebElement> openingBalAccountListInGrid;

	@FindBy(xpath="//input[@id='id_body_18']")
	private static WebElement  enter_DebitTxt;

	@FindBy(xpath="//input[@id='id_body_19']")
	private static WebElement  enter_CreditTxt;

	@FindBy(xpath="//*[@id='lblAccount']")
	private static WebElement  billRefPartyName;

	@FindBy(xpath="//table[@id='id_Adjustment_Grid']")
	private static WebElement billRefAdjustBillsGrid;

	@FindBy(xpath="//*[@id='id_Adjustment_Grid_body']/tr")
	private static List<WebElement> billRefAdjustBillsGridList;

	@FindBy(xpath="//input[@id='id_header_268435471']")
	private static WebElement  jurisdictionTxt;

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


	public boolean checkLOADRuleOptionAtVoucherLevel() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		System.out.println("*******************checkLOADRuleOptionAtVoucherLevel***************");

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATMenu));
		recepitsVATMenu.click();

		Thread.sleep(2000);


		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		//Thread.sleep(2000);

		checkValidationMessage("Screen opened");
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCashBankAccountTxt));
		newCashBankAccountTxt.click();

		newCashBankAccountTxt.sendKeys(Keys.SPACE);

		int cashAndBAnkAccountListCount  = cashAndBAnkAccountList.size();

		System.err.println("cashAndBAnkAccountListCount   : "+cashAndBAnkAccountListCount);

		for (int i = 0; i < cashAndBAnkAccountListCount; i++) 
		{
			String data		  = cashAndBAnkAccountList.get(i).getText();
			System.err.println(data);

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 31, 5)))
			{
				cashAndBAnkAccountList.get(i).click();

				break;
			}
		}

		newCashBankAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));

		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SHIFT,Keys.HOME,Keys.BACK_SPACE);
		departmentTxt.sendKeys(Keys.SPACE);

		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 32, 5));

		Thread.sleep(2000);

		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeofSupplyTxt));
		placeofSupplyTxt.click();
		placeofSupplyTxt.sendKeys(Keys.SPACE);

		int placeOfSupplyListCount=placeofSupplyList.size();

		System.err.println(placeOfSupplyListCount);

		for(int i=0 ; i < placeOfSupplyListCount ;i++)
		{
			String data=placeofSupplyList.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 33, 5)))
			{
				placeofSupplyList.get(i).click();

				break;
			}
		}

		placeofSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(Keys.SPACE);

		int jurisdictionListCount=jurisdictionList.size();

		System.err.println(jurisdictionListCount);

		for(int i=0 ; i < jurisdictionListCount ;i++)
		{
			String data=jurisdictionList.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 34, 5)))
			{
				jurisdictionList.get(i).click();

				break;
			}
		}

		jurisdictionTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
		enter_AccountTxt.click();
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 35, 5));


		getFluentWebDriverWait().until(ExpectedConditions.visibilityOfAllElements(bodyAccountListInGrid));
		int accountCount=bodyAccountListInGrid.size();

		System.err.println(accountCount);

		for(int i=0 ; i < accountCount ;i++)
		{
			String data=bodyAccountListInGrid.get(i).getText();


			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 36, 5)))
			{
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOfAllElements(bodyAccountListInGrid));
				bodyAccountListInGrid.get(i).click();

				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));	
		enter_AccountTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enterReceiptsVATTaxCode));	
		enterReceiptsVATTaxCode.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Amount));
		enter_Amount.sendKeys(excelReader.getCellData(xlSheetName, 37, 5));
		enter_Amount.sendKeys(Keys.TAB);


		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();

		Thread.sleep(1999);
		String docno= documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(entryPageBodySecoundColTxt));	
		String actBeforeLoadTxt=entryPageBodySecoundColTxt.getText();
		String expBeforeLoad=excelReader.getCellData(xlSheetName, 31, 6);
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actBeforeLoadTxt);

		System.out.println("Before Save Voucher: "+actBeforeLoadTxt +" Value Exp : "+expBeforeLoad);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));	
		voucherSaveBtn.click();

		Thread.sleep(2000);
		/*HashSet<String> actMessage = new HashSet<String>();

		for (int i = 0; i < 3; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
			
			if (errorMessage.isDisplayed()==true)
			{
				String Message=errorMessage.getText();
				System.out.println(" Message Displayed  : "+Message);
				errorMessageCloseBtn.click();
			}
			
		}
*/

	//	HashSet<String> expMessage = new HashSet<String>();
/*
		expMessage.add(excelReader.getCellData(xlSheetName, 32, 6));
		expMessage.add(excelReader.getCellData(xlSheetName, 33, 6)+docno);
		expMessage.add(excelReader.getCellData(xlSheetName, 34, 6));

		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actMessage.toString());
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
*/
	String expMessage1 = excelReader.getCellData(xlSheetName, 32, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		Thread.sleep(2000);
		String expMessage2 = excelReader.getCellData(xlSheetName, 33, 6);
		Thread.sleep(2000);
		String actMessage1 = checkValidationMessage(expMessage2);
		String expMessage3= excelReader.getCellData(xlSheetName, 34, 6) + docno;
		Thread.sleep(2000);
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

	//	if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		
		Thread.sleep(2500);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));	
		previousBtn.click();

		checkValidationMessage("Voucher Loaded Successfully");

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(entryPageBodySecoundColTxt));	
		String actAfterLoadTxt=entryPageBodySecoundColTxt.getText();
		String expAfterLoad=excelReader.getCellData(xlSheetName, 35, 6);
		excelReader.setCellData(xlfile, xlSheetName, 35, 7, actAfterLoadTxt);
		

		System.out.println("After Save Voucher: "+actAfterLoadTxt +" Value Exp : "+expAfterLoad);

		if(actMessage.equals(expMessage1)  && actMessage1.startsWith(expMessage2)&& actMessage1.endsWith(expMessage3)
				&& actBeforeLoadTxt.equalsIgnoreCase(expBeforeLoad) && 
				actAfterLoadTxt.equalsIgnoreCase(expAfterLoad))
		{
			excelReader.setCellData(xlfile, xlSheetName, 31, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 31, 8, resFail);
			return false;
		}	
	}



	@FindBy(xpath="//*[@id='id_transactionentry_previous']")
	private static WebElement  previousBtn;

	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_control_heading_2']/div[1]")
	private static WebElement entryPageBodySecoundColTxt;;


	@FindBy(xpath = "//input[@id='id_header_1']")
	private static WebElement documentNumberTxt;

	@FindBy(xpath="//input[@id='txtNewReference']")
	private static WebElement  billRefNewReferenceTxt;

	@FindBy(xpath="//*[@id='id_Pick']/a/span")
	private static WebElement  billRefPickIcon;		

	@FindBy(xpath="//*[@id='id_Ok']/a/span")
	private static WebElement  billRefOkBtn;		

	@FindBy(xpath="//*[@id='id_Cancel']/a/span")
	private static WebElement  billRefcancel;	


	public boolean checkSavingRulesWithOnLeave() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		NavigateToRulesScreen();

		Thread.sleep(1999);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleNameTxt));
		ruleNameTxt.click();
		
		//ruleNameTxt.clear();
		//Thread.sleep(2000);
		rulesTabRuleNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);

		rulesTabRuleNameTxt.sendKeys(Keys.BACK_SPACE);		

		Thread.sleep(1000);
		rulesTabRuleNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 38, 5));
		Thread.sleep(1999);

		ruleNameTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesNewRecordChekbox));
		rulesNewRecordChekbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesOnLeaveChekbox));
		rulesOnLeaveChekbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesActiveChekbox));
		rulesActiveChekbox.click();

		rulesTabIFSelectTabTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectTabDrpdwn));
		Select account=new Select(rulesTabIFSelectTabDrpdwn);
		account.selectByVisibleText(excelReader.getCellData(xlSheetName, 39, 5));
		rulesTabIFSelectTabDrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		rulesTabExtraFieldDrpdwn.click();

		rulesTabExtraFieldDrpdwn.sendKeys(Keys.ARROW_DOWN);
		rulesTabExtraFieldDrpdwn.sendKeys(Keys.TAB);

		rulesTabOperatorsDrpdwn.click();

		rulesTabOperatorsDrpdwn.sendKeys(Keys.ARROW_DOWN);

		rulesTabOperatorsDrpdwn.sendKeys(Keys.TAB);

		rulesTabValueTxt.click();

		Thread.sleep(1000);
		Select Value=new Select(rulesTabCompareWithDrpdwn);
		Value.selectByValue(excelReader.getCellData(xlSheetName, 40, 5));
		rulesTabCompareWithDrpdwn.sendKeys(Keys.TAB);
		rulesTabCompareWithDrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		rulesTabConditionTxt.sendKeys(Keys.SPACE);
		rulesTabConditionTxt.sendKeys(Keys.BACK_SPACE);
		Thread.sleep(2000);

		rulesTabConditionTxt.sendKeys(excelReader.getCellData(xlSheetName, 41, 5));

		Thread.sleep(4000);

		rulesTabConditionTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getAction().moveToElement(rulesIfTab).contextClick().build().perform();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleAddCndBtn));
		ruleAddCndBtn.click();

		int count=ruleIfElseList.size();

		ArrayList<String>ifelse=new ArrayList<String>();
		for (int i = 0; i < count; i++)
		{
			String data=ruleIfElseList.get(i).getText();
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 42, 5))) 
			{
				ruleIfElseList.get(i).click();
			} 
			ifelse.add(data);
		}
		String actruleIfElseList=ifelse.toString();
		String expruleIfElseList=excelReader.getCellData(xlSheetName, 38, 6);
		excelReader.setCellData(xlfile, xlSheetName, 38, 7, actruleIfElseList);

		System.out.println(" ruleIfElseList : "+actruleIfElseList +" Value  "+expruleIfElseList);

		Thread.sleep(2000);

		for (int i = 0; i < count; i++)
		{
			String data=ruleIfElseList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 43, 5))) 
			{
				ruleIfElseList.get(i).click();
			} 
		}
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid1stRow_1stcol));
		rulesGrid1stRow_1stcol.click();

		rulesIFEnterFiled.sendKeys(excelReader.getCellData(xlSheetName, 44, 5));
		rulesIFEnterFiled.sendKeys(Keys.TAB);
		rulesIFEnterChangeValueDrpdwn.sendKeys(excelReader.getCellData(xlSheetName, 45, 5));

		rulesIFEnterChangeValueDrpdwn.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFEnterTaxcodeValue));

		rulesIFEnterTaxcodeValue.sendKeys(Keys.TAB);


		Thread.sleep(1999);
		rulesIFEnterCaptionTXT.sendKeys(excelReader.getCellData(xlSheetName, 46, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabGrid1st_10thCol));
		rulesTabGrid1st_10thCol.click();

		rulesTabGridMandatoryChkbox.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessageTab));
		rulesIFMessageTab.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessagesGenralMessageTxt));
		rulesIFMessagesGenralMessageTxt.click();
		rulesIFMessagesGenralMessageTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		rulesIFMessagesGenralMessageTxt.sendKeys(excelReader.getCellData(xlSheetName, 47, 5));

		if (rulesIFMessagesInformationRadio.isSelected()==false)
		{
			rulesIFMessagesInformationRadio.click();
		}

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesElseTab));
		getAction().moveToElement(rulesElseTab).build().perform();
		rulesElseTab.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGridElse1stRow_1stcol));
		rulesGridElse1stRow_1stcol.click();

		rulesElseEnterFiled.sendKeys(excelReader.getCellData(xlSheetName, 48, 5));
		rulesElseEnterFiled.sendKeys(Keys.TAB);
		rulesElseEnterChangeValueDrpdwn.sendKeys(excelReader.getCellData(xlSheetName, 49, 5));

		rulesElseEnterChangeValueDrpdwn.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesElseTabGrid1st_10thCol));
		rulesElseTabGrid1st_10thCol.click();
		Thread.sleep(2000);

		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabGridElseMandatoryChkbox));
		rulesTabGridElseMandatoryChkbox.click();
		Thread.sleep(2000);


		Thread.sleep(1999);
		getAction().moveToElement(rulesIfTab).build().perform();
		rulesElseIfTab.click();

		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(elseIfSelectFiled));
		elseIfSelectFiled.click();
		Select account1=new Select(elseIfEnterSelectFiled);
		account1.selectByVisibleText(excelReader.getCellData(xlSheetName, 50, 5));
		elseIfEnterSelectFiled.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		rulesTabExtraFieldDrpdwn2.click();

		rulesTabExtraFieldDrpdwn2.sendKeys(Keys.ARROW_DOWN);
		rulesTabExtraFieldDrpdwn2.sendKeys(Keys.TAB);

		rulesTabOperatorsDrpdwn2.click();

		rulesTabOperatorsDrpdwn2.sendKeys(Keys.ARROW_DOWN);

		rulesTabOperatorsDrpdwn2.sendKeys(Keys.TAB);

		rulesTabValueTxt2.click();

		Thread.sleep(1000);
		Select Value1=new Select(rulesTabCompareWithDrpdwn2);
		Value1.selectByValue(excelReader.getCellData(xlSheetName, 51, 5));
		rulesTabCompareWithDrpdwn2.sendKeys(Keys.TAB);
		rulesTabCompareWithDrpdwn2.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		rulesTabConditionTxt.sendKeys(Keys.SPACE);
		rulesTabConditionTxt.sendKeys(Keys.BACK_SPACE);
		Thread.sleep(2000);

		rulesTabConditionTxt.sendKeys(excelReader.getCellData(xlSheetName, 52, 5));

		Thread.sleep(4000);

		rulesTabConditionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseIFrulesGrid1stRow_1stcol));
		ElseIFrulesGrid1stRow_1stcol.click();

		ElseIFrulesIFEnterFiled.sendKeys(excelReader.getCellData(xlSheetName, 53, 5));
		ElseIFrulesIFEnterFiled.sendKeys(Keys.TAB);
		ElseIFrulesIFEnterChangeValueDrpdwn.sendKeys(excelReader.getCellData(xlSheetName, 54, 5));

		ElseIFrulesIFEnterChangeValueDrpdwn.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesElseIFEnterTaxcodeValue));
		rulesElseIFEnterTaxcodeValue.sendKeys(Keys.TAB);

		Thread.sleep(1999);
		
		rulesElseIFEnterCAPTIONValue.sendKeys(excelReader.getCellData(xlSheetName, 55, 5));

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseIfrulesTabGrid1st_10thCol));
		ElseIfrulesTabGrid1st_10thCol.click();

		ElseIFrulesTabGridMandatoryChkbox.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveRuleBtn));
		saveRuleBtn.click();

		String ExpMessage=excelReader.getCellData(xlSheetName, 39, 6);
		String actMessage=checkValidationMessage(ExpMessage);
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actMessage);

		Thread.sleep(1999);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		checkValidationMessage("");

		if (actMessage.equalsIgnoreCase(ExpMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 38, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 38, 8, resFail);
			return false;
		}

	}

	
	
	
	

	public boolean checkOnLeaveValiationATVoucherLevelWithDepartmentWales() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getDriver().navigate().refresh();
		Thread.sleep(2000);

		System.out.println("*******************checkOnLeaveValiationATVoucherLevelWithDepartmentWales***************");

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATMenu));
		recepitsVATMenu.click();

		Thread.sleep(2000);    

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));

		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SHIFT,Keys.HOME,Keys.BACK_SPACE);
		departmentTxt.sendKeys(Keys.SPACE);

		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 56, 5));

		Thread.sleep(1000);

		departmentTxt.sendKeys(Keys.TAB);

		String Message = checkValidationMessage(excelReader.getCellData(xlSheetName, 56, 6));
		excelReader.setCellData(xlfile, xlSheetName, 56, 7, Message);
		
		Thread.sleep(1999);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(entryPageBodyFirstColTxt));
		String actAccountName=entryPageBodyFirstColTxt.getText();
		String expAccountName=excelReader.getCellData(xlSheetName, 57, 6);
		excelReader.setCellData(xlfile, xlSheetName, 57, 7, actAccountName);

		System.out.println("AccountName  : "+actAccountName +" Value Exp : "+expAccountName);

		if (actAccountName.equalsIgnoreCase(expAccountName))
		{
			excelReader.setCellData(xlfile, xlSheetName, 56, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 56, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	


	public boolean checkChangingDepartmentAsAMERICAAndValidateColounmHeadingName() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SHIFT,Keys.HOME,Keys.BACK_SPACE);
		departmentTxt.sendKeys(Keys.SPACE);

		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 58, 5));

		Thread.sleep(2000);

		departmentTxt.sendKeys(Keys.TAB);

		Thread.sleep(1999);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(entryPageBodyFirstColTxt));
		String actAccountName=entryPageBodyFirstColTxt.getText();
		String expAccountName=excelReader.getCellData(xlSheetName, 59, 6);
		excelReader.setCellData(xlfile, xlSheetName, 59, 7, actAccountName);

		System.out.println("AccountName  : "+actAccountName +" Value Exp : "+expAccountName);

		if (actAccountName.equalsIgnoreCase(expAccountName))
		{
			excelReader.setCellData(xlfile, xlSheetName, 58, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 58, 8, resFail);
			return false;
		}
	}

	
	
	
	

	public boolean checkBodyHeadingColumnWithCustomerAccount() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		System.out.println("*******************checkBodyHeadingColumnWithCustomerAccount***************");

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATMenu));
		recepitsVATMenu.click();

		Thread.sleep(2000);    

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("");

		Thread.sleep(2000);   
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(entryPageBodyFirstColTxt));
		String actBeforeLoadAccountName=entryPageBodyFirstColTxt.getText();
		String expBeforeLoadAccountName=excelReader.getCellData(xlSheetName, 59, 6);
		excelReader.setCellData(xlfile, xlSheetName, 59, 7, actBeforeLoadAccountName);

		System.out.println("**********BeforeLoadAccountName  : "+actBeforeLoadAccountName +" Value Exp : "+expBeforeLoadAccountName);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
		enter_AccountTxt.click();
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 59, 5));
		Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(entryPageBodyFirstColTxt));
		String actAccountName=entryPageBodyFirstColTxt.getText();
		String expAccountName=excelReader.getCellData(xlSheetName, 60, 6);
		excelReader.setCellData(xlfile, xlSheetName, 60, 7, actBeforeLoadAccountName);

		System.out.println("************After AccountName  : "+actAccountName +" Value Exp : "+expAccountName);

		if (actAccountName.equalsIgnoreCase(expAccountName) && actBeforeLoadAccountName.equalsIgnoreCase(expBeforeLoadAccountName))
		{
			excelReader.setCellData(xlfile, xlSheetName, 59, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 59, 8, resFail);
			return false;
		}
	}




	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_control_heading_1']/div[1]")
	private static WebElement entryPageBodyFirstColTxt;;



	@FindBy(xpath="//input[@id='txtField_RC_1_2_DocCustRules']")
	private static WebElement elseIfSelectFiled;

	@FindBy(xpath="//select[@id='ddlMasterFields_RULES_DocCustRules']")
	private static WebElement elseIfEnterSelectFiled;


	@FindBy(xpath = "//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-11']")
	private static WebElement rulesTabGrid1st_11thCol;

	@FindBy(xpath = "//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-10']")
	private static WebElement rulesTabGrid1st_10thCol;

	@FindBy(xpath = "//td[@id='tblFormating_IF_Rule_DocCustRules_col_2-10']")
	private static WebElement rulesTabGrid2nd_10thCol;

	@FindBy(xpath = "//td[@id='tblFormating_Else_Rule_DocCustRules_col_1-10']")
	private static WebElement rulesElseTabGrid1st_10thCol;


	
	@FindBy(xpath = "//td[@id='tblFormating_ELSEIF_Rule_DocCustRules_2_col_1-10']")
	private static WebElement ElseIfrulesTabGrid1st_10thCol;


	@FindBy(xpath = "//input[@id='chkMandatoryFormat_ELSEIF_DocCustRules_2']")
	private static WebElement ElseIFrulesTabGridMandatoryChkbox;


	@FindBy(xpath = "//input[@id='chkMandatoryFormat_IF_DocCustRules']")
	private static WebElement rulesTabGridMandatoryChkbox;	

	@FindBy(xpath = "//input[@id='chkMandatoryFormat_Else_DocCustRules']")
	private static WebElement rulesTabGridElseMandatoryChkbox;	
	
	@FindBy(xpath="//span[@class='dropdown icon-menuicon1 icon-font6 theme_color-inverse pull-right']")
	private static WebElement  toogleExpandBtn;

	@FindBy(xpath="//*[@id='navigationtab4']")
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

	//@FindBy(xpath="//*[@id='navHidetab11']")
	@FindBy(xpath="//*[@id='navigationtab11']")
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

	@FindBy(xpath="//input[@id='misc_currencyInputLocalExchangeRate']")
	private static WebElement inputLocalExchangeRateChkBox;

	@FindBy(xpath="//input[@id='misc_currencyPostExgRateDiff']")
	private static WebElement postExchangeRateDifferrenceChkBox;


	@FindBy(xpath="//div[@id='idGlobalError']")
	public static WebElement validationConfirmationMessage;

	@FindBy(xpath="//*[@id='navigationtab1']")
	private static WebElement  documentsTab;

	@FindBy(xpath="//li[@id='navigationtab1']")
	private static WebElement  documentsTabForMaterialReceiptNotes;


	public boolean chekSavingRuleWithOnEnter() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		getDriver().navigate().refresh();
		
		NavigateToRulesScreen();
		
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleNameTxt));
		ruleNameTxt.click();
		
		//ruleNameTxt.clear();
		
		rulesTabRuleNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);

		rulesTabRuleNameTxt.sendKeys(Keys.BACK_SPACE);		

		Thread.sleep(1000);
		
		rulesTabRuleNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 61, 5));
		
		Thread.sleep(1999);

		ruleNameTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesNewRecordChekbox));
		rulesNewRecordChekbox.click();
		/*Thread.sleep(1000);
		rulesNewRecordChekbox.sendKeys(Keys.TAB);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesOnEnterChekbox));
		rulesOnEnterChekbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesActiveChekbox));
		rulesActiveChekbox.click();

		rulesTabIFSelectTabTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectTabDrpdwn));
		Select account=new Select(rulesTabIFSelectTabDrpdwn);
		account.selectByVisibleText(excelReader.getCellData(xlSheetName, 62, 5));
		rulesTabIFSelectTabDrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		rulesTabExtraFieldDrpdwn.click();

		rulesTabExtraFieldDrpdwn.sendKeys(Keys.ARROW_DOWN);
		rulesTabExtraFieldDrpdwn.sendKeys(Keys.TAB);

		rulesTabOperatorsDrpdwn.click();

		rulesTabOperatorsDrpdwn.sendKeys(Keys.ARROW_DOWN);

		rulesTabOperatorsDrpdwn.sendKeys(Keys.TAB);

		rulesTabValueTxt.click();

		Thread.sleep(1000);
		
		Select Value=new Select(rulesTabCompareWithDrpdwn);
		Value.selectByValue(excelReader.getCellData(xlSheetName, 63, 5));
		rulesTabCompareWithDrpdwn.sendKeys(Keys.TAB);
		rulesTabCompareWithDrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		rulesTabConditionTxt.sendKeys(Keys.SPACE);
		rulesTabConditionTxt.sendKeys(Keys.BACK_SPACE);
		Thread.sleep(2000);

		rulesTabConditionTxt.sendKeys(excelReader.getCellData(xlSheetName, 64, 5));

		Thread.sleep(4000);

		rulesTabConditionTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid1stRow_1stcol));
		rulesGrid1stRow_1stcol.click();

		rulesIFEnterFiled.sendKeys(excelReader.getCellData(xlSheetName, 65, 5));
		rulesIFEnterFiled.sendKeys(Keys.TAB);
		rulesIFEnterChangeValueDrpdwn.sendKeys(excelReader.getCellData(xlSheetName, 66, 5));

		rulesIFEnterChangeValueDrpdwn.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFEnterValue));

		rulesIFEnterValue.sendKeys(excelReader.getCellData(xlSheetName, 67, 5));

		rulesIFEnterValue.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabGrid1st_10thCol));
		rulesTabGrid1st_10thCol.click();

		rulesTabGridMandatoryChkbox.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessageTab));
		rulesIFMessageTab.click();

		Thread.sleep(1999);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessagesGenralMessageTxt));
		rulesIFMessagesGenralMessageTxt.click();
		rulesIFMessagesGenralMessageTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		rulesIFMessagesGenralMessageTxt.sendKeys(excelReader.getCellData(xlSheetName, 68, 5));

		if (rulesIFMessagesInformationRadio.isSelected()==false)
		{
			rulesIFMessagesInformationRadio.click();
		}

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesElseTab));
		getAction().moveToElement(rulesElseTab).build().perform();
		rulesElseTab.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGridElse1stRow_1stcol));
		rulesGridElse1stRow_1stcol.click();

		rulesElseEnterFiled.sendKeys(excelReader.getCellData(xlSheetName, 69, 5));
		rulesElseEnterFiled.sendKeys(Keys.TAB);
		rulesElseEnterChangeValueDrpdwn.sendKeys(excelReader.getCellData(xlSheetName, 70, 5));

		rulesElseEnterChangeValueDrpdwn.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveRuleBtn));
		saveRuleBtn.click();


		String ExpMessage=excelReader.getCellData(xlSheetName, 61, 6);
		String actMessage=checkValidationMessage(ExpMessage);
		excelReader.setCellData(xlfile, xlSheetName, 61, 7, actMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		checkValidationMessage("data saved succesfully");

		Thread.sleep(1999);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();

		if (actMessage.equalsIgnoreCase(ExpMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 61, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 61, 8, resFail);
			return false;
		}
	}
	
	
	


	@FindBy(xpath="//input[@id='id_header_67108921']")
	private static WebElement recepitsVATChequeNo;

	public boolean checkVoucherLevelOnEnterRuleValidation() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		System.out.println("*******************checkVoucherLevelOnEnterRuleValidation***************");

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATMenu));
		recepitsVATMenu.click();

		Thread.sleep(2000);    

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		Thread.sleep(2000);   
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATChequeNo));
		String actBeforeLoadAccountName=Boolean.toString(recepitsVATChequeNo.getText().isEmpty());
		String expBeforeLoadAccountName=excelReader.getCellData(xlSheetName, 71, 6);
		excelReader.setCellData(xlfile, xlSheetName, 71, 7, actBeforeLoadAccountName);
		
		System.out.println("**********BeforeLoadAccountName  : "+actBeforeLoadAccountName +" Value Exp : "+expBeforeLoadAccountName);

		Thread.sleep(1999);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCashBankAccountTxt));
		newCashBankAccountTxt.click();

		newCashBankAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 71, 5));
		Thread.sleep(2000);
		newCashBankAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(1999);
		
		newCashBankAccountTxt.click();

		checkValidationMessage("ONEnterRuleActivated");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATChequeNo));
		String actAccountName=recepitsVATChequeNo.getAttribute("value");
		String expAccountName=excelReader.getCellData(xlSheetName, 72, 6);
		excelReader.setCellData(xlfile, xlSheetName, 72, 7, actAccountName);

		System.out.println("************ After AccountName : "+actAccountName +" Value Exp : "+expAccountName);

		if (actAccountName.equalsIgnoreCase(expAccountName) && actBeforeLoadAccountName.equalsIgnoreCase(expBeforeLoadAccountName))
		{
			excelReader.setCellData(xlfile, xlSheetName, 71, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 71, 8, resFail);
			return false;
		}
	}


	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-7']")
	private static WebElement rulesIfRow1ErrorMessageCol;

	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-11']")
	private static WebElement rulesIfRow1HideenCol;

	@FindBy(xpath="//input[@id='txtErrorMsgForamt_IF_DocCustRules']")
	private static WebElement ruleIFEnterErrorMessage;

	@FindBy(xpath="//input[@id='txtBannertextFormat_IF_DocCustRules']")
	private static WebElement ruleIFEnterBannerText;

	@FindBy(xpath="//input[@id='chkHiddenFormat_IF_DocCustRules']")
	private static WebElement ruleIFHidedenChkBox;
	
	@FindBy(xpath="//input[@id='chkHiddenFormat_IF_DocCustRules']/following-sibling::span")
	private static WebElement ruleIFHidedenChkBoxSelected;


	@FindBy(xpath="//select[@id='ddlConjunction_RC_1_1_DocCustRules']")
	private static WebElement rulesConjuctionDropdwn;

	@FindBy(xpath="//input[@id='txtRegExFormat_IF_DocCustRules']")
	private static WebElement ruleIFEnterRegularExpTxt;

	@FindBy(xpath="//td[@id='tblFormating_IF_Rule_DocCustRules_col_1-6']")
	private static WebElement rulesIfRow1RegularExpCol;


	public boolean checkEditingAndAddingSavedRuleUnderRecepitsVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		NavigateToRulesScreen();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleNameTxt));
		ruleNameTxt.click();

		rulesTabRuleNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);

		rulesTabRuleNameTxt.sendKeys(Keys.BACK_SPACE);		

		Thread.sleep(1000);
		
		rulesTabRuleNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 73, 5));
		
		Thread.sleep(1999);

		ruleNameTxt.sendKeys(Keys.TAB);

		Thread.sleep(1999);
		
		getAction().moveToElement(rulesIfRow1ErrorMessageCol).build().perform();
		rulesIfRow1ErrorMessageCol.click();


		ruleIFEnterErrorMessage.click();
		
		ruleIFEnterErrorMessage.sendKeys(excelReader.getCellData(xlSheetName, 74, 5));
		Thread.sleep(1999);
		ruleIFEnterErrorMessage.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabGrid1st_10thCol));
		rulesTabGrid1st_10thCol.click();

		rulesTabGridMandatoryChkbox.click();

		Thread.sleep(2000);
		Select s1=new Select(rulesConjuctionDropdwn);
		s1.selectByVisibleText(excelReader.getCellData(xlSheetName, 75, 5));

		Thread.sleep(2000);
		
		rulesConjuctionDropdwn.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(condition2ndRow_1stcol));
		condition2ndRow_1stcol.click();

		Select account=new Select(rulesTabIFSelectTabDrpdwn);
		account.selectByVisibleText(excelReader.getCellData(xlSheetName, 76, 5));
		rulesTabIFSelectTabDrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		rulesTabRow2ExtraFieldDrpdwn.click();

		rulesTabRow2ExtraFieldDrpdwn.sendKeys(Keys.ARROW_DOWN);
		rulesTabRow2ExtraFieldDrpdwn.sendKeys(Keys.TAB);

		rulesTabRow2OperatorsDrpdwn.click();

		rulesTabRow2OperatorsDrpdwn.sendKeys(Keys.ARROW_DOWN);

		rulesTabRow2OperatorsDrpdwn.sendKeys(Keys.TAB);

		rulesTabRow2ValueTxt.click();

		Thread.sleep(1000);
		Select Value=new Select(rulesTabRow2CompareWithDrpdwn);
		Value.selectByValue(excelReader.getCellData(xlSheetName, 77, 5));
		rulesTabRow2CompareWithDrpdwn.sendKeys(Keys.TAB);
		rulesTabRow2CompareWithDrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		rulesTabConditionTxt.sendKeys(Keys.SPACE);
		rulesTabConditionTxt.sendKeys(Keys.BACK_SPACE);
		Thread.sleep(2000);

		rulesTabConditionTxt.sendKeys(excelReader.getCellData(xlSheetName, 78, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid2ndRow_1stcol));
		rulesGrid2ndRow_1stcol.click();

		rulesIFEnterFiled.sendKeys(excelReader.getCellData(xlSheetName, 79, 5));
		rulesIFEnterFiled.sendKeys(Keys.TAB);
		rulesIFEnterChangeValueDrpdwn.sendKeys(excelReader.getCellData(xlSheetName, 80, 5));

		rulesIFEnterChangeValueDrpdwn.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFEnterValue));

		rulesIFEnterValue.sendKeys(excelReader.getCellData(xlSheetName, 81, 5));
		Thread.sleep(1999);

		rulesIFEnterValue.sendKeys(Keys.TAB);

		Thread.sleep(1999);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabGrid2nd_10thCol));
		rulesTabGrid2nd_10thCol.click();

		rulesTabGridMandatoryChkbox.click();


		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveRuleBtn));
		saveRuleBtn.click();

		String expMessage=excelReader.getCellData(xlSheetName, 73, 6);
		String actMesage=checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 73, 7, actMesage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		String expMessage1=excelReader.getCellData(xlSheetName, 74, 6);
		String actMesage1=checkValidationMessage(expMessage1);
		excelReader.setCellData(xlfile, xlSheetName, 74, 7, actMesage1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();

		if (actMesage.equalsIgnoreCase(expMessage) && actMesage1.equals(expMessage1))
		{
			excelReader.setCellData(xlfile, xlSheetName, 73, 8, resPass);
			return true;
		}
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 73, 8, resFail);
			return false;
		}
	}



	public boolean checkRuleConditionAfterEditingINEnterRule() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATMenu));
		recepitsVATMenu.click();

		Thread.sleep(2000);    

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCashBankAccountTxt));
		newCashBankAccountTxt.click();
		newCashBankAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 82, 5));

		Thread.sleep(1999);

		newCashBankAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		newCashBankAccountTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
		enter_AccountTxt.click();
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 83, 5));
		Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);

		select1stRow_1stColumn.click();

		checkValidationMessage("ONEnterRuleActivated");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATChequeNo));
		String actAccountName=recepitsVATChequeNo.getAttribute("value");
		String expAccountName=excelReader.getCellData(xlSheetName, 82, 6);
		excelReader.setCellData(xlfile, xlSheetName, 82, 7, actAccountName);

		System.err.println("************After AccountName  : "+actAccountName +" Value Exp : "+expAccountName);

		String actAmount=select1stRow_3rdColumn.getText();
		String expAmount=excelReader.getCellData(xlSheetName, 83, 6);
		excelReader.setCellData(xlfile, xlSheetName, 83, 7, actAmount);
		
		System.err.println("************After Amount  : "+actAmount +" Value Exp : "+expAmount);

		if (actAccountName.equalsIgnoreCase(expAccountName) && actAmount.equalsIgnoreCase(expAmount))
		{
			excelReader.setCellData(xlfile, xlSheetName, 82, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 82, 8, resFail);
			return false;
		}
	}
	
	
	


	public boolean checkClickOnHiddingOptionInRulesAndClickOnSave() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		NavigateToRulesScreen();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleNameTxt));
		ruleNameTxt.click();

		rulesTabRuleNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);

		rulesTabRuleNameTxt.sendKeys(Keys.BACK_SPACE);		

		Thread.sleep(1000);
		
		rulesTabRuleNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 84, 5));

		Thread.sleep(2999);
		
		rulesTabRuleNameTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(1000);
		
		getAction().moveToElement(rulesTabGrid1st_11thCol).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabGrid1st_11thCol));
		rulesTabGrid1st_11thCol.click();

		Thread.sleep(1999);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleIFHidedenChkBox));
		if (ruleIFHidedenChkBox.isSelected()==false)
		{
			ruleIFHidedenChkBox.click();
			ruleIFHidedenChkBox.sendKeys(Keys.TAB);
		}

		Thread.sleep(1999);

		String actHideStatus=rulesTabGrid1st_11thCol.getText();
		String expHideStatus=excelReader.getCellData(xlSheetName, 84, 6);
		excelReader.setCellData(xlfile, xlSheetName, 84, 7, actHideStatus);
		System.out.println("Hide Status : "+actHideStatus +" Value Exp : "+expHideStatus);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveRuleBtn));
		saveRuleBtn.click();

		String expMessage=excelReader.getCellData(xlSheetName, 85, 6);
		String actMesage=checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 85, 7, actMesage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		String expMessage1=excelReader.getCellData(xlSheetName, 86, 6);
		String actMesage1=checkValidationMessage(expMessage1);
		excelReader.setCellData(xlfile, xlSheetName, 86, 7, actMesage1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();

		if (actMesage.equalsIgnoreCase(expMessage) && actMesage1.equals(expMessage1) &&
				actHideStatus.equalsIgnoreCase(expHideStatus))
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
	
	
	
	


	public boolean checkAfterEnableHiddenOptionAtVoucherLevel() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(recepitsVATMenu));
		recepitsVATMenu.click();

		Thread.sleep(2000);    

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		Thread.sleep(1999);
		
		String actBeforeLoadAccount=Boolean.toString(recepitsVATChequeNo.isDisplayed());
		String expBeforeLoadAccount=excelReader.getCellData(xlSheetName, 87, 6);
		excelReader.setCellData(xlfile, xlSheetName, 87, 7, actBeforeLoadAccount);

		System.out.println(" BeforeLoadAccount : "+actBeforeLoadAccount +" Value Exp : "+expBeforeLoadAccount);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCashBankAccountTxt));
		newCashBankAccountTxt.click();
		newCashBankAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 87, 5));

		Thread.sleep(1999);

		newCashBankAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		newCashBankAccountTxt.click();

		Thread.sleep(1999);
		
		String actAfterLoadAccount=Boolean.toString(recepitsVATChequeNo.isDisplayed());
		String expAfterLoadAccount=excelReader.getCellData(xlSheetName, 88, 6);
		excelReader.setCellData(xlfile, xlSheetName, 88, 7, actAfterLoadAccount);

		System.out.println(" AfterLoadAccount : "+actAfterLoadAccount +" Value Exp : "+expAfterLoadAccount);

		if (actAfterLoadAccount.equalsIgnoreCase(expAfterLoadAccount) &&actBeforeLoadAccount.equalsIgnoreCase(expBeforeLoadAccount)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 87, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 87, 8, resPass);
			return false;
		}
	}
	
	
	
	
	
	

	public boolean chekSavingRuleWithBeforeDelete() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		Thread.sleep(2000);

	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionBtn));
		configureTransactionBtn.click();*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTxt));
		searchMenuTxt.click();
		searchMenuTxt.sendKeys("Configure Transactions");
		searchMenuTxt.sendKeys(Keys.ENTER);
		
	
		Thread.sleep(2000);

	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));
		documentCustomization45Btn.click();
*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(paymentsVATVoucher));
		paymentsVATVoucher.click();

		Thread.sleep(1999);
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentsTab));
		documentsTab.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toogleExpandBtn));
		toogleExpandBtn.click();
*/
		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesBtn));
		rulesBtn.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickhereToAddRuleBtn));
		clickhereToAddRuleBtn.click();


		Thread.sleep(1999);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleNameTxt));
		ruleNameTxt.click();

		rulesTabRuleNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);

		rulesTabRuleNameTxt.sendKeys(Keys.BACK_SPACE);		

		Thread.sleep(1000);
		
		rulesTabRuleNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 89, 5));
		
		Thread.sleep(1999);

		ruleNameTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesNewRecordChekbox));
		rulesNewRecordChekbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesLoadChekbox));
		rulesLoadChekbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesOnLeaveChekbox));
		rulesOnLeaveChekbox.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesBeforeDeleteChekbox));
		rulesBeforeDeleteChekbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesActiveChekbox));
		rulesActiveChekbox.click();


		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesNoConditionChekbox));
		rulesNoConditionChekbox.click();


		Thread.sleep(2000);


		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessageTab));
		rulesIFMessageTab.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessagesGenralMessageTxt));
		rulesIFMessagesGenralMessageTxt.click();
		rulesIFMessagesGenralMessageTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		rulesIFMessagesGenralMessageTxt.sendKeys(excelReader.getCellData(xlSheetName, 90, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessagesWarnAndStopRadio));
		if (rulesIFMessagesWarnAndStopRadio.isSelected()==false)
		{
			rulesIFMessagesWarnAndStopRadio.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveRuleBtn));
		saveRuleBtn.click();

		String ExpMessage=excelReader.getCellData(xlSheetName, 89, 6);
		String actMessage=checkValidationMessage(ExpMessage);
		excelReader.setCellData(xlfile, xlSheetName, 89, 7, actMessage);

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		String expMessage1=excelReader.getCellData(xlSheetName, 90, 6);
		String actMessage1=checkValidationMessage(expMessage1);
		excelReader.setCellData(xlfile, xlSheetName, 90, 7, actMessage1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();

		if(actMessage1.equalsIgnoreCase(expMessage1) && actMessage.equalsIgnoreCase(ExpMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 89, 8, resPass);
			return true;

		} else {

			excelReader.setCellData(xlfile, xlSheetName, 89, 8, resFail);
			return false;
		}

	}

	@FindBy(xpath="//li[@id='id_transactionentry_delete']")
	private static WebElement  new_DeleteBtn;

	@FindBy(xpath="//span[contains(text(),'Payments VAT')]")
	private static WebElement  paymentsVATMenu;


	@FindBy(xpath="//li[@id='id_transactionentry_save']")
	private static WebElement  MRsaveBtn;

	@FindBy(xpath="//input[@id='id_header_10']")
	private static WebElement  voucherHeaderCurrency;

	@FindBy(xpath="//tbody[@id='id_header_10_table_body']/tr/td[2]")
	private static List<WebElement> currencyListCount;	

	public boolean checkValidationRuleWithOnDeleteAtVoucherLevel() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(paymentsVATMenu));
		paymentsVATMenu.click();

		Thread.sleep(2000);    
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		Thread.sleep(2000);  
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCashBankAccountTxt));
		newCashBankAccountTxt.click();
		newCashBankAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 91, 5));
		Thread.sleep(2000);
		newCashBankAccountTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.click();
		voucherHeaderCurrency.sendKeys(Keys.SHIFT,Keys.HOME);

		voucherHeaderCurrency.sendKeys(Keys.SPACE);

		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 92, 5));
		
		Thread.sleep(2000);
		
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SHIFT,Keys.HOME,Keys.BACK_SPACE);
		departmentTxt.sendKeys(Keys.SPACE);

		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 93, 5));

		Thread.sleep(2000);

		departmentTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 94, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 95, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
		enter_AccountTxt.click();
		enter_AccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 96, 5));
		Thread.sleep(2000);
		enter_AccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enterpayVATTaxCode));	
		enterpayVATTaxCode.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Amount));
		enter_Amount.sendKeys(excelReader.getCellData(xlSheetName, 97, 5));
		enter_Amount.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();


		String docno=documentNumberTxt.getAttribute("value");
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRsaveBtn));
		MRsaveBtn.click();

		Thread.sleep(2000);

		/*Thread.sleep(1000);
		HashSet<String> actMesage = new HashSet<String>();

		for (int i = 0; i < 3; i++) 
		{
			String data = checkValidationMessage("");
			actMesage.add(data);
		}


		HashSet<String> expMesage = new HashSet<String>();

		expMesage.add(excelReader.getCellData(xlSheetName, 91, 6));
		Thread.sleep(1000);
		expMesage.add(excelReader.getCellData(xlSheetName, 92, 6)+docno);
		Thread.sleep(1000);
		expMesage.add(excelReader.getCellData(xlSheetName, 93, 6));
		
		String actMessage = actMesage.toString();
		String expMessage = expMesage.toString();

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 91, 7, actMessage);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();
		Thread.sleep(2000);

		checkValidationMessage("Voucher loaded");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_DeleteBtn));
		new_DeleteBtn.click();
		Thread.sleep(2000);*/

		String expAlert=excelReader.getCellData(xlSheetName, 94, 6);
		String actAlert=checkValidationMessage(expAlert);
		excelReader.setCellData(xlfile, xlSheetName, 94, 7, actAlert);

		if (actAlert.equalsIgnoreCase(expAlert) /*&&  
				actMessage.equalsIgnoreCase(expMessage)*/)
		{
			excelReader.setCellData(xlfile, xlSheetName, 91, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 91, 8, resFail);
			return false;

		}
	}

	
	
	
	
	
	public boolean chekSavingRuleWithOnEdit() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionBtn));
		configureTransactionBtn.click();
*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTxt));
		searchMenuTxt.click();
		searchMenuTxt.sendKeys("Configure Transactions");
		searchMenuTxt.sendKeys(Keys.ENTER);
		
	
		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));
		documentCustomization45Btn.click();

*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(paymentsVATVoucher));
		paymentsVATVoucher.click();

		Thread.sleep(1999);
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentsTab));
		documentsTab.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toogleExpandBtn));
		toogleExpandBtn.click();
*/
		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesBtn));
		rulesBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleNameTxt));
		ruleNameTxt.click();
		ruleNameTxt.clear();
		Thread.sleep(2000);

		rulesTabRuleNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);

		rulesTabRuleNameTxt.sendKeys(Keys.BACK_SPACE);		

		Thread.sleep(1000);
		rulesTabRuleNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 98, 5));
		Thread.sleep(1999);

		ruleNameTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesLoadChekbox));
		rulesLoadChekbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesEditChekbox));
		rulesEditChekbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesActiveChekbox));
		rulesActiveChekbox.click();


		Thread.sleep(1000);

		rulesTabIFSelectTabTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectTabDrpdwn));
		Select account=new Select(rulesTabIFSelectTabDrpdwn);
		account.selectByVisibleText(excelReader.getCellData(xlSheetName, 99, 5));
		rulesTabIFSelectTabDrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		rulesTabExtraFieldDrpdwn.click();

		rulesTabExtraFieldDrpdwn.sendKeys(Keys.ARROW_DOWN);
		rulesTabExtraFieldDrpdwn.sendKeys(Keys.TAB);

		rulesTabOperatorsDrpdwn.click();

		rulesTabOperatorsDrpdwn.sendKeys(Keys.ARROW_DOWN);

		rulesTabOperatorsDrpdwn.sendKeys(Keys.TAB);

		rulesTabValueTxt.click();

		Thread.sleep(1000);
		Select Value=new Select(rulesTabCompareWithDrpdwn);
		Value.selectByValue(excelReader.getCellData(xlSheetName, 100, 5));
		rulesTabCompareWithDrpdwn.sendKeys(Keys.TAB);
		rulesTabCompareWithDrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		rulesTabConditionTxt.sendKeys(Keys.SPACE);
		rulesTabConditionTxt.sendKeys(Keys.BACK_SPACE);
		Thread.sleep(2000);

		rulesTabConditionTxt.sendKeys(excelReader.getCellData(xlSheetName, 101, 5));

		Thread.sleep(4000);

		rulesTabConditionTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);


		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessageTab));
		rulesIFMessageTab.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessagesGenralMessageTxt));
		rulesIFMessagesGenralMessageTxt.click();
		rulesIFMessagesGenralMessageTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		rulesIFMessagesGenralMessageTxt.sendKeys(excelReader.getCellData(xlSheetName, 102, 5));

		if (rulesIFMessagesWarnAndProceddRadio.isSelected()==false)
		{
			rulesIFMessagesWarnAndProceddRadio.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveRuleBtn));
		saveRuleBtn.click();

		String ExpMessage=excelReader.getCellData(xlSheetName, 98, 6);
		String actMessage=checkValidationMessage(ExpMessage);
		excelReader.setCellData(xlfile, xlSheetName, 98, 7, actMessage);

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();

		if (actMessage.equalsIgnoreCase(ExpMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 98, 8, resPass);
			return true;

		} else {

			excelReader.setCellData(xlfile, xlSheetName, 98, 8, resFail);
			return false;
		}
	}


	@FindBy(xpath="//*[@id='2002']/span")
	private static WebElement recepitsVoucher;

	@FindBy(xpath="//*[@id='navigation_menu']/li[2]")
	private static WebElement  finacinalsMenu;

	@FindBy(xpath="//a[@id='2001']//span[contains(text(),'Cash and Bank')]")
	private static WebElement  cashAndBankMenu; 

	@FindBy(xpath="//*[@id='61']/span")
	private static WebElement  financialsTransactionMenu; 

	@FindBy(xpath="(//*[@id='navbarSupportedContent2']/ul/li[1])[1]")
	private static WebElement  newBtn;

	@FindBy(xpath="//div[@id='id_transaction_homescreen_edit']//span[@class='icon-edit icon-font6 ImagesinArabic']")
	private static WebElement  editBtn;

	@FindBy(xpath="//input[@id='id_header_268435459']")
	private static WebElement  departmentTxt;

	@FindBy(xpath="//tbody[@id='id_header_268435459_table_body']/tr/td[2]")
	private static List<WebElement> departmentListCount;

	@FindBy(xpath="//input[@id='id_header_268435470']")
	private static WebElement  placeOFSupplyTxt;


	public boolean checkValidationOFEditRuleWithWarnAndProcedueInRecepitsVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		checkValidationRuleWithOnDeleteAtVoucherLevel();

		Thread.sleep(1999);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SHIFT,Keys.HOME,Keys.BACK_SPACE);
		departmentTxt.sendKeys(Keys.SPACE);

		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 103, 5));

		Thread.sleep(2000);
		
		departmentTxt.sendKeys(Keys.TAB);


		Thread.sleep(1999);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();

		enter_Amount.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();


		String docno=documentNumberTxt.getAttribute("value");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRsaveBtn));
		MRsaveBtn.click();

		Thread.sleep(2000);

		Thread.sleep(1000);
		/*HashSet<String> actMesage = new HashSet<String>();

		for (int i = 0; i < 3; i++) 
		{
			String data = checkValidationMessage("");
			actMesage.add(data);
		}


		HashSet<String> expMesage = new HashSet<String>();

		expMesage.add(excelReader.getCellData(xlSheetName, 103, 6));
		expMesage.add(excelReader.getCellData(xlSheetName, 104, 5)+docno);
		expMesage.add(excelReader.getCellData(xlSheetName, 105, 5));

		System.out.println("Actual Message    : "+actMesage);
		System.out.println("Expected Message  : "+expMesage);

		String actMessage = actMesage.toString();
		String expMessage = expMesage.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 103, 7, actMessage);*/
		
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 103, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 104, 6);
		String expMessage3 = excelReader.getCellData(xlSheetName, 105, 6);
		String actMessage1 = checkValidationMessage(expMessage2);
		excelReader.setCellData(xlfile, xlSheetName, 103, 7, actMessage);
		excelReader.setCellData(xlfile, xlSheetName, 104, 7, actMessage1);
		
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.equalsIgnoreCase(expMessage1)&& actMessage1.startsWith(expMessage2)&&actMessage1.endsWith(expMessage3))
			
		//if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 103, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 103, 8, resFail);
			return false;
		}
	}



	@FindBy(xpath="(//a[@id='3333'])[2]")
	private static WebElement salesINvoiceVATVoucher;



	@FindBy(xpath="//input[@id='id_header_4']")
	private static WebElement newCashBankAccountTxt;



	@FindBy(xpath="//select[@id='editScreen_FieldsCustomization_Position']")
	private static WebElement extraFiledPOsitionDrpdwn;

	public boolean checkSavingExtraFiledInSalesINvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionBtn));
		configureTransactionBtn.click();
		Thread.sleep(2000);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTxt));
		searchMenuTxt.click();
		searchMenuTxt.sendKeys("Configure Transactions");
		searchMenuTxt.sendKeys(Keys.ENTER);
		
	
		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(preferencesBtn));
		preferencesBtn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));
*/
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));
		documentCustomization45Btn.click();

*/
		Thread.sleep(1999);
		
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", salesINvoiceVATVoucher);
		Thread.sleep(2000);
		getAction().moveToElement(salesINvoiceVATVoucher).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesINvoiceVATVoucher));
		salesINvoiceVATVoucher.click();

		Thread.sleep(1999);
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentsTab));
		documentsTab.click();

*/		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
		editScreenTab.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenaddBtn));
		editScreenaddBtn.click();

		Thread.sleep(1999);
		getAlert().accept();


		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenCaptionTxt));
		editScreenCaptionTxt.click();
		editScreenCaptionTxt.sendKeys(excelReader.getCellData(xlSheetName, 106, 5));

		Thread.sleep(1000);

		editScreenNoOfDecimalsTxt.click();
		editScreenNoOfDecimalsTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		editScreenNoOfDecimalsTxt.sendKeys(excelReader.getCellData(xlSheetName, 107, 5));

		editScreenNoOfDecimalsTxt.sendKeys(Keys.TAB);

		Thread.sleep(1000);
		Select s1= new Select(extraFiledPOsitionDrpdwn);
		s1.selectByValue(excelReader.getCellData(xlSheetName, 108, 5));
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
		editScreenApplyBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 106, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 106, 7, actMessage);

		if (actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 106, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 106, 8, resFail);
			return false;
		}



	}

	@FindBy(xpath="//*[@value='Apply']")
	private static WebElement editScreenApplyBtn;  

	@FindBy(xpath="//span[@id='editScreen_FieldsCustomization_Close']")
	private static WebElement editScreencloseBtn;  

	@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_Caption']")
	private static WebElement  editScreenCaptionTxt;

	@FindBy(xpath="//select[@id='editScreen_FieldsCustomization_DataType']")
	private static WebElement  editScreenCaptionDataTypeDropdown;

	@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_DefaultValue']")
	private static WebElement editScreenDefaultValueTxt;

	@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_noofDecimals']")
	private static WebElement editScreenNoOfDecimalsTxt;

	@FindBy(xpath="//*[@id='navigationtab13']")
	private static WebElement  editScreenTab; 

	@FindBy(xpath="//i[@class='icon-new hiconright2']")
	private static WebElement editScreenaddBtn;



	public boolean checkSavingRuleWithFormulaUnderSalesINvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(1999);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toogleExpandBtn));
		toogleExpandBtn.click();
*/
		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesBtn));
		rulesBtn.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickhereToAddRuleBtn));
		clickhereToAddRuleBtn.click();


		Thread.sleep(1999);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleNameTxt));
		ruleNameTxt.click();
		

		rulesTabRuleNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);

		rulesTabRuleNameTxt.sendKeys(Keys.BACK_SPACE);		

		Thread.sleep(1000);
		rulesTabRuleNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 109, 5));
		Thread.sleep(1999);

		ruleNameTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesNewRecordChekbox));
		rulesNewRecordChekbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesEditChekbox));
		rulesEditChekbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesOnLeaveChekbox));
		rulesOnLeaveChekbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesActiveChekbox));
		rulesActiveChekbox.click();


		Thread.sleep(1999);
		rulesTabIFSelectTabTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectTabDrpdwn));
		Select account=new Select(rulesTabIFSelectTabDrpdwn);
		account.selectByVisibleText(excelReader.getCellData(xlSheetName, 110, 5));
		rulesTabIFSelectTabDrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		Select s1= new Select(rulesTabOperatorsDrpdwn);
		s1.selectByValue(excelReader.getCellData(xlSheetName, 111, 5));

		Thread.sleep(1999);

		rulesTabOperatorsDrpdwn.sendKeys(Keys.TAB);

		rulesTabValueTxt.click();

		Thread.sleep(1000);
		Select Value=new Select(rulesTabCompareWithDrpdwn);
		Value.selectByValue(excelReader.getCellData(xlSheetName, 112, 5));
		rulesTabCompareWithDrpdwn.sendKeys(Keys.TAB);
		rulesTabCompareWithDrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		rulesIf_ValueTxt.sendKeys(Keys.SPACE);
		rulesIf_ValueTxt.sendKeys(Keys.BACK_SPACE);
		Thread.sleep(2000);

		rulesIf_ValueTxt.sendKeys(excelReader.getCellData(xlSheetName, 113, 5));

		Thread.sleep(4000);

		rulesIf_ValueTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid1stRow_1stcol));
		rulesGrid1stRow_1stcol.click();

		rulesIFEnterFiled.sendKeys(excelReader.getCellData(xlSheetName, 114, 5));
		rulesIFEnterFiled.sendKeys(Keys.TAB);
		rulesIFEnterChangeValueDrpdwn.sendKeys(excelReader.getCellData(xlSheetName, 115, 5));

		rulesIFEnterChangeValueDrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(1999);
		rulesIFEnterValue.sendKeys(excelReader.getCellData(xlSheetName, 116, 5));

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabGrid1st_10thCol));
		rulesTabGrid1st_10thCol.click();

		rulesTabGridMandatoryChkbox.click();


		Thread.sleep(2000);


		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessageTab));
		rulesIFMessageTab.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessagesGenralMessageTxt));
		rulesIFMessagesGenralMessageTxt.click();
		rulesIFMessagesGenralMessageTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		rulesIFMessagesGenralMessageTxt.sendKeys(excelReader.getCellData(xlSheetName, 117, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFMessagesInformationRadio));
		if (rulesIFMessagesInformationRadio.isSelected()==false)
		{
			rulesIFMessagesInformationRadio.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveRuleBtn));
		saveRuleBtn.click();

		String ExpMessage=excelReader.getCellData(xlSheetName, 109, 6);
		String actMessage=checkValidationMessage(ExpMessage);
		excelReader.setCellData(xlfile, xlSheetName, 109, 7, actMessage);

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		String expMessage1=excelReader.getCellData(xlSheetName, 110, 6);
		String actMessage1=checkValidationMessage(expMessage1);
		excelReader.setCellData(xlfile, xlSheetName, 110, 7, actMessage1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();

		if(actMessage1.equalsIgnoreCase(expMessage1) && actMessage.equalsIgnoreCase(ExpMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 109, 8, resPass);
			return true;

		} else {

			excelReader.setCellData(xlfile, xlSheetName, 109, 8, resFail);
			return false;
		}
	}

	@FindBy(xpath="//*[@id='60']")
	private static WebElement  financialsMenu; 

	@FindBy(xpath="//a[@id='2014']//span[contains(text(),'Sales')]")
	private static WebElement  financialTransactionSalesMenu;

	@FindBy(xpath="//span[contains(text(),'Sales invoice VAT')]")
	private static WebElement  salesInvoiceVATVoucher;

	@FindBy(xpath="//input[@id='id_body_536870916']")
	private static WebElement  pvWareHouseTxt;


	@FindBy(xpath="//tbody[@id='id_body_536870916_table_body']/tr/td[2]")
	private static List<WebElement> pvwareHouseListCount;

	@FindBy(xpath="//input[@id='id_body_33554476']")
	private static WebElement enter_AQTxt;

	@FindBy(xpath="//input[@id='id_body_33554477']")
	private static WebElement enter_FQTxt;

	@FindBy(xpath="//input[@id='id_body_23']")
	private static WebElement  enter_ItemTxt;

	@FindBy(xpath="//input[@id='id_body_12']")
	private static WebElement  enter_SalesAccountTxt;

	@FindBy(xpath="//input[@id='id_body_24']")
	private static WebElement  enter_UnitTxt;

	@FindBy(xpath="//input[@id='id_body_26']")
	private static WebElement  enter_Quantity;

	@FindBy(xpath="//input[@id='id_body_27']")
	private static WebElement  enter_Rate;

	@FindBy(xpath="//input[@id='id_body_28']")
	private static WebElement  enter_Gross;

	@FindBy(xpath="//tbody[@id='id_body_23_table_body']/tr/td[2]")
	private static List<WebElement> itemListCount;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[8]")
	private static WebElement  select1stRow_7thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[9]")
	private static WebElement  select1stRow_8thColumn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[19]")
	private static WebElement  select1stRow_18thColumn;

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

	@FindBy(xpath="//*[@id='2028']/span")
	private static WebElement  openingBalancesVoucher;

	@FindBy(xpath="//*[@id='135']")
	private static WebElement  inventoryMenu; 

	@FindBy(xpath="//*[@id='137']")
	private static WebElement  inventoryTransactionsMenu; 

	@FindBy(xpath="//*[@id='2033']")
	private static WebElement  inventoryTransactionsStocksMenu; 

	@FindBy(xpath="//a[@id='2037']//span[contains(text(),'Opening Stocks')]")
	private static WebElement  openingStocksVoucher;

	@FindBy(xpath="//input[@id='id_header_268435460']")
	private static WebElement  warehouseTxt;


	@FindBy(xpath="//tbody[@id='id_header_268435460_table_body']/tr/td[2]")
	private static List<WebElement> warehouseHeaderComboList;

	@FindBy(xpath="//tbody[@id='id_body_23_table_body']/tr/td[2]")
	private static List<WebElement> itemComboList;

	public boolean checkSavingOpeningStockVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksVoucher));
		openingStocksVoucher.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.click();
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 118, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 119, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 120, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 121, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);


		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*HashSet<String> actMesage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMesage.add(data);
		}

		HashSet<String> expMesage = new HashSet();

		expMesage.add(excelReader.getCellData(xlSheetName, 118, 6));
		expMesage.add(excelReader.getCellData(xlSheetName, 119, 6));
		
		String actMessage = actMesage.toString();
		String expMessage = expMesage.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 118, 7, actMessage);

		System.out.println("Actual Message    : "+actMesage);
		System.out.println("Expected Message  : "+expMesage);

		if(actMessage.equalsIgnoreCase(expMessage))*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 118, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 119, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 118, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		
		
		{
			excelReader.setCellData(xlfile, xlSheetName, 118, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 118, 8, resFail);
			return false;
		}
	}



	@FindBy(xpath="//input[@id='id_body_33554474']")
	private static WebElement enterAvgRate;


	public boolean checkValidationOFFOrmulaControlAtVoucherLevel() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
		financialTransactionSalesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
		salesInvoiceVATVoucher.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("");


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.click();
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 122, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(Keys.SPACE);

		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 123, 5));

		Thread.sleep(2000);

		enter_ItemTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
		select1stRow_6thColumn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
		select1stRow_8thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
		enter_AQTxt.sendKeys(excelReader.getCellData(xlSheetName, 124, 5));
		enter_AQTxt.sendKeys(Keys.TAB);

		enter_FQTxt.sendKeys(excelReader.getCellData(xlSheetName, 125, 5));

		enter_FQTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
		select1stRow_6thColumn.click();

		enterAvgRate.sendKeys(Keys.TAB);

		String ExpMessage=excelReader.getCellData(xlSheetName, 122, 6);
		String actMessage=checkValidationMessage(ExpMessage);
		excelReader.setCellData(xlfile, xlSheetName, 122, 7, actMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
		String actData=select1stRow_11thColumn.getText();
		String expDate=excelReader.getCellData(xlSheetName, 123, 6);
		excelReader.setCellData(xlfile, xlSheetName, 123, 7, actData);
		
		System.err.println(" Coloumn Updated Value  : "+actData +" VALUE EXP "+expDate);

		if (actData.equalsIgnoreCase(expDate) && actMessage.equalsIgnoreCase(ExpMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 122, 8, resPass);
			return true;
		} 
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 122, 8, resFail);
			return false;
		}
	}


	@FindBy (xpath="//*[@id='1']")
	public static WebElement homeMenu;

	//Masters Menu	
	@FindBy (xpath="//a[@id='1000']//span[contains(text(),'Masters')]")
	public static WebElement mastersMenu;

	@FindBy (xpath="//*[@id='25']")
	private static WebElement dataMangementMenu;

	@FindBy (xpath="//*[@id='26']")
	private static WebElement backup;

	@FindBy (xpath="//*[@id='27']")
	private static WebElement restore;

	@FindBy (xpath="//*[@id='111']/span")
	private static WebElement reindex;

	@FindBy (xpath="//*[@id='115']")
	private static WebElement eraseAll;

	@FindBy (xpath="//button[@class='btn Fbutton']")
	private static WebElement folderpathExpandBtn;

	@FindBy (xpath="//*[@class='icon-restore hiconright2']")
	private static WebElement restoreCompanyBtn;

	@FindBy (xpath="//*[contains(text(),'Yes')]")
	private static WebElement overRideYesBtn;

	@FindBy (xpath="//button[@class='Fbutton'][contains(text(),'No')]")
	private static WebElement overRideNoBtn;

	@FindBy (xpath="//div[@id='ReIndexingError_Modal']/div/div")
	private static WebElement reindexingPopup;

	@FindBy (xpath="(//div[@id='ReIndexingError_Modal']//div/button)[1]")
	private static WebElement reindexingPopupCloseBtn;

	@FindBy (xpath="//*[@id='ReIndexingError_Modal']//div[3]/button")
	private static WebElement reindexingPopupCancelBtn;

	public boolean chckRestoreBackUpForSQLFunction() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getDriver().navigate().refresh();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataMangementMenu));
		dataMangementMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restore));
		restore.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(folderpathExpandBtn));
		folderpathExpandBtn.click();

		Thread.sleep(3000);
		//FOR SQL FUNTION IN SANITY BAck Up Restore 
		Runtime.getRuntime().exec(getBaseDir()+excelReader.getCellData(xlSheetName, 126, 5));

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

		LoginPage lp= new LoginPage(getDriver());

		String unamelt=excelReader.getCellData(xlSheetName, 127, 5);

		String pawslt=excelReader.getCellData(xlSheetName, 128, 5);

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);

		Thread.sleep(2000);

		/*Select s = new Select(companyDropDownList);
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
			}*/

		lp.clickOnSignInBtn();

		Thread.sleep(40000);

		if (reindexingPopup.isDisplayed()==false) 
		{
			System.out.println("Reindexing Log is Not Displaying");
		} 
		else 
		{
			System.out.println("Reindexing Log is Displaying");
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reindexingPopupCancelBtn));
			reindexingPopupCancelBtn.click();
			Thread.sleep(6000);
			lp.clickOnSignInBtn();
		}

		String actUserInfo1=userNameDisplay.getText();

		System.out.println("User Info  : "+actUserInfo1);

		System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt1=companyName.getText();
		String getLoginCompanyName1=getCompanyTxt1.substring(0, 19);
		System.out.println("company name  :  "+ getLoginCompanyName1);
		companyLogo.click();

		String expUserInfo1           =excelReader.getCellData(xlSheetName, 126, 6);
		String expLoginCompanyName1   =excelReader.getCellData(xlSheetName, 127, 6);
		excelReader.setCellData(xlfile, xlSheetName, 126, 7, actUserInfo1);
		excelReader.setCellData(xlfile, xlSheetName, 127, 7, getLoginCompanyName1);

		System.out.println("UserInfo1             : "+actUserInfo1            +" Value Expected : "+expUserInfo1);
		System.out.println("LoginCompanyName1     : "+getLoginCompanyName1    +" Value Expected : "+expLoginCompanyName1);

		if(actUserInfo1.equalsIgnoreCase(expUserInfo1) && getLoginCompanyName1.contains(expLoginCompanyName1))
		{
			excelReader.setCellData(xlfile, xlSheetName, 126, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 126, 8, resFail);
			return false;
		}
	}

	@FindBy(xpath="//span[contains(text(),'Excesses in Stocks')]")
	private static WebElement  excessesInStocksVoucher;

	@FindBy(xpath="//input[@id='id_body_13']")
	private static WebElement  enter_Batch;


	@FindBy(xpath="//input[@id='id_body_37']")
	private static WebElement enter_Expirydate;


	public boolean checkSQLFuntionATExcessStockVoucherLevel() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(excessesInStocksVoucher));
		excessesInStocksVoucher.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();


		checkValidationMessage("Screen Opened");
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
		pvWareHouseTxt.click();
		pvWareHouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 129, 5));
		Thread.sleep(2000);
		pvWareHouseTxt.sendKeys(Keys.TAB);

		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 130, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 131, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 132, 5));
		enter_Rate.sendKeys(Keys.TAB);

		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.sendKeys(excelReader.getCellData(xlSheetName, 133, 5));
		enter_Batch.sendKeys(Keys.TAB);

		Thread.sleep(1000);
		
		enter_Expirydate.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();

		Thread.sleep(1000);
		getAction().moveToElement(select1stRow_12thColumn).build().perform();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));

		String actBarcode=select1stRow_12thColumn.getText();
		String expBarcode=excelReader.getCellData(xlSheetName, 129, 6);
		excelReader.setCellData(xlfile, xlSheetName, 129, 7, actBarcode);

		System.err.println(" Barcode DIsplayed : "+actBarcode +" Value Exp : "+expBarcode);

		if (actBarcode.equalsIgnoreCase(expBarcode))
		{
			excelReader.setCellData(xlfile, xlSheetName, 129, 8, resPass);
			System.out.println(" Test Pass: VaLUES DISPLAYED from SQl ");
			return true;
		} 
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 129, 8, resFail);
			System.out.println(" Test FAIl: VaLUES DISPLAYED from SQl ");
			return false;
		}

	}

	@FindBy(xpath="//input[@id='views_existingViews']")
	private static WebElement  ViewExistingViewTxt;


	@FindBy(xpath="//*[@id='views_FieldsTable']/tbody/tr/td[3]")
	private static List<WebElement>  viewsGridFiledsList;

	@FindBy(xpath="//*[@id='views_FieldsTable']/tbody/tr/td[1]")
	private static List<WebElement>  viewsGridFiledsEditList;


	@FindBy(xpath = "//*[@id='EditLayout_FieldsCustomization_FieldDetails_DataType']")
	private static WebElement fieldDetails_DataTypeDrpdwn;

	@FindBy(xpath = "//*[@id='editScreen_FieldsCustomization_DefaultValue']")
	private static WebElement fieldDetails_DefaultValueTxt;

	@FindBy(xpath = "//*[@value='Apply']")
	private static WebElement fieldDetails_ApplyBtn;
	
	@FindBy(xpath = "(//*[@id='dvCustomizeEditLayoutField']//div[3]/input[2])[2]")
	private static WebElement fieldDetails_CloseBtn;
	
	

	@FindBy(xpath = "//div[@id='views_tabContent']//tr[2]//td[1]//span[1]")
	private static WebElement viewGridRow2EditBtn;

	@FindBy(xpath="//input[@id='views_User_3']/following-sibling::span")
	private static WebElement  viewsUserAllOptionsSTChkbox;

	@FindBy(xpath="//*[@id='spnSaveView']")
	private static WebElement  viewSaveView;


	public boolean checkCreatingViewOptionInPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getDriver().navigate().refresh();

		Thread.sleep(2000);
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionBtn));
		configureTransactionBtn.click();*/
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTxt));
		searchMenuTxt.click();
		searchMenuTxt.sendKeys("Configure Transactions");
		searchMenuTxt.sendKeys(Keys.ENTER);
		
	
		Thread.sleep(2000);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));
		documentCustomization45Btn.click();
*/

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucher));
		purchaseVoucher.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentsTab));
		documentsTab.click();

	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toogleExpandBtn));
		toogleExpandBtn.click();
*/
		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsBtn));
		viewsBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewExistingViewTxt));
		ViewExistingViewTxt.click();

		ViewExistingViewTxt.sendKeys(Keys.SHIFT,Keys.HOME);

		ViewExistingViewTxt.sendKeys(Keys.SPACE);
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys(excelReader.getCellData(xlSheetName, 134, 5));
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		int Count = viewsGridFiledsList.size();


		System.out.println("Count    : "+Count);

		ArrayList<String> actviewsGridFiledsList  = new ArrayList<String>();

		for (int i = 0; i < Count; i++) 
		{
			String data		  = viewsGridFiledsList.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 135, 5)))
			{
				viewsGridFiledsEditList.get(i).click();

				break;
			}
		}



		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewGridRow2EditBtn));
		getAction().doubleClick(viewGridRow2EditBtn).build().perform();


		Thread.sleep(6000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldDetails_DataTypeDrpdwn));			
		fieldDetails_DataTypeDrpdwn.click();

		Select data=new Select(fieldDetails_DataTypeDrpdwn);
		data.selectByVisibleText(excelReader.getCellData(xlSheetName, 136, 5));

		String actDataType=data.getFirstSelectedOption().getText();
		String expDataType=excelReader.getCellData(xlSheetName, 134, 6);
		excelReader.setCellData(xlfile, xlSheetName, 134, 7, actDataType);


		System.err.println("FiledCaption For date In Create View  : "+actDataType+" Value Exp : "+expDataType);

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldDetails_ApplyBtn));
		fieldDetails_ApplyBtn.click();


		Thread.sleep(2000);
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldDetails_CloseBtn));
		fieldDetails_CloseBtn.click();
		Thread.sleep(2000);
		*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
		viewTabLoginBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserAllOptionsSTChkbox));
		viewsUserAllOptionsSTChkbox.click();

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSaveView));
		viewSaveView.click();

		String expMessage= excelReader.getCellData(xlSheetName, 135, 6);

		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 135, 7, actMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		checkValidationMessage("data saved succesfully");

		Thread.sleep(1999);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();

		if (actMessage.equalsIgnoreCase(expMessage) && actDataType.equalsIgnoreCase(expDataType)) 
		{ 
			excelReader.setCellData(xlfile, xlSheetName, 134, 8, resPass);
			System.out.println("Test Pass : VIew Created In Purchase Voucher");
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 134, 8, resFail);
			System.out.println("Test Fail : VIew Created In Purchase Voucher");
			return false;
		}
	}


	@FindBy(xpath="//a[@id='768']")
	private static WebElement  purchaseVoucher;



	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[2]")
	private static WebElement  select2ndRow_1stColumn;



	public boolean checkViewOptionCreatedInPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(1999);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();


		Thread.sleep(1999);

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt=excelReader.getCellData(xlSheetName, 137, 5);

		String pawslt=excelReader.getCellData(xlSheetName, 138, 5);

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);

		lp.clickOnSignInBtn();

		Thread.sleep(8000);


		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesMenu));
		purchasesMenu.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucherMenu));
		purchaseVoucherMenu.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("");


		String actDateTxt=dateTxt.getAttribute("disabled");
		String expDateTxt=excelReader.getCellData(xlSheetName, 137, 6);
		excelReader.setCellData(xlfile, xlSheetName, 137, 7, actDateTxt);

		System.out.println(" DATE TXT :"+actDateTxt +" Value Exp : "+expDateTxt);

		Thread.sleep(1999);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(1999);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();


		if (actDateTxt.equalsIgnoreCase(expDateTxt)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 137, 8, resPass);
			System.out.println(" Test Pass: Date Field is Read Only ");
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 137, 8, resFail);
			System.out.println(" Test Fail: Date Field is Read Only ");
			return false;
		}



	}
	@FindBy(xpath="//*[@id='2007']")
	private static WebElement purchasesMenu;

	@FindBy(xpath="//input[@id='id_header_2']")
	private static WebElement dateTxt;

	@FindBy(xpath="//*[@id='2008']/span[contains(text(),'Purchases Vouchers')]")
	private static WebElement purchaseVoucherMenu;

	@FindBy(xpath="//*[@id='dashName']")
	private static WebElement dashboard;


	public boolean checkLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt=excelReader.getCellData(xlSheetName, 139, 5);

		String pawslt=excelReader.getCellData(xlSheetName, 140, 5);

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);

		lp.clickOnSignInBtn();

		//checkRefershPopOnlogin();

		//checkPopUpWindow();

		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		String actUserInfo=userNameDisplay.getText();

		System.out.println("User Info  : "+actUserInfo);

		System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String getLoginCompanyName=getCompanyTxt.substring(0, 19);
		System.out.println("company name  :  "+ getLoginCompanyName);
		companyLogo.click();

		String actDashboard = Boolean.toString(dashboard.isDisplayed());

		String expUserInfo           =excelReader.getCellData(xlSheetName, 139, 6);
		excelReader.setCellData(xlfile, xlSheetName, 139, 7, actUserInfo);
		
		String expLoginCompanyName   =excelReader.getCellData(xlSheetName, 140, 6);
		excelReader.setCellData(xlfile, xlSheetName, 140, 7, getLoginCompanyName);
		
		String expDashboard         = excelReader.getCellData(xlSheetName, 141, 6);
		excelReader.setCellData(xlfile, xlSheetName, 141, 7, actDashboard);

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.equalsIgnoreCase(expLoginCompanyName) && actDashboard.equalsIgnoreCase(expDashboard)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 139, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 139, 8, resFail);
			return false;
		}
	}








	public boolean checkITEMWithF5Key() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesMenu));
		purchasesMenu.click();
		Thread.sleep(2000);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucherMenu));
		purchaseVoucherMenu.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("");


		departmentTxt.click();
		
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 142, 5));

		Thread.sleep(1999);
		
		departmentTxt.sendKeys(Keys.TAB);


		getAction().moveToElement(departmentTxt).doubleClick().build().perform();


		departmentTxt.sendKeys(Keys.chord(Keys.CONTROL,"c"));

		Thread.sleep(2000);
		
		departmentTxt.sendKeys(Keys.BACK_SPACE);// Here erasing dep and paste takes palce 
		
		departmentTxt.sendKeys(Keys.chord(Keys.CONTROL, "v"));
		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		enter_ItemTxt.click();

		enter_ItemTxt.sendKeys(Keys.F5);

		Thread.sleep(1999);
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchPopUpSearchTXT));
		searchPopUpSearchTXT.sendKeys(Keys.chord(Keys.CONTROL, "v"));
		Thread.sleep(1999);
		searchPopUpSearchTXT.sendKeys(Keys.TAB);

		String actsearchPopUpSearchTXT=searchPopUpSearchTXT.getAttribute("value");
		String expsearchPopUpSearchTXT=excelReader.getCellData(xlSheetName, 142, 6);
		excelReader.setCellData(xlfile, xlSheetName, 142, 7, actsearchPopUpSearchTXT);

		System.out.println("Copy and Paste value : "+actsearchPopUpSearchTXT +"value exp :"+expsearchPopUpSearchTXT);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchPopUpSearchTXT));
		searchPopUpSearchTXT.click();
		searchPopUpSearchTXT.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		searchPopUpSearchTXT.sendKeys(excelReader.getCellData(xlSheetName, 143, 5));

		Thread.sleep(2000);
		searchPopUpSearchTXT.sendKeys(Keys.TAB);

		String actSendKeys=searchPopUpSearchTXT.getAttribute("value");
		String expSendKeys=excelReader.getCellData(xlSheetName, 143, 6);
		excelReader.setCellData(xlfile, xlSheetName, 143, 7, actSendKeys);

		System.out.println("SendKeys :"+actSendKeys +" value exp :"+expSendKeys);


		if (actSendKeys.equalsIgnoreCase(expSendKeys) && actsearchPopUpSearchTXT.equalsIgnoreCase(expsearchPopUpSearchTXT))
		{
			excelReader.setCellData(xlfile, xlSheetName, 142, 8, resPass);
			return true;
		} 
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 142, 8, resFail);
			return false;
		}
	}


	public void enterName(String item) throws InterruptedException
	{
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchPopUpSearchTXT));
		searchPopUpSearchTXT.click();
		searchPopUpSearchTXT.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		searchPopUpSearchTXT.sendKeys(item);

		Thread.sleep(2000);
		searchPopUpSearchTXT.sendKeys(Keys.TAB);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchPopUpRefreshbtn));
		searchPopUpRefreshbtn.click();
	}



	public boolean checkLoadingFieldsIntoPopOnClickOnRefreshBtn() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{

		Thread.sleep(1999);
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchPopUpRefreshbtn));
		searchPopUpRefreshbtn.click();

		enterName(excelReader.getCellData(xlSheetName, 144, 5));

		try {

			if (getIsAlertPresent())
			{
				String alert=getAlert().getText();
				System.out.println(" ALERT Displayed  : "+alert);

				if (errorMessage.isDisplayed()==true)
				{
					String Message=errorMessage.getText();
					System.out.println(" Message Displayed  : "+Message);
					errorMessageCloseBtn.click();
				}
			}
		} catch (Exception e) 
		{
		}

		enterName(excelReader.getCellData(xlSheetName, 145, 5));

		Thread.sleep(1999);
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchPopUpSelectAllChkBox));
		searchPopUpSelectAllChkBox.click();

		Thread.sleep(1999);

		int count =searchPopUpNameList.size();
		ArrayList<String >array1= new ArrayList<>();
		for (int i = 0; i < count; i++) 
		{
			String data=searchPopUpNameList.get(i).getText();
			array1.add(data);
		}

		String actsearchPopUpNameList=array1.toString();
		String expsearchPopUpNameList=excelReader.getCellData(xlSheetName, 144, 6);
		System.out.println(" ACT searchPopUpNameList : "+actsearchPopUpNameList);
		System.out.println(" EXP searchPopUpNameList : "+expsearchPopUpNameList);
		excelReader.setCellData(xlfile, xlSheetName, 144, 7, actsearchPopUpNameList);


		int count2 =searchPopUpCodeList.size();
		ArrayList<String >array2= new ArrayList<>();
		for (int i = 0; i < count2; i++) 
		{
			String data=searchPopUpCodeList.get(i).getText();
			array2.add(data);
		}

		String actsearchPopUpCodeList=array2.toString();
		String expsearchPopUpCodeList=excelReader.getCellData(xlSheetName, 145, 6);
		System.out.println(" ACT searchPopUpCodeList : "+actsearchPopUpCodeList);
		System.out.println(" EXP searchPopUpCodeList : "+expsearchPopUpCodeList);
		excelReader.setCellData(xlfile, xlSheetName, 145, 7, actsearchPopUpCodeList);

		int count3 =searchPopUpAliasList.size();
		ArrayList<String >array3= new ArrayList<>();
		for (int i = 0; i < count3; i++) 
		{
			String data=searchPopUpAliasList.get(i).getText();
			array3.add(data);
		}

		String actsearchPopUpAliasList=array3.toString();
		String expsearchPopUpAliasList=excelReader.getCellData(xlSheetName, 146, 6);
		System.out.println(" ACT searchPopUpAliasList : "+actsearchPopUpAliasList);
		System.out.println(" EXp searchPopUpAliasList : "+expsearchPopUpAliasList);
		excelReader.setCellData(xlfile, xlSheetName, 146, 7, actsearchPopUpAliasList);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchPopUpOkBtn));
		searchPopUpOkBtn.click();

		if (actsearchPopUpCodeList.equalsIgnoreCase(expsearchPopUpCodeList) && 
				actsearchPopUpNameList.equalsIgnoreCase(expsearchPopUpNameList) && 
				actsearchPopUpAliasList.equalsIgnoreCase(expsearchPopUpAliasList))
		{
			excelReader.setCellData(xlfile, xlSheetName, 144, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 144, 8, resFail);
			return false;
		}
	}

	@FindBy(xpath="//*[@id='id_body_23_search_popup_grid_body']/tr/td[4]")
	private static List<WebElement> searchPopUpNameList;

	@FindBy(xpath="//*[@id='id_body_23_search_popup_grid_body']/tr/td[5]")
	private static List<WebElement> searchPopUpCodeList;

	@FindBy(xpath="//*[@id='id_body_23_search_popup_grid_body']/tr/td[6]")
	private static List<WebElement> searchPopUpAliasList;

	@FindBy(xpath="//*[@id='id_body_23_search_popup']/div/div/div[3]/div/div[2]/input[1]")
	private static WebElement searchPopUpOkBtn;

	@FindBy(xpath="//*[@id='id_header_3_search_popup_grid_body']/tr/td[2]/input")
	private static List<WebElement> searchAccPopUpRadioBtnList;

	@FindBy(xpath="//*[@id='id_header_3_search_popup_grid_body']/tr/td[4]")
	private static List<WebElement> searchAccPopUpNameList;

	@FindBy(xpath="//*[@id='id_header_3_search_popup_grid_body']/tr/td[5]")
	private static List<WebElement> searchAccPopUpCodeList;

	@FindBy(xpath="//*[@id='id_header_3_search_popup_grid_body']/tr/td[6]")
	private static List<WebElement> searchAccPopUpAliasList;

	@FindBy(xpath="//*[@id='id_header_3_search_popup']/div/div/div[3]/div/div[2]/input[1]")
	private static WebElement searchAccPopUpOkBtn;


	@FindBy(xpath="//input[@id='id_body_23_search_popup_grid_control_heading_ctrl_1']")
	private static WebElement searchPopUpSelectAllChkBox;

	@FindBy(xpath="//input[@id='id_body_23_search_popup_input']")
	private static WebElement searchPopUpSearchTXT;

	@FindBy(xpath="//input[@id='id_body_23_search_popup_category']")
	private static WebElement searchPopUpCategoryTxt;

	@FindBy(xpath="//button[contains(text(),'Refresh')]")
	private static WebElement searchPopUpRefreshbtn;

	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr/td[2]")
	private static List<WebElement> vouEntryPageBodyCol1List;



	public boolean checkItemValuesintoVoucherLevelOnClickOnOkBtnInSearchPopUp() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());	
		
		Thread.sleep(1999);

		int count3 =vouEntryPageBodyCol1List.size();
		ArrayList<String >array3= new ArrayList<>();
		for (int i = 0; i < count3; i++) 
		{
			String data=vouEntryPageBodyCol1List.get(i).getText();
			array3.add(data);
		}

		String actsearchPopUpAliasList=array3.toString();
		String expsearchPopUpAliasList=excelReader.getCellData(xlSheetName, 147, 6);
		System.out.println(" ACT searchPopUpAliasList : "+actsearchPopUpAliasList);
		System.out.println(" EXP searchPopUpAliasList : "+expsearchPopUpAliasList);
		excelReader.setCellData(xlfile, xlSheetName, 147, 7, actsearchPopUpAliasList);

		Thread.sleep(1999);

		getDriver().navigate().refresh();

		if (actsearchPopUpAliasList.equalsIgnoreCase(expsearchPopUpAliasList))
		{
			excelReader.setCellData(xlfile, xlSheetName, 147, 8, resPass);
			System.out.println(" Test Pass : Item Values are Loaded From Search Pop Up Selected " );
			return true;

		}
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 147, 8, resFail);
			System.out.println(" Test FAIl : Item Values are Loaded From Search Pop Up Selected " );
			return false;
		}

	}



	@FindBy(xpath="//input[@id='id_header_3']")
	private static WebElement  MRpurchaseAccountTxt;


	@FindBy(xpath="//input[@id='id_header_3_search_popup_input']")
	private static WebElement  accSearchPopSearchTXt;



	public boolean checkAccountTxtWithClickOnF5() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(finacinalsMenu));
		finacinalsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesMenu));
		purchasesMenu.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucherMenu));
		purchaseVoucherMenu.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("");

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 148, 5));

		Thread.sleep(1999);
		departmentTxt.sendKeys(Keys.TAB);



		getAction().moveToElement(departmentTxt).doubleClick().build().perform();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(Keys.chord(Keys.CONTROL,"c"));

		Thread.sleep(1999);
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		MRpurchaseAccountTxt.click();
		MRpurchaseAccountTxt.sendKeys(Keys.F5);

		Thread.sleep(1999);
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(accSearchPopSearchTXt));
		accSearchPopSearchTXt.sendKeys(Keys.chord(Keys.CONTROL,"v"));
		Thread.sleep(1999);

		accSearchPopSearchTXt.sendKeys(Keys.TAB);
		try {

			if (getIsAlertPresent())
			{
				String alert=getAlert().getText();
				System.out.println(" ALERT Displayed  : "+alert);

				if (errorMessage.isDisplayed()==true)
				{
					String Message=errorMessage.getText();
					System.out.println(" Message Displayed  : "+Message);
					errorMessageCloseBtn.click();
				}
			}
		} catch (Exception e) 
		{
			System.err.println(e.getMessage());
		}

		String actaccSearchPopSearchTXt=accSearchPopSearchTXt.getAttribute("value");
		String expaccSearchPopSearchTXt=excelReader.getCellData(xlSheetName, 148, 6);
		excelReader.setCellData(xlfile, xlSheetName, 148, 7, actaccSearchPopSearchTXt);

		System.out.println(" accSearchPopSearchTXt TXT : "+actaccSearchPopSearchTXt +" Value Exp : "+expaccSearchPopSearchTXt);

		Thread.sleep(1999);
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchPopUpRefreshbtn));
		searchPopUpRefreshbtn.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(accSearchPopSearchTXt));
		accSearchPopSearchTXt.click();
		accSearchPopSearchTXt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		accSearchPopSearchTXt.sendKeys(excelReader.getCellData(xlSheetName, 149, 5));

		Thread.sleep(2000);
		accSearchPopSearchTXt.sendKeys(Keys.TAB);

		try {

			if (getIsAlertPresent())
			{
				String alert=getAlert().getText();
				System.out.println(" ALERT Displayed  : "+alert);

				if (errorMessage.isDisplayed()==true)
				{
					String Message=errorMessage.getText();
					System.out.println(" Message Displayed  : "+Message);
					errorMessageCloseBtn.click();
				}
			}
		} catch (Exception e) 
		{
		}

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchPopUpRefreshbtn));
		searchPopUpRefreshbtn.click();

		Thread.sleep(1999);

		int count =searchAccPopUpNameList.size();
		ArrayList<String >array1= new ArrayList<>();
		for (int i = 0; i < count; i++) 
		{
			String data=searchAccPopUpNameList.get(i).getText();
			array1.add(data);
		}

		String actsearchPopUpNameList=array1.toString();
		String expsearchPopUpNameList=excelReader.getCellData(xlSheetName, 149, 6);
		System.out.println(" ACT searchAccPopUpNameList : "+actsearchPopUpNameList);
		System.out.println(" EXP searchAccPopUpNameList : "+expsearchPopUpNameList);
		excelReader.setCellData(xlfile, xlSheetName, 149, 7, actsearchPopUpNameList);


		int count2 =searchAccPopUpCodeList.size();
		ArrayList<String >array2= new ArrayList<>();
		for (int i = 0; i < count2; i++) 
		{
			String data=searchAccPopUpCodeList.get(i).getText();
			array2.add(data);
		}

		String actsearchPopUpCodeList=array2.toString();
		String expsearchPopUpCodeList=excelReader.getCellData(xlSheetName, 150, 6);
		System.out.println(" ACT searchAccPopUpCodeList : "+actsearchPopUpCodeList);
		System.out.println(" EXP searchAccPopUpCodeList : "+expsearchPopUpCodeList);
		excelReader.setCellData(xlfile, xlSheetName, 150, 7, actsearchPopUpCodeList);

		int count3 =searchAccPopUpAliasList.size();
		ArrayList<String >array3= new ArrayList<>();
		for (int i = 0; i < count3; i++) 
		{
			String data=searchAccPopUpAliasList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 150, 5)))
			{
				searchAccPopUpRadioBtnList.get(i).click();
			}
			array3.add(data);
		}

		String actsearchPopUpAliasList=array3.toString();
		String expsearchPopUpAliasList=excelReader.getCellData(xlSheetName, 151, 6);
		System.out.println(" ACT searchAccPopUpAliasList : "+actsearchPopUpAliasList);
		System.out.println(" EXp searchAccPopUpAliasList : "+expsearchPopUpAliasList);
		excelReader.setCellData(xlfile, xlSheetName, 151, 7, actsearchPopUpAliasList);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchAccPopUpOkBtn));
		searchAccPopUpOkBtn.click();

		if (actsearchPopUpCodeList.equalsIgnoreCase(expsearchPopUpCodeList) && 
				actsearchPopUpNameList.equalsIgnoreCase(expsearchPopUpNameList) && 
				actsearchPopUpAliasList.equalsIgnoreCase(expsearchPopUpAliasList))
		{
			excelReader.setCellData(xlfile, xlSheetName, 148, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 148, 8, resFail);
			return false;
		}
	}
	
	
	

	public boolean checkLoadingAccountValuesFromSearchPopUp() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		MRpurchaseAccountTxt.click();

		String actMRpurchaseAccountTxt=MRpurchaseAccountTxt.getAttribute("value");
		String expMRpurchaseAccountTxt=excelReader.getCellData(xlSheetName, 152, 6);
		excelReader.setCellData(xlfile, xlSheetName, 152, 7, actMRpurchaseAccountTxt);

		System.out.println(" actMRpurchaseAccountTxt : "+actMRpurchaseAccountTxt);
		System.out.println(" EXPMRpurchaseAccountTxt : "+expMRpurchaseAccountTxt);

		Thread.sleep(1999);
		
		getDriver().navigate().refresh();

		if (actMRpurchaseAccountTxt.equalsIgnoreCase(expMRpurchaseAccountTxt))
		{
			excelReader.setCellData(xlfile, xlSheetName, 152, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 152, 8, resFail);
			return false;
		}
	}
	
	
	

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

	public boolean checkF5KeyWithRMAPopUpInOpeningStocks() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(1999);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksVoucher));
		openingStocksVoucher.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.click();
		warehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 153, 5));
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		getAction().moveToElement(warehouseTxt).doubleClick().build().perform();

		warehouseTxt.sendKeys(Keys.chord(Keys.CONTROL,"c"));

		Thread.sleep(1999);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 154, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 155, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 156, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
		rmaSerialNumberTxtField.sendKeys(Keys.chord(Keys.CONTROL,"v"));

		rmaSerialNumberTxtField.sendKeys(Keys.TAB);

		String actrmaSerialNumberTxtField=rmaSerialNumberTxtField.getAttribute("value");
		String exprmaSerialNumberTxtField=excelReader.getCellData(xlSheetName, 153, 6);
		excelReader.setCellData(xlfile, xlSheetName, 153, 7, actrmaSerialNumberTxtField);

		System.out.println("_______rmaSerialNumberTxtField :"+actrmaSerialNumberTxtField +" value exp:"+exprmaSerialNumberTxtField);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupCancelBtn));
		RMAPopupCancelBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();

		Thread.sleep(1999);

		rmaSerialNumberTxtField.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);

		rmaSerialNumberTxtField.sendKeys(excelReader.getCellData(xlSheetName, 157, 5));

		rmaSerialNumberTxtField.sendKeys(Keys.TAB);

		String actrmaSerialNumberTxtFieldManual=rmaSerialNumberTxtField.getAttribute("value");
		String exprmaSerialNumberTxtFieldManual=excelReader.getCellData(xlSheetName, 154, 6);
		excelReader.setCellData(xlfile, xlSheetName, 154, 7, actrmaSerialNumberTxtFieldManual);

		System.out.println(" __________________rmaSerialNumberTxtField Manual:"+actrmaSerialNumberTxtFieldManual +" value exp:"+exprmaSerialNumberTxtFieldManual);

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
		rmaAddBtn.click();

		//actual
		String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
		String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
		String actRMA3=RMAPopupTableRow3Column1.getAttribute("data-value");
		String actRMA4=RMAPopupTableRow4Column1.getAttribute("data-value");
		String actRMA5=RMAPopupTableRow5Column1.getAttribute("data-value");
		String actRMA6=RMAPopupTableRow6Column1.getAttribute("data-value");

		//expected
		String expRMA1=excelReader.getCellData(xlSheetName, 155, 6);
		excelReader.setCellData(xlfile, xlSheetName, 155, 7, actRMA1);
		
		String expRMA2=excelReader.getCellData(xlSheetName, 156, 6);
		excelReader.setCellData(xlfile, xlSheetName, 156, 7, actRMA2);
		
		String expRMA3=excelReader.getCellData(xlSheetName, 157, 6);
		excelReader.setCellData(xlfile, xlSheetName, 157, 7, actRMA3);
		
		String expRMA4=excelReader.getCellData(xlSheetName, 158, 6);
		excelReader.setCellData(xlfile, xlSheetName, 158, 7, actRMA4);
		
		String expRMA5=excelReader.getCellData(xlSheetName, 159, 6);
		excelReader.setCellData(xlfile, xlSheetName, 159, 7, actRMA5);
		
		String expRMA6=excelReader.getCellData(xlSheetName, 160, 6);
		excelReader.setCellData(xlfile, xlSheetName, 160, 7, actRMA6);

		System.out.println("********* Row1        : "+actRMA1  +"  value expected  "+expRMA1);
		System.out.println("********* Row2        : "+actRMA2  +"  value expected  "+expRMA2);
		System.out.println("********* Row3        : "+actRMA3  +"  value expected  "+expRMA3);
		System.out.println("********* Row4        : "+actRMA4  +"  value expected  "+expRMA4);
		System.out.println("********* Row5        : "+actRMA5  +"  value expected  "+expRMA5);
		System.out.println("********* Row6        : "+actRMA6  +"  value expected  "+expRMA6);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
		RMAPopupOkBtn.click();

		Thread.sleep(2000);

		if(actRMA1.equalsIgnoreCase(expRMA1) && actRMA2.equalsIgnoreCase(expRMA2) && actRMA3.equalsIgnoreCase(expRMA3) && actRMA4.equalsIgnoreCase(expRMA4) 
				&& actRMA5.equalsIgnoreCase(expRMA5) && actRMA6.equalsIgnoreCase(expRMA6) && actrmaSerialNumberTxtField.equalsIgnoreCase(exprmaSerialNumberTxtField) &&
				actrmaSerialNumberTxtFieldManual.equalsIgnoreCase(exprmaSerialNumberTxtFieldManual))
		{
			excelReader.setCellData(xlfile, xlSheetName, 153, 8, resPass);
			System.out.println(" Test Pass : Values are displayed through copy and Paste in RMA POP ");
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 153, 8, resPass);
			System.out.println(" Test FAIl : Values are displayed through copy and Paste in RMA POP  ");
			return false;
		}
	}
	
	
	


	@FindBy (xpath="//*[@id='div_OK']")
	private static WebElement RMAPopupOkBtn;

	@FindBy (xpath="(//*[@id='div_Clear'])[2]")
	private static WebElement RMAPopupCancelBtn;


	@FindBy (xpath="//div[@id='Modal_Header']")
	private static WebElement rmaScreenTitle;

	@FindBy (xpath="//input[@id='txtSerialNo']")
	private static WebElement rmaSerialNumberTxtField;

	@FindBy (xpath="//input[@id='txtQuantity']")
	private static WebElement rmaQuantityTxtField;

	@FindBy (xpath="//*[@tabindex='4']")
	private static WebElement rmaAddBtn;

	@FindBy(xpath="//*[@id='81']/span")
	private static WebElement  financialsReportsMenu; 

	@FindBy(xpath="//*[@id='500']/span")
	private static WebElement  ledger;

	@FindBy(xpath="//input[@id='cmbUserTypeMaster']")
	private static WebElement  ledgerSearchBar;


	public static void click(WebElement element)
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(element));
		element.click();
	}


	public static void alerOrMessageDisplayed()
	{
		try {

			if (getIsAlertPresent())
			{
				String alert=getAlert().getText();
				System.err.println(" ALERT Displayed  : "+alert);

				if (errorMessage.isDisplayed()==true)
				{
					String Message=errorMessage.getText();
					System.err.println(" Message Displayed  : "+Message);
					errorMessageCloseBtn.click();
				}
			}
		} catch (Exception e) 
		{
		}

	}

	@FindBy(xpath="//td[contains(text(),'ASSETS')]")
	private static WebElement ledgerAssetsTXT;

//	@FindBy(xpath="//input[@id='txtSearchMenu_MainLayout']")
	@FindBy(xpath="//input[@id='cmbUserTypeMaster']")
	private static WebElement mainMenuTxtArea;



	public boolean checkCopyAndPasteWithControlOptionInLedgerReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsReportsMenu));
		financialsReportsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ledger));
		ledger.click();

		Thread.sleep(3000);

		click(mainMenuTxtArea);
		
		mainMenuTxtArea.sendKeys(excelReader.getCellData(xlSheetName, 161, 5));

		getAction().moveToElement(mainMenuTxtArea).doubleClick().build().perform();
		mainMenuTxtArea.sendKeys(Keys.chord(Keys.CONTROL,"c"));

		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ledgerSearchBar));

		System.out.println("*******Copied ");

		Thread.sleep(1000);

		click(ledgerSearchBar);
		ledgerSearchBar.clear();

		ledgerSearchBar.sendKeys(Keys.chord(Keys.CONTROL,"v"));
		
		System.out.println("*******Pasted ");

		alerOrMessageDisplayed();

		Thread.sleep(1000);
		String actledgerAssetsTXT=ledgerSearchBar.getAttribute("value");
		String expledgerAssetsTXT=excelReader.getCellData(xlSheetName, 161, 6);
		excelReader.setCellData(xlfile, xlSheetName, 161, 7, actledgerAssetsTXT);

		System.out.println(" Paste Text Area : "+actledgerAssetsTXT +" Value Exp : "+expledgerAssetsTXT);

		Thread.sleep(1000);
		ledgerSearchBar.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);

		ledgerSearchBar.sendKeys(excelReader.getCellData(xlSheetName, 162, 5));
		Thread.sleep(1000);
		ledgerSearchBar.sendKeys(Keys.TAB);

		Thread.sleep(1000);
		
		String actSendKeys=ledgerSearchBar.getAttribute("value");
		String expSendKeys=excelReader.getCellData(xlSheetName, 162, 6);
		excelReader.setCellData(xlfile, xlSheetName, 162, 7, actSendKeys);

		System.out.println(" SendKeys Text Area : "+actSendKeys+" Value Exp : "+expSendKeys);

		Thread.sleep(1000);
		
		alerOrMessageDisplayed();

		if (actledgerAssetsTXT.equalsIgnoreCase(expledgerAssetsTXT)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 161, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 161, 8, resFail);
			return false;
		}
	}



	@FindBy(xpath="//*[@id='idRecentMenus_Main']/i")
	private static WebElement recentMenusIcon;

	@FindBy(xpath="//*[@id='recentMenuUL']/li")
	private static List<WebElement> recentMenusList;

	public boolean checkEnterF5keyInLedgerReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		click(recentMenusIcon);

		Thread.sleep(12000);
		
		int count=recentMenusList.size();
		for (int i = 0; i < count; i++)
		{
			String data=recentMenusList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 163, 5)))
			{
				recentMenusList.get(i).click();
			}
		}

		Thread.sleep(12000);

		click(ledgerSearchBar);

		ledgerSearchBar.sendKeys(Keys.F5);


		click(reportSearchPopSearchTxt);

		reportSearchPopSearchTxt.sendKeys(excelReader.getCellData(xlSheetName, 164, 5));
		
		Thread.sleep(1000);

		click(searchPopUpRefreshbtn);

		Thread.sleep(1999);

		int count1 =reportSearchPopNameList.size();
		
		ArrayList<String >array1= new ArrayList<>();
		
		for (int i = 0; i < count1; i++) 
		{
			String data=reportSearchPopNameList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 165, 5))) 
			{
				reportSearchPopRadioBtn.get(i).click();
			}
			array1.add(data);
		}

		String actsearchPopUpNameList=array1.toString();
		String expsearchPopUpNameList=excelReader.getCellData(xlSheetName, 163, 6);
		System.out.println(" ACT searchAccPopUpNameList : "+actsearchPopUpNameList);
		System.out.println(" EXP searchAccPopUpNameList : "+expsearchPopUpNameList);
		excelReader.setCellData(xlfile, xlSheetName, 163, 7, actsearchPopUpNameList);

		click(reportSearchPopOkbtn);

		Thread.sleep(1000);
		String actSendKeys=ledgerSearchBar.getAttribute("value");
		String expSendKeys=excelReader.getCellData(xlSheetName, 165, 6);
		excelReader.setCellData(xlfile, xlSheetName, 165, 7, actSendKeys);
		
		System.out.println(" SendKeys Text Area : "+actSendKeys+" Value Exp : "+expSendKeys);

		Thread.sleep(1000);
		
		alerOrMessageDisplayed();

		if (actsearchPopUpNameList.equalsIgnoreCase(expsearchPopUpNameList) &&
				actSendKeys.equalsIgnoreCase(expSendKeys))
		{
			excelReader.setCellData(xlfile, xlSheetName, 163, 8, resPass);
			System.out.println("***************** Dispalyed Cogs Related Voucher ");
			return true;
		} else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 163, 8, resFail);
			System.out.println(" *******************NOT Dispalyed Cogs Related Voucher ");
			return false;
		}
	}

	@FindBy(xpath="//input[@id='cmbUserTypeMaster_search_popup_input']")
	private static WebElement reportSearchPopSearchTxt;

	@FindBy(xpath="//*[@id='cmbUserTypeMaster_search_popup']/div/div/div[3]/div/div[2]/input[1]")
	private static WebElement reportSearchPopOkbtn;

	@FindBy(xpath="//*[@id='cmbUserTypeMaster_search_popup_grid_body']/tr/td[4]")
	private static List<WebElement> reportSearchPopNameList;


	@FindBy(xpath="//*[@id='cmbUserTypeMaster_search_popup_grid_body']/tr/td[2]/input")
	private static List<WebElement> reportSearchPopRadioBtn;


	@FindBy(xpath="//*[@id='reportViewFilterBtn']")
	private static WebElement reportFilterBtn;

	@FindBy(xpath="//span[@id='idFilterCustomizeIcon']")
	private static WebElement reportFilterCustBtn;

	@FindBy(xpath="(//a[contains(text(),'Account')])[1]")
	private static WebElement filterCusAccExpandBtn;

	@FindBy(xpath="(//label[contains(text(),'Name')])[1]")
	private static WebElement filterCusAccNameChkbox;

	@FindBy(xpath="//*[@id='FOption_500_0_DefaultFilter_0']")
	private static WebElement filterAccTxt;


	@FindBy(xpath="//*[@id='filter_Okbtn_']")
	private static WebElement filterOkBtn;

	@FindBy(xpath="//input[@id='FOption_500_0_DefaultFilter_0_search_popup_input']")
	private static WebElement filterSearchPopUpSearchTxt;

	@FindBy(xpath="//*[@id='FOption_500_0_DefaultFilter_0_search_popup_grid_body']/tr/td[4]")
	private static List<WebElement> filterSearchPopUpNameList;

	@FindBy(xpath="//*[@id='FOption_500_0_DefaultFilter_0_search_popup_grid_body']/tr/td[2]/input")
	private static List<WebElement> filterSearchPopUpRadio;

	@FindBy(xpath="//*[@id='FOption_500_0_DefaultFilter_0_search_popup']/div/div/div[3]/div/div[2]/input[1]")
	private static WebElement filterSearchPopUpOkBtn;

	//@FindBy(xpath="//*[@id='FilterFieldCust_500_0']/div/div[3]/button[1]")
	@FindBy(xpath="(//*[@value='Ok'])[1]")
	private static WebElement cusFilterOkBtn;



	public boolean checkF5KeyInFilterPopUpScreen() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		click(recentMenusIcon);

		Thread.sleep(3000);
		int count=recentMenusList.size();
		for (int i = 0; i < count; i++)
		{
			String data=recentMenusList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 166, 5)))
			{
				recentMenusList.get(i).click();
			}
		}

		Thread.sleep(3000);


		click(reportFilterBtn);

		Thread.sleep(1000);
		click(reportFilterCustBtn);

		Thread.sleep(1000);
		click(filterCusAccExpandBtn);

		Thread.sleep(1000);
		click(filterCusAccNameChkbox);

		Thread.sleep(1000);
		click(cusFilterOkBtn);

		Thread.sleep(1000);
		click(filterAccTxt);

		filterAccTxt.sendKeys(excelReader.getCellData(xlSheetName, 167, 5));
		Thread.sleep(1000);
		filterAccTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		Thread.sleep(1000);
		filterAccTxt.sendKeys(Keys.chord(Keys.CONTROL,"c"));

		Thread.sleep(1999);
		filterAccTxt.sendKeys(Keys.F5);

		filterSearchPopUpSearchTxt.sendKeys(Keys.chord(Keys.CONTROL,"v"));


		String actfilterSearchPopUpSearchTxt=filterSearchPopUpSearchTxt.getAttribute("value");
		String expfilterSearchPopUpSearchTxt=excelReader.getCellData(xlSheetName, 166, 6);
		excelReader.setCellData(xlfile, xlSheetName, 166, 7, actfilterSearchPopUpSearchTxt);

		System.out.println(" filterSearchPopUpSearchTxt : "+actfilterSearchPopUpSearchTxt+" Value Exp : "+expfilterSearchPopUpSearchTxt);

		Thread.sleep(2000);
		click(filterSearchPopUpSearchTxt);
		filterSearchPopUpSearchTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		filterSearchPopUpSearchTxt.sendKeys(excelReader.getCellData(xlSheetName, 168, 5));

		Thread.sleep(2000);
		String actSendKeys=filterSearchPopUpSearchTxt.getAttribute("value");
		String expSendKeys=excelReader.getCellData(xlSheetName, 167, 6);
		System.out.println(" SendKeys : "+actSendKeys+" Value Exp : "+expSendKeys);
		excelReader.setCellData(xlfile, xlSheetName, 167, 7, actSendKeys);
		
		click(searchPopUpRefreshbtn);

		int count1 =filterSearchPopUpNameList.size();
		ArrayList<String >array1= new ArrayList<>();
		for (int i = 0; i < count1; i++) 
		{
			String data=filterSearchPopUpNameList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 169, 5))) 
			{
				filterSearchPopUpRadio.get(i).click();
			}
			array1.add(data);
		}

		String actsearchPopUpNameList=array1.toString();
		String expsearchPopUpNameList=excelReader.getCellData(xlSheetName, 168, 6);
		System.out.println(" ACT searchAccPopUpNameList : "+actsearchPopUpNameList);
		System.out.println(" EXP searchAccPopUpNameList : "+expsearchPopUpNameList);
		excelReader.setCellData(xlfile, xlSheetName, 168, 7, actsearchPopUpNameList);

		click(filterSearchPopUpOkBtn);

		Thread.sleep(1000);

		String actAfterF5=filterAccTxt.getAttribute("value");
		String expAfterF5=excelReader.getCellData(xlSheetName, 169, 6);
		excelReader.setCellData(xlfile, xlSheetName, 169, 7, actAfterF5);

		System.out.println(" AfterF5 : "+actAfterF5+" Value Exp : "+expAfterF5);

		click(filterOkBtn);

		if (actAfterF5.equalsIgnoreCase(expAfterF5) &&
				actsearchPopUpNameList.equalsIgnoreCase(expsearchPopUpNameList) &&
				actfilterSearchPopUpSearchTxt.equalsIgnoreCase(expfilterSearchPopUpSearchTxt))
		{
			excelReader.setCellData(xlfile, xlSheetName, 166, 8, resPass);
			return true;

		} else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 166, 8, resFail);
			return false;

		}
	}





















	// Miscellaneous Tab Restrictions

	@FindBy(xpath="//*[@id='2007']/span")
	private static WebElement  financialsTransactionsPurchaseMenu; 

	@FindBy(xpath="//span[contains(text(),'Purchase Voucher VAT')]")
	private static WebElement  purchaseVouchersVat;

//	@FindBy(xpath="//span[@class='icon-menu icon-font4']")
	@FindBy(xpath="//*[@id='id_transactionentry_more']")
	private static WebElement  toggleBtn;

	@FindBy(xpath="//*[@id='id_transactionentry_settings']")
	private static WebElement  settingsOption;

	@FindBy(xpath="//*[@id='navigationtab2']")
	private static WebElement  miscellaneousTab;
	
	@FindBy(xpath="//*[@class='icon-settings hiconright2']")
	private static WebElement  settingsIcon;

	
	@FindBy(xpath="//*[@id='panelsStayOpen-headingsix']")
	private static WebElement  restrictionsTab;

	@FindBy(xpath="//input[@id='misc_restrictionRdStop']")
	private static WebElement  stopRadio;

	@FindBy(xpath="//*[@id='panelsStayOpen-headingsix']")
	private static WebElement  RestrictionsTab;
	
	@FindBy(xpath="//input[@id='misc_restrictionRdwarnandAllow']")
	private static WebElement  warnAndAllowRadio;

	@FindBy(xpath="//input[@id='misc_restrictionentryCondition_textbox']")
	private static WebElement  entryRestrictConditionTxt;

	@FindBy(xpath="//input[@id='misc_restrictionentryCondition_formulaText']")
	private static WebElement  miscformulaTxt;

	@FindBy(xpath="//div[@id='misc_restrictionentryCondition_expandedFormula']")
	private static WebElement  miscexpandeFormula;

	@FindBy(xpath="//*[@id='misc_restrictionentryCondition_Ok']")
	private static WebElement miscexpandeFormulaOkBtn;

	@FindBy(xpath="//div[@id='misc_restrictionentryCondition_formulaDiv']/div[6]/button[text()='Cancel']")
	private static WebElement miscexpandeFormulaCancelBtn;

	@FindBy(xpath="//input[@id='misc_restrictionEntryMessage']")
	private static WebElement  entryRestrictMessageTxt;

	@FindBy(xpath="//*[@id='btnCustomizeClose']")
	private static WebElement  settings_closeBtn;

	public static boolean checkCreatingConditionForRestrictionsWithStopOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		financialsTransactionsPurchaseMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
		purchaseVouchersVat.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));
		toggleBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsOption));
		settingsOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousTab));
		miscellaneousTab.click();

		Thread.sleep(2000);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionsTab));
		restrictionsTab.click();
		Thread.sleep(2000);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(entryRestrictConditionTxt));
		getAction().moveToElement(entryRestrictConditionTxt).build().perform();

		Thread.sleep(1000);
		entryRestrictConditionTxt.click();

		entryRestrictConditionTxt.sendKeys(excelReader.getCellData(xlSheetName, 170, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscexpandeFormulaOkBtn));
		getAction().moveToElement(miscexpandeFormulaOkBtn).click().build().perform();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
		settings_closeBtn.click();
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsIcon));
		settingsIcon.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousTab));
		miscellaneousTab.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionsTab));
		restrictionsTab.click();
		Thread.sleep(2000);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(entryRestrictMessageTxt));
		entryRestrictMessageTxt.click();
		entryRestrictMessageTxt.sendKeys(excelReader.getCellData(xlSheetName, 171, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		String expMessage=excelReader.getCellData(xlSheetName, 170, 6);

		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 170, 7, actMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 170, 8, resPass);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();
			Thread.sleep(3000);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 170, 8, resFail);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();
			Thread.sleep(3000);
			return false;
		}
	}






	@FindBy(xpath="//input[@id='id_header_4']")
	private static WebElement  vendorAccountTxt;

	@FindBy(xpath="//input[@id='id_body_536870916']")
	private static WebElement enter_WarehouseTxt;

	@FindBy(xpath="//input[@id='id_body_16777331']")
	private static WebElement  enter_TaxCode;

	@FindBy(xpath="//input[@id='id_body_12']")
	private static WebElement  enter_PurchaseAccountTxt;

	@FindBy(xpath="//*[@class='icon-pick']")
	private static WebElement pickBtn;

	@FindBy(xpath="//input[@id='txtNewReference']")
	private static WebElement newReferenceTxt;

	@FindBy(xpath="//*[@class='icon-ok']")
	private static WebElement Bill_OkBtn;



	public static boolean checkSavingPurchaseVoucherVATWithRateBelowAverageRate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
					financialsMenu.click();

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
					financialsTransactionMenu.click();

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
					financialsTransactionsPurchaseMenu.click();

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
					purchaseVouchersVat.click();

					Thread.sleep(2000);*/
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		/*vendorAccountTxt.click();
		Thread.sleep(1000);*/
		vendorAccountTxt.sendKeys("Vendor B");
		//vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 172, 5));	
		Thread.sleep(3000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys("INDIA");
		//departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 173, 5));	
		Thread.sleep(3000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.sendKeys("Abu Dhabi");
		
		//placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 174, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.sendKeys("Dubai");
		//jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 175, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys("HYDERABAD");
		//enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 176, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys("STD RATE COGS ITEM");
		//enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 177, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("2");
		//enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 178, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
		String avgRate = select1stRow_7thColumn.getText();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("9");
		//enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 179, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*String expMessage = excelReader.getCellData(xlSheetName, 172, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 172, 7, actMessage);
*/
		
		String expMessage = "Rate is Less Than or Equal to Average Rate.";
		String actMessage = checkValidationMessage(expMessage);
		

		System.out.println(actMessage);
		System.out.println(expMessage);

		if(actMessage.equals(expMessage))
		{
			//excelReader.setCellData(xlfile, xlSheetName, 172, 8, resPass);
			return true;
		} 
		else 
		{
			//excelReader.setCellData(xlfile, xlSheetName, 172, 8, resFail);
			return false;
		}
	}






	public static boolean checkSavingPurchaseVoucherVATWithRateEqualToAverageRate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		//enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 180, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		/*String expMessage = excelReader.getCellData(xlSheetName, 180, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 180, 7, actMessage);
*/
		
		String expMessage = "Rate is Less Than or Equal to Average Rate.";
		String actMessage = checkValidationMessage(expMessage);
		

		System.out.println(actMessage);
		System.out.println(expMessage);

		if(actMessage.equals(expMessage))
		{
			//excelReader.setCellData(xlfile, xlSheetName, 180, 8, resPass);
			return true;
		} 
		else 
		{
			//excelReader.setCellData(xlfile, xlSheetName, 180, 8, resFail);
			return false;
		}
	}






	public static boolean checkSavingPurchaseVoucherVATWithRateGreaterThanAverageRate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("11");
		//enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 181, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		HashSet<String> actMsg = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMsg.add(data);
		}

		String actMessage = actMsg.toString();
		String expMessage = excelReader.getCellData(xlSheetName, 181, 6);
		excelReader.setCellData(xlfile, xlSheetName, 181, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 181, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 181, 8, resFail);
			return false;
		}
	}







	public static boolean checkCreatingConditionForRestrictionsWithWarnAndAllowOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		financialsTransactionsPurchaseMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
		purchaseVouchersVat.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));
		toggleBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsOption));
		settingsOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousTab));
		miscellaneousTab.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RestrictionsTab));
		RestrictionsTab.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warnAndAllowRadio));
		warnAndAllowRadio.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		String expMessage=excelReader.getCellData(xlSheetName, 182, 6);

		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 182, 7, actMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 182, 8, resPass);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();
			Thread.sleep(3000);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 182, 8, resFail);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();
			Thread.sleep(3000);
			return false;
		}
	}








	public static boolean checkSavingPurchaseVoucherVATWithRateBelowAverageRateForWarnAndAllow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 183, 5));	
		Thread.sleep(3000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 184, 5));	
		Thread.sleep(3000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 185, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 186, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 187, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 188, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 189, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
		String avgRate = select1stRow_7thColumn.getText();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 190, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getWaitForAlert();

		String actWarnMessage = getAlert().getText();
		String expWarnMessage = excelReader.getCellData(xlSheetName, 183, 6);
		excelReader.setCellData(xlfile, xlSheetName, 183, 7, actWarnMessage);
		
		getAlert().accept();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		HashSet<String> actMsg = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMsg.add(data);
		}

		String actMessage = actMsg.toString();
		String expMessage = excelReader.getCellData(xlSheetName, 185, 6);
		excelReader.setCellData(xlfile, xlSheetName, 185, 7, actMessage);
		
		System.out.println(actMessage);
		System.out.println(expMessage);

		if(actWarnMessage.equalsIgnoreCase(expWarnMessage) && actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 184, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 184, 8, resFail);
			return false;
		}
	}

	
	





	public static boolean checkSavingPurchaseVoucherVATWithRateEqualToAverageRateWithWarnAndAllow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 191, 5));	
		Thread.sleep(3000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 192, 5));	
		Thread.sleep(3000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 193, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 194, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 195, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 196, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 197, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 198, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getWaitForAlert();

		String actWarnMessage = getAlert().getText();
		String expWarnMessage = excelReader.getCellData(xlSheetName, 191, 6);
		excelReader.setCellData(xlfile, xlSheetName, 191, 7, actWarnMessage);

		getAlert().accept();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		HashSet<String> actMsg = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMsg.add(data);
		}
		
		String actMessage = actMsg.toString();
		String expMessage = excelReader.getCellData(xlSheetName, 193, 6);
		excelReader.setCellData(xlfile, xlSheetName, 193, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);

		if(actWarnMessage.equalsIgnoreCase(expWarnMessage) && actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 191, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 191, 8, resFail);
			return false;
		}
	}






	public static boolean checkSavingPurchaseVoucherVATWithRateGreaterThanAverageRateWithWarnAndAllow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 199, 5));	
		Thread.sleep(3000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 200, 5));	
		Thread.sleep(3000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 201, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 202, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 203, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 204, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 205, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 206, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		HashSet<String> actMsg = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMsg.add(data);
		}

		String actMessage = actMsg.toString();
		String expMessage = excelReader.getCellData(xlSheetName, 199, 6);
		excelReader.setCellData(xlfile, xlSheetName, 199, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);

		if(actMessage.equals(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 199, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 199, 8, resFail);
			return false;
		}
	}






	@FindBy(xpath="//*[@class='icon-custamize hiconright2']")
	private static WebElement reportScreenCustomizeBtn;

	@FindBy(xpath="//input[@id='liSelectAllMasters']")
	private static WebElement sl_HeaderSelectChkBox;

	@FindBy(xpath="//i[contains(@class,'icon-ok hiconright2')]")
	private static WebElement sl_OkBtn;

	@FindBy(xpath="//*[@id='id_customize_headerfooter_btn']")
	private static WebElement sl_cusHeaderAndFooter;

	@FindBy(xpath="//*[@id='forHf']/span[3]/span/i[1]")
	private static WebElement sl_cusHeaderAndFooterSaveBtn;

	@FindBy(xpath="//li[@id='forHf']/span[2]")
	private static WebElement sl_cusHeaderAndFooterOpenBtn;

	@FindBy(xpath="//li[@id='forHf']/span[2]/ul/li/a[2]")
	private static WebElement sl_cusHeaderAndFooterOpenFromXMLBtn;

	@FindBy(xpath="//span[@id='id_InvoiceDesingCancel']")
	private static WebElement sl_cusHeaderAndFooterExitBtn;




	@FindBy(xpath="//a[contains(text(),'Save')]")
	private static WebElement HFsave_SaveBtn;

	@FindBy(xpath="//input[@id='id_PopuplayoutName']")
	private static WebElement HFsavePop_SaveTxt;

	@FindBy(xpath="//button[@id='btn_PopUpOk']")
	private static WebElement HFsavePop_OKBtn;

	@FindBy(xpath="//select[@id='DateOptions_']")
	private static WebElement sl_DateOptionDropdown;

	@FindBy(xpath="//*[@id='trRender_0']/td[1]")
	private static WebElement sl_1stRow1stCol;

	@FindBy(xpath="//*[@id='Save']/div[1]/span")
	private static WebElement  osr_customizeSaveBtn;

	
	
	

	public boolean checkAddingHeadeAndFooterInLedgerReport() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsReportsMenu));
		financialsReportsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ledger));
		ledger.click();

		Thread.sleep(3999);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		sl_DateOptionDropdown.click();
		Select Date=new Select(sl_DateOptionDropdown);
		Date.selectByValue(excelReader.getCellData(xlSheetName, 207, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_HeaderSelectChkBox));
		sl_HeaderSelectChkBox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();

		Thread.sleep(1500);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reportScreenCustomizeBtn));
		reportScreenCustomizeBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_cusHeaderAndFooter));
		sl_cusHeaderAndFooter.click();

		//autoit

		Thread.sleep(4000);

		Runtime.getRuntime().exec(getBaseDir()+excelReader.getCellData(xlSheetName, 208, 5));

		Thread.sleep(18000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_cusHeaderAndFooterSaveBtn));
		sl_cusHeaderAndFooterSaveBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(HFsave_SaveBtn));
		HFsave_SaveBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(HFsavePop_SaveTxt));
		HFsavePop_SaveTxt.sendKeys(excelReader.getCellData(xlSheetName, 209, 5));

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(HFsavePop_OKBtn));
		HFsavePop_OKBtn.click();

		String expLayoutMessage = excelReader.getCellData(xlSheetName, 207, 6);
		String actLayoutMessage = checkValidationMessage(expLayoutMessage);
		excelReader.setCellData(xlfile, xlSheetName, 207, 7, actLayoutMessage);

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_cusHeaderAndFooterExitBtn));
		sl_cusHeaderAndFooterExitBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_customizeSaveBtn));
		osr_customizeSaveBtn.click();

		String expCustMessage = excelReader.getCellData(xlSheetName, 208, 6);
		String actCustMessage = checkValidationMessage(expCustMessage);
		excelReader.setCellData(xlfile, xlSheetName, 208, 7, actCustMessage);

		Thread.sleep(2000);

		if (actLayoutMessage.equalsIgnoreCase(expLayoutMessage) && actCustMessage.equalsIgnoreCase(expCustMessage)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 207, 8, resPass);
			return true;
		}
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 207, 8, resFail);
			return false;
		}
	}






	@FindBy(xpath="//div[@id='REPORTRENDERNEWControls']/ul/li/span[4]")
	private static WebElement sl_ExportBtn;

	@FindBy(xpath="//div[@id='REPORTRENDERNEWControls']/ul/li/span[4]/ul/li[2]")
	private static WebElement sl_ExportPDFBtn;

	public boolean checkPrintPDFFileAfterCreatingHeadeAndFooterLayoutInLedgerReport() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException, AWTException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_ExportBtn));
		sl_ExportBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_ExportPDFBtn));
		sl_ExportPDFBtn.click();

		Thread.sleep(5000);

		String Filename = checkDownloadedFileName(getDriver());

		System.err.println(Filename);

		String actPDF = getBaseDir()+"\\autoIt\\ExportFiles\\"+Filename;
		String expPDF = getBaseDir()+"\\autoIt\\ImportFiles\\LedgerReportWithCreatedHeaderAndFooterLayout.pdf";

		PDFUtil pdfutil = new PDFUtil();

		boolean result = pdfutil.compare(actPDF, expPDF);

		String data = pdfutil.getText(expPDF);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		Calendar cal = Calendar.getInstance();
		String date = df.format(cal.getTime());
		System.err.println(date);

		String oldDate = "08/10/2021";

		String actData = pdfutil.getText(actPDF);
		String expData = data.replace(oldDate, date);

		System.err.println(actData);
		System.err.println(expData);
		
		excelReader.setCellData(xlfile, xlSheetName, 210, 7, actData);

		System.out.println("Compared Result  : "+result);

		if (actData.equalsIgnoreCase(expData)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 210, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 210, 8, resFail);
			return false;
		}
	}






	@FindBy(xpath="//a[@id='200']//span[contains(text(),'Reports')]")
	private static WebElement  inventoryReportsMenu; 

	@FindBy(xpath="//span[contains(text(),'Stock Ledger')]")
	private static WebElement  stockLedger;



	public boolean checkStockLedgerReportImportingHeaderAndFooterLayout() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedger));
		stockLedger.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		sl_DateOptionDropdown.click();
		Select Date=new Select(sl_DateOptionDropdown);
		Date.selectByValue(excelReader.getCellData(xlSheetName, 211, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_HeaderSelectChkBox));
		sl_HeaderSelectChkBox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reportScreenCustomizeBtn));
		reportScreenCustomizeBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_cusHeaderAndFooter));
		sl_cusHeaderAndFooter.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_cusHeaderAndFooterOpenBtn));
		sl_cusHeaderAndFooterOpenBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_cusHeaderAndFooterOpenFromXMLBtn));
		sl_cusHeaderAndFooterOpenFromXMLBtn.click();

		Thread.sleep(2000);

		Runtime.getRuntime().exec(getBaseDir()+excelReader.getCellData(xlSheetName, 212, 5));

		Thread.sleep(10000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_cusHeaderAndFooterSaveBtn));
		sl_cusHeaderAndFooterSaveBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(HFsave_SaveBtn));
		HFsave_SaveBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(HFsavePop_SaveTxt));
		HFsavePop_SaveTxt.sendKeys(excelReader.getCellData(xlSheetName, 213, 5));

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(HFsavePop_OKBtn));
		HFsavePop_OKBtn.click();

		String expLayoutMessage = excelReader.getCellData(xlSheetName, 211, 6);
		String actLayoutMessage = checkValidationMessage(expLayoutMessage);
		excelReader.setCellData(xlfile, xlSheetName, 211, 7, actLayoutMessage);
		
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_cusHeaderAndFooterExitBtn));
		sl_cusHeaderAndFooterExitBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_customizeSaveBtn));
		osr_customizeSaveBtn.click();

		String expCustMessage = excelReader.getCellData(xlSheetName, 212, 6);
		String actCustMessage = checkValidationMessage(expCustMessage);
		excelReader.setCellData(xlfile, xlSheetName, 212, 7, actCustMessage);
		
		Thread.sleep(2000);

		if (actLayoutMessage.equalsIgnoreCase(expLayoutMessage) && actCustMessage.equalsIgnoreCase(expCustMessage))  
		{
			excelReader.setCellData(xlfile, xlSheetName, 211, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 211, 8, resFail);
			return false;
		}
	}







	public boolean checkPrintPDFFileAfterImportingHeaderAndFooterLayoutInStockLedgerReport() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException, AWTException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_ExportBtn));
		sl_ExportBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_ExportPDFBtn));
		sl_ExportPDFBtn.click();

		Thread.sleep(5000);

		String Filename = checkDownloadedFileName(getDriver());

		System.err.println(Filename);

		String actPDF = getBaseDir()+"\\autoIt\\ExportFiles\\"+Filename;
		String expPDF = getBaseDir()+"\\autoIt\\ImportFiles\\StockLedgerReportWithImportedHeaderAndFooterLayout.pdf";

		PDFUtil pdfutil = new PDFUtil();

		String data = pdfutil.getText(expPDF);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		Calendar cal = Calendar.getInstance();
		String date = df.format(cal.getTime());
		System.err.println(date);

		String oldDate = "12/10/2021";

		String actData = pdfutil.getText(actPDF);
		String expData = data.replace(oldDate, date);

		System.err.println(actData);
		System.err.println(expData);
		excelReader.setCellData(xlfile, xlSheetName, 214, 7, actData);

		if (actData.equalsIgnoreCase(expData)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 214, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 214, 8, resFail);
			return false;
		}
	}


public boolean checkSavedRule() throws InterruptedException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleNameTxt));
	ruleNameTxt.click();

	rulesTabRuleNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);

	rulesTabRuleNameTxt.sendKeys(Keys.BACK_SPACE);		

	Thread.sleep(1000);
	
	rulesTabRuleNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 20, 5));
	
	Thread.sleep(1999);

	ruleNameTxt.sendKeys(Keys.TAB);

	Thread.sleep(1999);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesNewRecordChekbox));
	String actrulesNewRecordChekbox=Boolean.toString(rulesNewRecordChekboxSelected.isSelected());
	String exprulesNewRecordChekbox="true";
	System.out.println(" rulesNewRecordChekbox: "+actrulesNewRecordChekbox +" Value "+exprulesNewRecordChekbox);


	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesLoadChekbox));
	String actrulesLoadChekbox=Boolean.toString(rulesLoadChekboxSelected.isSelected());
	String exprulesLoadChekbox="true";
	
	System.out.println(" rulesLoadChekbox: "+actrulesLoadChekbox +" Value "+exprulesLoadChekbox);

	Thread.sleep(1999);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesActiveChekbox));
	String actrulesActiveChekbox=Boolean.toString(rulesActiveChekboxSelected.isSelected());
	String exprulesActiveChekbox="true";
	System.out.println(" rulesActiveChekbox: "+actrulesActiveChekbox +" Value "+exprulesActiveChekbox);

	Thread.sleep(1999);
	
	rulesTabIFSelectTabTxt.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectTabDrpdwn));
	Select account=new Select(rulesTabIFSelectTabDrpdwn);

	String  actrulesTabIFSelectTabDrpdwn=account.getFirstSelectedOption().getText();
	String exprulesTabIFSelectTabDrpdwn="Department";

	System.out.println("rulesTabIFSelectTabDrpdwn : "+actrulesTabIFSelectTabDrpdwn +" Value : "+exprulesTabIFSelectTabDrpdwn);

	Thread.sleep(1999);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleConditionRow1Of2ndColbox));
	Select con1=new Select(ruleConditionRow1Of2ndColbox);

	String  actruleConditionRow1Of2ndColbox=con1.getFirstSelectedOption().getText();
	String expruleConditionRow1Of2ndColbox="sName";
	
	System.out.println("ruleConditionRow1Of2ndColbox : "+actruleConditionRow1Of2ndColbox +" Value : "+expruleConditionRow1Of2ndColbox);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleConditionRow1Of3rdColbox));
	Select con2=new Select(ruleConditionRow1Of3rdColbox);

	String  actruleConditionRow1Of3rdColbox=con2.getFirstSelectedOption().getText();
	String expruleConditionRow1Of3rdColbox="Equal To";

	System.out.println("ruleConditionRow1Of3rdColbox : "+actruleConditionRow1Of3rdColbox +" Value : "+expruleConditionRow1Of3rdColbox);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleConditionRow1Of4thColbox));
	Select con3=new Select(ruleConditionRow1Of4thColbox);

	String  actruleConditionRow1Of4thColbox=con3.getFirstSelectedOption().getText();
	String expruleConditionRow1Of4thColbox="Value";

	System.out.println("ruleConditionRow1Of4thColbox : "+actruleConditionRow1Of4thColbox +" Value : "+expruleConditionRow1Of4thColbox);

	Thread.sleep(1999);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ruleConditionRow1Of5thColbox));

	String  actruleConditionRow1Of5thColbox=ruleConditionRow1Of5thColbox.getAttribute("value");
	String expruleConditionRow1Of5thColbox="DUBAI";
	System.out.println("ruleConditionRow1Of5thColbox : "+actruleConditionRow1Of5thColbox +" Value : "+expruleConditionRow1Of5thColbox);

	Thread.sleep(1999);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid1stRow_1stcol));
	getAction().moveToElement(rulesGrid1stRow_1stcol).build().perform();
	String actrulesGrid1stRow_1stcol=rulesGrid1stRow_1stcol.getText();
	String exprulesGrid1stRow_1stcol="TaxCode";

	System.out.println(" rulesGrid1stRow_1stcol : "+actrulesGrid1stRow_1stcol +" Value : "+exprulesGrid1stRow_1stcol);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid1stRow_2ndcol));
	String actrulesGrid1stRow_2ndcol=rulesGrid1stRow_2ndcol.getText();
	String exprulesGrid1stRow_2ndcol="Value";

	System.out.println(" rulesGrid1stRow_2ndcol : "+actrulesGrid1stRow_2ndcol +" Value : "+exprulesGrid1stRow_2ndcol);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid1stRow_5thcol));
	String actrulesGrid1stRow_4thcol=rulesGrid1stRow_5thcol.getText();
	String exprulesGrid1stRow_4thcol="DubaiTaxcode";


	System.out.println(" rulesGrid1stRow_4thcol : "+actrulesGrid1stRow_4thcol +" Value : "+exprulesGrid1stRow_4thcol);

	Thread.sleep(1999);		
	
	getDriver().navigate().refresh();

	if (actrulesNewRecordChekbox.equalsIgnoreCase(exprulesNewRecordChekbox) && actrulesLoadChekbox.equalsIgnoreCase(exprulesLoadChekbox) &&
			actrulesActiveChekbox.equalsIgnoreCase(exprulesActiveChekbox) && actrulesTabIFSelectTabDrpdwn.equalsIgnoreCase(exprulesTabIFSelectTabDrpdwn) &&
			actrulesGrid1stRow_1stcol.equalsIgnoreCase(exprulesGrid1stRow_1stcol) && 
			actrulesGrid1stRow_4thcol.equalsIgnoreCase(exprulesGrid1stRow_4thcol) && 

			actruleConditionRow1Of2ndColbox.equalsIgnoreCase(expruleConditionRow1Of2ndColbox) &&
			actruleConditionRow1Of3rdColbox.equalsIgnoreCase(expruleConditionRow1Of3rdColbox) &&
			actruleConditionRow1Of4thColbox.equalsIgnoreCase(expruleConditionRow1Of4thColbox) &&
			actruleConditionRow1Of5thColbox.equalsIgnoreCase(expruleConditionRow1Of5thColbox) ) 
	{
		
		return true;
	} 
	else 
	{
		
		return false;
	}
}









	public RulesPage(WebDriver driver)
	{

		PageFactory.initElements(driver, this);

	}





}
