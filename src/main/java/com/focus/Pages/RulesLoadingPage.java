package com.focus.Pages;

import java.awt.AWTException;

import java.io.IOException;
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

public class RulesLoadingPage extends BaseEngine {
	
	@FindBy(xpath="//*[@id='id_menu_search_input']")
	public static WebElement  SearchText;
	
	@FindBy(xpath="//*[@id='60']")
	public static WebElement  FinancialMenu;
	
	@FindBy(xpath="//*[@id='61']")
	public static WebElement  FinancialTransactionsMenu;
	
	@FindBy(xpath="//*[@id='2007']")
	public static WebElement  TransactionPurchaseMenu;
	
	@FindBy(xpath="//*[@id='2008']")
	public static WebElement  PurchaseVoucher;
	
	@FindBy(xpath="//*[@id='id_transaction_homescreen_new']//i")
	public static WebElement  PurchaseVoucher_NewBut;
	
	@FindBy(xpath="//*[@id='id_transaction_homescreen_Delete']//i")
	public static WebElement  PurchaseVoucher_DeleteBut;
	
	@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[15]/a/i")
	public static WebElement  PuchaseVoucher_Settingsbut;
	
	@FindBy(xpath="//*[@id='navigationtab11']")
	public static WebElement  Rules_Tab;
	
	@FindBy(xpath="//*[@id='lnkRule_div_DocCustRules']//input")
	public static WebElement  Addrule;
	
	@FindBy(xpath="//*[@id='txtMasterRuleName_DocCustRules']")
	public static WebElement  Rulenametext;
	
	@FindBy(xpath="//*[@id='chkListApplyOn_DocCustRules']/div/div[2]/div/label/span")
	public static WebElement  ApplyNewRecord;
	
	@FindBy(xpath="//*[@id='chkListApplyOn_DocCustRules']/div/div[3]/div/label/span")
	public static WebElement  ApplyEdit;
	
	@FindBy(xpath="//*[@id='chkListEvaluateOn_DocCustRules']/div/div[1]/div/label/span")
	public static WebElement  ApplyLoad;
	@FindBy(xpath="//*[@id='chkListEvaluateOn_DocCustRules']/div/div[2]/div/label/span")
	public static WebElement  ApplyLeave;
	
	@FindBy(xpath="//*[@id='chkListEvaluateOn_DocCustRules']/div/div[3]/div/label/span")
	public static WebElement  ApplyONEnter;
	
	@FindBy(xpath="//*[@id='chkListStatus_DocCustRules']/div/div/div/label/span")
	public static WebElement  ApplyActiveStatus;
	
	@FindBy(xpath="//*[@id='chkNoRuleCondition_DocCustRules']//..//span")
	public static WebElement  NOCondition;
	
	@FindBy(xpath = "//*[@id='ddlMasterFields_RULES_DocCustRules']")
	public static WebElement rulesTabIFSelectFieldText;

	@FindBy(xpath = "//*[@id='txtField_RC_1_1_DocCustRules']")
	public static WebElement rulesTabIFSelectField;
	
	@FindBy(xpath = "//*[@id='AddedCondition_1_Rule_ELSEIF_2_DocCustRules']/div[1]/div[2]")
	public static WebElement selectfield;

	@FindBy(xpath = "//select[@id='ddlOperators_RC_1_1_DocCustRules']")
	public static WebElement rulesTabOperatorsDrpdwn;
	
	@FindBy(xpath = "//select[@id='ddlOperators_RC_1_2_DocCustRules']")
	public static WebElement ElseifrulesTabOperatorsDrpdwn;
	
	@FindBy(xpath = "//select[@id='ddlCompareWith_RC_1_1_DocCustRules']")
	public static WebElement rulesTabComparewithDrpdwn;
	
	@FindBy(xpath = "//select[@id='ddlCompareWith_RC_1_2_DocCustRules']")
	public static WebElement ElseifrulesTabComparewithDrpdwn;
	
	@FindBy(xpath = "//*[@id='optControl_RulesCondition_DocCustRules_input_container']/div[2]/table/tbody/tr/td[1]/i")
	public static WebElement rulesTabCompareTypeDrpdwn;
	
	@FindBy(xpath = "//*[@id='optControl_RulesCondition_DocCustRules']")
	public static WebElement rulesTabConditionTxt;
	
	@FindBy(xpath = "//*[@id='txtCompareType_RC_1_1_DocCustRules']")
	public static WebElement rulesValueConditionTxt;
	
	@FindBy(xpath = "//*[@id='txtCompareType_RC_1_1_DocCustRules']//..//select")
	public static WebElement Valueconditiondrpdown;
	
	@FindBy(xpath = "//*[@id='optControl_RulesCondition_DocCustRules']")
	public static WebElement ElseifrulesTabConditionTxt;
	
	@FindBy(xpath = "//*[@id='txtCompareType_RC_1_2_DocCustRules']//..//select")
	public static WebElement ElseifselectrulesvalueConditionTxt;
	
	@FindBy(xpath = "//*[@id='AddedCondition_1_Rule_IF_DocCustRules']/div[1]/div[6]")
	public static WebElement selectrulesTabConditionTxt;
	
	@FindBy(xpath = "//*[@id='AddedCondition_1_Rule_IF_DocCustRules']/div[1]/div[5]")
	public static WebElement selectrulesValueConditionTxt;
	
	@FindBy(xpath = "//*[@id='AddedCondition_1_Rule_ELSEIF_2_DocCustRules']/div[1]/div[6]")
	public static WebElement ElseifselectrulesTabConditionTxt;
	
	@FindBy(xpath = "//*[@id='AddedCondition_1_Rule_ELSEIF_2_DocCustRules']/div[1]/div[5]")
	public static WebElement ElseifselectValueConditionTxt;
	
	@FindBy(xpath = "//*[@id='AddedCondition_1_Rule_IF_']/div[1]/div[5]")
	public static WebElement ElseifselectrulesValueConditionTxt;
	
	@FindBy(xpath = "//*[@id='tblFormating_IF_Rule_DocCustRules_col_1-1']")
	public static WebElement IfFieldName;
	
	@FindBy(xpath = "//*[@id='tblFormating_IF_Rule_DocCustRules_col_1-12']")
	public static WebElement Disable_Column;
	
	@FindBy(xpath = "//*[@id='chkDisableFormat_IF_DocCustRules']")
	public static WebElement Chk_Disable;
	
	@FindBy(xpath = "//*[@id='tblFormating_ELSEIF_Rule_DocCustRules_2_col_1-1']")
	public static WebElement ElseIfFieldName;
	
	@FindBy(xpath = "//*[@id='ddlFieldsDiv_Rule_IF_DocCustRules']")
	public static WebElement IfFieldNameText;
	
	@FindBy(xpath = "//*[@id='ddlFieldsDiv_Rule_ELSEIF_DocCustRules_2']")
	public static WebElement ElseIfFieldNameText;
	
	@FindBy(xpath = "//*[@id='tblFormating_IF_Rule_DocCustRules_col_1-2']")
	public static WebElement ChangeValue_Column;
	
	@FindBy(xpath = "//*[@id='ddlChangeValueFormat_IF_DocCustRules']")
	public static WebElement IfChangeValueTxt;
	
	@FindBy(xpath = "//*[@id='tblFormating_IF_Rule_DocCustRules_col_1-16']")
	public static WebElement SelectIfFontField;
	
	@FindBy(xpath = "//*[@id='tblFormating_Else_Rule_DocCustRules_col_1-16']")
	public static WebElement SelectElseFontField;
	
	@FindBy(xpath = "//*[@id='btnFontFormat_IF_DocCustRules']")
	public static WebElement IfFontTxt;
	
	@FindBy(xpath = "//*[@id='btnFontFormat_Else_DocCustRules']")
	public static WebElement ElseFontTxt;
	
	@FindBy(xpath = "//*[@id='fontBackColor_ctrlFontControl_IF']")
	public static WebElement IfBackColourFiled;
	
	@FindBy(xpath = "//*[@id='fontBackColor_ctrlFontControl_ELSE']")
	public static WebElement ElseBackColourFiled;
	
	@FindBy(xpath = "//*[@id='myFontControl_IF_Save']")
	public static WebElement If_FontSave;
	
	@FindBy(xpath = "//*[@id='myFontControl_ELSE_Save']")
	public static WebElement Else_FontSave;
	
	@FindBy(xpath = "//*[@id='ddlChangeValueFormat_ELSEIF_DocCustRules_2']")
	public static WebElement ElseIfChangeValueTxt;
	
	@FindBy(xpath = "//*[@id='ValuesDivFormat_IF_DocCustRules']")
	public static WebElement IfValuecheck;
	
	@FindBy(xpath = "//select[@id='ddlStrListNew']")
	public static WebElement IfValuecheck2;
	
	@FindBy(xpath = "//*[@id='txtTextField']")
	public static WebElement IfValue2;
	
	@FindBy(xpath = "//*[@id='ValuesDivFormat_ELSEIF_DocCustRules_2']")
	public static WebElement ElseIfValuecheck;
	
	@FindBy(xpath = "//*[@id='ValuesDivFormat_IF_DocCustRules']")
	public static WebElement IfValueTxt;
	
	@FindBy(xpath = "//*[@id='tab_Rule_Else_DocCustRules']//a")
	public static WebElement Elseoption;
	
	@FindBy(xpath = "//*[@id='tab_Rule_IF_DocCustRules']//a")
	public static WebElement Ifoption;
	
	@FindBy(xpath = "//*[@id='tab_Rule_ELSEIF_2_DocCustRules']//a")
	public static WebElement ElseIfoption;
	
	@FindBy(xpath = "//*[@id='btnAddNewCondition']")
	public static WebElement Addnewcondition;
	
	@FindBy(xpath = "//*[@id='tblFormating_Else_Rule_DocCustRules_col_1-1']")
	public static WebElement ElseFieldName;
	
	@FindBy(xpath = "//*[@id='ddlFieldsDiv_Rule_ELSE_DocCustRules']")
	public static WebElement ElseFieldNameText;
	
	@FindBy(xpath = "//*[@id='ddlChangeValueFormat_ELSE_DocCustRules']")
	public static WebElement ElseChangeValueTxt;
	
	@FindBy(xpath = "//*[@id='tblFormating_Else_Rule_DocCustRules_col_1-4']")
	public static WebElement ElseValuecheck;
	
	@FindBy(xpath = "//*[@id='ValuesDivFormat_ELSE_DocCustRules']")
	public static WebElement ElseValueTxt;
	
	@FindBy(xpath = "//*[@id='txtTextField']")
	public static WebElement ElseTxt;
	
	@FindBy(xpath = "//*[@id='btnAddRule']")
	public static WebElement saverule;
	
	@FindBy(xpath = "//*[@id='dvRuleButtons']/button[2]")
	public static WebElement deleterule;
	
	@FindBy(xpath = "//*[@id='updateButton']")
	public static WebElement SettingsUpdatebut;
	
	@FindBy(xpath = "//*[@id='btnCustomizeClose']")
	public static WebElement SettingsClosebut;
	
	@FindBy(xpath = "//*[@id='id_transaction_homescreen_new']/i")
	public static WebElement NewBut;
	
	@FindBy(xpath = "//input[@id='id_header_3']")
	public static WebElement PurAccTxt;
	
	@FindBy(xpath = "//input[@id='id_header_4']")
	public static WebElement VendorAccTxt;
	
	@FindBy(xpath = "//*[@id='id_header_268435459']")
	public static WebElement DepartmentTxt;
	
	@FindBy(xpath = "//*[@id='id_header_268435460']")
	public static WebElement WarehouseTxt;
	
	@FindBy(xpath = "//*[@id='id_footer_134218922']")
	public static WebElement Footerfield;
	
	//@FindBy(xpath = "//*[@id='id_header_67108927']")
	@FindBy(xpath = "//*[@id='id_transactionentry_header1_section']//div[13]//select")
	public static WebElement NumText_Column;
	
	@FindBy(xpath = "//*[@id='id_transactionentry_header1_section']//div[14]//select")
	public static WebElement stringTxt_Column;
	
	@FindBy(xpath = "//*[@id='id_transactionentry_header1_section']//div[13]//select//option")
	public static WebElement H_NumText;
	
	@FindBy(xpath = "//*[@id='id_transactionentry_header1_section']//div[14]//select//option")
	public static WebElement HstringTxt;
	
	@FindBy(xpath = "//*[@id='id_body_33558787']")
	public static WebElement Observe1_txt;
	
	@FindBy(xpath = "//*[@id='id_body_33558788']")
	public static WebElement Observe2_txt;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[2]")
	public static WebElement Select1stColumn;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[2]")
	public static WebElement Select2ndrow_1stColumn;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[5]")
	public static WebElement Select2ndrow_4thColumn;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[4]")
	public static WebElement Select3rdColumn;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[14]")
	public static WebElement Select13thColumn;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[15]")
	public static WebElement Select14thColumn;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[16]")
	public static WebElement Select15thColumn;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[17]")
	public static WebElement Select16thColumn;
	
	
	@FindBy(xpath = "//*[@id='id_body_16782595']")
	public static WebElement B_stringTxt;
	
	@FindBy(xpath = "//*[@id='id_body_16782596']")
	public static WebElement B_numTxt;
	
	@FindBy(xpath = "//*[@id='id_body_23']")
	public static WebElement ItemTxt;
	
	@FindBy(xpath = "//*[@id='id_body_24']")
	public static WebElement UnitTxt;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[5]")
	public static WebElement Select4thColumn;
	
	@FindBy(xpath = "//*[@id='id_body_26']")
	public static WebElement QuantityTxt;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[6]")
	public static WebElement Select5thColumn;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[7]")
	public static WebElement Select6thColumn;
	
	@FindBy(xpath = "//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[8]")
	public static WebElement Select7thColumn;
	
	@FindBy(xpath = "//*[@id='id_body_27']")
	public static WebElement RateTxt;
	
	@FindBy(xpath = "//*[@id='id_body_28']")
	public static WebElement GrossTxt;
	
	@FindBy(xpath = "//*[@id='id_body_33554435']")
	public static WebElement DiscountTxt;
	
	@FindBy(xpath = "//*[@id='id_transactionentry_save']//span")
	public static WebElement VoucherSaveBut;
	
	@FindBy(xpath = "//*[@id='id_transactionentry_previous']")
	public static WebElement Previous_But;
	
	@FindBy(xpath = "//*[@id='id_transactionentry_close']//span")
	public static WebElement VoucherCloseBut;
	
	@FindBy(xpath = "//*[@id='id_Pick']//a")
	public static WebElement Pickbut;
	
	@FindBy(xpath = "//*[@id='id_Ok']/a")
	public static WebElement OkBut;
	
	@FindBy(xpath = "//*[@id='id_fconfirm_btn2']")
	public static WebElement PopUp_Yesbut;
	
	
	public static boolean CheckLoginRulesLoadingPage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	        Thread.sleep(3000);

	        getDriver().navigate().refresh();
	        Thread.sleep(3000);

	        LoginPage lp=new LoginPage(getDriver()); 

	        lp.checkLoginPageTitleByURLInputInBrowser();

	        String unamelt="su";

	        String pawslt="su";

	        lp.enterUserName(unamelt);
	        
	        Thread.sleep(2000);

	        lp.enterPassword(pawslt);

	        Thread.sleep(2000);

	        String compname = "automation";

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

	        //getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	       // companyLogo.click();

	        String getCompanyTxt1=Company_Name.getText();
	        String getLoginCompanyName1=getCompanyTxt1.substring(0, 19);
	        System.out.println("company name  :  "+ getLoginCompanyName1);
	        //companyLogo.click();

	        String expUserInfo1           ="SU";
	        String expLoginCompanyName1   ="automation";

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


	public static boolean checkRestoreRuleLoading() throws InterruptedException, IOException, AWTException
	{
		
		getDriver().navigate().refresh();
		Thread.sleep(6000);
		String actMessage=BaseEngine.restoreCompany("automation","automation");
		String expMessage="Restore company code : 0M0";
		
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
	
	public static boolean CreatingRuleWithDisableConditionInPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(2000);
		
		ClickUsingJs(FinancialMenu);
		
		Thread.sleep(1000);
		
		ClickUsingJs(FinancialTransactionsMenu);
		
		Thread.sleep(1000);
		
		ClickUsingJs(TransactionPurchaseMenu);
		
		Thread.sleep(1000);
		
		ClickUsingJs(PurchaseVoucher);
		
		Thread.sleep(1000);
		
		ClickUsingJs(PuchaseVoucher_Settingsbut);
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rules_Tab));
		Rules_Tab.click();
		
		Thread.sleep(1500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rulenametext));
		Rulenametext.click();
		Thread.sleep(1000);
		Rulenametext.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
		Thread.sleep(1000);
		Rulenametext.sendKeys("NewRule");
		Thread.sleep(1000);
		Rulenametext.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyNewRecord));
		ApplyNewRecord.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyEdit));
		ApplyEdit.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLoad));
		ApplyLoad.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLeave));
		ApplyLeave.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyONEnter));
		ApplyONEnter.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyActiveStatus));
		ApplyActiveStatus.click();
		
		rulesTabIFSelectField.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectFieldText));
		Select d1=new Select(rulesTabIFSelectFieldText);
		d1.selectByVisibleText("Observe1");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabOperatorsDrpdwn));
		rulesTabOperatorsDrpdwn.click();
		Thread.sleep(1000);
		Select O1=new Select(rulesTabOperatorsDrpdwn);
		O1.selectByVisibleText("Equal To");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabComparewithDrpdwn));
		rulesTabComparewithDrpdwn.click();
		Thread.sleep(1000);
		Select V1=new Select(rulesTabComparewithDrpdwn);
		V1.selectByVisibleText("Value");
		Thread.sleep(1000);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectrulesValueConditionTxt));
		selectrulesValueConditionTxt.click();
		Thread.sleep(1500);
		rulesValueConditionTxt.sendKeys("2");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfFieldName));
		IfFieldName.click();
		Select W1=new Select(IfFieldNameText);
		W1.selectByVisibleText("Observe2");
		IfFieldNameText.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfChangeValueTxt));
		IfChangeValueTxt.click();
		Thread.sleep(1000);
		IfChangeValueTxt.sendKeys("None");
		Thread.sleep(1000);
		IfChangeValueTxt.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		IfChangeValueTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Disable_Column));
		Disable_Column.click();
		
		if(Chk_Disable.isSelected()==false)
		{
			Chk_Disable.click();
		}
		boolean ActDisableOption=Chk_Disable.isSelected();
		boolean ExpDisableoption=true;
		
		System.out.println("ActDisableOption is " + ActDisableOption);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Elseoption));
		Elseoption.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseFieldName));
		ElseFieldName.click();
		Select W3=new Select(ElseFieldNameText);
		W3.selectByVisibleText("Observe2");
		ElseFieldNameText.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseChangeValueTxt));
		ElseChangeValueTxt.click();
		Thread.sleep(1000);
		ElseChangeValueTxt.sendKeys("none");
		Thread.sleep(1000);
		ElseChangeValueTxt.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		ElseChangeValueTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(1500);
		
		ClickUsingJs(saverule);
		
		String expMessage = "Rule Saved Successfully";

		String actMessage = checkValidationMessage(expMessage);
		
		Thread.sleep(2000);
		
		ClickUsingJs(SettingsUpdatebut);
		
		Thread.sleep(1500);
		
		String expMessage1 = "Data saved Successfully";

		String actMessage1 = checkValidationMessage(expMessage1);

		
		ClickUsingJs(SettingsClosebut);
		
		if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1))
		{
			return true;
		}
		
		else
		{
			return false;
		}
		

		
	
	}
	
	
	
	public static boolean CheckPurchaseVoucherwithcreatedRuleConditionnumberlistandstringlistdatatypeinBody() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
		
		Thread.sleep(4000);
		
		ClickUsingJs(FinancialMenu);
		
		Thread.sleep(1000);
		
		ClickUsingJs(FinancialTransactionsMenu);
		
		Thread.sleep(1000);
		
		ClickUsingJs(TransactionPurchaseMenu);
		
		Thread.sleep(1000);
		
		ClickUsingJs(PurchaseVoucher);
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		Thread.sleep(4000);
		
		ClickUsingJs(Select1stColumn);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemTxt));
		ItemTxt.sendKeys("Item 1");
		Thread.sleep(1000);
		ItemTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Select15thColumn));
		Select15thColumn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(B_stringTxt));
		Select s1=new Select(B_stringTxt);
		s1.selectByVisibleText("sec");
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(B_stringTxt));
		B_stringTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(B_numTxt));
		B_numTxt.sendKeys(Keys.TAB);
		
		String acttxt1=Select16thColumn.getText();
		String exptxt1="twitter";
		
		System.out.println("Act numtxt1 = " + acttxt1);
		System.out.println("Exp numtxt1 = " + exptxt1);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Select15thColumn));
		Select15thColumn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(B_stringTxt));
		Select s2=new Select(B_stringTxt);
		s2.selectByVisibleText("vzy");
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(B_stringTxt));
		B_stringTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(B_numTxt));
		B_numTxt.sendKeys(Keys.TAB);
		
		String acttxt2=Select16thColumn.getText();
		String exptxt2="yahoo";
		
		System.out.println("Act numtxt2 = " + acttxt2);
		System.out.println("Exp numtxt2 = " + exptxt2);
		
		Thread.sleep(1000);
		
		ClickUsingJs(VoucherCloseBut);
		
		Thread.sleep(2000);
		
		ClickUsingJs(PopUp_Yesbut);
		Thread.sleep(2000);
		
		

		
		if(acttxt1.equalsIgnoreCase(exptxt1) && acttxt2.equalsIgnoreCase(exptxt2))
				
		{
			return true;
		}
		else
		{
			return false;
		}

	}
	
	public static boolean CheckCreatingRuleWithConditionHeadtoHeadInPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(2000);
		
		ClickUsingJs(FinancialMenu);
		
		Thread.sleep(1000);
		
		ClickUsingJs(FinancialTransactionsMenu);
		
		Thread.sleep(1000);
		
		ClickUsingJs(TransactionPurchaseMenu);
		
		Thread.sleep(1000);
		
		ClickUsingJs(PurchaseVoucher);
		
		Thread.sleep(1000);
		
		ClickUsingJs(PuchaseVoucher_Settingsbut);
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rules_Tab));
		Rules_Tab.click();
		
		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Addrule));
		//Addrule.click();
		
		Thread.sleep(1500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rulenametext));
		Rulenametext.click();
		Thread.sleep(1000);
		Rulenametext.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
		Thread.sleep(1000);
		Rulenametext.sendKeys("TestRule1");
		Thread.sleep(1000);
		Rulenametext.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyNewRecord));
		ApplyNewRecord.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyEdit));
		ApplyEdit.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLoad));
		ApplyLoad.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLeave));
		ApplyLeave.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyONEnter));
		ApplyONEnter.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyActiveStatus));
		ApplyActiveStatus.click();
		
		rulesTabIFSelectField.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectFieldText));
		Select d1=new Select(rulesTabIFSelectFieldText);
		d1.selectByVisibleText("Department");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabOperatorsDrpdwn));
		rulesTabOperatorsDrpdwn.click();
		Thread.sleep(1000);
		Select O1=new Select(rulesTabOperatorsDrpdwn);
		O1.selectByVisibleText("Equal To");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabComparewithDrpdwn));
		rulesTabComparewithDrpdwn.click();
		Thread.sleep(1000);
		Select V1=new Select(rulesTabComparewithDrpdwn);
		V1.selectByVisibleText("Value");
		Thread.sleep(1000);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectrulesTabConditionTxt));
		selectrulesTabConditionTxt.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabConditionTxt));
		rulesTabConditionTxt.click();
		Thread.sleep(1000);
		rulesTabConditionTxt.sendKeys("Dep 1");
		Thread.sleep(1000);
		rulesTabConditionTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfFieldName));
		IfFieldName.click();
		Select W1=new Select(IfFieldNameText);
		W1.selectByVisibleText("Warehouse");
		IfFieldNameText.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfChangeValueTxt));
		IfChangeValueTxt.click();
		Thread.sleep(1000);
		IfChangeValueTxt.sendKeys("Value");
		Thread.sleep(1000);
		IfChangeValueTxt.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		IfChangeValueTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfValuecheck));
		IfValuecheck.click();
		Thread.sleep(1000);
		IfValueTxt.sendKeys("Wh 1");
		Thread.sleep(1000);
		IfValueTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(1500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Ifoption));
		ClickUsingJs(Ifoption);
		Thread.sleep(1500);
		contextClick(Ifoption);
		Thread.sleep(1000);
		Addnewcondition.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseIfoption));
		ClickUsingJs(ElseIfoption);
		
		selectfield.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectFieldText));
		Select D2=new Select(rulesTabIFSelectFieldText);
		D2.selectByVisibleText("Department");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseifrulesTabOperatorsDrpdwn));
		ElseifrulesTabOperatorsDrpdwn.click();
		Thread.sleep(1000);
		Select O2=new Select(ElseifrulesTabOperatorsDrpdwn);
		O2.selectByVisibleText("Equal To");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseifrulesTabComparewithDrpdwn));
		ElseifrulesTabComparewithDrpdwn.click();
		Thread.sleep(1000);
		Select V2=new Select(ElseifrulesTabComparewithDrpdwn);
		V2.selectByVisibleText("Value");
		Thread.sleep(1000);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseifselectrulesTabConditionTxt));
		ElseifselectrulesTabConditionTxt.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseifrulesTabConditionTxt));
		ElseifrulesTabConditionTxt.click();
		Thread.sleep(1000);
		ElseifrulesTabConditionTxt.sendKeys("Dep 2");
		Thread.sleep(1000);
		ElseifrulesTabConditionTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseIfFieldName));
		ElseIfFieldName.click();
		Select W2=new Select(ElseIfFieldNameText);
		W2.selectByVisibleText("Warehouse");
		ElseIfFieldNameText.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseIfChangeValueTxt));
		ElseIfChangeValueTxt.click();
		Thread.sleep(1000);
		ElseIfChangeValueTxt.sendKeys("Value");
		Thread.sleep(1000);
		ElseIfChangeValueTxt.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		ElseIfChangeValueTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseIfValuecheck));
		ElseIfValuecheck.click();
		Thread.sleep(1000);
		ElseIfValuecheck.sendKeys("Wh 2");
		Thread.sleep(1000);
		ElseIfValuecheck.sendKeys(Keys.TAB);
		
		Thread.sleep(1500);
		
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Elseoption));
		Elseoption.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseFieldName));
		ElseFieldName.click();
		Select W3=new Select(ElseFieldNameText);
		W3.selectByVisibleText("Warehouse");
		ElseFieldNameText.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseChangeValueTxt));
		ElseChangeValueTxt.click();
		Thread.sleep(1000);
		ElseChangeValueTxt.sendKeys("Value");
		Thread.sleep(1000);
		ElseChangeValueTxt.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		ElseChangeValueTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseValuecheck));
		ElseValuecheck.click();
		Thread.sleep(1000);
		ElseValueTxt.sendKeys("wh3");
		
		Thread.sleep(1500);
		
		ClickUsingJs(saverule);
		
		String expMessage = "Rule Saved Successfully";

		String actMessage = checkValidationMessage(expMessage);
		
		Thread.sleep(2000);
		
		ClickUsingJs(SettingsUpdatebut);
		
		Thread.sleep(1500);
		
		String expMessage1 = "Data saved Successfully";

		String actMessage1 = checkValidationMessage(expMessage1);

		
		ClickUsingJs(SettingsClosebut);
		
		if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1))
		{
			return true;
		}
		
		else
		{
			return false;
		}
		

		
	}
	
	public static boolean SavingPurchaseVoucherwithcreatedRuleConditionHeadtoHead() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
		Thread.sleep(5000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		financialsTransactionsPurchaseMenu.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PurchaseVoucher));
		PurchaseVoucher.click();

		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		Thread.sleep(4000);
		
		String actDep=WarehouseTxt.getAttribute("value");
		String expDep="wh3";
		
		System.out.println("act warehouse Txt:" + actDep );
		System.out.println("exp warehouse Txt:" + expDep );
		
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PurAccTxt));
		PurAccTxt.click();
		Thread.sleep(1000);
		PurAccTxt.sendKeys("Purchase");
		Thread.sleep(1000);
		PurAccTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(VendorAccTxt));
		VendorAccTxt.sendKeys("Vendor A");
		Thread.sleep(1000);
		VendorAccTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepartmentTxt));
		DepartmentTxt.click();
		DepartmentTxt.sendKeys("Dep 1");
		Thread.sleep(1000);
		DepartmentTxt.sendKeys(Keys.TAB);
		
		String actDep1=WarehouseTxt.getAttribute("value");
		String expDep1="Wh 1";
		
		System.out.println("act warehouse Txt1:" + actDep1 );
		System.out.println("exp warehouse Txt1:" + expDep1 );
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepartmentTxt));
		DepartmentTxt.click();
		Thread.sleep(1000);
		DepartmentTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
		Thread.sleep(1000);
		DepartmentTxt.sendKeys("Dep 2");
		Thread.sleep(1000);
		DepartmentTxt.sendKeys(Keys.TAB);
		
		
		String actDep2=WarehouseTxt.getAttribute("value");
		String expDep2="Wh 2";
		
		System.out.println("act warehouse Txt2:" + actDep2 );
		System.out.println("exp warehouse Txt2:" + expDep2 );
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepartmentTxt));
		DepartmentTxt.click();
		Thread.sleep(1000);
		DepartmentTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
		Thread.sleep(1000);
		DepartmentTxt.sendKeys("dep3");
		Thread.sleep(1000);
		DepartmentTxt.sendKeys(Keys.TAB);
		
		
		String actDep3=WarehouseTxt.getAttribute("value");
		String expDep3="wh3";
		
		System.out.println("act warehouse Txt3:" + actDep3 );
		System.out.println("exp warehouse Txt3:" + expDep3 );
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Select1stColumn));
		ClickUsingJs(Select1stColumn);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemTxt));
		ItemTxt.sendKeys("s1");
		Thread.sleep(1000);
		ItemTxt.sendKeys(Keys.TAB);
		
		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(UnitTxt));
		//UnitTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Select5thColumn));
		Select5thColumn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(QuantityTxt));
		QuantityTxt.sendKeys("1");
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Select7thColumn));
		Select7thColumn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RateTxt));
		RateTxt.sendKeys("10");
		Thread.sleep(1000);
		RateTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GrossTxt));
		GrossTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DiscountTxt));
		DiscountTxt.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(VoucherSaveBut));
		ClickUsingJs(VoucherSaveBut);
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Pickbut));
		Pickbut.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(OkBut));
		OkBut.click();
		
		Thread.sleep(1000);
		
		String expMessage = "Voucher saved successfully";
		String actMessage = checkValidationMessage(expMessage);
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
		
		ClickUsingJs(VoucherCloseBut);
		Thread.sleep(2000);
		ClickUsingJs(PopUp_Yesbut);
		Thread.sleep(2000);
		
		if(actMessage.startsWith(expMessage) && actDep.equalsIgnoreCase(expDep)
				&& actDep1.equalsIgnoreCase(expDep1)
				&&  actDep2.equalsIgnoreCase(expDep2) && actDep3.equalsIgnoreCase(expDep3))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
	
public static boolean CreatingRuleWithConditionStringandNumberListextrafieldsInPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
	Thread.sleep(2000);
	
	ClickUsingJs(FinancialMenu);
	
	Thread.sleep(1000);
	
	ClickUsingJs(FinancialTransactionsMenu);
	
	Thread.sleep(1000);
	
	ClickUsingJs(TransactionPurchaseMenu);
	
	Thread.sleep(1000);
	
	ClickUsingJs(PurchaseVoucher);
	
	Thread.sleep(1000);
	
	ClickUsingJs(PuchaseVoucher_Settingsbut);
	
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rules_Tab));
	Rules_Tab.click();
	
	Thread.sleep(1500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rulenametext));
	Rulenametext.click();
	Thread.sleep(1000);
	Rulenametext.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	Thread.sleep(1000);
	Rulenametext.sendKeys("TestRule2");
	Thread.sleep(1000);
	Rulenametext.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyNewRecord));
	ApplyNewRecord.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyEdit));
	ApplyEdit.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLoad));
	ApplyLoad.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLeave));
	ApplyLeave.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyONEnter));
	ApplyONEnter.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyActiveStatus));
	ApplyActiveStatus.click();
	
	rulesTabIFSelectField.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectFieldText));
	Select d1=new Select(rulesTabIFSelectFieldText);
	d1.selectByVisibleText("H_num");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabOperatorsDrpdwn));
	rulesTabOperatorsDrpdwn.click();
	Thread.sleep(1000);
	Select O1=new Select(rulesTabOperatorsDrpdwn);
	O1.selectByVisibleText("Equal To");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabComparewithDrpdwn));
	rulesTabComparewithDrpdwn.click();
	Thread.sleep(1000);
	Select V1=new Select(rulesTabComparewithDrpdwn);
	V1.selectByVisibleText("Value");
	
	//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectrulesValueConditionTxt));
	//selectrulesValueConditionTxt.click();
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectrulesValueConditionTxt));
	selectrulesValueConditionTxt.click();
	Thread.sleep(1500);
	Select V12=new Select(Valueconditiondrpdown);
	V12.selectByVisibleText("pqr");
	/*rulesValueConditionTxt.sendKeys("pqr");
	Thread.sleep(1500);
	rulesValueConditionTxt.sendKeys(Keys.TAB);
	Thread.sleep(1000);*/
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfFieldName));
	IfFieldName.click();
	Select W1=new Select(IfFieldNameText);
	W1.selectByVisibleText("Hstring");
	IfFieldNameText.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfChangeValueTxt));
	IfChangeValueTxt.click();
	Thread.sleep(1000);
	IfChangeValueTxt.sendKeys("Value");
	Thread.sleep(1000);
	IfChangeValueTxt.sendKeys(Keys.TAB);
	Thread.sleep(1000);
	IfChangeValueTxt.sendKeys(Keys.TAB);
	Thread.sleep(1500);
	Select V13=new Select(IfValuecheck2);
	V13.selectByVisibleText("yes");
	
	/*
	 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
	 * IfValuecheck2)); IfValuecheck2.sendKeys("yes"); Thread.sleep(1000);
	 * IfValuecheck2.sendKeys(Keys.TAB);
	 */
	
	Thread.sleep(1500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Ifoption));
	ClickUsingJs(Ifoption);
	Thread.sleep(1500);
	contextClick(Ifoption);
	Thread.sleep(1000);
	Addnewcondition.click();
	Thread.sleep(1000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseIfoption));
	ClickUsingJs(ElseIfoption);
	
	selectfield.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectFieldText));
	Select D2=new Select(rulesTabIFSelectFieldText);
	D2.selectByVisibleText("H_num");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseifrulesTabOperatorsDrpdwn));
	ElseifrulesTabOperatorsDrpdwn.click();
	Thread.sleep(1000);
	Select O2=new Select(ElseifrulesTabOperatorsDrpdwn);
	O2.selectByVisibleText("Equal To");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseifrulesTabComparewithDrpdwn));
	ElseifrulesTabComparewithDrpdwn.click();
	Thread.sleep(1000);
	Select V2=new Select(ElseifrulesTabComparewithDrpdwn);
	V2.selectByVisibleText("Value");
	Thread.sleep(1000);
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseifselectValueConditionTxt));
	ElseifselectValueConditionTxt.click();
	Thread.sleep(1000);
	Select V21=new Select(ElseifselectrulesvalueConditionTxt);
	V21.selectByVisibleText("xyz");
	
	
	/*
	 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
	 * ElseifselectrulesvalueConditionTxt));
	 * ElseifselectrulesvalueConditionTxt.click(); Thread.sleep(1000);
	 * ElseifselectrulesvalueConditionTxt.sendKeys("xyz"); Thread.sleep(1000);
	 * ElseifselectrulesvalueConditionTxt.sendKeys(Keys.TAB);
	 */
	 
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseIfFieldName));
	ElseIfFieldName.click();
	Select W2=new Select(ElseIfFieldNameText);
	W2.selectByVisibleText("Hstring");
	ElseIfFieldNameText.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseIfChangeValueTxt));
	ElseIfChangeValueTxt.click();
	Thread.sleep(1000);
	ElseIfChangeValueTxt.sendKeys("Value");
	Thread.sleep(1000);
	ElseIfChangeValueTxt.sendKeys(Keys.TAB);
	Thread.sleep(1000);
	ElseIfChangeValueTxt.sendKeys(Keys.TAB);
	Thread.sleep(1500);
	Select W21=new Select(IfValuecheck2);
	W21.selectByVisibleText("no");
	/*
	 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
	 * IfValuecheck2)); IfValuecheck2.click(); Thread.sleep(1000);
	 * IfValuecheck2.sendKeys("no"); Thread.sleep(1000);
	 * IfValuecheck2.sendKeys(Keys.TAB);
	 */
	
	Thread.sleep(1500);
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Elseoption));
	Elseoption.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseFieldName));
	ElseFieldName.click();
	Select W3=new Select(ElseFieldNameText);
	W3.selectByVisibleText("Hstring");
	ElseFieldNameText.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseChangeValueTxt));
	ElseChangeValueTxt.click();
	Thread.sleep(1000);
	ElseChangeValueTxt.sendKeys("Value");
	Thread.sleep(1000);
	ElseChangeValueTxt.sendKeys(Keys.TAB);
	Thread.sleep(1500);
	ElseChangeValueTxt.sendKeys(Keys.TAB);

	Thread.sleep(1500);
	
	ClickUsingJs(saverule);
	
	String expMessage = "Rule Saved Successfully";

	String actMessage = checkValidationMessage(expMessage);
	
	Thread.sleep(2000);
	
	ClickUsingJs(SettingsUpdatebut);
	
	Thread.sleep(1500);
	
	String expMessage1 = "Data saved Successfully";

	String actMessage1 = checkValidationMessage(expMessage1);

	
	ClickUsingJs(SettingsClosebut);
	
	if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1))
	{
		return true;
	}
	
	else
	{
		return false;
	}
	

}

		
		
	
	
	
	
	


		public RulesLoadingPage(WebDriver driver)
		{

				PageFactory.initElements(driver, this);

		}


		public static boolean CheckPurchaseVoucherwithcreatedRuleConditionStringandNumberListextrafields() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
		
			Thread.sleep(2000);
			
			ClickUsingJs(FinancialMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(FinancialTransactionsMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(TransactionPurchaseMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(PurchaseVoucher);
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NumText_Column));
			NumText_Column.click();
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NumText_Column));
			//Select s1=new Select(NumText_Column);
			//s1.selectByVisibleText("pqr");
			NumText_Column.sendKeys("pqr");
			
			Thread.sleep(1000);
			NumText_Column.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			NumText_Column.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			stringTxt_Column.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			String actString1=stringTxt_Column.getAttribute("value");
			String expString1="yes";
			
			System.out.println("act String Txt1 :" + actString1 );
			System.out.println("exp String Txt1 :" + expString1);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NumText_Column));
			NumText_Column.click();
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NumText_Column));
			//Select s2=new Select(NumText_Column);
			//s2.selectByVisibleText("xyz");
			NumText_Column.sendKeys("xyz");
			
			Thread.sleep(1000);
			NumText_Column.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			NumText_Column.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			stringTxt_Column.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			String actString2=stringTxt_Column.getAttribute("value");
			String expString2="no";
			
			System.out.println("act String Txt2 :" + actString2);
			System.out.println("exp String Txt2 :" + expString2);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NumText_Column));
			NumText_Column.click();
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NumText_Column));
			//Select s3=new Select(NumText_Column);
			//s3.selectByVisibleText("abc");
			NumText_Column.sendKeys("abc");
			Thread.sleep(1000);
			
			NumText_Column.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			NumText_Column.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			stringTxt_Column.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			Select s1=new Select(stringTxt_Column);
			s1.getFirstSelectedOption();
			
			
			boolean actString3=s1.getFirstSelectedOption().getText().isEmpty();
			boolean expString3=true;
			
			System.out.println("act String Txt3 :" + actString3);
			System.out.println("exp String Txt3 :" + expString3);
			
			
			Thread.sleep(2000);
			
			ClickUsingJs(VoucherCloseBut);
			
			Thread.sleep(2000);
			
			ClickUsingJs(PopUp_Yesbut);
			Thread.sleep(2000);
			
			

			
			if(actString1.equalsIgnoreCase(expString1) && actString2.equalsIgnoreCase(expString2)
					&& actString3==(expString3))
					
			{
				return true;
			}
			else
			{
				return false;
			}

		}


		public static boolean CreatingRuleWithConditionPurAccandFootervalInPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			Thread.sleep(2000);
			
			ClickUsingJs(FinancialMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(FinancialTransactionsMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(TransactionPurchaseMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(PurchaseVoucher);
			
			Thread.sleep(1000);
			
			ClickUsingJs(PuchaseVoucher_Settingsbut);
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rules_Tab));
			Rules_Tab.click();
			
			Thread.sleep(1500);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rulenametext));
			Rulenametext.click();
			Thread.sleep(1000);
			Rulenametext.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
			Thread.sleep(1000);
			Rulenametext.sendKeys("TestRule3");
			Thread.sleep(1000);
			Rulenametext.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyNewRecord));
			ApplyNewRecord.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyEdit));
			ApplyEdit.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLoad));
			ApplyLoad.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLeave));
			ApplyLeave.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyONEnter));
			ApplyONEnter.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyActiveStatus));
			ApplyActiveStatus.click();
			
			rulesTabIFSelectField.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectFieldText));
			Select d1=new Select(rulesTabIFSelectFieldText);
			d1.selectByVisibleText("PurchaseAC");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabOperatorsDrpdwn));
			rulesTabOperatorsDrpdwn.click();
			Thread.sleep(1000);
			Select O1=new Select(rulesTabOperatorsDrpdwn);
			O1.selectByVisibleText("Equal To");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabComparewithDrpdwn));
			rulesTabComparewithDrpdwn.click();
			Thread.sleep(1000);
			Select V1=new Select(rulesTabComparewithDrpdwn);
			V1.selectByVisibleText("Value");
			Thread.sleep(1000);
			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectrulesTabConditionTxt));
			selectrulesTabConditionTxt.click();
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabConditionTxt));
			rulesTabConditionTxt.click();
			Thread.sleep(1000);
			rulesTabConditionTxt.sendKeys("Purchase");
			Thread.sleep(1000);
			rulesTabConditionTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfFieldName));
			IfFieldName.click();
			Select W1=new Select(IfFieldNameText);
			W1.selectByVisibleText("f1");
			IfFieldNameText.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfChangeValueTxt));
			IfChangeValueTxt.click();
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys("Value");
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfValue2));
			IfValue2.sendKeys("10");
			Thread.sleep(1000);
			IfValue2.sendKeys(Keys.TAB);
			
			Thread.sleep(1500);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Elseoption));
			Elseoption.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseFieldName));
			ElseFieldName.click();
			Select W3=new Select(ElseFieldNameText);
			W3.selectByVisibleText("f1");
			ElseFieldNameText.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseChangeValueTxt));
			ElseChangeValueTxt.click();
			Thread.sleep(1000);
			ElseChangeValueTxt.sendKeys("Value");
			Thread.sleep(1000);
			ElseChangeValueTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			ElseChangeValueTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseValuecheck));
			ElseValuecheck.click();
			Thread.sleep(1000);
			ElseTxt.sendKeys("20");

			Thread.sleep(1500);
			
			ClickUsingJs(saverule);
			
			String expMessage = "Rule Saved Successfully";

			String actMessage = checkValidationMessage(expMessage);
			
			Thread.sleep(2000);
			
			ClickUsingJs(SettingsUpdatebut);
			
			Thread.sleep(1500);
			
			String expMessage1 = "Data saved Successfully";

			String actMessage1 = checkValidationMessage(expMessage1);

			
			ClickUsingJs(SettingsClosebut);
			
			if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1))
			{
				return true;
			}
			
			else
			{
				return false;
			}
			

		}


		public static boolean CheckPurchaseVoucherwithcreatedRuleConditionPurAccandFooterval() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			
			Thread.sleep(2000);
			
			ClickUsingJs(FinancialMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(FinancialTransactionsMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(TransactionPurchaseMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(PurchaseVoucher);
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PurAccTxt));
			PurAccTxt.click();
			Thread.sleep(1000);
			PurAccTxt.sendKeys("Purchase");
			Thread.sleep(1000);
			PurAccTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			String actfootervalue1=Footerfield.getAttribute("value");
			String expfootervalue1="10";
			
			System.out.println("Act footer value 1: " + actfootervalue1);
			System.out.println("Exp footer value 1: " + expfootervalue1);
			
			Thread.sleep(2000);
			
			ClickUsingJs(PurAccTxt);
			Thread.sleep(1000);
			PurAccTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
			Thread.sleep(1000);
			PurAccTxt.sendKeys("Inventory  Trade");
			Thread.sleep(1000);
			PurAccTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			String actfootervalue2=Footerfield.getAttribute("value");
			String expfootervalue2="20";
			
			System.out.println("Act footer value 2: " + actfootervalue2);
			System.out.println("Exp footer value 2: " + expfootervalue2);
			
			Thread.sleep(2000);
			
			ClickUsingJs(VoucherCloseBut);
			
			Thread.sleep(2000);
			
			ClickUsingJs(PopUp_Yesbut);
			Thread.sleep(2000);
			
			

			
			if(actfootervalue1.equalsIgnoreCase(expfootervalue1) && actfootervalue2.equalsIgnoreCase(expfootervalue2))
					
			{
				return true;
			}
			else
			{
				return false;
			}

		}
		
		public static boolean CreatingRuleWithConditionItemandQuantityValueInPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			Thread.sleep(2000);
			
			ClickUsingJs(FinancialMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(FinancialTransactionsMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(TransactionPurchaseMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(PurchaseVoucher);
			
			Thread.sleep(1000);
			
			ClickUsingJs(PuchaseVoucher_Settingsbut);
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rules_Tab));
			Rules_Tab.click();
			
			Thread.sleep(1500);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rulenametext));
			Rulenametext.click();
			Thread.sleep(1000);
			Rulenametext.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
			Thread.sleep(1000);
			Rulenametext.sendKeys("TestRule4");
			Thread.sleep(1000);
			Rulenametext.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyNewRecord));
			ApplyNewRecord.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyEdit));
			ApplyEdit.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLoad));
			ApplyLoad.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLeave));
			ApplyLeave.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyONEnter));
			ApplyONEnter.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyActiveStatus));
			ApplyActiveStatus.click();
			
			rulesTabIFSelectField.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectFieldText));
			Select d1=new Select(rulesTabIFSelectFieldText);
			d1.selectByVisibleText("Item");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabOperatorsDrpdwn));
			rulesTabOperatorsDrpdwn.click();
			Thread.sleep(1000);
			Select O1=new Select(rulesTabOperatorsDrpdwn);
			O1.selectByVisibleText("Equal To");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabComparewithDrpdwn));
			rulesTabComparewithDrpdwn.click();
			Thread.sleep(1000);
			Select V1=new Select(rulesTabComparewithDrpdwn);
			V1.selectByVisibleText("Value");
			Thread.sleep(1000);
			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectrulesTabConditionTxt));
			selectrulesTabConditionTxt.click();
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabConditionTxt));
			rulesTabConditionTxt.click();
			Thread.sleep(1000);
			rulesTabConditionTxt.sendKeys("Item 1");
			Thread.sleep(1000);
			rulesTabConditionTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfFieldName));
			IfFieldName.click();
			Select W1=new Select(IfFieldNameText);
			W1.selectByVisibleText("Quantity");
			IfFieldNameText.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfChangeValueTxt));
			IfChangeValueTxt.click();
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys("Value");
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfValue2));
			IfValue2.sendKeys("10");
			Thread.sleep(1000);
			IfValue2.sendKeys(Keys.TAB);
			
			Thread.sleep(1500);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Elseoption));
			Elseoption.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseFieldName));
			ElseFieldName.click();
			Select W3=new Select(ElseFieldNameText);
			W3.selectByVisibleText("Quantity");
			ElseFieldNameText.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseChangeValueTxt));
			ElseChangeValueTxt.click();
			Thread.sleep(1000);
			ElseChangeValueTxt.sendKeys("Value");
			Thread.sleep(1000);
			ElseChangeValueTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			ElseChangeValueTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseValuecheck));
			ElseValuecheck.click();
			Thread.sleep(1000);
			ElseTxt.sendKeys("20");

			Thread.sleep(1500);
			
			ClickUsingJs(saverule);
			
			String expMessage = "Rule Saved Successfully";

			String actMessage = checkValidationMessage(expMessage);
			
			Thread.sleep(2000);
			
			ClickUsingJs(SettingsUpdatebut);
			
			Thread.sleep(1500);
			
			String expMessage1 = "Data saved Successfully";

			String actMessage1 = checkValidationMessage(expMessage1);

			
			ClickUsingJs(SettingsClosebut);
			
			if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1))
			{
				return true;
			}
			
			else
			{
				return false;
			}
			

		}


		public static boolean CheckPurchaseVoucherwithcreatedRuleConditionItemandQuantityValue() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			
			Thread.sleep(2000);
			
			ClickUsingJs(FinancialMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(FinancialTransactionsMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(TransactionPurchaseMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(PurchaseVoucher);
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			Thread.sleep(4000);
			
			ClickUsingJs(Select1stColumn);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemTxt));
			ItemTxt.sendKeys("Item 1");
			Thread.sleep(1000);
			ItemTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Select5thColumn));
			Select5thColumn.click();
			
			
			String actQuantity1=QuantityTxt.getAttribute("value");
			String expQuantity1="10.00";
			
			System.out.println("Act Quantity1 = " + actQuantity1);
			System.out.println("Exp Quantity1 = " + expQuantity1);
			
			Thread.sleep(2000);
			
			ClickUsingJs(Select1stColumn);
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemTxt));
			ItemTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
			Thread.sleep(1000);
			ItemTxt.sendKeys("item3");
			Thread.sleep(1000);
			ItemTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Select5thColumn));
			Select5thColumn.click();
			
			String actQuantity2=QuantityTxt.getAttribute("value");
			String expQuantity2="20.00";
			
			System.out.println("Act Quantity2 = " + actQuantity2);
			System.out.println("Exp Quantity2 = " + expQuantity2);
			
			Thread.sleep(2000);
			
			ClickUsingJs(VoucherCloseBut);
			
			Thread.sleep(2000);
			
			ClickUsingJs(PopUp_Yesbut);
			Thread.sleep(2000);
			
			

			
			if(actQuantity1.equalsIgnoreCase(expQuantity1) && actQuantity2.equalsIgnoreCase(expQuantity2))
					
			{
				return true;
			}
			else
			{
				return false;
			}

		}

		public static boolean CreatingRuleConditionLoadingwithColours() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			Thread.sleep(2000);
			
			ClickUsingJs(FinancialMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(FinancialTransactionsMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(TransactionPurchaseMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(PurchaseVoucher);
			
			Thread.sleep(1000);
			
			ClickUsingJs(PuchaseVoucher_Settingsbut);
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rules_Tab));
			Rules_Tab.click();
			
			Thread.sleep(1500);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rulenametext));
			Rulenametext.click();
			Thread.sleep(1000);
			Rulenametext.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
			Thread.sleep(1000);
			Rulenametext.sendKeys("TestRule5");
			Thread.sleep(1000);
			Rulenametext.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyNewRecord));
			ApplyNewRecord.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyEdit));
			ApplyEdit.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLoad));
			ApplyLoad.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLeave));
			ApplyLeave.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyONEnter));
			ApplyONEnter.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyActiveStatus));
			ApplyActiveStatus.click();
			
			rulesTabIFSelectField.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectFieldText));
			Select d1=new Select(rulesTabIFSelectFieldText);
			d1.selectByVisibleText("Quantity");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabOperatorsDrpdwn));
			rulesTabOperatorsDrpdwn.click();
			Thread.sleep(1000);
			Select O1=new Select(rulesTabOperatorsDrpdwn);
			O1.selectByVisibleText("Less Than or Equal To");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabComparewithDrpdwn));
			rulesTabComparewithDrpdwn.click();
			Thread.sleep(1000);
			Select V1=new Select(rulesTabComparewithDrpdwn);
			V1.selectByVisibleText("Value");
			Thread.sleep(1000);
			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectrulesTabConditionTxt));
			selectrulesTabConditionTxt.click();
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesValueConditionTxt));
			rulesValueConditionTxt.click();
			Thread.sleep(1000);
			rulesValueConditionTxt.sendKeys("50");
			Thread.sleep(1000);
			rulesValueConditionTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfFieldName));
			IfFieldName.click();
			Select W1=new Select(IfFieldNameText);
			W1.selectByVisibleText("Rate");
			IfFieldNameText.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfChangeValueTxt));
			IfChangeValueTxt.click();
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys("None");
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SelectIfFontField));
			ClickUsingJs(SelectIfFontField);
			Thread.sleep(1000);
			IfFontTxt.click();
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfBackColourFiled));
			IfBackColourFiled.click();
			
			Select s1=new Select(IfBackColourFiled);
			s1.selectByVisibleText("Green");
			
			ClickUsingJs(If_FontSave);
			
			Thread.sleep(1500);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Elseoption));
			Elseoption.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseFieldName));
			ElseFieldName.click();
			Select W3=new Select(ElseFieldNameText);
			W3.selectByVisibleText("Rate");
			ElseFieldNameText.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseChangeValueTxt));
			ElseChangeValueTxt.click();
			Thread.sleep(1000);
			ElseChangeValueTxt.sendKeys("None");
			Thread.sleep(1000);
			ElseChangeValueTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			ElseChangeValueTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SelectElseFontField));
			ClickUsingJs(SelectElseFontField);
			Thread.sleep(1000);
			ElseFontTxt.click();
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseBackColourFiled));
			ElseBackColourFiled.click();
			
			Select s2=new Select(ElseBackColourFiled);
			s2.selectByVisibleText("Red");
			
			ClickUsingJs(Else_FontSave);
			
			Thread.sleep(1500);
			
			ClickUsingJs(saverule);
			
			String expMessage = "Rule Saved Successfully";

			String actMessage = checkValidationMessage(expMessage);
			
			Thread.sleep(2000);
			
			ClickUsingJs(SettingsUpdatebut);
			
			Thread.sleep(1500);
			
			String expMessage1 = "Data saved Successfully";

			String actMessage1 = checkValidationMessage(expMessage1);

			
			ClickUsingJs(SettingsClosebut);
			
			if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1))
			{
				return true;
			}
			
			else
			{
				return false;
			}
			

		}


		public static boolean CheckPurchaseVoucherwithcreatedRuleConditionLoadingwithColour() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			Thread.sleep(2000);
			
			ClickUsingJs(FinancialMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(FinancialTransactionsMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(TransactionPurchaseMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(PurchaseVoucher);
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			Thread.sleep(4000);
			
			ClickUsingJs(Select1stColumn);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemTxt));
			ItemTxt.sendKeys("Item 1");
			Thread.sleep(1000);
			ItemTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Select5thColumn));
			Select5thColumn.click();
			Thread.sleep(1000);
			QuantityTxt.sendKeys("30");
			Thread.sleep(1000);
			QuantityTxt.sendKeys(Keys.TAB);
			
			String actratecolour1=Select7thColumn.getCssValue("background-color");
			String expratecolour1="rgba(0, 128, 0, 1)";
			
			System.out.println("Act colour 1 = " + actratecolour1);
			System.out.println("Exp colour 1 = " + expratecolour1);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Select5thColumn));
			Select5thColumn.click();
			Thread.sleep(1000);
			QuantityTxt.sendKeys("55");
			Thread.sleep(1000);
			QuantityTxt.sendKeys(Keys.TAB);
			
			String actratecolour2=Select7thColumn.getCssValue("background-color");
			String expratecolour2="rgba(255, 0, 0, 1)";
			
			System.out.println("Act colour 2 = " + actratecolour2);
			System.out.println("Exp colour 2 = " + expratecolour2);
			
			Thread.sleep(2000);
			
			ClickUsingJs(VoucherCloseBut);
			
			Thread.sleep(2000);
			
			ClickUsingJs(PopUp_Yesbut);
			Thread.sleep(2000);
			
			

			
			if(actratecolour1.equalsIgnoreCase(expratecolour1) && actratecolour2.equalsIgnoreCase(expratecolour2))
					
			{
				return true;
			}
			else
			{
				return false;
			}

		}
		
		public static boolean CreatingRuleConditionwithExtrafieldsinbodywithnumberlistandstringlistdatatype() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			Thread.sleep(5000);
			
			ClickUsingJs(PuchaseVoucher_Settingsbut);
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rules_Tab));
			Rules_Tab.click();
			
			Thread.sleep(1500);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rulenametext));
			Rulenametext.click();
			Thread.sleep(1000);
			Rulenametext.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
			Thread.sleep(1000);
			Rulenametext.sendKeys("TestRule6");
			Thread.sleep(1000);
			Rulenametext.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyNewRecord));
			ApplyNewRecord.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyEdit));
			ApplyEdit.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLoad));
			ApplyLoad.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLeave));
			ApplyLeave.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyONEnter));
			ApplyONEnter.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyActiveStatus));
			ApplyActiveStatus.click();
			
			rulesTabIFSelectField.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectFieldText));
			Select d1=new Select(rulesTabIFSelectFieldText);
			d1.selectByVisibleText("B_Str");
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabOperatorsDrpdwn));
			rulesTabOperatorsDrpdwn.click();
			Thread.sleep(1000);
			Select O1=new Select(rulesTabOperatorsDrpdwn);
			O1.selectByVisibleText("Equal To");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabComparewithDrpdwn));
			rulesTabComparewithDrpdwn.click();
			Thread.sleep(1000);
			Select V1=new Select(rulesTabComparewithDrpdwn);
			V1.selectByVisibleText("Value");
			Thread.sleep(1000);
			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectrulesTabConditionTxt));
			selectrulesTabConditionTxt.click();
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabConditionTxt));
			rulesTabConditionTxt.click();
			Thread.sleep(1000);
			rulesTabConditionTxt.sendKeys("sec");
			Thread.sleep(1000);
			rulesTabConditionTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfFieldName));
			IfFieldName.click();
			Select W1=new Select(IfFieldNameText);
			W1.selectByVisibleText("B_num");
			IfFieldNameText.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfChangeValueTxt));
			IfChangeValueTxt.click();
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys("Value");
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfValuecheck2));
			IfValuecheck2.sendKeys("twitter");
			Thread.sleep(1000);
			IfValuecheck2.sendKeys(Keys.TAB);
			
			Thread.sleep(1500);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Elseoption));
			Elseoption.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseFieldName));
			ElseFieldName.click();
			Select W3=new Select(ElseFieldNameText);
			W3.selectByVisibleText("B_num");
			ElseFieldNameText.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseChangeValueTxt));
			ElseChangeValueTxt.click();
			Thread.sleep(1000);
			ElseChangeValueTxt.sendKeys("Value");
			Thread.sleep(1000);
			ElseChangeValueTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			ElseChangeValueTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ElseValuecheck));
			ElseValuecheck.click();
			Thread.sleep(1000);
			ElseTxt.sendKeys("yahoo");

			Thread.sleep(1500);
			
			ClickUsingJs(saverule);
			
			String expMessage = "Rule Saved Successfully";

			String actMessage = checkValidationMessage(expMessage);
			
			Thread.sleep(2000);
			
			ClickUsingJs(SettingsUpdatebut);
			
			Thread.sleep(1500);
			
			String expMessage1 = "Data saved Successfully";

			String actMessage1 = checkValidationMessage(expMessage1);

			
			ClickUsingJs(SettingsClosebut);
			
			if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1))
			{
				return true;
			}
			
			else
			{
				return false;
			}
			

		}


		public static boolean checkLogoutRulesLoadingPage() throws InterruptedException {
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


		public static boolean SavingPurchaseVoucherwithcreatedRuleWithDisableCondition() throws InterruptedException {
			
			Thread.sleep(2000);
			
			ClickUsingJs(FinancialMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(FinancialTransactionsMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(TransactionPurchaseMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(PurchaseVoucher);
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Select1stColumn));
			Select1stColumn.click();
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemTxt));
			ItemTxt.sendKeys("Item 1");
			Thread.sleep(1000);
			ItemTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			Thread.sleep(1000);
			UnitTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			Observe1_txt.sendKeys("1");
			Thread.sleep(1000);
			Observe1_txt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			
			Observe2_txt.sendKeys("20");
			Thread.sleep(1000);
			Observe2_txt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			QuantityTxt.sendKeys("12");
			Thread.sleep(1000);
			QuantityTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			String actString1=Observe2_txt.getAttribute("value");
			String expString1="20";
			
			System.out.println("act String Txt1 :" + actString1 );
			System.out.println("exp String Txt1 :" + expString1);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Select2ndrow_1stColumn));
			Select2ndrow_1stColumn.click();
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemTxt));
			ItemTxt.sendKeys("Item 1");
			Thread.sleep(1000);
			ItemTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			UnitTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			Observe1_txt.sendKeys("2");
			Thread.sleep(1000);
			Observe1_txt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
		
			QuantityTxt.sendKeys("12");
			Thread.sleep(1000);
			QuantityTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			String actString2=Select2ndrow_4thColumn.getAttribute("data-brulereadonly");
			String expString2="true";
			
			/*boolean actString2=Observe2_txt.isEnabled();
			boolean expString2=false;*/
			
			System.out.println("act String Txt2 :" + actString2 );
			System.out.println("exp String Txt2 :" + expString2);
			
			Thread.sleep(2000);
			
			ClickUsingJs(VoucherCloseBut);
			
			Thread.sleep(2000);
			
			ClickUsingJs(PopUp_Yesbut);
			Thread.sleep(2000);
			
			

			
			if(actString1.equalsIgnoreCase(expString1) && actString2.equalsIgnoreCase(expString2))
			{
				return true;
			}
			else
			{
				return false;
			}

		}


		public static boolean CreatingRuleWithFootervalueValidationWithFORMULAE() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
				
			Thread.sleep(2000);
			
			ClickUsingJs(FinancialMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(FinancialTransactionsMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(TransactionPurchaseMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(PurchaseVoucher);
			
			Thread.sleep(1000);
			
			ClickUsingJs(PuchaseVoucher_Settingsbut);
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rules_Tab));
			Rules_Tab.click();
			
			Thread.sleep(1500);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Rulenametext));
			Rulenametext.click();
			Thread.sleep(1000);
			Rulenametext.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
			Thread.sleep(1000);
			Rulenametext.sendKeys("FormulaeRule");
			Thread.sleep(1000);
			Rulenametext.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyNewRecord));
			ApplyNewRecord.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyEdit));
			ApplyEdit.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLoad));
			ApplyLoad.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyLeave));
			ApplyLeave.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyONEnter));
			ApplyONEnter.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ApplyActiveStatus));
			ApplyActiveStatus.click();
			
			Thread.sleep(1000);
			
			NOCondition.click();
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfFieldName));
			IfFieldName.click();
			Select W1=new Select(IfFieldNameText);
			W1.selectByVisibleText("f1");
			IfFieldNameText.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfChangeValueTxt));
			IfChangeValueTxt.click();
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys("Value");
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfValue2));
			IfValue2.sendKeys("10");
			Thread.sleep(1000);
			IfValue2.sendKeys(Keys.TAB);
			
			Thread.sleep(1500);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ChangeValue_Column));
			ChangeValue_Column.click();
			Thread.sleep(1000);
			IfChangeValueTxt.click();
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys("Formula");
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			IfChangeValueTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(IfValue2));
			IfValue2.sendKeys("SBO3");
			Thread.sleep(1000);
			IfValue2.sendKeys(Keys.TAB);
			
			ClickUsingJs(saverule);
			
			String expMessage = "Rule Saved Successfully";

			String actMessage = checkValidationMessage(expMessage);
			
			Thread.sleep(2000);
			
			ClickUsingJs(SettingsUpdatebut);
			
			Thread.sleep(1500);
			
			String expMessage1 = "Data saved Successfully";

			String actMessage1 = checkValidationMessage(expMessage1);

			
			ClickUsingJs(SettingsClosebut);
			
			if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1))
			{
				return true;
			}
			
			else
			{
				return false;
			}
			
		}


		public static boolean CheckPurchaseVoucherwithcreatedRuleWithFootervalueValidationWithFORMULAE() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			Thread.sleep(2000);
			
			ClickUsingJs(FinancialMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(FinancialTransactionsMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(TransactionPurchaseMenu);
			
			Thread.sleep(1000);
			
			ClickUsingJs(PurchaseVoucher);
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			Thread.sleep(2500);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PurAccTxt));
			PurAccTxt.click();
			Thread.sleep(1000);
			PurAccTxt.sendKeys("Purchase");
			Thread.sleep(1000);
			PurAccTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(VendorAccTxt));
			VendorAccTxt.sendKeys("Vendor A");
			Thread.sleep(1000);
			VendorAccTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepartmentTxt));
			DepartmentTxt.click();
			Thread.sleep(1000);
			DepartmentTxt.sendKeys("Dep 1");
			Thread.sleep(1000);
			DepartmentTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Select1stColumn));
			Select1stColumn.click();
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemTxt));
			ItemTxt.sendKeys("Item 1");
			Thread.sleep(1000);
			ItemTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			Thread.sleep(1000);
			UnitTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			Observe1_txt.sendKeys("15");
			Thread.sleep(1000);
			Observe1_txt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			
			Observe2_txt.sendKeys("20");
			Thread.sleep(1000);
			Observe2_txt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			QuantityTxt.sendKeys("12");
			Thread.sleep(1000);
			QuantityTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			//Before Saving
			
			String actfootervalueBeforesave=Footerfield.getAttribute("value");
			String expfootervalueBeforesave="0";
			
			System.out.println("Act footer value Beforesave : " + actfootervalueBeforesave);
			System.out.println("Exp footer value Beforesave : " + expfootervalueBeforesave);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(VoucherSaveBut));
			ClickUsingJs(VoucherSaveBut);
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Pickbut));
			Pickbut.click();
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(OkBut));
			OkBut.click();
			
			Thread.sleep(1000);
			
			String expMessage = "Voucher saved successfully";
			String actMessage = checkValidationMessage(expMessage);
			
			System.out.println("Actual Message    : "+actMessage);
			System.out.println("Expected Message  : "+expMessage);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Previous_But));
			Previous_But.click();
			
			Thread.sleep(1000);
			
			ClickUsingJs(PopUp_Yesbut);
			Thread.sleep(2000);
			
			String expMessage1 = "Voucher loaded successfully";
			String actMessage1 = checkValidationMessage(expMessage1);
			
			System.out.println("Actual Message 1   : "+actMessage1);
			System.out.println("Expected Message 1 : "+expMessage1);
			
			//After Saving
			String actfootervalueAfterSave=Footerfield.getAttribute("value");
			String expfootervalueAfterSave="15";
			
			System.out.println("Act footer value AfterSave : " + actfootervalueAfterSave);
			System.out.println("Exp footer value AfterSave : " + expfootervalueAfterSave);
			
			if(actfootervalueBeforesave.equalsIgnoreCase(expfootervalueBeforesave) && actMessage.startsWith(expMessage)
					&& actMessage1.startsWith(expMessage1) && actfootervalueAfterSave.equalsIgnoreCase(expfootervalueAfterSave) )
			{
				return true;
			}
			else
			{
				return false;
			}
		}

}

