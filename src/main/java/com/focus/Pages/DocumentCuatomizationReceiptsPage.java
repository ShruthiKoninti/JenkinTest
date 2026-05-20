package com.focus.Pages;

import java.io.IOException;
import java.util.ArrayList;
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

import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

public class DocumentCuatomizationReceiptsPage extends BaseEngine{
	
	public boolean checkSavingDocumentCustomizationReceiptsVoucherDocumentsTabAndVerify() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		Thread.sleep(3000);
		
		getAction().moveToElement(SettingsMenu).build().perform();
		Thread.sleep(2000);
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();*/
		
		ClickUsingJs(SettingsMenu);
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionBtn));
		configureTransactionBtn.click();
		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receiptsBtn));
		receiptsBtn.click();
		Thread.sleep(10000);
		
		getAction().moveToElement(docCustCreateVoucherBtn).build().perform();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustCreateVoucherBtn));
		docCustCreateVoucherBtn.click();

		Thread.sleep(8000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(titleTxt));
		titleTxt.click();
		titleTxt.clear();
		titleTxt.sendKeys("Receipts N VAT");
		titleTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masters1));
		masters1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDropDown));
		masterDropDown.click();
		masterDropDown.sendKeys("Warehouse");
		masterDropDown.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(position1));
		position1.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(positionDropdown));
		positionDropdown.click();
		positionDropdown.sendKeys("Header");
		positionDropdown.sendKeys(Keys.TAB);	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masters2));
		masters2.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDropDown));
		masterDropDown.click();
		masterDropDown.sendKeys("Department");
		masterDropDown.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(position2));
		position2.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(positionDropdown));
		positionDropdown.click();
		positionDropdown.sendKeys("Header");
		Thread.sleep(2000);
		positionDropdown.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		String expMessage = "Data saved successfully";
		String actMessage = checkValidationMessage(expMessage);
		
		

		String acttitleTxt					 =titleTxt.getAttribute("value");
		String exptitleTxt					 ="Receipts N VAT";

		String actMastersR1C1                = masters1.getText();
		String actPositionR1C1               = position1.getText();

		String expMastersR1C1                = "Warehouse";
		String expPositionR1C1               = "Header";
		
		String actMastersR2C1                = masters2.getText();
		String actPositionR2C1               = position2.getText();

		String expMastersR2C1                = "Department";
		String expPositionR2C1               = "Header";


		System.out.println("**************************checkSavingDocumentCustomizationPurchaseVoucherDocumentsTabAndVerify*****************************");

		System.out.println("Error Message                                 : " + actMessage       +     "     Value Expected : " + expMessage);
		System.out.println("Documents Tab Title                           : " + acttitleTxt      +     "	 Value Expected : " + exptitleTxt);
		System.out.println("Documents Tab Master First Row                : " + actMastersR1C1   +     "	 Value Expected : " + expMastersR1C1);
		System.out.println("Documents Tab Master Position First Row       : " + actPositionR1C1  +     "	 Value Expected : " + expPositionR1C1);
		System.out.println("Documents Tab Master First Row                : " + actMastersR2C1   +     "	 Value Expected : " + expMastersR2C1);
		System.out.println("Documents Tab Master Position First Row       : " + actPositionR2C1  +     "	 Value Expected : " + expPositionR2C1);

		if(actMessage.equalsIgnoreCase(expMessage) && acttitleTxt.equalsIgnoreCase(exptitleTxt) && actMastersR1C1.equalsIgnoreCase(expMastersR1C1) 
				&& actPositionR1C1.equalsIgnoreCase(expPositionR1C1))
		{
			
			return true;
		}
		else
		{
			
			return false;
		} 
	}
		
	

	
	 public boolean checkmiscellaneousTabInDCInReceiptNewVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
	
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousTab));
		 miscellaneousTab.click();
		 Thread.sleep(2000);
	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(abbreviationTxt));
			abbreviationTxt.click();
			abbreviationTxt.clear();
			abbreviationTxt.sendKeys("RNVAT");
			abbreviationTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneusTabPostingDetailsMenu));
			getAction().moveToElement(miscellaneusTabPostingDetailsMenu).build().perform();
			miscellaneusTabPostingDetailsMenu.click();
			
			Thread.sleep(1000);
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateFAChkBox));
			if(updateFAChkBoxIsSelected.isSelected()==false)
			{
				updateFAChkBox.click();
			}
			
			//((JavascriptExecutor)getDriver()).executeScript("document.getElementById('misc_PostingUpdateFA').checked=false;");
		
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(doNotCheckCreditLimit));
			if(doNotCheckCreditLimitIsSelected.isSelected()==false)
			{
				doNotCheckCreditLimit.click();
			}
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousTabCurrencyMenu));
			getAction().moveToElement(miscellaneousTabCurrencyMenu).build().perform();
			miscellaneousTabCurrencyMenu.click();
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inputExchangeRateChkBox));
			if(inputExchangeRateChkBoxIsSelected.isSelected()==false)
			{
				inputExchangeRateChkBox.click();
			}
		
			Thread.sleep(2000);
		
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inputLocalExchangeRateChkBox));
			if(inputLocalExchangeRateChkBoxIsSelected.isSelected()==false)
			{
				inputLocalExchangeRateChkBox.click();
			}
			
		
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneusTabARAPBtn));
			miscellaneusTabARAPBtn.click();
			
			Thread.sleep(1000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dueDateDropdown));
			Select s1=new Select(dueDateDropdown);
			s1.selectByVisibleText("Header");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneusTabARAPBtn));
			miscellaneusTabARAPBtn.click();
			
			Thread.sleep(1000);

			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			Thread.sleep(2000);

			String expMessage = "Data saved successfully";
			String actMessage = checkValidationMessage(expMessage);
			
			


			String actabbreviationTxt				= abbreviationTxt.getAttribute("value");
			String expabbreviationTxt				= "RNVAT";
			
			String actupdateFAChkBox				= Boolean.toString(updateFAChkBoxIsSelected.isSelected());
			String expupdateFAChkBox				="true";
			
			
			String actupdateStockChkBox			= Boolean.toString(doNotCheckCreditLimitIsSelected.isSelected());
			String expupdateStockChkBox			= "true";
			
			String actinputExchangeRateChkBox		= Boolean.toString(inputExchangeRateChkBoxIsSelected.isSelected());
			String expinputExchangeRateChkBox		= "true";
			
			String actinputLocalExchangeRateChkBox	= Boolean.toString(inputLocalExchangeRateChkBoxIsSelected.isSelected());		
			String expinputLocalExchangeRateChkBox	= "true";

			System.out.println("********************************checkUpdatingMiscellaneousTabInPurchaseVoucher**************************************");

			System.out.println("ErrorMessage                  : " + actMessage                       + "  value expected  " + expMessage);
			System.out.println("AbbreviationTxt               : " + actabbreviationTxt               + "  value expected  " + expabbreviationTxt);
			System.out.println("UpdateFAChkBox                : " + actupdateFAChkBox                + "  value expected  " + expupdateFAChkBox);
			System.out.println("UpdateStockChkBox             : " + actupdateStockChkBox             + "  value expected  " + expupdateStockChkBox);
			System.out.println("InputExchangeRateChkBox       : " + actinputExchangeRateChkBox       + "  value expected  " + expinputExchangeRateChkBox);
			System.out.println("InputLocalExchangeRateChkBox  : " + actinputLocalExchangeRateChkBox  + "  value expected  " + expinputLocalExchangeRateChkBox);

			if(actMessage.equalsIgnoreCase(expMessage) && actabbreviationTxt.equalsIgnoreCase(actabbreviationTxt) && actupdateFAChkBox.equalsIgnoreCase(expupdateFAChkBox) 
					&& actupdateStockChkBox.equalsIgnoreCase(expupdateStockChkBox) && actinputExchangeRateChkBox.equalsIgnoreCase(expinputExchangeRateChkBox) && actinputLocalExchangeRateChkBox.equalsIgnoreCase(expinputLocalExchangeRateChkBox))
			{
				
				return true;
			}
			else
			{
				
				return false;
			}	
			
	 }
	
	 
	 public boolean checktheNarrationinHeaderTabEditLayoutinReceiptsNVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(4000);		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTab));
			editLayoutTab.click();

			Thread.sleep(4000);

			String actdata,expdata;

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutHeader1stRowCaption));
			editLayoutHeader1stRowCaption.click();

			Thread.sleep(3000);

			actdata=editLayoutHeader1stRowCaption.getText();
			expdata="Narration";

			System.out.println("actdata"+actdata);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayouEditBtn));
			editLayouEditBtn.click();

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutDefaultValueTxt));
			editLayoutDefaultValueTxt.click();
			editLayoutDefaultValueTxt.clear();
			
			editLayoutDefaultValueTxt.sendKeys("Test");//332

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutPropertiesTab));
			editLayoutPropertiesTab.click();

			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(readOnlyChkBox));
			readOnlyChkBox.click();

			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutApplyBtn));
			editLayoutApplyBtn.click();
			Thread.sleep(2000);

			String expMessageClickOnApply ="Data saved successfully";

			String actMessageClickOnApply = checkValidationMessage(expMessageClickOnApply);

			

			if(actMessageClickOnApply.equalsIgnoreCase(expMessageClickOnApply))
			{
				System.out.println("Test Pass : Data Saved In Edit Screeen Field Details");
			
				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT Saved In Edit Screeen Field Details");
				
				return false;

			}
		}


	 public boolean checkSelectingApplyRateSchemesInInventoryOptionsTabInReceiptsNVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryoptionsTab));
			inventoryoptionsTab.click();
			Thread.sleep(2000);
			
			
			
			
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(applyrateSchemesChkBox));
			applyrateSchemesChkBox.click();
			
			Thread.sleep(1000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTab));
			editLayoutTab.click();	

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryoptionsTab));
			inventoryoptionsTab.click();	


			String actapplyrateSchemesChkBox = Boolean.toString(applyrateSchemesChkBoxIsSelected.isSelected());
			String expapplyrateSchemesChkBox = "";

			System.out.println("*******************************************checkSelectingApplyRateSchemesInInventoryOptionsTabInTrialPurchaseUpdate*******************************************");

			System.out.println("ApplyrateSchemesChkBox : "+actapplyrateSchemesChkBox+"  Value Expected  "+expapplyrateSchemesChkBox);

			if(actapplyrateSchemesChkBox.equalsIgnoreCase(expapplyrateSchemesChkBox))
			{	
				

				return  true;
			}
			else
			{
				

				return  false;
			}
		}
	
	
	 public boolean checkAddFieldInEditLayoutBodyInReceiptsNVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		 Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBodyTab));
			editLayoutBodyTab.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutAddFieldsBtn));
			editLayoutAddFieldsBtn.click();

			Thread.sleep(2000);
			
			//getAlert().accept();
			//Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCaptionTxt));
			editLayoutCaptionTxt.click();
			editLayoutCaptionTxt.sendKeys("TrialReceiptBody");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutApplyBtn));
			editLayoutApplyBtn.click();

			System.out.println("*******************************************checkAddFieldInEditLayoutBodyInTrialPurchaseUpdate*******************************************");

			String expMessage ="Data saved successfully";
			String actMessage = checkValidationMessage(expMessage);
		
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBodyTab));
			editLayoutBodyTab.click();

			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutbody2ndRowCaption));
			String actCaption = editLayoutbody2ndRowCaption.getText();
			String expCaption ="TrialReceiptBody";

			System.out.println("Caption        : "+actCaption+"  Value expected  "+expCaption);

			if(actMessage.equalsIgnoreCase(expMessage) && actCaption.equalsIgnoreCase(expCaption))
			{
				
				
				return true;
			}
			else
			{
				
				return false;
			}
		}







		public boolean checkEditFieldInEditLayoutBodyInReceiptsNVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBodyTab));
			editLayoutBodyTab.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBody2ndRowEditBtn));
			editLayoutBody2ndRowEditBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutCaptionTxt));
			editLayoutCaptionTxt.click();
			editLayoutCaptionTxt.clear();
			editLayoutCaptionTxt.sendKeys("TrialReceiptBodyUpdate");

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutApplyBtn));
			editLayoutApplyBtn.click();

			System.out.println("*******************************************checkEditFieldInEditLayoutBodyInTrialPurchaseUpdate*******************************************");

			String expMessage = "Data saved successfully";
			String actMessage = checkValidationMessage(expMessage);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBodyTab));
			editLayoutBodyTab.click();

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutbody2ndRowCaption));
			String actCaption = editLayoutbody2ndRowCaption.getText();
			String expCaption ="TrialReceiptBodyUpdate";

			System.out.println("Caption        : "+actCaption+"  Value expected  "+expCaption);

			if(actMessage.equalsIgnoreCase(expMessage) && actCaption.equalsIgnoreCase(expCaption))
			{	
				

				return  true;
			}
			else
			{
				

				return  false;
			}
		}







		public boolean checkDeleteFieldInEditLayoutBodyInReceiptsNVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBodyTab));
			editLayoutBodyTab.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutBody2ndRowDeleteBtn));
			editLayoutBody2ndRowDeleteBtn.click();

			Thread.sleep(2000);

			getWaitForAlert();
			getAlert().accept();

			System.out.println("*******************************************checkDeleteFieldInEditLayoutBodyInTrialPurchaseUpdate*******************************************");

			String expMessage ="Data deleted successfully";
			String actMessage = checkValidationMessage(expMessage);
		

			if(actMessage.equalsIgnoreCase(expMessage))
			{	
				
				return  true;
			}
			else
			{
				
				return  false;
			}
		}

		
	

		public boolean checkCreatingRuleinRuleTAbUnderReceiptNVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			

			getDriver().navigate().refresh();
			Thread.sleep(2000);

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
			searchTxt.click();
			searchTxt.sendKeys("Configure Transactions");
			searchTxt.sendKeys(Keys.ENTER);
			Thread.sleep(8000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receiptsBtn));
			receiptsBtn.click();
			Thread.sleep(12000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receiptsNVATBtn));
			receiptsNVATBtn.click();
			Thread.sleep(8000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesBtn));
			rulesBtn.click();

			Thread.sleep(10000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabAddRuleBtn));

			boolean actRuleAdd=rulesTabAddRuleBtn.isDisplayed();

			rulesTabAddRuleBtn.click();

			String actRuleAddBtnDisplay=Boolean.toString(actRuleAdd);
			String expRuleAddBtnDisplay="true";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabRuleNameTxt));
			rulesTabRuleNameTxt.click();

			rulesTabRuleNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);

			rulesTabRuleNameTxt.sendKeys(Keys.BACK_SPACE);		

			Thread.sleep(1000);

			rulesTabRuleNameTxt.sendKeys("NewReceiptRule");

			Thread.sleep(2000);
			
			rulesTabRuleNameTxt.sendKeys(Keys.TAB);

			String actRuleName=rulesTabRuleNameTxt.getText();
			String expRuleName="";

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabNewRecordChkbox));
			rulesTabNewRecordChkbox.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabEditChkbox));
			rulesTabEditChkbox.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabOnLeaveChkbox));
			rulesTabOnLeaveChkbox.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabActiveChkbox));
			rulesTabActiveChkbox.click();

			boolean RulesTabNewRecordChkbox=rulesTabNewRecordChkbox.isSelected();
			String actRulesTabNewRecordChkbox=Boolean.toString(RulesTabNewRecordChkbox);
			String expRulesTabNewRecordChkbox="true";

			boolean RulesTabEditChkbox=rulesTabEditChkbox.isSelected();
			String actRulesTabEditChkbox=Boolean.toString(RulesTabEditChkbox);
			String expRulesTabEditChkbox="true";

			boolean RulesTabOnLeaveChkbox=rulesTabOnLeaveChkbox.isSelected();
			String actRulesTabOnLeaveChkbox=Boolean.toString(RulesTabOnLeaveChkbox);
			String exRrulesTabOnLeaveChkbox="true";

			boolean RulesTabActiveChkbox=rulesTabNewRecordChkbox.isSelected();
			String actRulesTabActiveChkbox=Boolean.toString(RulesTabActiveChkbox);
			String expRulesTabActiveChkbox="true";

			rulesTabIFSelectTabTxt.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectTabDrpdwn));
			Select account=new Select(rulesTabIFSelectTabDrpdwn);
			account.selectByVisibleText("CashBankAC");//6
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
			Value.selectByValue("1");
			rulesTabCompareWithDrpdwn.sendKeys(Keys.TAB);
			rulesTabCompareWithDrpdwn.sendKeys(Keys.TAB);

			Thread.sleep(2000);
			rulesTabConditionTxt.sendKeys(Keys.SPACE);
			rulesTabConditionTxt.sendKeys(Keys.BACK_SPACE);
			Thread.sleep(2000);

			rulesTabConditionTxt.sendKeys("Bank");

			Thread.sleep(4000);

			rulesTabConditionTxt.sendKeys(Keys.TAB);

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid1stRow_1stcol));
			rulesGrid1stRow_1stcol.click();

			rulesIFEnterFiled.sendKeys("sChequeNo");
			rulesIFEnterFiled.sendKeys(Keys.TAB);
			rulesIFEnterChangeValueDrpdwn.sendKeys("Value");

			rulesIFEnterChangeValueDrpdwn.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFEnterValue));

			rulesIFEnterValue.sendKeys("12345");

			rulesIFEnterValue.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabGrid1st_10thCol));
			rulesTabGrid1st_10thCol.click();

			rulesTabGridMandatoryChkbox.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid2ndRow_1stcol));
			rulesGrid2ndRow_1stcol.click();

			rulesIFEnterFiled.sendKeys("Account");
			rulesIFEnterFiled.sendKeys(Keys.TAB);
			rulesIFEnterChangeValueDrpdwn.sendKeys("Value");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid2ndRow_5thcol));
			rulesGrid2ndRow_5thcol.click();
			Thread.sleep(2000);

			rulesTabIFGridEnterCaption.sendKeys("ACCCaption");

			rulesTabIFGridEnterCaption.sendKeys(Keys.TAB);			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTab_IFMesssageTab));
			rulesTab_IFMesssageTab.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rules_MessageTabGeneralTXT));
			rules_MessageTabGeneralTXT.sendKeys("ReceiptsReule");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabElseTab));
			rulesTabElseTab.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGridElse1stRow_1stcol));
			rulesGridElse1stRow_1stcol.click();

			rulesElseEnterFiled.sendKeys("sChequeNo");
			rulesElseEnterFiled.sendKeys(Keys.TAB);
			rulesElseEnterChangeValueDrpdwn.sendKeys("None");

			rulesElseEnterChangeValueDrpdwn.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGridElse2ndRow_1stcol));
			rulesGridElse2ndRow_1stcol.click();

			rulesElseEnterFiled.sendKeys("Account");
			rulesElseEnterFiled.sendKeys(Keys.TAB);
			rulesElseEnterChangeValueDrpdwn.sendKeys("None");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGridElse2ndRow_5thcol));
			rulesGridElse2ndRow_5thcol.click();
			Thread.sleep(2000);

			rulesTabElseGridEnterCaption.sendKeys("Account");

			rulesTabElseGridEnterCaption.sendKeys(Keys.TAB);	
			
			//((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", rulesTab_SaveRuleBtn);
		getAction().moveToElement(rulesTab_SaveRuleBtn).build().perform();
		Thread.sleep(4000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTab_SaveRuleBtn));
			rulesTab_SaveRuleBtn.click();

			String expMessageOnSaveRule = "Rule Saved Successfully";
			String actMessageOnSaveRule=checkValidationMessage(expMessageOnSaveRule);

			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();

			Thread.sleep(2000);
			
			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Settings_closeBtn));
			Settings_closeBtn.click();*/

			String expMessageOnClickOnUpdate = "Data saved successfully";
			String actMessageOnClickOnUpdate=checkValidationMessage(expMessageOnSaveRule);

			

			if (actMessageOnSaveRule.equalsIgnoreCase(expMessageOnSaveRule) && 
					actMessageOnClickOnUpdate.equalsIgnoreCase(expMessageOnClickOnUpdate))
			{
				System.out.println(" Test Pass:  Rule Created Successful");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail:  Rule not Created ");
				
				return false;
			}
		}

		
		
		public boolean 	checkSaveTriggerWithRaiseDocumentOptionInTriggers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{	
			Thread.sleep(5000);
		

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
			triggersBtn.click();

			
			Thread.sleep(4000);
			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
			editLayoutTriggerName.click();
			Thread.sleep(1500);
			editLayoutTriggerName.sendKeys("Receipts Trigger");
			Thread.sleep(4000);
			editLayoutTriggerName.sendKeys(Keys.TAB);			
			Thread.sleep(4000);
			
			triggersConjuctionDrpdwn.click();
			Thread.sleep(2000);		
			triggersConjuctionDrpdwn.sendKeys(Keys.ARROW_DOWN);			
			triggersConjuctionDrpdwn.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
			triggersSelectFieldTxt.click();

			Thread.sleep(3000);
		

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerVendorACExpansionBtn));
			triggerVendorACExpansionBtn.click();

			Thread.sleep(2000);
			
			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerVendorAcName));
			triggerVendorAcName.click();

			Thread.sleep(2000);
			
			//triggerVendorAcName.sendKeys(Keys.TAB);
			triggersSelectFieldTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectOperdrpdwn));
			triggersSelectOperdrpdwn.click();
			triggersSelectOperdrpdwn.sendKeys("Equal");


			Thread.sleep(5000);
			
			triggersSelectOperdrpdwn.sendKeys(Keys.TAB);

			Thread.sleep(5000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersCompareWithdrpdwn));
			triggersCompareWithdrpdwn.click();

			/*triggersCompareWithdrpdwn.sendKeys(Keys.ARROW_DOWN);
			Thread.sleep(2000);*/
			triggersCompareWithdrpdwn.sendKeys("value");
			Thread.sleep(3000);
			triggersCompareWithdrpdwn.sendKeys(Keys.TAB);

			Thread.sleep(1000);

			//triggersValueToEnterTxt.click();
			//triggersValueTxt.click();
			triggersValueToEnterTxt.sendKeys("Bank");	
			Thread.sleep(2000);
			triggersValueToEnterTxt.sendKeys(Keys.TAB);
			Thread.sleep(2000);


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersRiseANewDocChkbox));
			triggersRiseANewDocChkbox.click();

			
		/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabActionBtn));
			triggerTabActionBtn.click();
			
			Thread.sleep(1000);*/
			
			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabAlertBtn));
			triggerTabAlertBtn.click();*/
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
			boolean actdata=editLayoutTriggersAllVouchers.isDisplayed();
			boolean expdata=true;

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersAllVouchers));
			editLayoutTriggersAllVouchers.click();
			editLayoutTriggersAllVouchers.sendKeys("Job Orders");
			editLayoutTriggersAllVouchers.sendKeys(Keys.ENTER);


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
			editLayoutTriggersSaveBtn.click();
			Thread.sleep(4000);

			String expSaveMessage = "Data saved successfully";
			String actSaveMessage = checkValidationMessage(expSaveMessage);
			


			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();


			String expUpdateMessage = "Data saved successfully";
			String actUpdateMessage = checkValidationMessage(expUpdateMessage);
			
			
			
			if(actdata==expdata && actSaveMessage.equalsIgnoreCase(expSaveMessage))
			{
				System.out.println("Test Pass : Trigger  Added Successfully");
			
				return true;
			}
			else
			{
				System.out.println("Test Fail : Trigger NOT Added Successfully");
				
				return false;
			}
		}
		
		
		public boolean 	checkEditingInTriggersTabUnderReceiptsNVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			Thread.sleep(4000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
			editLayoutTriggerName.click();

			editLayoutTriggerName.sendKeys(Keys.SPACE);
		
			
			editLayoutTriggerName.sendKeys(Keys.TAB);

			Thread.sleep(4000);
			
			
			
		
			
			triggersValueTxtGetValue.click();

			Thread.sleep(2000);
			triggersValueToEnterTxt.click();
			triggersValueToEnterTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			

			triggersValueToEnterTxt.sendKeys("Cash");	

			Thread.sleep(2000);

			triggersValueToEnterTxt.sendKeys(Keys.TAB);

			Thread.sleep(10000);
			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersConjuctionDrpdwn));
			Select triggersConjuctionDrpdwnSelect = new Select(triggersConjuctionDrpdwn);
			String acttriggersConjuctionDrpdwn=triggersConjuctionDrpdwnSelect.getFirstSelectedOption().getText();
			String exptriggersConjuctionDrpdwn="Where";

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
			String acttriggersSelectFieldTxt= triggersSelectFieldTxt.getAttribute("value");
			String exptriggersSelectFieldTxt="Name";

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectOperdrpdwn));
			Select triggersSelectOperdrpdwnSelect = new Select(triggersSelectOperdrpdwn);
			String acttriggersSelectOperdrpdwn=triggersSelectOperdrpdwnSelect.getFirstSelectedOption().getText();
			String exptriggersSelectOperdrpdwn="Equal to";

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersCompareWithdrpdwn));
			Select triggersCompareWithdrpdwnSelect = new Select(triggersCompareWithdrpdwn);
			String acttriggersCompareWithdrpdwn=triggersCompareWithdrpdwnSelect.getFirstSelectedOption().getText();
			String exptriggersCompareWithdrpdwn="value";

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
			String acttriggersValueToEnterTxt= triggersValueTxtGetValue.getAttribute("value");
			String exptriggersValueToEnterTxt="Cash";
			
			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabAlertBtn));
			triggerTabAlertBtn.click();*/
			
			Thread.sleep(1000);
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
			boolean actdata=editLayoutTriggersAllVouchers.isDisplayed();
			boolean expdata=true;

			System.out.println("triggersConjuctionDrpdwn  : "+acttriggersConjuctionDrpdwn  +" Value Expected :"+exptriggersConjuctionDrpdwn);
			System.out.println("triggersSelectFieldTxt    : "+acttriggersSelectFieldTxt    +" Value Expected :"+exptriggersSelectFieldTxt);
			System.out.println("triggersSelectOperdrpdwn  : "+acttriggersSelectOperdrpdwn  +" Value Expected :"+exptriggersSelectOperdrpdwn);
			System.out.println("triggersCompareWithdrpdwn : "+acttriggersCompareWithdrpdwn +" Value Expected :"+exptriggersCompareWithdrpdwn);
			System.out.println("triggersValueToEnterTxt   : "+acttriggersValueToEnterTxt   +" Value Expected :"+exptriggersValueToEnterTxt);

			getAction().moveToElement(editLayoutTriggersSaveBtn).build().perform();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
			editLayoutTriggersSaveBtn.click();
			Thread.sleep(2000);
			String expSaveMessage = "Data saved successfully";
			String actSaveMessage = checkValidationMessage(expSaveMessage);
					

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			Thread.sleep(2000);
			String expUpdateMessage = "Data saved successfully";
			String actUpdateMessage = checkValidationMessage(expUpdateMessage);
			

			if(actdata==expdata && actSaveMessage.equalsIgnoreCase(expSaveMessage) && 
					acttriggersCompareWithdrpdwn.equalsIgnoreCase(exptriggersCompareWithdrpdwn) && 
					acttriggersConjuctionDrpdwn.equalsIgnoreCase(exptriggersConjuctionDrpdwn) && 
					acttriggersSelectFieldTxt.equalsIgnoreCase(exptriggersSelectFieldTxt) && 
					acttriggersSelectOperdrpdwn.equalsIgnoreCase(exptriggersSelectOperdrpdwn) && 
					acttriggersValueToEnterTxt.equalsIgnoreCase(exptriggersValueToEnterTxt)
					&& actUpdateMessage.equalsIgnoreCase(expUpdateMessage))
			{
				System.out.println("Test Pass : Trigger  Edited Successfully");
				
				return true;
			}
			else
			{
				System.out.println("Test Fail : Trigger NOT Edited Successfully");
			
				return false;
			}
		}


		public boolean checkSavingViewinReceiptsNVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
		Thread.sleep(20000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsBtn));
		viewsBtn.click();

		Thread.sleep(2000);
		
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewExistingViewTxt));
		ViewExistingViewTxt.click();
		ViewExistingViewTxt.sendKeys(Keys.END);
		ViewExistingViewTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		ViewExistingViewTxt.sendKeys(Keys.SPACE);
		ViewExistingViewTxt.sendKeys("Receipts View");
		Thread.sleep(4000);
		ViewExistingViewTxt.sendKeys(Keys.TAB);

		Thread.sleep(8000);
		
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)","");
		Thread.sleep(2000);
		
		
		
		
		if(viewsUserAllOptionsSTChkboxSelected.isDisplayed()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
		viewTabLoginBtn.click();
		}
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserAllOptionsSTChkbox));
		viewsUserAllOptionsSTChkbox.click();

		Thread.sleep(2000);
		
		getAction().moveToElement(viewSaveView).build().perform();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSaveView));
		viewSaveView.click();
		

		int Count = viewsGridFiledsEditList.size();

		System.out.println("Count    : "+Count);

		ArrayList<String> actviewsGridFiledsList  = new ArrayList<String>();

		for (int i = 0; i < Count; i++) 

		{
			Thread.sleep(1000);
			String data		  = viewsGridFiledsEditList.get(i).getText();

			if(data.equalsIgnoreCase("Date"))

			{
				viewsGridFiledsEditList.get(i).click();

				break;
			}
		}


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewGridRow2EditBtn));
		getAction().doubleClick(viewGridRow2EditBtn).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldDetails_DataTypeDrpdwn));			
		fieldDetails_DataTypeDrpdwn.click();

		Select data=new Select(fieldDetails_DataTypeDrpdwn);
		data.selectByVisibleText("Read Only");

		String actDataType=data.getFirstSelectedOption().getText();
		String expDataType="Read Only";

		System.err.println("FiledCaption For date In Create View  : "+actDataType);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldDetails_ApplyBtn));
		fieldDetails_ApplyBtn.click();

		Thread.sleep(2000);
		
		if(fieldDetails_CloseBtn.isDisplayed()==true)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldDetails_CloseBtn));
			fieldDetails_CloseBtn.click();
		}
		
		

		/*((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)","");
		Thread.sleep(2000);
		
		
		
		
		if(viewsUserAllOptionsSTChkboxSelected.isDisplayed()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
		viewTabLoginBtn.click();
		}
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserAllOptionsSTChkbox));
		viewsUserAllOptionsSTChkbox.click();
*/
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSaveView));
		viewSaveView.click();

		String expMessage= "Data saved successfully";

		String actMessage=checkValidationMessage(expMessage);

		

		if (actMessage.equalsIgnoreCase(expMessage) && actDataType.equalsIgnoreCase(expDataType)) 
		{ 
			System.out.println("Test Pass : Toogle Options in Sales Orders As Expected");
			
			return true;
		}
		else
		{
			System.out.println("Test Fail : Toogle Options in Sales Orders As Expected");
			
			return false;
		}
	}


		public boolean checkNavigateToDocumentNumberingTabInPurchaseVoucherVatView() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{	
			
			getDriver().navigate().refresh();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
			searchTxt.click();
			searchTxt.sendKeys("Configure Transactions");
			searchTxt.sendKeys(Keys.ENTER);
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receiptsBtn));
			receiptsBtn.click();
			Thread.sleep(2000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receiptsNVATBtn));
			receiptsNVATBtn.click();
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberingBtn));
			documentNumberingBtn.click();

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow1));
			typeRow1.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
			Select typeDropDownSelect = new Select(typeDropDown);

			typeDropDownSelect.selectByVisibleText("Login");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow1));
			fieldValueRow1.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueDropDown));
			Select valueDropDownSelect= new Select(fieldValueDropDown);

			valueDropDownSelect.selectByVisibleText("Login Name");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingCharacterRow1));
			startingCharacterRow1.click();

			startingCharacterTxt.sendKeys(Keys.BACK_SPACE);
			
			startingCharacterTxt.sendKeys("0");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow1));
			noofCharactersRow1.click();

			noOfCharacterTxt.sendKeys("2");

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow2));
			typeRow2.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
			typeDropDownSelect.selectByVisibleText("Character Input");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow2));
			fieldValueRow2.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueDropDown));
			Select fieldValueDropDownSelect = new Select(fieldValueDropDown);

			fieldValueDropDownSelect.selectByVisibleText("/");


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow2));
			noofCharactersRow2.click();

			noOfCharacterTxt.sendKeys("1");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow3));
			typeRow3.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
			typeDropDownSelect.selectByVisibleText("Department");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow3));
			fieldValueRow3.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueDropDown));
			fieldValueDropDownSelect.selectByVisibleText("Name");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingCharacterRow3));
			startingCharacterRow3.click();

			startingCharacterTxt.sendKeys(Keys.BACK_SPACE);
			startingCharacterTxt.sendKeys("0");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow3));
			noofCharactersRow3.click();

			noOfCharacterTxt.sendKeys("3");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow4));
			typeRow4.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
			typeDropDownSelect.selectByVisibleText("Character Input");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow4));
			fieldValueRow4.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueDropDown));
			fieldValueDropDownSelect.selectByVisibleText("/");


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow4));
			noofCharactersRow4.click();

			noOfCharacterTxt.sendKeys("1");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow5));
			typeRow5.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
			typeDropDownSelect.selectByVisibleText("Input");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow5));
			fieldValueRow5.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueTxt));
			fieldValueTxt.sendKeys("Text");


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow5));
			noofCharactersRow5.click();

			noOfCharacterTxt.sendKeys("1");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow6));
			typeRow6.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
			typeDropDownSelect.selectByVisibleText("Increment");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow6));
			fieldValueRow6.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueTxt));
			fieldValueTxt.sendKeys("0");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingCharacterRow6));
			startingCharacterRow6.click();

			startingCharacterTxt.sendKeys("0");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow6));
			noofCharactersRow6.click();

			noOfCharacterTxt.sendKeys("1");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow7));
			typeRow7.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docNumberAppearAsValue));

			String actdocNumberAppearAsValue = docNumberAppearAsValue.getText();

			String expdocNumberAppearAsValue = "SU/Tag/Text0";

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();

			String expSaveMessage = "Data saved successfully";

			String actSaveMessage = checkValidationMessage(expSaveMessage);

			


			System.out.println("docNumberAppearAsValue Display Value Actual      : " + actdocNumberAppearAsValue);
			System.out.println("docNumberAppearAsValue Display Value Expected    : " + expdocNumberAppearAsValue);

			if(actdocNumberAppearAsValue.equalsIgnoreCase(expdocNumberAppearAsValue) && actSaveMessage.equalsIgnoreCase(expSaveMessage)) 
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

		        String compname = "Automation Company : 08/10/2020";

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
		        String getLoginCompanyName1=getCompanyTxt1.substring(0, 31);
		        System.out.println("company name  :  "+ getLoginCompanyName1);
		       // companyLogo.click();

		        String expUserInfo1           ="SU";
		        String expLoginCompanyName1   ="Automation Company : 08/10/2020";

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


		public boolean checkLogoutDocuCustReceiptsPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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



		
		
 
	 
	public DocumentCuatomizationReceiptsPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}


	
	

}
