package com.focus.testcases;

import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.SmokeBudgetPage;
import com.focus.base.BaseEngine;


public class SmokeBudgetTest extends BaseEngine
{
	
	static SmokeBudgetPage sbp;

    //@Test(priority=100000)
  	public void checkSignInToCheckBudgets() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkSignInToCheckBudgets(), true);
  	}
	
    
    
    
    @Test(priority=100001)
  	public void checkBudgetsMenu() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkBudgetsMenu(), true);
  	}
    
    
    
    
    // DEFINE BUDGET PAGE
    
    @Test(priority=100002)
  	public void checkDefineBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkDefineBudget(), true);
  	}
    
    @Test(priority=100003)
  	public void checkSaveConfirmPlanBudgetDefineBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkSaveConfirmPlanBudgetDefineBudget(), true);
  	}
    
    @Test(priority=100005)
  	public void checkSaveEditPlanBudgetDefineBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkSaveEditPlanBudgetDefineBudget(), true);
  	}
    
    @Test(priority=100006)
  	public void checkSaveToDeletePlanBudgetDefineBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkSaveToDeletePlanBudgetDefineBudget(), true);
  	}
    
    @Test(priority=100007)
  	public void checkCloseButtonInDefineBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkCloseButtonInDefineBudget(), true);
  	}
    
    
    
    //CONFIRM BUDGET PAGE
    
    @Test(priority=100008)
  	public void checkConfirmBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkConfirmBudget(), true);
  	}
    
    @Test(priority=100009)
  	public void checkConfirmBudgetValueIsDisplay() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkConfirmBudgetValueIsDisplay(), true);
  	}
    
    @Test(priority=100010)
  	public void checkNonConfirmBudgetValueIsDisplay() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkNonConfirmBudgetValueIsDisplay(), true);
  	}
    
    @Test(priority=100011)
  	public void checkCloseButtonInConfirmBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkCloseButtonInConfirmBudget(), true);
  	}
    
    
    // APPEND CONFIRM BUDGET
    
    @Test(priority=100012)
  	public void checkReviseAppendBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkReviseAppendBudget(), true);
  	}
    
    @Test(priority=100013)
  	public void checkSaveConfirmPlanBudgetAppendBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkSaveConfirmPlanBudgetAppendBudget(), true);
  	}
    
    @Test(priority=100014)
  	public void checkCloseButtonInBudgetAppend() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkCloseButtonInBudgetAppend(), true);
  	}
    
    
    //ADD OR REDUCE BUDGET
    
    @Test(priority=100015)
  	public void checkReviseADDOrREDUCEBudgetMenu() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkReviseADDOrREDUCEBudgetMenu(), true);
  	}
    
    @Test(priority=100016)
  	public void checkConfirmPlanBudgetAddOrReduceBudgetLoadingValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkConfirmPlanBudgetAddOrReduceBudgetLoadingValues(), true);
  	}
      
    @Test(priority=100017)
  	public void checkSaveConfirmPlanBudgetAddOrReduce() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkSaveConfirmPlanBudgetAddOrReduce(), true);
  	}
    
    
    @Test(priority=100018)
  	public void checkConfirmPlanBudgetAddOrReduceBudgetUpdateValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkConfirmPlanBudgetAddOrReduceBudgetUpdateValues(), true);
  	}

    @Test(priority=100019)
  	public void checkCloseButtonInBudgetADDORREDUCEValue() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkCloseButtonInBudgetADDORREDUCEValue(), true);
  	}
  	
    
    
    
   // TRANSFER BUDGET
    @Test(priority=100020)
  	public void checkReviseMenuTransferBudgets() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkReviseMenuTransferBudgets(), true);
  	}
      
    @Test(priority=100021)
  	public void checkTransferBudgerFromOptionsValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkTransferBudgerFromOptionsValues(), true);
  	}
       
    @Test(priority=100022)
  	public void checkAddToGridOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkAddToGridOptions(), true);
  	}
      
    @Test(priority=100023)
  	public void checkCloseButtonInBudgetTransfer() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbp=new SmokeBudgetPage(getDriver());
  		Assert.assertEquals(sbp.checkCloseButtonInBudgetTransfer(), true);
  	}
    
    
    
    // BUDGET AUTHORIZATION
    
    @Test(priority=100024)
	public void checkOpenBudgetAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkOpenBudgetAuthorization(), true);	
	}
	
	
	
	@Test(priority=100025)
	public void checkBudgetDropdownOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());	
		Assert.assertEquals(sbp.checkBudgetDropdownOptions(), true);
	}
	
	
	
	@Test(priority=100026)
	public void checkSaveWithoutSelectionOfBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkSaveWithoutSelectionOfBudget(), true);
	}
	

	@Test(priority=100027)
	public void checkSelectingDefineBudgetinBudgetDropdown() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkSelectingDefineBudgetinBudgetDropdown(), true);
	}
	

	@Test(priority=100028)
	public void checkSaveWithoutName() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkSaveWithoutName(), true);
	}
	
	@Test(priority=100029)
	public void checkActiveCheckBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkActiveCheckBox(), true);
	}
	
	
	@Test(priority=100030)
	public void checkInputNameAndClickOnSaveWithoutAuthorizationSettings() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkInputNameAndClickOnSaveWithoutAuthorizationSettings(), true);
	}
	
	
	@Test(priority=100031)
	public void checkStartPoint() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkStartPoint(), true);
	}
	
	@Test(priority=100032)
	public void checkConditionTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkConditionTabOptions(), true);
	}
	
	@Test(priority=100033)
	public void checkUserSelectionTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkUserSelectionTabOptions(), true);
	}
	
	
	@Test(priority=100034)
	public void checkUserDropdownOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkUserDropdownOptions(), true);
	}
	
	
	@Test(priority=100035)
	public void checkAlertsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkAlertsTabOptions(), true);
	}
	
	
	
	@Test(priority=100036)
	public void checkCancelButtonInDefinitionScreen() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkCancelButtonInDefinitionScreen(), true);
	}
	
	
	@Test(priority=100037)
	public void checkDescriptionAsMandatory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkDescriptionAsMandatory(), true);
	}
	
	
	
	@Test(priority=100038)
	public void checkConditionAsMandatory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkConditionAsMandatory(), true);
	}
	
	
	@Test(priority=100039)
	public void checkUserSelectionIsManadatory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkUserSelectionIsManadatory(), true);
	}
	
	
	@Test(priority=100040)
	public void checkOkButtonOnSelectingUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkOkButtonOnSelectingUser(), true);
	}
	
	
	@Test(priority=100041)
	public void checkSavingDefineBudgetAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkSavingDefineBudgetAuthorization(), true);
	}
	
	
	@Test(priority=100042)
	public void checkCancelButton() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkCancelButton(), true);
	}
	
	
    
	
	@Test(priority=100043)
	public void checkDisplayingMastersNameInListOfMasters() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkDisplayingMastersNameInListOfMasters(), true);
	}
	
	
	@Test(priority=100044)
	public void checkSavedNameDisplayInNameTextBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkSavedNameDisplayInNameTextBox(), true);
	}

	
	@Test(priority=100045)
	public void checkEditBudgetAuthorizationSettings() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkEditBudgetAuthorizationSettings(), true);
	}
	
	//@Test(priority=100046) (NOT REUIRED CONDITION AS OF NOW IN BUDGET SO COMMENT IT)
	public void checkConditionCheckbox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkConditionCheckbox(), true);
	}
	
	
	@Test(priority=100047)
	public void checkEditOnSelectAlertsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkEditOnSelectAlertsTabOptions(), true);
	}
	
	
	@Test(priority=100048)
	public void checkSaveWithAllMandatoryInputs() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkSaveWithAllMandatoryInputs(), true);
	}
	
	
	@Test(priority=100049)
	public void checkCreatedEmailInEmailHedaerTemplate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkCreatedEmailInEmailHedaerTemplate(), true);
	}
	
	
	@Test(priority=100050)
	public void checkEmailTemplateUserExpandInFieldsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkEmailTemplateUserExpandInFieldsTabOptions(), true);
	}
	

	@Test(priority=100051)
	public void checkFieldsTabBudgetFieldsOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkFieldsTabBudgetFieldsOptions(), true);
	}
	
	@Test(priority=100052)
	public void checkEmailTemplateCompanyExpandInFieldsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkEmailTemplateCompanyExpandInFieldsTabOptions(), true);
	}
	
	
	@Test(priority=100053)
	public void checkFieldsTabMiscellaneousOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkFieldsTabMiscellaneousOptions(), true);
	} 	
	
	
	@Test(priority=100054)
	public void checkDesignAreaPrintScreen() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkDesignAreaPrintScreen(), true);
	}
	
	
	@Test(priority=100055)
	public void checkFieldsDragInHeaderOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkFieldsDragInHeaderOptions(), true);
	}
	
	
	
	@Test(priority=100056)
	public void checkExitOptionInLayout() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkExitOptionInLayout(), true);
	}
	
	
	
	@Test(priority=100057)
	public void checkSavedEmailTemplateInEmailTemplate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		sbp=new SmokeBudgetPage(getDriver());		
		Assert.assertEquals(sbp.checkSavedEmailTemplateInEmailTemplate(), true);
	}
	
	
	
	 
	 @Test(priority=100058)
	 public void checkSavedAlertTemplateInTemplate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkSavedAlertTemplateInTemplate(), true);
			
	 }
	
	 
	@Test(priority=100059)
	 public void checkSaveButtonWithCondition() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkSaveButtonWithCondition(), true);
			
	 }
	 
	 
	 @Test(priority=100060)
	 public void checkSaveAuthorizationWithConditionAndAlerts() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkSaveAuthorizationWithConditionAndAlerts(), true);
			
	 }
	 
	 
	@Test(priority=100061)
	 public void checkClearButton() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkClearButton(), true);
			
	 }
	 
	 
	 @Test(priority=100062)
	 public void checkAlertMessageOnDeleteAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkAlertMessageOnDeleteAuthorization(), true);
			
	 }
	 
	 
	 @Test(priority=100063)
	 public void checkCancelInAlertDeleteOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkCancelInAlertDeleteOption(), true);
			
	 }
	 
	 
	 @Test(priority=100064)
	 public void checkOkInAlertDeleteOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkOkInAlertDeleteOption(), true);
			
	 }
	 
	 
	 
	 @Test(priority=100065)
	 public void checkSplitRowOneAndRowTwoInDefineBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkSplitRowOneAndRowTwoInDefineBudget(), true);	
	 }
	
	 @Test(priority=100066)
	 public void checSavedBudgetInfoThroughSplitOptionDefineBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checSavedBudgetInfoThroughSplitOptionDefineBudget(), true);	
	 }
	 
	 @Test(priority=100067)
	 public void checkAlertsInWrapperLayoutAuthorizeUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkAlertsInWrapperLayoutAuthorizeUser(), true);	
	 }
	 	 
	 @Test(priority=100068)
	 public void checkClickOnAlertsInWrapperLayoutAuthorizeUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkClickOnAlertsInWrapperLayoutAuthorizeUser(), true);	
	 }
	 
	 
     @Test(priority=100069)
	 public void checkPendingBudgetInAuthorizeUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkPendingBudgetInAuthorizeUser(), true);	
	 }
	 
	 @Test(priority=100070)
	 public void checkCloseButtonInAuthorizeBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkCloseButtonInAuthorizeBudget(), true);	
	 }
	 
	 
	 @Test(priority=100071)
	 public void checkReviseBudgetDetails() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkReviseBudgetDetails(), true);	
	 }
	 	  
	 @Test(priority=100072)
	 public void checkAddReduceReviseBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkAddReduceReviseBudget(), true);	
	 }
	 
	 @Test(priority=100073)
	 public void checkViewOptionInReviseBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkViewOptionInReviseBudget(), true);	
	 }
	
	// @Test(priority=100074)
	 public void checkTransferReviseBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkTransferReviseBudget(), true);	
	 }
	 	 
	 @Test(priority=100075)
	 public void checkCloseButtonInReviseBudget() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			sbp=new SmokeBudgetPage(getDriver());		
			Assert.assertEquals(sbp.checkCloseButtonInReviseBudget(), true);	
	 }
}
