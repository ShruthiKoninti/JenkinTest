package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.RulesPage;
import com.focus.base.BaseEngine;

public class RulesTest extends BaseEngine
{

	RulesPage rp;
	
	@Test(priority=99)
	public void CheckNavigateToRulesTAB() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.CheckNavigateToRulesTAB(), true);
	}

	@Test(priority=100)
	public void checkSavingRuleWithOutInput() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkSavingRuleWithOutInput(), true);
	}
	
	
	@Test(priority=101)
	public void checkSavingRulesWithOnLoadOptionEnableInVoucherLevel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkSavingRulesWithOnLoadOptionEnableInVoucherLevel(), true);
	}
	
	@Test(priority=102)
	public void checkSavedRuleWithLoadOptionEnable() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkSavedRuleWithLoadOptionEnable(), true);
	}
	
	@Test(priority=103)
	public void checkLOADRuleOptionAtVoucherLevel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkLOADRuleOptionAtVoucherLevel(), true);
	}
	
	@Test(priority=104)
	public void checkSavingRulesWithOnLeave() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkSavingRulesWithOnLeave(), true);
	}
	
	@Test(priority=105)
	public void checkOnLeaveValiationATVoucherLevelWithDepartmentWales() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkOnLeaveValiationATVoucherLevelWithDepartmentWales(), true);
	}
	
	@Test(priority=106)// Expected When We Rever Back Remains Same Account Name 
	public void checkChangingDepartmentAsAMERICAAndValidateColounmHeadingName() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkChangingDepartmentAsAMERICAAndValidateColounmHeadingName(), true);
	}
	

	@Test(priority=107)
	public void checkBodyHeadingColumnWithCustomerAccount() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkBodyHeadingColumnWithCustomerAccount(), true);
	}
	
	@Test(priority=108)
	public void chekSavingRuleWithOnEnter() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.chekSavingRuleWithOnEnter(), true);
	}
	
	@Test(priority=109)
	public void checkVoucherLevelOnEnterRuleValidation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkVoucherLevelOnEnterRuleValidation(), true);
	}
	
	
	
	@Test(priority=110)
	public void checkEditingAndAddingSavedRuleUnderRecepitsVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkEditingAndAddingSavedRuleUnderRecepitsVAT(), true);
	}

	@Test(priority=115)//amount as 250.00
	public void checkRuleConditionAfterEditingINEnterRule() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkRuleConditionAfterEditingINEnterRule(), true);
	}
	
	@Test(priority=116)//else value missing with pronghorn
	public void checkClickOnHiddingOptionInRulesAndClickOnSave() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkClickOnHiddingOptionInRulesAndClickOnSave(), true);
	}
	
	@Test(priority=117)
	public void checkAfterEnableHiddenOptionAtVoucherLevel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkAfterEnableHiddenOptionAtVoucherLevel(), true);
	}
	
	
	
	
	@Test(priority=118)
	public void chekSavingRuleWithBeforeDelete() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.chekSavingRuleWithBeforeDelete(), true);
	}
	
	
	@Test(priority=119)
	public void checkValidationRuleWithOnDeleteAtVoucherLevel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkValidationRuleWithOnDeleteAtVoucherLevel(), true);
	}
		
	
	@Test(priority=120)
	public void chekSavingRuleWithOnEdit() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.chekSavingRuleWithOnEdit(), true);
	}

	@Test(priority=121)
	public void checkValidationOFEditRuleWithWarnAndProcedueInRecepitsVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkValidationOFEditRuleWithWarnAndProcedueInRecepitsVoucher(), true);
	}

	
	@Test(priority=122)
	public void checkSavingExtraFiledInSalesINvoiceVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkSavingExtraFiledInSalesINvoiceVAT(), true);
	}
	
	@Test(priority=123)
	public void checkSavingRuleWithFormulaUnderSalesINvoiceVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkSavingRuleWithFormulaUnderSalesINvoiceVAT(), true);
	}
	
	
	@Test(priority=126)
	public void checkSavingOpeningStockVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkSavingOpeningStockVoucher(), true);
	}
	
	@Test(priority=127)
	public void checkValidationOFFOrmulaControlAtVoucherLevel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkValidationOFFOrmulaControlAtVoucherLevel(), true);
	}
	
	@Test(priority=128)//For Sql Updated Back Up--->D:\FocusSanityWithSaveInBackground fbak 
	public void chckRestoreBackUpForSQLFunction() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.chckRestoreBackUpForSQLFunction(), true);
	}
	
	@Test(priority=129)
	public void checkSQLFuntionATExcessStockVoucherLevel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkSQLFuntionATExcessStockVoucherLevel(), true);
	}
	
	@Test(priority=130)
	public void checkCreatingViewOptionInPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkCreatingViewOptionInPurchaseVoucher(), true);
	}
	
	@Test(priority=131)
	public void checkViewOptionCreatedInPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkViewOptionCreatedInPurchaseVoucherVAT(), true);
	}
	
	@Test(priority=145)
	public void checkLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkLogin(), true);
	}
	
	@Test(priority=150)
	public void checkITEMWithF5Key() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkITEMWithF5Key(), true);
	}
	
	@Test(priority=151)
	public void checkLoadingFieldsIntoPopOnClickOnRefreshBtn() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkLoadingFieldsIntoPopOnClickOnRefreshBtn(), true);
	}
	
	
	@Test(priority=152)
	public void checkItemValuesintoVoucherLevelOnClickOnOkBtnInSearchPopUp() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkItemValuesintoVoucherLevelOnClickOnOkBtnInSearchPopUp(), true);
	}
	
	@Test(priority=153)
	public void checkAccountTxtWithClickOnF5() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkAccountTxtWithClickOnF5(), true);
	}
	
	
	@Test(priority=154)
	public void checkLoadingAccountValuesFromSearchPopUp() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkLoadingAccountValuesFromSearchPopUp(), true);
	}
	
	@Test(priority=155) // Issue, Should Display Hyderabad.
	public void checkF5KeyWithRMAPopUpInOpeningStocks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkF5KeyWithRMAPopUpInOpeningStocks(), true);
	}
	
	
	@Test(priority=156)
	public void checkCopyAndPasteWithControlOptionInLedgerReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkCopyAndPasteWithControlOptionInLedgerReport(), true);
	}
	

	@Test(priority=157)
	public void checkEnterF5keyInLedgerReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkEnterF5keyInLedgerReport(), true);
	}
	
	@Test(priority=158)///search issuse
	public void checkF5KeyInFilterPopUpScreen() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkF5KeyInFilterPopUpScreen(), true);
	}
	
	
	
	
	
	

	@Test(priority=200)
	public void checkCreatingConditionForRestrictionsWithStopOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkCreatingConditionForRestrictionsWithStopOption(), true);
	}
	
	@Test(priority=201)//document number missing
	public void checkSavingPurchaseVoucherVATWithRateBelowAverageRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkSavingPurchaseVoucherVATWithRateBelowAverageRate(), true);
	}
	
	@Test(priority=202)
	public void checkSavingPurchaseVoucherVATWithRateEqualToAverageRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkSavingPurchaseVoucherVATWithRateEqualToAverageRate(), true);
	}
	
	@Test(priority=203)
	public void checkSavingPurchaseVoucherVATWithRateGreaterThanAverageRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkSavingPurchaseVoucherVATWithRateGreaterThanAverageRate(), true);
	}
	
	@Test(priority=205)
	public void checkCreatingConditionForRestrictionsWithWarnAndAllowOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkCreatingConditionForRestrictionsWithWarnAndAllowOption(), true);
	}
	
	@Test(priority=206)
	public void checkSavingPurchaseVoucherVATWithRateBelowAverageRateForWarnAndAllow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkSavingPurchaseVoucherVATWithRateBelowAverageRateForWarnAndAllow(), true);
	}
	
	@Test(priority=207)
	public void checkSavingPurchaseVoucherVATWithRateEqualToAverageRateWithWarnAndAllow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkSavingPurchaseVoucherVATWithRateEqualToAverageRateWithWarnAndAllow(), true);
	}
	
	@Test(priority=208)
	public void checkSavingPurchaseVoucherVATWithRateGreaterThanAverageRateWithWarnAndAllow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkSavingPurchaseVoucherVATWithRateGreaterThanAverageRateWithWarnAndAllow(), true);
	}
	
	
	
	
	
	@Test(priority=300)
	public void checkAddingHeadeAndFooterInLedgerReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkAddingHeadeAndFooterInLedgerReport(), true);
	}
	
	@Test(priority=301)
	public void checkPrintPDFFileAfterCreatingHeadeAndFooterLayoutInLedgerReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkPrintPDFFileAfterCreatingHeadeAndFooterLayoutInLedgerReport(), true);
	}
	
	@Test(priority=302)
	public void checkStockLedgerReportImportingHeaderAndFooterLayout() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkStockLedgerReportImportingHeaderAndFooterLayout(), true);
	}
	
	@Test(priority=303)
	public void checkPrintPDFFileAfterImportingHeaderAndFooterLayoutInStockLedgerReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		 rp=new RulesPage(getDriver());
		 Assert.assertEquals(rp.checkPrintPDFFileAfterImportingHeaderAndFooterLayoutInStockLedgerReport(), true);
	}
	
	
	
	
	
	
}
