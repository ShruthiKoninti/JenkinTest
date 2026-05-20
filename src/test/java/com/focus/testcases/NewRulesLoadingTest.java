package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.RulesLoadingPage;
import com.focus.Pages.RulesPage;
import com.focus.base.BaseEngine;

public class NewRulesLoadingTest extends BaseEngine {
	
	RulesLoadingPage RLP;
	
	@Test(priority=101)
	public void CheckLoginRulesLoadingPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.CheckLoginRulesLoadingPage(), true);
	}

	@Test(priority=102)
	public void RestoreRuleLoadingwithAccountaryandInventoryTag() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.checkRestoreRuleLoading(), true);
	}
	
	//@Test(priority=103)//Body to Body with extra field//exist in Backup
	public void CreatingRuleConditionwithExtrafieldsinbodywithnumberlistandstringlistdatatype() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.CreatingRuleConditionwithExtrafieldsinbodywithnumberlistandstringlistdatatype(), true);
	}
	
	@Test(priority=104)
	public void CheckPurchaseVoucherwithcreatedRuleConditionnumberlistandstringlistdatatypeinBody() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.CheckPurchaseVoucherwithcreatedRuleConditionnumberlistandstringlistdatatypeinBody(), true);
	}
	
	@Test(priority=105)//Disable condition
	public void CreatingRuleWithDisableConditionInPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.CreatingRuleWithDisableConditionInPurchaseVoucher(), true);
	}
	
	@Test(priority=106)
	public void SavingPurchaseVoucherwithcreatedRuleWithDisableCondition() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.SavingPurchaseVoucherwithcreatedRuleWithDisableCondition(), true);
	}
	
	@Test(priority=107)//Head to head
	public void CreatingRuleWithConditionHeadtoHeadInPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.CheckCreatingRuleWithConditionHeadtoHeadInPurchaseVoucher(), true);
	}
	
	@Test(priority=108)
	public void SavingPurchaseVoucherwithcreatedRuleConditionHeadtoHead() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.SavingPurchaseVoucherwithcreatedRuleConditionHeadtoHead(), true);
	}
	
	
	
	@Test(priority=109)//head to head with extra field
	public void CreatingRuleWithConditionStringandNumberListextrafieldsInPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.CreatingRuleWithConditionStringandNumberListextrafieldsInPurchaseVoucher(), true);
	}
	
	@Test(priority=110)
	public void CheckPurchaseVoucherwithcreatedRuleConditionStringandNumberListextrafields() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.CheckPurchaseVoucherwithcreatedRuleConditionStringandNumberListextrafields(), true);
	}
	
	@Test(priority=111)//Formuale Validation//issue exist
	public void CreatingRuleWithFootervalueValidationWithFORMULAE() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.CreatingRuleWithFootervalueValidationWithFORMULAE(), true);
	}
	
	@Test(priority=112)
	public void CheckPurchaseVoucherwithcreatedRuleWithFootervalueValidationWithFORMULAE() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.CheckPurchaseVoucherwithcreatedRuleWithFootervalueValidationWithFORMULAE(), true);
		 
	}
	
	@Test(priority=113)//Head to footer
	public void CreatingRuleWithConditionPurAccandFootervalInPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.CreatingRuleWithConditionPurAccandFootervalInPurchaseVoucher(), true);
	}
	
	@Test(priority=114)
	public void CheckPurchaseVoucherwithcreatedRuleConditionPurAccandFooterval() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.CheckPurchaseVoucherwithcreatedRuleConditionPurAccandFooterval(), true);
		 
	}
	
	@Test(priority=115)//Body to Body
	public void CreatingRuleWithConditionItemandQuantityValueInPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.CreatingRuleWithConditionItemandQuantityValueInPurchaseVoucher(), true);
	}
	
	@Test(priority=116)
	public void CheckPurchaseVoucherwithcreatedRuleConditionItemandQuantityValue() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.CheckPurchaseVoucherwithcreatedRuleConditionItemandQuantityValue(), true);
	}
	
	@Test(priority=117)//Based on colour
	public void CreatingRuleConditionLoadingwithColours() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.CreatingRuleConditionLoadingwithColours(), true);
	}
	
	@Test(priority=118)
	public void CheckPurchaseVoucherwithcreatedRuleConditionLoadingwithColour() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.CheckPurchaseVoucherwithcreatedRuleConditionLoadingwithColour(), true);
	}
	
	@Test(priority=119)
	public void checkLogoutRulesLoadingPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RLP=new RulesLoadingPage(getDriver());
		 Assert.assertEquals(RLP.checkLogoutRulesLoadingPage(), true);
	}

}
