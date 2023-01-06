package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.DocumentCustomizationPage;
import com.focus.base.BaseEngine;

public class DocumentCustomizationTest extends BaseEngine
{
	DocumentCustomizationPage DCP;
	
	
	 @Test(priority=177)
	 public void checkSavingDocumentCustomizationPurchaseVoucherDocumentsTabAndVerify() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkSavingDocumentCustomizationPurchaseVoucherDocumentsTabAndVerify(), true);
	 }
	 
	 @Test(priority=178)
	 public void checkEditingPurchaseVoucherDocumentAndUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkEditingPurchaseVoucherDocumentAndUpdate(), true);
	 }
	 
	 @Test(priority=179)
	 public void checkUpdatingMiscellaneousTabInPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkUpdatingMiscellaneousTabInPurchaseVoucher(), true);
	 }
	
	 @Test(priority=180)
	 public void checkEditingMiscellaneousTabInPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkEditingMiscellaneousTabInPurchaseVoucher(), true);
	 }
	 
	
	 
	 @Test(priority=181)
	 public void checkAddFieldInEditScreenBodyInTrialPurchaseUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkAddFieldInEditScreenBodyInTrialPurchaseUpdate(), true);
	 }
	 
	 @Test(priority=182)
	 public void checkEditFieldInEditScreenBodyInTrialPurchaseUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkEditFieldInEditScreenBodyInTrialPurchaseUpdate(), true);
	 }
	 
	 @Test(priority=183)
	 public void checkDeleteFieldInEditScreenBodyInTrialPurchaseUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkDeleteFieldInEditScreenBodyInTrialPurchaseUpdate(), true);
	 }
	 
	
	 
	 @Test(priority=184)//issuse
	 public void checkAddFieldInEditScreenFooterInTrialPurchaseUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkAddFieldInEditScreenFooterInTrialPurchaseUpdate(), true);
	 }
	 
	 @Test(priority=185)
	 public void checkEditFieldInEditScreenFooterInTrialPurchaseUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkEditFieldInEditScreenFooterInTrialPurchaseUpdate(), true);
	 }
	 
	 @Test(priority=186)
	 public void checkDeleteFieldInEditScreenFooterInTrialPurchaseUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkDeleteFieldInEditScreenFooterInTrialPurchaseUpdate(), true);
	 }
	 
	
	
	 @Test(priority=187)
	 public void checkSelectingApplyRateSchemesInInventoryOptionsTabInTrialPurchaseUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkSelectingApplyRateSchemesInInventoryOptionsTabInTrialPurchaseUpdate(), true);
	 }
	 
	 @Test(priority=188)
	 public void checkUnSelectingApplyRateSchemesInInventoryOptionsTabInTrialPurchaseUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkUnSelectingApplyRateSchemesInInventoryOptionsTabInTrialPurchaseUpdate(), true);
	 }
	 
	 
	 
	 @Test(priority=189)
	 public void checkAddFieldInEditLayoutHeaderInTrialPurchaseUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkAddFieldInEditLayoutHeaderInTrialPurchaseUpdate(), true);
	 }
	 
	 @Test(priority=190)
	 public void checkEditFieldInEditLayoutHeaderInTrialPurchaseUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkEditFieldInEditLayoutHeaderInTrialPurchaseUpdate(), true);
	 }
	 
	 @Test(priority=191)
	 public void checkDeleteFieldInEditLayoutHeaderInTrialPurchaseUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkDeleteFieldInEditLayoutHeaderInTrialPurchaseUpdate(), true);
	 }
	 
	 
	
	 @Test(priority=192)
	 public void checkAddFieldInEditLayoutBodyInTrialPurchaseUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkAddFieldInEditLayoutBodyInTrialPurchaseUpdate(), true);
	 }
	 
	 @Test(priority=193)
	 public void checkEditFieldInEditLayoutBodyInTrialPurchaseUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkEditFieldInEditLayoutBodyInTrialPurchaseUpdate(), true);
	 }
	 
	 @Test(priority=194)
	 public void checkDeleteFieldInEditLayoutBodyInTrialPurchaseUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkDeleteFieldInEditLayoutBodyInTrialPurchaseUpdate(), true);
	 }
	 

	 @Test(priority=195)
	 public void checkSavingViewInTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkSavingViewInTrialVoucherUpdate(), true);
	 }
	 
	 @Test(priority=196)
	 public void checkEditingViewInTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkEditingViewInTrialVoucherUpdate(), true);
	 }
	 
	 @Test(priority=197)
	 public void checkDeletingViewInTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkDeletingViewInTrialVoucherUpdate(), true);
	 }
	
	
	 
	 
	 
	 @Test(priority=198)
	 public void checkSavingTriggerInTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkSavingTriggerInTrialVoucherUpdate(), true);
	 }
	 
	 @Test(priority=199)
	 public void checkEditingTriggerInTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkEditingTriggerInTrialVoucherUpdate(), true);
	 }
	 
	 @Test(priority=200)
	 public void checkDeletingTriggerInTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkDeletingTriggerInTrialVoucherUpdate(), true);
	 }
	 
	 
	
	
	 
	@Test(priority=201)
	 public void checkSavingDocumentNumberingInTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkSavingDocumentNumberingInTrialVoucherUpdate(), true);
	 }
	 
	 @Test(priority=202)
	 public void checkEditingDocumentNumberingInTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkEditingDocumentNumberingInTrialVoucherUpdate(), true);
	 }
	 
	 @Test(priority=203)
	 public void checkDeletingDocumentNumberingInTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkDeletingDocumentNumberingInTrialVoucherUpdate(), true);
	 }
	 
	
	
	 
	 @Test(priority=204)
	 public void checkSavingRuleInTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkSavingRuleInTrialVoucherUpdate(), true);
	 }
	 
	 @Test(priority=205)
	 public void checkEditingRuleInTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkEditingRuleInTrialVoucherUpdate(), true);
	 }
	 
	 @Test(priority=206)
	 public void checkDeletingRuleInTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkDeletingRuleInTrialVoucherUpdate(), true);
	 }
	 
	 
	  
	 @Test(priority=207)
	 public void checkDeletingTrialVoucherUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkDeletingTrialVoucherUpdate(), true);
	 }
	 
	 
	 @Test(priority=208)
	 public void checkCreatingNewVoucherPurchaseVoucheVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCP=new DocumentCustomizationPage(getDriver());
		 Assert.assertEquals(DCP.checkCreatingNewVoucherPurchaseVoucheVAT(), true);
	 }
	 
	 
	 
	 
	 
	 
	 
}
