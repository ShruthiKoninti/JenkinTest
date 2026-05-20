package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.DocumentCuatomizationReceiptsPage;
import com.focus.Pages.DocumentCustomizationPage;
import com.focus.base.BaseEngine;

public class DocumentCuatomizationReceiptsTest extends BaseEngine{
	
	DocumentCuatomizationReceiptsPage DCRP;
	
	
	 @Test(priority=149)
	 public void CheckLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCRP=new DocumentCuatomizationReceiptsPage(getDriver());
		 Assert.assertEquals(DCRP.CheckLogin(), true);
	 }
	
	 @Test(priority=150)
	 public void checkSavingDocumentCustomizationReceiptsVoucherDocumentsTabAndVerify() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 
		 DCRP=new DocumentCuatomizationReceiptsPage(getDriver());
		 Assert.assertEquals(DCRP.checkSavingDocumentCustomizationReceiptsVoucherDocumentsTabAndVerify(), true);
	 }
	 
	 @Test(priority=151)
	 public void checkmiscellaneousTabInDCInReceiptNewVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCRP=new DocumentCuatomizationReceiptsPage(getDriver());
		 Assert.assertEquals(DCRP.checkmiscellaneousTabInDCInReceiptNewVoucher(), true);
	 }
	 
	 @Test(priority=153)
	 public void checktheNarrationinHeaderTabEditLayoutinReceiptsNVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCRP=new DocumentCuatomizationReceiptsPage(getDriver());
		 Assert.assertEquals(DCRP.checktheNarrationinHeaderTabEditLayoutinReceiptsNVAT(), true);
	 }
	 
	 @Test(priority=154)
	 public void checkAddFieldInEditLayoutBodyInReceiptsNVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCRP=new DocumentCuatomizationReceiptsPage(getDriver());
		 Assert.assertEquals(DCRP.checkAddFieldInEditLayoutBodyInReceiptsNVAT(), true);
	 }
	 
	 @Test(priority=155)
	 public void checkEditFieldInEditLayoutBodyInReceiptsNVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCRP=new DocumentCuatomizationReceiptsPage(getDriver());
		 Assert.assertEquals(DCRP.checkEditFieldInEditLayoutBodyInReceiptsNVAT(), true);
	 }
	 
	 @Test(priority=156)
	 public void checkDeleteFieldInEditLayoutBodyInReceiptsNVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCRP=new DocumentCuatomizationReceiptsPage(getDriver());
		 Assert.assertEquals(DCRP.checkDeleteFieldInEditLayoutBodyInReceiptsNVAT(), true);
	 }
	 
	 @Test(priority=157)
	 public void checkCreatingRuleinRuleTAbUnderReceiptNVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCRP=new DocumentCuatomizationReceiptsPage(getDriver());
		 Assert.assertEquals(DCRP.checkCreatingRuleinRuleTAbUnderReceiptNVAT(), true);
	 }
	
	 @Test(priority=158)
	 public void checkSaveTriggerWithRaiseDocumentOptionInTriggers() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCRP=new DocumentCuatomizationReceiptsPage(getDriver());
		 Assert.assertEquals(DCRP.checkSaveTriggerWithRaiseDocumentOptionInTriggers(), true);
	 }
	 
	 @Test(priority=159)
	 public void checkEditingInTriggersTabUnderReceiptsNVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCRP=new DocumentCuatomizationReceiptsPage(getDriver());
		 Assert.assertEquals(DCRP.checkEditingInTriggersTabUnderReceiptsNVAT(), true);
	 }
	 
	 @Test(priority=160)
	 public void checkSavingViewinReceiptsNVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCRP=new DocumentCuatomizationReceiptsPage(getDriver());
		 Assert.assertEquals(DCRP.checkSavingViewinReceiptsNVAT(), true);
	 }
	 
	 @Test(priority=161)
	 public void checkNavigateToDocumentNumberingTabInPurchaseVoucherVatView() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCRP=new DocumentCuatomizationReceiptsPage(getDriver());
		 Assert.assertEquals(DCRP.checkNavigateToDocumentNumberingTabInPurchaseVoucherVatView(), true);
	 }
	 
	@Test(priority=163)
	 public void checkLogoutDocuCustReceiptsPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DCRP=new DocumentCuatomizationReceiptsPage(getDriver());
		 Assert.assertEquals(DCRP.checkLogoutDocuCustReceiptsPage(), true);
	 }
	 
}
