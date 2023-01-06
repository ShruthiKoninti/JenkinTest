package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.PurchaseVouchersVATPage;
import com.focus.base.BaseEngine;

public class PurchaseVouchersVATTest extends BaseEngine
{
	 PurchaseVouchersVATPage PVV;
	 
	 @Test(priority=257)
	 public void checkRulesInPurchaseVoucherVat() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkRulesInPurchaseVoucherVat(), true);
	 }
	 
	 @Test(priority=258)
	 public void checkSavingPurchaseVoucherVATVoucher1WithPurchaseOrderLink() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkSavingPurchaseVoucherVATVoucher1WithPurchaseOrderLink(), true);
	 }
	 
	 
	 @Test(priority=259)
	 public void checkSavingPurchaseVoucherVATVoucher2WithBatchItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkSavingPurchaseVoucherVATVoucher2WithBatchItem(), true);
	 }
	 
	 @Test(priority=260)
	 public void checkEditingPurchaseVoucherVATVoucher2AddRow2AndUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkEditingPurchaseVoucherVATVoucher2AddRow2AndUpdate(), true);
	 }
	 
	 @Test(priority=261)
	 public void checkVerifyingPurchaseVoucherVATVoucher2SavingVoucher3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkVerifyingPurchaseVoucherVATVoucher2SavingVoucher3(), true);
	 }
	 
	 @Test(priority=262)
	 public void checkVerifyingAndDeletingPurchaseVoucherVATVoucher3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkVerifyingAndDeletingPurchaseVoucherVATVoucher3(), true);
	 }
	 
	 
	 
	 
	 
	 
	 
	 @Test(priority=263)
	 public void checkSavingPurchaseVoucherVATVoucher3WithStockItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkSavingPurchaseVoucherVATVoucher3WithStockItem(), true);
	 }
	 
	 @Test(priority=264)
	 public void checkEditingPurchaseVoucherVATVoucher3AddRow2AndUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkEditingPurchaseVoucherVATVoucher3AddRow2AndUpdate(), true);
	 }
	 
	 @Test(priority=265)
	 public void checkVerifyingPurchaseVoucherVATVoucher3SavingVoucher4() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkVerifyingPurchaseVoucherVATVoucher3SavingVoucher4(), true);
	 }
	 
	 @Test(priority=266)
	 public void checkVerifyingAndDeletingPurchaseVoucherVATVoucher4() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkVerifyingAndDeletingPurchaseVoucherVATVoucher4(), true);
	 }	 
	 
	 
	 
	 
	 
	 
	 
	 @Test(priority=267) // Bin Popup Total to be Allocated issuse
	 public void checkSavingPurchaseVoucherVATVoucher4WithBinItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkSavingPurchaseVoucherVATVoucher4WithBinItem(), true);
	 }	
	 
	 @Test(priority=268)
	 public void checkSavingPurchaseVoucherVATVoucher4With2RowsAndBinItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkSavingPurchaseVoucherVATVoucher4With2RowsAndBinItem(), true);
	 }
	 
	 @Test(priority=269) // Bin Popup Total to be Allocated
	 public void checkEditingPurchaseVoucherVATVoucher4AddRow3AndUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkEditingPurchaseVoucherVATVoucher4AddRow3AndUpdate(), true);
	 }
	 
	 @Test(priority=270)//Expected
	 public void checkVerifyingPurchaseVoucherVATVoucher4SavingVoucher5() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkVerifyingPurchaseVoucherVATVoucher4SavingVoucher5(), true);
	 }
	 
	 @Test(priority=271)
	 public void checkVerifyingAndDeletingPurchaseVoucherVATVoucher5() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkVerifyingAndDeletingPurchaseVoucherVATVoucher5(), true);
	 }
	 
	 
	 
	 
	 
	 @Test(priority=272)
	 public void checkSavingPurchaseVoucherVATVoucher5WithRMAItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkSavingPurchaseVoucherVATVoucher5WithRMAItem(), true);
	 }
	 
	 @Test(priority=273)
	 public void checkEditingPurchaseVoucherVATVoucher5AddRow2AndUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkEditingPurchaseVoucherVATVoucher5AddRow2AndUpdate(), true);
	 }
	 
	 @Test(priority=274)
	 public void checkVerifyingPurchaseVoucherVATVoucher5SavingVoucher6() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkVerifyingPurchaseVoucherVATVoucher5SavingVoucher6(), true);
	 }
	 
	 @Test(priority=275)
	 public void checkVerifyingAndDeletingPurchaseVoucherVATVoucher6() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkVerifyingAndDeletingPurchaseVoucherVATVoucher6(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=276)
	 public void checkSavingAndSuspendingPurchaseVoucherVATVoucher6WithStockItemInEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkSavingAndSuspendingPurchaseVoucherVATVoucher6WithStockItemInEntryPage(), true);
	 }
	 
	 @Test(priority=277)
	 public void checkResavingPurchaseVoucherVAT6AndSuspendingFromHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkResavingPurchaseVoucherVAT6AndSuspendingFromHomepage(), true);
	 }
	 
	 @Test(priority=278)
	 public void checkDeleteOptionInPurchaseVoucherVATHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkDeleteOptionInPurchaseVoucherVATHomePage(), true);
	 }
	 
	 @Test(priority=279)
	 public void checkVerifingDetailsOfSavedPurchaseVoucherVATInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkVerifingDetailsOfSavedPurchaseVoucherVATInHomepage(), true);
	 }
	 
	 
	 @Test(priority=280)
	 public void checkVerifingTriggersOptionInPurchaseVAT() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkVerifingTriggersOptionInPurchaseVAT(), true);
	 }
	 
	 
	 //////Verifying Info button//////////////////
	 
	 @Test(priority=281)
	 public void checkAuditLogInformationOfSavedPurchaseVoucherVATInHomepage() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkAuditLogInformationOfSavedPurchaseVoucherVATInHomepage(), true);
	 }
	 
	 
	 @Test(priority=282)
	 public void checkAuthorizeHistoryInformationOfSavedPurchaseVoucherVATInHomepage() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkAuthorizeHistoryInformationOfSavedPurchaseVoucherVATInHomepage(), true);
	 }
	 
	 @Test(priority=283)
	 public void checkSettingDetailsInformationOfSavedPurchaseVoucherVATInHomepage() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkSettingDetailsInformationOfSavedPurchaseVoucherVATInHomepage(), true);
	 }
	 
	 @Test(priority=284)
	 public void checkVoucherDetailsInformationOfSavedPurchaseVoucherVATInHomepage() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkVoucherDetailsInformationOfSavedPurchaseVoucherVATInHomepage(), true);
	 }
	 
	 
	 
	 
	 
	 @Test(priority=285)
	 public void checkVoucherFooterDetailsatHomePage() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkVoucherFooterDetailsatHomePage(), true);
	 }
	 
	
	 
	 @Test(priority=286)
	 public void checkSearchingVendorandItemsListwithoutAstrickSigninPurchaseVoucherVAT() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkSearchingVendorandItemsListwithoutAstrickSigninPurchaseVoucherVAT(), true);
	 }
	 
	 
	 
	 @Test(priority=287)
	 public void checkSearchingofVendorandItemsListusingAstricksigninPurchaseVoucherVAT() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkSearchingofVendorandItemsListusingAstricksigninPurchaseVoucherVAT(), true);
	 }
	 
	 
	 
	 ////Customization at voucher entry level/////////////////
	 
	 
	 
	 
	 @Test(priority=288)
	 public void checkCreatingEditLayoutHeaderatPurchaseVoucherVAT() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkCreatingEditLayoutHeaderatPurchaseVoucherVAT(), true);
	 }
	 
	 
	 
	 @Test(priority=289)
	 public void checkEdittheCreatedFieldinPurchaseVoucherVAT() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkEdittheCreatedFieldinPurchaseVoucherVAT(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=290)
	 public void checkDeletetheCreatedFieldatPurchaseVoucherVAT() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkDeletetheCreatedFieldatPurchaseVoucherVAT(), true);
	 }
	 
	 
	 
	 @Test(priority=291)
	 public void checkCustomisationScreenBodyatPurchaseVoucherVAT() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkCustomisationScreenBodyatPurchaseVoucherVAT(), true);
	 }
	 
	 
	 @Test(priority=292)
	 public void checkEditCustomisationScreenBodyatPurchaseVoucherVAT() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkEditCustomisationScreenBodyatPurchaseVoucherVAT(), true);
	 }
	 
	 
	 @Test(priority=293)
	 public void checkDeleteAddedFieldinBodyatPurchaseVoucherVAT() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkDeleteAddedFieldinBodyatPurchaseVoucherVAT(), true);
	 }
	 
	 
	 @Test(priority=294)
	 public void checkCustomizationLayoutBodyatPurchaseVoucherVAT() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkCustomizationLayoutBodyatPurchaseVoucherVAT(), true);
	 }
	 
	 @Test(priority=295)
	 public void  checkEditCreatedLayoutBodyatPurchaseVoucherVAT() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV. checkEditCreatedLayoutBodyatPurchaseVoucherVAT(), true);
	 }
	 
	 @Test(priority=296)
	 public void  checkDeleteAddedFieldinLayoutBodyatPurchaseVoucherVAT() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV. checkDeleteAddedFieldinLayoutBodyatPurchaseVoucherVAT(), true);
	 }
	 
	 
	 @Test(priority=297)
	 public void   checkCustomizationFooteratPurchaseVoucherVAT() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV. checkCustomizationFooteratPurchaseVoucherVAT(), true);
	 }
	 
	 
	 @Test(priority=298)
	 public void   checkEditCustomizedFooterFieldatPurchaseVoucherVAT() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV.checkEditCustomizedFooterFieldatPurchaseVoucherVAT(), true);
	 }
	 
	 
	 @Test(priority=299)
	 public void  checkDeleteAddedFieldinFooteratPurchaseVoucherVAT() throws Exception
	 {
		 PVV=new PurchaseVouchersVATPage(getDriver());
		 Assert.assertEquals(PVV. checkDeleteAddedFieldinFooteratPurchaseVoucherVAT(), true);
	 }
	 
	 
	
	
	
	
	 
	 
}
