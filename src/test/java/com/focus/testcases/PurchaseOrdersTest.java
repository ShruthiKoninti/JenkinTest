package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.PurchaseOrdersPage;
import com.focus.base.BaseEngine;

public class PurchaseOrdersTest extends BaseEngine
{
	PurchaseOrdersPage PO;
	
	 @Test(priority=250)
	 public void checkPurchasesOrderVoucherToDisplayLinksWithoutAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkPurchasesOrderVoucherToDisplayLinksWithoutAuthorization(), true);
	 }
	 
	 @Test(priority=251)
	 public void checkLogoutAndLoginWithUserAllOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkLogoutAndLoginWithUserAllOptions(), true);
	 }
	 
	 @Test(priority=252)
	 public void checkAuthorizationInMaterialRequisitionWithUserAllOptionsST() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkAuthorizationInMaterialRequisitionWithUserAllOptionsST(), true);
	 }
	 
	 @Test(priority=253)
	 public void checkLogoutAndLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkLogoutAndLoginWithSU(), true);
	 }
	 
	 @Test(priority=254)
	 public void checkPurchasesOrderVoucherToDisplayLinksAfterAuthorizationAndSavePurchaseOrder() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkPurchasesOrderVoucherToDisplayLinksAfterAuthorizationAndSavePurchaseOrder(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=255)
	 public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizePurchaseOrder() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizePurchaseOrder(), true);
	 }
	 
	 @Test(priority=256)
	 public void checkAuthorizationInPurchaseOrderWithUserAllOptionsST() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkAuthorizationInPurchaseOrderWithUserAllOptionsST(), true);
	 }
	
	 
	 @Test(priority=257)
	 public void checkLogoutAndLoginWithSUForPurchasesVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkLogoutAndLoginWithSUForPurchasesVoucherVAT(), true);
	 }
	 
	 @Test(priority=258)
	 public void checkVerifingDetailsOfSavedPurchaseOrdersVouchersInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkVerifingDetailsOfSavedPurchaseOrdersVouchersInHomepage(), true);
	 }
	 
	 
	 
	 @Test(priority=259)
	 public void checkSavingPurchasesOrderVoucher2WithTwoRows() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkSavingPurchasesOrderVoucher2WithTwoRows(), true);
	 }
	 
	 @Test(priority=260)
	 public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizePurchaseOrderVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizePurchaseOrderVoucher2(), true);
	 }
	 
	 @Test(priority=261)
	 public void checkAuthorizationInPurchaseOrderVoucher2WithUserAllOptionsST() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkAuthorizationInPurchaseOrderVoucher2WithUserAllOptionsST(), true);
	 }
	 
	 @Test(priority=262)
	 public void checkLogoutAndLoginWithSUAfterAuthorizingPurchaseOrderVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkLogoutAndLoginWithSUAfterAuthorizingPurchaseOrderVoucher2(), true);
	 }
	
	 @Test(priority=263)
	 public void checkDataInPurchaseVoucherVATVoucher1ConvertedLink() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkDataInPurchaseVoucherVATVoucher1ConvertedLink(), true);
	 }
	
	 @Test(priority=264)
	 public void checkCloseRow2LinkInPurchaseOrderHavingTwoLinkDocuments() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkCloseRow2LinkInPurchaseOrderHavingTwoLinkDocuments(), true);
	 }
	
	 @Test(priority=265)
	 public void checkSavingPurchaseVoucherVATVoucher1WithPurchaseOrderLink() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkSavingPurchaseVoucherVATVoucher1WithPurchaseOrderLink(), true);
	 }
	
	 @Test(priority=266)
	 public void checkLPurchaseOrderPopupInPurchaseVoucherVATVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkLPurchaseOrderPopupInPurchaseVoucherVATVoucher2(), true);
	 }
	 
	 @Test(priority=267)
	 public void checkSuspendingPurchasesOrdersVoucher2FromEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkSuspendingPurchasesOrdersVoucher2FromEntryPage(), true);
	 }
	 
	 @Test(priority=268)
	 public void checkResavingPurchasesOrdersVoucher2AndSuspendingFromHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkResavingPurchasesOrdersVoucher2AndSuspendingFromHomepage(), true);
	 }
	 
	 @Test(priority=269)
	 public void checkDeletingVoucher2PurchaseOrder() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 PO=new PurchaseOrdersPage(getDriver());
		 Assert.assertEquals(PO.checkDeletingVoucher2PurchaseOrder(), true);
	 }
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
}


