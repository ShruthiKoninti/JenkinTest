package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.SellerPriceBookPage;
import com.focus.base.BaseEngine;

public class SellerPriceBookTest extends BaseEngine
{

	SellerPriceBookPage SPB;
	
	 @Test(priority=169)
	 public void checkSavingSellerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkSavingSellerPriceBook(), true);
	 }
	
	 @Test(priority=170)
	 public void checkUpdatingSellerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkUpdatingSellerPriceBook(), true);
	 }
	
	 @Test(priority=171)
	 public void checkCopyPasteAndSavingSellerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkCopyPasteAndSavingSellerPriceBook(), true);
	 }
	
	 @Test(priority=172)
	 public void checkDeletingSellerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkDeletingSellerPriceBook(), true);
	 }
	 
	 @Test(priority=173)
	 public void checkExportingSellerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkExportingSellerPriceBook(), true);
	 }
	 
	 @Test(priority=174)
	 public void checkImportingSellerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkImportingSellerPriceBook(), true);
	 }
	 
	 
	 
	 
	 //////////////////////////////////////////////////////////////////////////////////////////////////
	 
	 
	 
	 
	  @Test(priority=176)
	 public void  checkSavingItemGroupalongwithItemsinItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkSavingItemGroupalongwithItemsinItemMaster(), true);
	 }
	 
	 @Test(priority=177)
	 public void checkSavingSellerPriceBookforItemGroupCustomerWisePricebook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkSavingSellerPriceBookforItemGroupCustomerWisePricebook(), true);
	 }
	 
		 
	 @Test(priority=178)//issue with pronghorn rate is not loading
	 public void checkSavedSellerPriceBookatVoucherSalesInvoice() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkSavedSellerPriceBookatVoucherSalesInvoice(), true);
	 }
	 
	
	 @Test(priority=179)
	 public void checkDeleteSellerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkDeleteSellerPriceBook(), true);
	 }

	 @Test(priority=180)
	 public void checkCreatingTwodiffBuyerPricebookforSameProductatDiffDates() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkCreatingTwodiffBuyerPricebookforSameProductatDiffDates(), true);
	 }
	 
	  
	 
	 @Test(priority=181)
	 public void checkUpdatedBuyerPriceBooksRateValueatVoucherPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkUpdatedBuyerPriceBooksRateValueatVoucherPurchaseVoucher(), true);
	 }
	
	 
	 @Test(priority=182)
	 public void checkDeleteBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkDeleteBuyerPriceBook(), true);
	 }
	
	 @Test(priority=183)
	 public void checkCreatingBuyerPriceBookwithMaxandMinQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkCreatingBuyerPriceBookwithMaxandMinQuantity(), true);
	 }
	
	 @Test(priority=184)//loading val1 value
	 public void checkPurchaseVoucherRateFieldgivingMinorMaxQuantitygiveninBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkPurchaseVoucherRateFieldgivingMinorMaxQuantitygiveninBuyerPriceBook(), true);
	 }
	
	 @Test(priority=185)
	 public void checkDeleteCreatedBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkDeleteCreatedBuyerPriceBook(), true);
	 }
	 
	 
	 
	 @Test(priority=186)
	 public void checkCreateTwoDiffBuyerPriceBooksatDiffDateswithSameItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkCreateTwoDiffBuyerPriceBooksatDiffDateswithSameItem(), true);
	 }
	
	 @Test(priority=187)//issue with pronghorn
	 public void checkValidatingtheRateFieldatPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB.checkValidatingtheRateFieldatPurchaseVoucher(), true);
	 }
	 

	 @Test(priority=188)
	 public void  checkDeletingtheBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkDeletingtheBuyerPriceBook(), true);
	 }
	 
	 
	 @Test(priority=189)
	 public void  checkCreatingBuyerPriceBookwithVendorandRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkCreatingBuyerPriceBookwithVendorandRate(), true);
	 }
	 
	
	 @Test(priority=190)
	 public void  checkPurchaseVoucherRateFieldforVendorAccount() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkPurchaseVoucherRateFieldforVendorAccount(), true);
	 }
	 
	 
	 @Test(priority=191)
	 public void  checkDeletetheBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkDeletetheBuyerPriceBook(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=192)
	 public void  checkCreatingBuyerPriceBookforMinandMaxQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkCreatingBuyerPriceBookforMinandMaxQuantity(), true);
	 }
	 
	 @Test(priority=193)//issue 
	 public void  checkVerifyingPurchaseVoucherforRateQuantityApplied() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkVerifyingPurchaseVoucherforRateQuantityApplied(), true);
	 }
	 
	 @Test(priority=194)
	 public void  checkDeletetheQuantityBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkDeletetheQuantityBuyerPriceBook(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=195)
	 public void  checkSavingStatesinStateMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkSavingStatesinStateMaster(), true);
	 }
	 
	 
	 
	 @Test(priority=196)
	 public void  checkCreatingBuyerPriceBookforState() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkCreatingBuyerPriceBookforState(), true);
	 }
	 
	 @Test(priority=197)//issue with pronghorn
	 public void  checkPurchaseVoucherforLoadingRateforStateApplied() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkPurchaseVoucherforLoadingRateforStateApplied(), true);
	 }
	 

	 @Test(priority=198)
	 public void  checkDeletetheStateBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkDeletetheStateBuyerPriceBook(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=199)
	 public void  checkCreatingTwoPriceBooksforSameItemwithDiffRates() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkCreatingTwoPriceBooksforSameItemwithDiffRates(), true);
	 }
	 
	
	 @Test(priority=200)
	 public void  checkRateValueatPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkRateValueatPurchaseVoucher(), true);
	 }
	
	 @Test(priority=201)
	 public void  checkUpdateCreatedPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkUpdateCreatedPriceBook(), true);
	 }

	 
	 @Test(priority=202)
	 public void  checkRateFieldatPurchaseVoucherafterUpdatingPricebook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkRateFieldatPurchaseVoucherafterUpdatingPricebook(), true);
	 }
	 
	
	 
	 
	 @Test(priority=203)
	 public void  checkDeleteTwoBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPB=new SellerPriceBookPage(getDriver());
		 Assert.assertEquals(SPB. checkDeleteTwoBuyerPriceBook(), true);
	 }
	 
	 

	 
	 
	 
}
