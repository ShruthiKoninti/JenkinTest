package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage;
import com.focus.Pages.PurchaseVouchersVATPage;
import com.focus.base.BaseEngine;

public class BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsTest extends BaseEngine
{
	BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage BSP;
	
	 
	 // Buyer Pricebook Validations
	
	 @Test(priority=10050)
	 public void checkSavingPurchaseVoucherVATVoucher1WithoutRateFormula() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingPurchaseVoucherVATVoucher1WithoutRateFormula(), true);
	 }
	 
	 @Test(priority=10051)
	 public void checkVerifyingPurchaseVoucherVATVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkVerifyingPurchaseVoucherVATVoucher1(), true);
	 }

	 @Test(priority=10052)
	 public void checkUpdatingBuyerPricebookFormulaForPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkUpdatingBuyerPricebookFormulaForPurchaseVoucherVAT(), true);
	 }
	 
	 @Test(priority=10053)
	 public void checkSavingPurchaseVoucherVATVoucher2WithRateFormula() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingPurchaseVoucherVATVoucher2WithRateFormula(), true);
	 }
	
	 @Test(priority=10054)
	 public void checkVerifyingPurchaseVoucherVATVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkVerifyingPurchaseVoucherVATVoucher2(), true);
	 }

	 @Test(priority=10055)
	 public void checkUpdatingBuyerPricebookFormulaWithAbbrivationForPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkUpdatingBuyerPricebookFormulaWithAbbrivationForPurchaseVoucherVAT(), true);
	 }
	 
	 @Test(priority=10056)
	 public void checkSavingPurchaseVoucherVATVoucher3WithRateFormulaWithAbbrevation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingPurchaseVoucherVATVoucher3WithRateFormulaWithAbbrevation(), true);
	 }
	 
	 @Test(priority=10057)
	 public void checkVerifyingPurchaseVoucherVATVoucher3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkVerifyingPurchaseVoucherVATVoucher3(), true);
	 }
	 
	 @Test(priority=10058)
	 public void checkChangingRatePreloadBehaviourToRatePurchaseVoucherVAT() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkChangingRatePreloadBehaviourToRatePurchaseVoucherVAT(), true);
	 }
	 
	 
	 
	 // Purchase Variance Validations
	
	 @Test(priority=10059)
	 public void checkSavingPurchaseVoucherVATVoucher4WithStockItem() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingPurchaseVoucherVATVoucher4WithStockItem(), true);
	 }
	 
	 @Test(priority=10060)
	 public void checkPostingDetailsWithEqualsToAndUpdatingRateToGreaterThanStandardRate() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkPostingDetailsWithEqualsToAndUpdatingRateToGreaterThanStandardRate(), true);
	 }
	 
	 @Test(priority=10061)
	 public void checkPostingDetailsWithGreaterThanAndUpdatingRateToLessThanStandardRate() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkPostingDetailsWithGreaterThanAndUpdatingRateToLessThanStandardRate(), true);
	 }
	 
	 @Test(priority=10062)
	 public void checkPostingDetailsWithRateLessThanStandardRate() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkPostingDetailsWithRateLessThanStandardRate(), true);
	 }
	 
	 @Test(priority=10063)
	 public void checkSavingPurchaseVoucherVATVoucher5WithStockItem() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingPurchaseVoucherVATVoucher5WithStockItem(), true);
	 }
	 
	 @Test(priority=10064)
	 public void checkPurchaseVoucherVATWithBinItemAutoAllocateInBinPopup() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkPurchaseVoucherVATWithBinItemAutoAllocateInBinPopup(), true);
	 }
	 
	 
	

	  
	
	 // Voucher Print Validations
	
	  @Test(priority=10101)
	  public void checkPrintScreenOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkPrintScreenOptions(), true);
	  }
	  
	  @Test(priority=10102)
	  public void checkLayoutOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkLayoutOptions(), true);
	  }
	  
	  @Test(priority=10103)
	  public void checkInvoiceDesignOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkInvoiceDesignOptions(), true);
	  }
	  
	  @Test(priority=10104)
	  public void checkFieldsTabOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkFieldsTabOptions(), true);
	  }
	
	  @Test(priority=10105)
	  public void checkDragandDropHeaderFiledsAndBodyFiledsIntoLayout() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkDragandDropHeaderFiledsAndBodyFiledsIntoLayout(), true);
	  }
	  
	  @Test(priority=10106)
	  public void checkChangeFieldPropertiesInLyout() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkChangeFieldPropertiesInLyout(), true);
	  }
	  
	  @Test(priority=10107)
	  public void checkSavingPurchasesVatDesignLayout() throws AWTException, Exception
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkSavingPurchasesVatDesignLayout(), true);
	  }
	  
	  @Test(priority=10108)
	  public void checkEditingLayoutByAddingFields() throws AWTException, Exception
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkEditingLayoutByAddingFields(), true);
	  }

	  @Test(priority=10109)
	  public void checkUpdatedLayoutAndDeleteTheLayout() throws AWTException, Exception
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkUpdatedLayoutAndDeleteTheLayout(), true);
	  }
	  
	  
	  
	  
	  @Test(priority=10110)
	  public void checkPrintPdfOfPurchaseVoucherVATWithCreatedLayout() throws AWTException, Exception
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkPrintPdfOfPurchaseVoucherVATWithCreatedLayout(), true);
	  }
	  
	  @Test(priority=10111)
	  public void checkPrintPdfOfPurchaseVoucherVAT() throws Exception
	  {
	 	  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkPrintPdfOfPurchaseVoucherVAT(), true);
	  }
	
	  @Test(priority=10112)
	  public void checkPrintPdfOfPurchaseVoucherVATWithImportedXML() throws Exception
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkPrintPdfOfPurchaseVoucherVATWithImportedXML(), true);
	  }
	
	  @Test(priority=10113)
	  public void checkEmailOptionInPurchaseVoucherVAT() throws Exception
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkEmailOptionInPurchaseVoucherVAT(), true);
	  }
	
  
	
	  // Document Set
	  
	  //@Test(enabled=false)
	  public void checkRestoreCompanyForDocumentSet() throws Exception
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkRestoreCompanyForDocumentSet(), true);
	  }
	
	  
	  @Test(priority=11001)//issue at filter
	  public void checkCreatingDocumentSet() throws Exception
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkCreatingDocumentSet(), true);
	  }
	  
	  @Test(priority=11002)
	  public void checkCreatingDocumentSetForDelete() throws Exception
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkCreatingDocumentSetForDelete(), true);
	  }
	  
	  @Test(priority=11003)
	  public void checkEditingAndUpdatingDocumentSetForDelete() throws Exception
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkEditingAndUpdatingDocumentSetForDelete(), true);
	  }
	  
	  @Test(priority=11004)
	  public void checkDeletingDocumentSetForDelete() throws Exception
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkDeletingDocumentSetForDelete(), true);
	  }
	  
	  @Test(priority=11005)
	  public void checkPrintPdfOfPurchaseVoucherVATWithDocumentSetVendorA() throws Exception
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkPrintPdfOfPurchaseVoucherVATWithDocumentSetVendorA(), true);
	  }
	
	  @Test(priority=11006)
	  public void checkPrintPdfOfPurchaseVoucherVATWithDocumentSetVendorB() throws Exception
	  {
		  BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		  Assert.assertEquals(BSP.checkPrintPdfOfPurchaseVoucherVATWithDocumentSetVendorB(), true);
	  }
	  
	
	
	 
	
	 // Seller Pricebook Validations
	
	 @Test(priority=20010)
	 public void checksalesInvoiceVatWithBuyerSellerPriceBooklist() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checksalesInvoiceVatWithBuyerSellerPriceBooklist(), true);
	 }
	 
	 @Test(priority=20011)
	 public void checkSavedVoucherInSalesINvoiceVAT() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavedVoucherInSalesINvoiceVAT(), true);
	 }
	 
	 @Test(priority=20012)
	 public void checkUpdatingBuyerPricebookFormulaForSalesInvoiceVAT() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkUpdatingBuyerPricebookFormulaForSalesInvoiceVAT(), true);
	 }
	 
	 @Test(priority=20013)
	 public void checkSavingSalesInvoiceVATVoucher2WithRateFormula() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingSalesInvoiceVATVoucher2WithRateFormula(), true);
	 }
	 
	 @Test(priority=20014)
	 public void checkSavedVoucher2InSalesINvoiceVATWithFormulaControl() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavedVoucher2InSalesINvoiceVATWithFormulaControl(), true);
	 }
	 
	 @Test(priority=20015)
	 public void checkUpdatingBuyerPricebookFormulaForSalesInvoiceVATVoucher3() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkUpdatingBuyerPricebookFormulaForSalesInvoiceVATVoucher3(), true);
	 }
	 
	 @Test(priority=20016)
	 public void checkSavingSalesInvoiceVATVoucher3WithRateFormula() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingSalesInvoiceVATVoucher3WithRateFormula(), true);
	 }
	 
	 @Test(priority=20017)
	 public void checkSavedVoucher3InSalesINvoiceVATWithFormulaControl() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavedVoucher3InSalesINvoiceVATWithFormulaControl(), true);
	 }
	 
	 @Test(priority=20018)
	 public void checkChangingRatePreloadBehaviourToRateSalesInvoiceVAT() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkChangingRatePreloadBehaviourToRateSalesInvoiceVAT(), true);
	 }
	 
	 
	 
	 
	 // Credit Limit Validations
	 
	 @Test(priority=20019)
	 public void checkSavingSalesInvoiceVATVoucher4WithRateAsLessThanCreditLimit() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingSalesInvoiceVATVoucher4WithRateAsLessThanCreditLimit(), true);
	 }
	 
	 @Test(priority=20020)
	 public void checkSavedVoucher4AndTryingToUpdateVoucherRateGreaterThanCreditLimit() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavedVoucher4AndTryingToUpdateVoucherRateGreaterThanCreditLimit(), true);
	 }
	 
	 
	 
	 // Voucher Print Validations
	   
	 @Test(priority=20021)
	 public void checkPrintPdfOfSalesInvoiceVAT() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkPrintPdfOfSalesInvoiceVAT(), true);
	 }
	 
	 @Test(priority=20022)
	 public void checkPrintPdfOfSalesInvoiceVATWithImportedXML() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkPrintPdfOfSalesInvoiceVATWithImportedXML(), true);
	 }
	 
	 @Test(priority=20023)
	 public void checkEmailOptionInSalesInvoiceVAT() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkEmailOptionInSalesInvoiceVAT(), true);
	 }
	
	
	 
	 
	 
	 
	
	 @Test(priority=20030)
	 public void checkSavingMaterialReceiptNoteVoucher1WithNoTagsAndValidatePricebookValues() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingMaterialReceiptNoteVoucher1WithNoTagsAndValidatePricebookValues(), true);
	 }
	
	 @Test(priority=20031)
	 public void checkSavingMaterialReceiptNoteVoucher2WithPricebookVendorAccountOneAndValidatePricebookValues() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingMaterialReceiptNoteVoucher2WithPricebookVendorAccountOneAndValidatePricebookValues(), true);
	 }
	
	 @Test(priority=20032)
	 public void checkSavingMaterialReceiptsNoteVoucher3WithPricebookVendorAccountTwoAndValidatePricebookValues() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingMaterialReceiptsNoteVoucher3WithPricebookVendorAccountTwoAndValidatePricebookValues(), true);
	 }
	
	 @Test(priority=20033)
	 public void checkSavingMaterialReceiptNoteVoucher4WithGroupOneAccountOneAndValidatePricebookValues() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingMaterialReceiptNoteVoucher4WithGroupOneAccountOneAndValidatePricebookValues(), true);
	 }
	
	 @Test(priority=20034)
	 public void checkSavingMaterialReceiptNoteVoucher5WithGroupOneAccountTwoAndValidatePricebookValues() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingMaterialReceiptNoteVoucher5WithGroupOneAccountTwoAndValidatePricebookValues(), true);
	 }
	
	 @Test(priority=20035)
	 public void checkSavingMaterialReceiptNoteVoucher6WithGroupTwoAccountOneAndValidatePricebookValues() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingMaterialReceiptNoteVoucher6WithGroupTwoAccountOneAndValidatePricebookValues(), true);
	 }
	
	 @Test(priority=20036)
	 public void checkSavingMaterialReceiptNoteVoucher7WithGroupTwoAccountTwoAndValidatePricebookValues() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingMaterialReceiptNoteVoucher7WithGroupTwoAccountTwoAndValidatePricebookValues(), true);
	 }
	 
	
	
	
	 
	 @Test(enabled=false)
	 public void checkSavingPurchaseVoucherVATVoucherForDebitNoteSellerPricebookValidations() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingPurchaseVoucherVATVoucherForDebitNoteSellerPricebookValidations(), true);
	 }
	
	 @Test(priority=20101)
	 public void checkSavingDeliveryNoteVoucher1WithNoTagsAndValidatePricebookValues() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingDeliveryNoteVoucher1WithNoTagsAndValidatePricebookValues(), true);
	 }
	
	 @Test(priority=20102)
	 public void checkSavingDeliveryNoteVoucher2WithPricebookAccountOneAndValidatePricebookValues() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingDeliveryNoteVoucher2WithPricebookAccountOneAndValidatePricebookValues(), true);
	 }
	 
	 @Test(priority=20103)
	 public void checkSavingDeliveryNoteVoucher3WithPricebookAccountTwoAndValidatePricebookValues() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingDeliveryNoteVoucher3WithPricebookAccountTwoAndValidatePricebookValues(), true);
	 }
	
	 @Test(priority=20104)
	 public void checkSavingDeliveryNoteVoucher4WithGroupOneAccountOneAndValidatePricebookValues() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingDeliveryNoteVoucher4WithGroupOneAccountOneAndValidatePricebookValues(), true);
	 }
	 
	 @Test(priority=20105)
	 public void checkSavingDeliveryNoteVoucher5WithGroupOneAccountTwoAndValidatePricebookValues() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingDeliveryNoteVoucher5WithGroupOneAccountTwoAndValidatePricebookValues(), true);
	 }
	 
	 @Test(priority=20106)
	 public void checkSavingDeliveryNoteVoucher6WithGroupTwoAccountOneAndValidatePricebookValues() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingDeliveryNoteVoucher6WithGroupTwoAccountOneAndValidatePricebookValues(), true);
	 }
	
	 @Test(priority=20107)
	 public void checkSavingDeliveryNoteVoucher7WithGroupTwoAccountTwoAndValidatePricebookValues() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingDeliveryNoteVoucher7WithGroupTwoAccountTwoAndValidatePricebookValues(), true);
	 }
	
	
	
	
	
	
	 @Test(priority=30000)
	 public void checkEraseAllTransactionsAfterCompletingAllValidations() throws Exception
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkEraseAllTransactionsAfterCompletingAllValidations(), true);
	 }
	 
	 
	 
	/* @Test(priority=100)
	 public void checkSavingPurchaseVoucherVATVoucher1WithRMA() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BSP=new BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage(getDriver());
		 Assert.assertEquals(BSP.checkSavingPurchaseVoucherVATVoucher1WithRMA(), true);
	 }
	*/

}
