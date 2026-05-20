package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.BuyerPriceBookValidationPage;
import com.focus.Pages.BuyerPriceBookPage;
import com.focus.Pages.BuyerPriceBookValidationPage;
import com.focus.base.BaseEngine;

public class BuyerPriceBookValidationTest extends BaseEngine{
	
	BuyerPriceBookValidationPage BPBV;
	
	
	
	
	 @Test(priority=1)
	 public void  checkLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPBV=new BuyerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(BPBV. checkLogin(), true);
	 }
	
	 
	
	 
	
	
	 @Test(priority=148)
	 public void  checkRestoreAutomationCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPBV=new BuyerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(BPBV. checkRestoreAutomationCompany(), true);
	 }
	 
	
	  @Test(priority=150)
		 public void  checkBuyerPriceBookwithNoTagRateValueatPurchasesVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV. checkBuyerPriceBookwithNoTagRateValueatPurchasesVoucher(), true);
		 }
	
	  
	  
	  @Test(priority=151)
		 public void  checkRatewithPricebookVendorAccountOneatPurchasesVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV. checkRatewithPricebookVendorAccountOneatPurchasesVoucher(), true);
		 }
	
	  @Test(priority=152)
		 public void  checkRatewithPricebookVendorAccountTwoatPurchasesVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV. checkRatewithPricebookVendorAccountTwoatPurchasesVoucher(), true);
		 }
	  
	  

	  @Test(priority=153)
		 public void  checkRatewithPricebookGroupAccountOneatPurchasesVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkRatewithPricebookGroupAccountOneatPurchasesVoucher(), true);
		 }
	  
	  
	  @Test(priority=154)
		 public void  checkRatewithPricebookGroupAccountTwoatPurchasesVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkRatewithPricebookGroupAccountTwoatPurchasesVoucher(), true);
		 }
	 
	  @Test(priority=155)
		 public void  checkRatewithPricebookGroupAccountOneandGroupItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkRatewithPricebookGroupAccountOneandGroupItem(), true);
		 }
	  
	  @Test(priority=156)
		 public void  checkRatewithPricebookGroupAccountTwoandGroupItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkRatewithPricebookGroupAccountTwoandGroupItem(), true);
		 }
	  
	  
	  
	  @Test(priority=157)
		 public void  checkRateLoadingforPricebookWarehouseGroupinPurchasesVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkRateLoadingforPricebookWarehouseGroupinPurchasesVoucher(), true);
		 }
	  
	  
	  
	  @Test(priority=158)
		 public void  checkRateLoadingforItemhavingMaxandMinQuantityinPurchasesVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkRateLoadingforItemhavingMaxandMinQuantityinPurchasesVoucher(), true);
		 }
	  

	  
	  @Test(priority=159)
		 public void  checkCreatingBuyerPriceBookforDateRange() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkCreatingBuyerPriceBookforDateRange(), true);
		 }
	  
	  
	  
	  
	  @Test(priority=160)
		 public void  checkRateLoadingforItemhavingDateRangeinPurchasesVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkRateLoadingforItemhavingDateRangeinPurchasesVoucher(), true);
		 }
	  
	 
	  
	  
	  @Test(priority=161)
		 public void  checkItemFilterinBuyerPricebook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkItemFilterinBuyerPriceBook(), true);
		 }
	  
	  
	  
	
	  
	  @Test(priority=162)
		 public void  checkVendorFilterinBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkVendorFilterinBuyerPriceBook(), true);
		 }
	  
	  
	  
	  
	  @Test(priority=163)
		 public void  checkCurrencyFilterinBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkCurrencyFilterinBuyerPriceBook(), true);
		 }
	  
	  
	  @Test(priority=164)
		 public void  checkItemandCurrencyFilterinBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkItemandCurrencyFilterinBuyerPriceBook(), true);
		 }
	  
	  @Test(priority=165)
		 public void  checkItemandCustomerFilterinBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkItemandCustomerFilterinBuyerPriceBook(), true);
		 }
	  
	  @Test(priority=166)
		 public void  checkItemCurrencyandCustomerFilterinBuyerPricebook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkItemCurrencyandCustomerFilterinBuyerPriceBook(), true);
		 }
	  
	  
	  
	  @Test(priority=167)
		 public void  
		 
		 checkItemCurrencyCustomerDepartmentFilterinBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkItemCurrencyCustomerDepartmentFilterinBuyerPriceBook(), true);
		 }
	  
	  
	  
	   
	  
	  @Test(priority=168)
		 public void  checkExportFileafterApplyingFilter() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkExportFileafterApplyingFilter(), true);
		 }
	  
	  
	  
	//  @Test(priority=169)
		 public void  checkSavingStatesinStateMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkSavingStatesinStateMaster(), true);
		 }
	  
	  
	  @Test(priority=170)
		 public void  checkCreatingBuyerPriceBookforState() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkCreatingBuyerPriceBookforState(), true);
		 }
	  
	  
	  
	
	  
	@Test(priority=171)
	 public void  checkCreatedMasterRateValueinPurchasesVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPBV=new BuyerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(BPBV.checkCreatedMasterRateValueinPurchasesVoucher(), true);
	 }
	
	
	
	@Test(priority=172)
	 public void  checkAdvExcelImportingBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPBV=new BuyerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(BPBV.checkAdvExcelImportingBuyerPriceBook(), true);
	 }
	
	
	
	
	@Test(priority=173)
	 public void  checkDeleteBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPBV=new BuyerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(BPBV.checkDeleteBuyerPriceBook(), true);
	 }
	
	
	@Test(priority=174)
	 public void  checkExportingXMLBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPBV=new BuyerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(BPBV.checkExportingXMLBuyerPriceBook(), true);
	 }

	@Test(priority=175)
	 public void  checkImportingXMLBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPBV=new BuyerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(BPBV.checkImportingXMLBuyerPriceBook(), true);
	 }
	
	
	
	
	
	

	@Test(priority=176)
	 public void  checkPrintRateHistoryBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPBV=new BuyerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(BPBV.checkPrintRateHistoryBuyerPriceBook(), true);
	 }
	
	
	@Test(priority=177)
	 public void  checkPreviewandClearBtninPriceRateHistory() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPBV=new BuyerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(BPBV.checkPreviewandClearBtninPriceRateHistory(), true);
	 }
	
	@Test(priority=178)
	 public void  checkViewPriceCustomizationBtnBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPBV=new BuyerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(BPBV.checkViewPriceCustomizationBtnBuyerPriceBook(), true);
	 }
	
	
	

	  
	  ///Material Receipts Notes
	 
	  
	  @Test(priority=179)
		 public void  checkSavingMaterialReceiptNoteVoucher1WithNoTagsAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkSavingMaterialReceiptNoteVoucher1WithNoTagsAndValidatePricebookValues(), true);
		 }
	
	  
	  @Test(priority=180)
		 public void  checkSavingMaterialReceiptNoteVoucher2WithPricebookVendorAccountOneAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkSavingMaterialReceiptNoteVoucher2WithPricebookVendorAccountOneAndValidatePricebookValues(), true);
		 }
	  
	  
	  @Test(priority=181)
		 public void  checkSavingMaterialReceiptsNoteVoucher3WithPricebookVendorAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkSavingMaterialReceiptsNoteVoucher3WithPricebookVendorAccountTwoAndValidatePricebookValues(), true);
		 }
	  
	  
	  
	  @Test(priority=182)
		 public void  checkSavingMaterialReceiptNoteVoucher4WithGroupOneAccountOneAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkSavingMaterialReceiptNoteVoucher4WithGroupOneAccountOneAndValidatePricebookValues(), true);
		 }
	  
	  
	  
	  @Test(priority=183)
		 public void  checkSavingMaterialReceiptNoteVoucher5WithGroupOneAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkSavingMaterialReceiptNoteVoucher5WithGroupOneAccountTwoAndValidatePricebookValues(), true);
		 }
	  
	  
	  
	  
	  @Test(priority=184)
		 public void  checkSavingMaterialReceiptNoteVoucher6WithGroupTwoAccountOneAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkSavingMaterialReceiptNoteVoucher6WithGroupTwoAccountOneAndValidatePricebookValues(), true);
		 }
	  
	  
	  
	  @Test(priority=185)
		 public void  checkSavingMaterialReceiptNoteVoucher7WithGroupTwoAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkSavingMaterialReceiptNoteVoucher7WithGroupTwoAccountTwoAndValidatePricebookValues(), true);
		 }
	  
	
	
	  
	  
	  
	  
	  
	  @Test(priority=186)
		 public void  checkRateValueofanItemBasedonUnitsinPurchasesVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkRateValueofanItemBasedonUnitsinPurchasesVoucher(), true);
		 }
	  
	  
	  
	  
	  @Test(priority=187)
		 public void  checkRatevalueforGroupItemsbasedonUnitsinPurchasesVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkRatevalueforGroupItemsbasedonUnitsinPurchasesVoucher(), true);
		 }
	  
	  
	  
	  
	  
	  @Test(priority=188)
		 public void  checkRateValuewhenquantiyisZerobasedonUnitsinPurchasesVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkRateValuewhenquantiyisZerobasedonUnitsinPurchasesVoucher(), true);
		 }
		 
	  
	  
	// @Test(priority=187)
		 public void  checkSavingaPricebookforCurrency() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkSavingaPricebookforCurrency(), true);
		 }
		 
	  
	  //@Test(priority=188)
		 public void  checkRateValueisLoadingforCurrencyandIteminPurchasesVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkRateValueisLoadingforCurrencyandIteminPurchasesVoucher(), true);
		 }
		 
	  
	  //Abbrevations
	  
	  
	
	  
	  @Test(priority=189)
		 public void  checkRateValueinPurchasesVoucherBasedonAbbrevationofPricebookforLessQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkRateValueinPurchasesVoucherBasedonAbbrevationofPricebookforLessQuantity(), true);
		 }
		 
	  @Test(priority=190)
		 public void  checkRateValueinPurchasesVoucherBasedonAbbrevationofPricebookforMediumQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkRateValueinPurchasesVoucherBasedonAbbrevationofPricebookforMediumQuantity(), true);
		 }
	  
	  @Test(priority=191)
		 public void  checkRateValueinPurchasesVoucherBasedonAbbrevationofPricebookforHighQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkRateValueinPurchasesVoucherBasedonAbbrevationofPricebookforHighQuantity(), true);
		 }
	  
	  
	  
	  
	  
	  @Test(priority=192)
		 public void  checkItemQueryReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkItemQueryReport(), true);
		 }
	  
	  
	  
	  
	  @Test(priority=193)
		 public void  checkViewSellerPriceinItemQueryReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkViewSellerPriceinItemQueryReport(), true);
		 }
	  
	  
	  @Test(priority=194)
		 public void  checkExportExcelforViewPriceSellerwithoutFilter() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkExportExcelforViewPriceSellerwithoutFilter(), true);
		 }
	  
	  
	  
	 @Test(priority=195)
		 public void  checkFilterwithDepartmentinViewSellerPriceinItemQueryReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkFilterwithDepartmentinViewSellerPriceinItemQueryReport(), true);
		 }
	  
	  
	  
	  @Test(priority=196)
		 public void  checkAllFilterOptionsinViewSellerPriceinItemQueryReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkAllFilterOptionsinViewSellerPriceinItemQueryReport(), true);
		 }
	  
	  
	  
	  @Test(priority=197)
		 public void  checkFilteroptionasVendorandCurrencyViewBuyerPriceinItemQuery() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkFilteroptionasVendorandCurrencyViewBuyerPriceinItemQuery(), true);
		 }
	  
	  
	//Stock Valuation
	  @Test(priority=215)
		 public void  checkSavingBuyerPricebookforSatockValuationReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkSavingBuyerPricebookforSatockValuationReport(), true);
		 }
	  
	  
	  
	  
	  
	  @Test(priority=216)
		 public void  checkSavingSellerPricebookforStockValuationReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkSavingSellerPricebookforStockValuationReport(), true);
		 }
	  
	  
	  
	  
	  @Test(priority=217)
		 public void  checkSavingOpeningStockVoucherwithPricebookItems() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkSavingOpeningStockVoucherwithPricebookItems(), true);
		 }
	  
	  
	  
		 
		 
		 @Test(priority=218)
		 public void  checkStockValuationReportforRatesinSellerandBuyerPricebook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkStockValuationReportforRatesinSellerandBuyerPricebook(), true);
		 }
		 
	  
	  
	  
	  @Test(priority=250)
		 public void  checkLogoutBuyerPriceBookPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPBV=new BuyerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(BPBV.checkLogoutBuyerPriceBookPage(), true);
		 }
	  
	  
}
