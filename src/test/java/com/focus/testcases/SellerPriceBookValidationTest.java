package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.SellerPriceBookPage;
import com.focus.Pages.SellerPriceBookValidationPage;
import com.focus.base.BaseEngine;

public class SellerPriceBookValidationTest extends BaseEngine{
	
	SellerPriceBookValidationPage SPBV;
	
	
	
	  @Test(priority=150)
		 public void  checkSellerPriceBookwithNoTagRateValueatSalesInvoiceVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV. checkSellerPriceBookwithNoTagRateValueatSalesInvoiceVoucher(), true);
		 }
	
	  
	  
	  @Test(priority=151)//Br Cogs Rate is not loading 
		 public void  checkRatewithPricebookAccountOneatSalesInvoiceVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV. checkRatewithPricebookAccountOneatSalesInvoiceVoucher(), true);
		 }
	
	  @Test(priority=152)
		 public void  checkRatewithPricebookAccountTwoatSalesInvoiceVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV. checkRatewithPricebookAccountTwoatSalesInvoiceVoucher(), true);
		 }
	  
	  

	  @Test(priority=153)
		 public void  checkRatewithPricebookGroupAccountOneatSalesInvoiceVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkRatewithPricebookGroupAccountOneatSalesInvoiceVoucher(), true);
		 }
	  
	  
	  @Test(priority=154)
		 public void  checkRatewithPricebookGroupAccountTwoatSalesInvoiceVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkRatewithPricebookGroupAccountTwoatSalesInvoiceVoucher(), true);
		 }
	  
	  @Test(priority=155)
		 public void  checkRatewithPricebookGroupAccountOneandGroupItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkRatewithPricebookGroupAccountOneandGroupItem(), true);
		 }
	  
	  @Test(priority=156)
		 public void  checkRatewithPricebookGroupAccountTwoandGroupItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkRatewithPricebookGroupAccountTwoandGroupItem(), true);
		 }
	  
	  
	  
	  @Test(priority=157)
		 public void  checkRateLoadingforPricebookDepartmentGroupinSalesInvoice() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkRateLoadingforPricebookDepartmentGroupinSalesInvoice(), true);
		 }
	  
	  
	  
	  @Test(priority=158)
		 public void  checkRateLoadingforItemhavingMaxandMinQuantityinSalesInvoiceVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkRateLoadingforItemhavingMaxandMinQuantityinSalesInvoiceVoucher(), true);
		 }
	  

	  
	  @Test(priority=159)
		 public void  checkCreatingSellerPriceBookforDateRange() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkCreatingSellerPriceBookforDateRange(), true);
		 }
	  
	  
	  
	  
	  @Test(priority=160)
		 public void  checkRateLoadingforItemhavingDateRangeinSalesInvoiceVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkRateLoadingforItemhavingDateRangeinSalesInvoiceVoucher(), true);
		 }
	  
	 
	  
	  
	  @Test(priority=161)
		 public void  checkItemFilterinSellerPricebook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkItemFilterinSellerPricebook(), true);
		 }
	  
	  
	  
	  
	  
	  @Test(priority=162)
		 public void  checkCustomerFilterinSellerPricebook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkCustomerFilterinSellerPricebook(), true);
		 }
	  
	  
	  
	  
	  @Test(priority=163)
		 public void  checkCurrencyFilterinSellerPricebook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkCurrencyFilterinSellerPricebook(), true);
		 }
	  
	  
	  @Test(priority=164)
		 public void  checkItemandCurrencyFilterinSellerPricebook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkItemandCurrencyFilterinSellerPricebook(), true);
		 }
	  
	  @Test(priority=165)
		 public void  checkItemandCustomerFilterinSellerPricebook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkItemandCustomerFilterinSellerPricebook(), true);
		 }
	  
	  @Test(priority=166)
		 public void  checkItemCurrencyandCustomerFilterinSellerPricebook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkItemCurrencyandCustomerFilterinSellerPricebook(), true);
		 }
	  
	  
	  
	  @Test(priority=167)
		 public void  checkItemCurrencyCustomerDepartmentFilterinSellerPricebook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkItemCurrencyCustomerDepartmentFilterinSellerPricebook(), true);
		 }
	  
	  
	  
	   
	  
	  @Test(priority=168)
		 public void  checkExportFileafterApplyingFilter() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkExportFileafterApplyingFilter(), true);
		 }
	  
	  
	  
	//  @Test(priority=169)
		 public void  checkSavingStatesinStateMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkSavingStatesinStateMaster(), true);
		 }
	  
	  
	  @Test(priority=170)
		 public void  checkCreatingSellerPriceBookforState() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkCreatingSellerPriceBookforState(), true);
		 }
	  
	  
	  
	
	  
	@Test(priority=171)
	 public void  checkCreatedMasterRateValueinSalesInvoiceVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPBV=new SellerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(SPBV.checkCreatedMasterRateValueinSalesInvoiceVoucher(), true);
	 }
	
	
	
	@Test(priority=172)
	 public void  checkAdvExcelImportingSellerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPBV=new SellerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(SPBV.checkAdvExcelImportingSellerPriceBook(), true);
	 }
	
	
	
	
	@Test(priority=173)
	 public void  checkDeleteSellerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPBV=new SellerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(SPBV.checkDeleteSellerPriceBook(), true);
	 }
	
	
	@Test(priority=174)
	 public void  checkExportingXMLBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPBV=new SellerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(SPBV.checkExportingXMLBuyerPriceBook(), true);
	 }
	
	@Test(priority=175)
	 public void  checkImportingXMLBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPBV=new SellerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(SPBV.checkImportingXMLBuyerPriceBook(), true);
	 }
	
	
	
	
	
	

	@Test(priority=176)
	 public void  checkPrintRateHistorySellerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPBV=new SellerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(SPBV.checkPrintRateHistorySellerPriceBook(), true);
	 }
	
	
	@Test(priority=177)
	 public void  checkPreviewandClearBtninPriceRateHistory() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPBV=new SellerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(SPBV.checkPreviewandClearBtninPriceRateHistory(), true);
	 }
	
	@Test(priority=178)
	 public void  checkViewPriceCustomizationBtnSellerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SPBV=new SellerPriceBookValidationPage(getDriver());
		 Assert.assertEquals(SPBV.checkViewPriceCustomizationBtnSellerPriceBook(), true);
	 }
	
	
	

	  
	  ///Delivery Notes
	 
	  
	  @Test(priority=179)
		 public void  checkSavingDeliveryNoteVoucher1WithNoTagsAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkSavingDeliveryNoteVoucher1WithNoTagsAndValidatePricebookValues(), true);
		 }
	
	  
	  @Test(priority=180)
		 public void  checkSavingDeliveryNoteVoucher2WithPricebookAccountOneAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkSavingDeliveryNoteVoucher2WithPricebookAccountOneAndValidatePricebookValues(), true);
		 }
	  
	  
	  @Test(priority=181)
		 public void  checkSavingDeliveryNoteVoucher3WithPricebookAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkSavingDeliveryNoteVoucher3WithPricebookAccountTwoAndValidatePricebookValues(), true);
		 }
	  
	  
	  
	  @Test(priority=182)
		 public void  checkSavingDeliveryNoteVoucher4WithGroupOneAccountOneAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkSavingDeliveryNoteVoucher4WithGroupOneAccountOneAndValidatePricebookValues(), true);
		 }
	  
	  
	  
	  @Test(priority=183)
		 public void  checkSavingDeliveryNoteVoucher5WithGroupOneAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkSavingDeliveryNoteVoucher5WithGroupOneAccountTwoAndValidatePricebookValues(), true);
		 }
	  
	  
	  
	  
	  @Test(priority=184)
		 public void  checkSavingDeliveryNoteVoucher6WithGroupTwoAccountOneAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkSavingDeliveryNoteVoucher6WithGroupTwoAccountOneAndValidatePricebookValues(), true);
		 }
	  
	  
	  
	  @Test(priority=185)//check it
		 public void  checkSavingDeliveryNoteVoucher7WithGroupTwoAccountTwoAndValidatePricebookValues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkSavingDeliveryNoteVoucher7WithGroupTwoAccountTwoAndValidatePricebookValues(), true);
		 }
	  
	
	
	  
	  
	  
	  
	  
	  @Test(priority=186)
		 public void  checkRateValueofanItemBasedonUnitsinSalesInvoiceVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkRateValueofanItemBasedonUnitsinSalesInvoiceVoucher(), true);
		 }
	  
	  
	  
	  
	  @Test(priority=187)
		 public void  checkRatevalueforGroupItemsbasedonUnitsinSalesInvoiceVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkRatevalueforGroupItemsbasedonUnitsinSalesInvoiceVoucher(), true);
		 }
	  
	  
	  
	  
	  
	  @Test(priority=188)
		 public void  checkRateValuewhenquantiyisZerobasedonUnitsinSalesInvoiceVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkRateValuewhenquantiyisZerobasedonUnitsinSalesInvoiceVoucher(), true);
		 }
		 
	  
	  
	  @Test(priority=189)
		 public void  checkSavingaPricebookforCurrency() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkSavingaPricebookforCurrency(), true);
		 }
		 
	  
	  @Test(priority=190)
		 public void  checkRateValueisLoadingforCurrencyandIteminSalesInvoiceVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkRateValueisLoadingforCurrencyandIteminSalesInvoiceVoucher(), true);
		 }
		 
	  
	  //Abbrevations
	  
	  
	  
	  
	  @Test(priority=191)
		 public void  checkRateValueinSalesInvoiceVoucherBasedonAbbrevationofPricebookforLessQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkRateValueinSalesInvoiceVoucherBasedonAbbrevationofPricebookforLessQuantity(), true);
		 }
		 
	  @Test(priority=192)
		 public void  checkRateValueinSalesInvoiceVoucherBasedonAbbrevationofPricebookforMediumQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkRateValueinSalesInvoiceVoucherBasedonAbbrevationofPricebookforMediumQuantity(), true);
		 }
	  
	  @Test(priority=193)
		 public void  checkRateValueinSalesInvoiceVoucherBasedonAbbrevationofPricebookforHighQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 SPBV=new SellerPriceBookValidationPage(getDriver());
			 Assert.assertEquals(SPBV.checkRateValueinSalesInvoiceVoucherBasedonAbbrevationofPricebookforHighQuantity(), true);
		 }
	  
	  
}
