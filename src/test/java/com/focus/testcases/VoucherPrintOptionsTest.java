package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.VoucherPrintOptionsPage;
import com.focus.base.BaseEngine;

public class VoucherPrintOptionsTest extends BaseEngine
{
	VoucherPrintOptionsPage VPO;
	
	
	 // Buyer Pricebook Validations
	
	/* @Test(priority=10050)
	 public void checkSavingPurchaseVoucherVATVoucher1WithAllItems() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 VPO=new VoucherPrintOptionsPage(getDriver());
		 Assert.assertEquals(VPO.checkSavingPurchaseVoucherVATVoucher1WithAllItems(), true);
	 }
	 
	  @Test(priority=10051)
	  public void checkSavingPurchaseVoucherVATVoucher2() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkSavingPurchaseVoucherVATVoucher2(), true);
	  }
	*/
	  
	
	  //
	
	 /* @Test(priority=10051)
	  public void checkAccouncreatedForPrintEmailOne() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkAccouncreatedForPrintEmailOne(), true);
	  }
	
	  @Test(priority=10052)
	  public void checkAccouncreatedForPrintEmailTwo() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkAccouncreatedForPrintEmailTwo(), true);
	  }
	  
	  @Test(priority=10053)
	  public void checkAccouncreatedForPrintEmailThree() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkAccouncreatedForPrintEmailThree(), true);
	  }
	
	  @Test(priority=10054)
	  public void checkSavingReceiptsVoucherForEmailAccountOne() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkSavingReceiptsVoucherForEmailAccountOne(), true);
	  }
	  
	  @Test(priority=10055)
	  public void checkSavingReceiptsvoucherForEmailAccountTwo() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkSavingReceiptsvoucherForEmailAccountTwo(), true);
	  }
	  
	  @Test(priority=10056)
	  public void checkSavingReceiptsVoucherForEmailAccountThree() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkSavingReceiptsVoucherForEmailAccountThree(), true);
	  }
	*/
	
	
	  @Test(priority=10057)
	  public void checkEmailOptionInReceiptsVoucherVAT() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkEmailOptionInReceiptsVoucherVAT(), true);
	  }
	  
	
	
	
	  
	
	
	 // Voucher Print Validations
	
	
	
	/*
	
	  @Test(priority=10101)
	  public void checkPrintScreenOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintScreenOptions(), true);
	  }
	  
	  @Test(priority=10102)
	  public void checkLayoutOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkLayoutOptions(), true);
	  }
	  
	  @Test(priority=10103)
	  public void checkInvoiceDesignOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkInvoiceDesignOptions(), true);
	  }
	  
	  @Test(priority=10104)
	  public void checkFieldsTabOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkFieldsTabOptions(), true);
	  }
	
	  @Test(priority=10105)
	  public void checkDragandDropHeaderFiledsAndBodyFiledsIntoLayout() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkDragandDropHeaderFiledsAndBodyFiledsIntoLayout(), true);
	  }
	  
	  @Test(priority=10106)
	  public void checkChangeFieldPropertiesInLyout() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkChangeFieldPropertiesInLyout(), true);
	  }
	
	  @Test(priority=10107)
	  public void checkPrintPdfOfPurchaseVoucherVATWithCreatedLayout() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfOfPurchaseVoucherVATWithCreatedLayout(), true);
	  }
	  */
	/*
	  @Test(priority=10108)
	  public void checkCreatingLayoutWithTwoPages() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkCreatingLayoutWithTwoPages(), true);
	  }
	
	  
	  
	  @Test(priority=10109)
	  public void checkChangeFieldPropertiesInLayoutForMultiplePages() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkChangeFieldPropertiesInLayoutForMultiplePages(), true);
	  }	  
	  
	  
	  @Test(priority=10110)
	  public void checkPrintPdfOfLayoutForMultiplePages() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfOfLayoutForMultiplePages(), true);
	  }
	  
	  
	  
	  @Test(priority=10111)
	  public void checkCreatingLayoutForShapesLayout() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkCreatingLayoutForShapesLayout(), true);
	  }
	  
	   @Test(priority=10112)
	  public void CheckShapesPropertiesInPurchaseVatDesignLayout() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.CheckShapesPropertiesInPurchaseVatDesignLayout(), true);
	  }
	  
	  @Test(priority=10113)
	  public void checkPrintPdfofShapesLayout() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfofShapesLayout(), true);
	  }
	  
	  
	  @Test(priority=10114)
	  public void checkCreatingLayoutforGridPropertiesForClubByItem() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkCreatingLayoutforGridProperties(), true);
	  }
	  
	  
	  @Test(priority=10115)
	  public void checkLayoutFieldProperties() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkLayoutFieldProperties(), true);
	  }
	  
	  
	  @Test(priority=10116)
	  public void checkGridProperties_ClubByItem() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkGridProperties_ClubByItem(), true);
	  }
	  
	  
	 @Test(priority=10117)
	  public void checkPrintPdfWithEnablingDontClubIfRatesAreDiff() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfWithEnablingDontClubIfRatesAreDiff(), true);
	  }
	  
	  @Test(priority=10118)
	  public void checkDisablingIfRatesAreNotDifferentInLayout() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkDisablingIfRatesAreNotDifferentInLayout(), true);
	  }
	  
	  @Test(priority=10119)
	  public void checkPrintPdfAfterDisablingRatesCheckBox() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfAfterDisablingRatesCheckBox(), true);
	  }
	  
	  
	 @Test(priority=10120)
	  public void checkGridProperties_ClubByWarehouse() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkGridProperties_ClubByWarehouse(), true);
	  }
	  
	 
	  @Test(priority=10121)
	  public void checkPrintpdfForLayoutWithClubByWarehouse() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintpdfForLayoutWithClubByWarehouse(), true);
	  }
	  
	  @Test(priority=10122)
	  public void checkLayoutForGridPropertiesGroupByItem() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkLayoutForGridPropertiesGroupBy(), true);
	  }
	  
	  
	  @Test(priority=10123)
	  public void checkLayoutForFieldPropertiesGroupByItem() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkLayoutForFieldPropertiesGroupBy(), true);
	  }
	  
	  @Test(priority=10124)
	  public void checkGridproperties_GroupByItem() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkGridproperties_GroupByItem(), true);
	  }
	  
	 
	  @Test(priority=10125)
	  public void checkPrintPdfGridPropertiesGroupByItem() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfGridPropertiesGroupByItem(), true);
	  }
	  
	  @Test(priority=10126)
	  public void checkEnableCheckboxesInGroupBy() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkEnableCheckboxesInGroupBy(), true);
	  }
	 
	  @Test(priority=10127)
	  public void checkPrintPdfAfterEnablingCheckboexInGroupBy() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfAfterEnablingCheckboexInGroupBy(), true);
	  }
	  
	  @Test(priority=10128)
	  public void checkLayoutForGridPropertiesOrderByItem() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkLayoutForGridPropertiesOrderByItem(), true);
	  }
	  
	  
	  @Test(priority=10129)
	  public void checkLayoutForFieldPropertiesOrderByItem() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkLayoutForFieldPropertiesOrderByItem(), true);
	  }
	  
	  @Test(priority=10130)
	  public void checkGridproperties_OrderbyItem() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkGridproperties_OrderbyItem(), true);
	  }
	  
	  @Test(priority=10131)
	  public void checkPrintPdfGridPropertiesOrderByItem() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfGridPropertiesOrderByItem(), true);
	  }
	  
	  //issue
	//  @Test(priority=10132)
	  public void checkEnableCheckboxesInOderBy() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkEnableCheckboxesInOderBy(), true);
	  }
	  
	  
	  //@Test(priority=10133)
	  public void checkPrintPdfAfterEnablingCheckboexInOderBy() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfAfterEnablingCheckboexInOderBy(), true);
	  }
	  
	  
	 @Test(priority=10134)
	  public void checkLayoutGridProperties() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkLayoutGridProperties(), true);
	  }
	 
	  @Test(priority=10135)
	  public void checkLayoutFieldPropries() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkLayoutFieldPropries(), true);
	  }
	    
	  
	  @Test(priority=10136)
	  public void checkGridPropertiesInPvVatLayout() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkGridPropertiesInPvVatLayout(), true);
	  }
	  
	
	  @Test(priority=10137)
	  public void checkPrintPdfAfterGridPropertiesEnabledToTrue() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfAfterGridPropertiesEnabledToTrue(), true);
	  }
	  
	  
	  //area
	  @Test(priority=10138)
	  public void checkAreaInPurchaseVatLayout() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkAreaInPurchaseVatLayout(), true);
	  }
	  
	  
	  @Test(priority=10139)
	  public void checkFiledProperitesForArea() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkFiledProperitesForArea(), true);
	  }
	  
	  
	  @Test(priority=10140)
	  public void checkDragnetVouchernetAndAreaItoLayout() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkDragnetVouchernetAndAreaItoLayout(), true);
	  }
	  
	  
	  @Test(priority=10141)
	  public void checkPrintPdfForAreaLayout() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfForAreaLayout(), true);
	  }
	  
	  
	  @Test(priority=10142)
	  public void checkChangeAreaPropertiesToRelative() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkChangeAreaPropertiesToRelative(), true);
	  }
	  
	  
	  @Test(priority=10143)
	  public void checkPrintPDFForAreaLayoutAfterChangingToRealtive() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPDFForAreaLayoutAfterChangingToRealtive(), true);
	  }
	  
	  
	
	 
	  @Test(priority=10144)
	  public void checkAccountsCreationForMulipleandSinglePrintLayoutAssign() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkAccountsCreationForMulipleandSinglePrintLayoutAssign(), true);
	  }
	
	  

	  @Test(priority=10145)
	  public void checkAccountRadiobtnAndPrintUsingCheckboxinbothPrefAndDocCust() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkAccountRadiobtnAndPrintUsingCheckboxinbothPrefAndDocCust(), true);
	  }
	
	  
	  
	  @Test(priority=10146)
	  public void checkSavingPvVatForMultiplePrintLayoutAccount() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkSavingPvVatForMultiplePrintLayoutAccount(), true);
	  }
	  
	  @Test(priority=10147)
	  public void checkSavingPvVatForSinglePrintLayoutAccount() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkSavingPvVatForSinglePrintLayoutAccount(), true);
	  }
	  
	  
	
	  @Test(priority=10148)
	  public void chechPrintPdfForMultipleLayoutVoucher() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.chechPrintPdfForMultipleLayoutVoucher(), true);
	  }

	
	 @Test(priority=10149)
	  public void checkAddingPrintLayoutForNoofCopies() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkAddingPrintLayoutForNoofCopies(), true);
	  }
	
	 
	 @Test(priority=10150)
	  public void checkLayoutPropertiesPageType() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkLayoutPropertiesPageType(), true);
	  }
	 
	 
	
	 @Test(priority=10151)
	  public void checkPrintPdfWithoutChangingNamesInReports() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfWithoutChangingNamesInReports(), true);
	  }
	
	

	 @Test(priority=10152)
	  public void checkSettingLayoutLayoutNamesInReports() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkSettingLayoutLayoutNamesInReports(), true);
	  }
	 
	  
	  
	
	 @Test(priority=10153)
	  public void checkPrintPDFAfterChangingNamesInReports() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPDFAfterChangingNamesInReports(), true);
	  }
	
	
	
	 @Test(priority=10154)
	  public void checkAddingFieldsInPvVatHeader() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkAddingFieldsInPvVatHeader(), true);
	  }
	
	 
	 @Test(priority=10155)
	  public void checkAddingFieldsInPvVATBody() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkAddingFieldsInPvVATBody(), true);
	  }
	 
	
	
	@Test(priority=10156)
	  public void checkSavingPvVATWithDocAndPictures() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkSavingPvVATWithDocAndPictures(), true);
	  }
	
	
	
	@Test(priority=10157)
	  public void checkSavingDesignLayourForAttachments() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkSavingDesignLayourForAttachments(), true);
	  }
	
	
	@Test(priority=10158)
	  public void checkPropertiesForLayoutAttachments() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPropertiesForLayoutAttachments(), true);
	  }
	
	
	@Test(priority=10159)
	  public void checkPrintPDFForAttachmentsLayout() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPDFForAttachmentsLayout(), true);
	  }
	
	
	  */
	  ///////////    
	  
	
	  /*@Test(priority=10107)
	  public void checkSavingPurchasesVatDesignLayout() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkSavingPurchasesVatDesignLayout(), true);
	  }
	  
	  @Test(priority=10108)
	  public void checkEditingLayoutByAddingFields() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkEditingLayoutByAddingFields(), true);
	  }

	  @Test(priority=10109)
	  public void checkUpdatedLayoutAndDeleteTheLayout() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkUpdatedLayoutAndDeleteTheLayout(), true);
	  }
	  
	  
	  @Test(priority=10110)
	  public void checkPrintPdfOfPurchaseVoucherVATWithCreatedLayout() throws AWTException, Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfOfPurchaseVoucherVATWithCreatedLayout(), true);
	  }
	  
	  @Test(priority=10111)
	  public void checkPrintPdfOfPurchaseVoucherVAT() throws Exception
	  {
	 	  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfOfPurchaseVoucherVAT(), true);
	  }
	
	  @Test(priority=10112)
	  public void checkPrintPdfOfPurchaseVoucherVATWithImportedXML() throws Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfOfPurchaseVoucherVATWithImportedXML(), true);
	  }
	
	  @Test(priority=10113)
	  public void checkEmailOptionInPurchaseVoucherVAT() throws Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkEmailOptionInPurchaseVoucherVAT(), true);
	  }
	 
	   
	
	 /* // Document Set
	  
	  @Test(priority=11000)
	  public void checkRestoreCompanyForDocumentSet() throws Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkRestoreCompanyForDocumentSet(), true);
	  }
	
	  @Test(priority=11001)
	  public void checkPrintPdfOfPurchaseVoucherVATWithDocumentSetVendorA() throws Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfOfPurchaseVoucherVATWithDocumentSetVendorA(), true);
	  }
	
	  @Test(priority=11002)
	  public void checkPrintPdfOfPurchaseVoucherVATWithDocumentSetVendorB() throws Exception
	  {
		  VPO=new VoucherPrintOptionsPage(getDriver());
		  Assert.assertEquals(VPO.checkPrintPdfOfPurchaseVoucherVATWithDocumentSetVendorB(), true);
	  }*/
	 
	
	
	 
	 
	  
	/* 
	 // Seller Pricebook Validations
	
	 @Test(priority=20010)
	 public void checksalesInvoiceVatWithBuyerSellerPriceBooklist() throws Exception
	 {
		 VPO=new VoucherPrintOptionsPage(getDriver());
		 Assert.assertEquals(VPO.checksalesInvoiceVatWithBuyerSellerPriceBooklist(), true);
	 }
	 
	 @Test(priority=20011)
	 public void checkSavedVoucherInSalesINvoiceVAT() throws Exception
	 {
		 VPO=new VoucherPrintOptionsPage(getDriver());
		 Assert.assertEquals(VPO.checkSavedVoucherInSalesINvoiceVAT(), true);
	 }
	 
	 @Test(priority=20012)
	 public void checkUpdatingBuyerPricebookFormulaForSalesInvoiceVAT() throws Exception
	 {
		 VPO=new VoucherPrintOptionsPage(getDriver());
		 Assert.assertEquals(VPO.checkUpdatingBuyerPricebookFormulaForSalesInvoiceVAT(), true);
	 }
	 
	 @Test(priority=20013)
	 public void checkSavingSalesInvoiceVATVoucher2WithRateFormula() throws Exception
	 {
		 VPO=new VoucherPrintOptionsPage(getDriver());
		 Assert.assertEquals(VPO.checkSavingSalesInvoiceVATVoucher2WithRateFormula(), true);
	 }
	 
	 @Test(priority=20014)
	 public void checkSavedVoucher2InSalesINvoiceVATWithFormulaControl() throws Exception
	 {
		 VPO=new VoucherPrintOptionsPage(getDriver());
		 Assert.assertEquals(VPO.checkSavedVoucher2InSalesINvoiceVATWithFormulaControl(), true);
	 }
	 
	 @Test(priority=20015)
	 public void checkUpdatingBuyerPricebookFormulaForSalesInvoiceVATVoucher3() throws Exception
	 {
		 VPO=new VoucherPrintOptionsPage(getDriver());
		 Assert.assertEquals(VPO.checkUpdatingBuyerPricebookFormulaForSalesInvoiceVATVoucher3(), true);
	 }
	 
	 @Test(priority=20016)
	 public void checkSavingSalesInvoiceVATVoucher3WithRateFormula() throws Exception
	 {
		 VPO=new VoucherPrintOptionsPage(getDriver());
		 Assert.assertEquals(VPO.checkSavingSalesInvoiceVATVoucher3WithRateFormula(), true);
	 }
	 
	 @Test(priority=20017)
	 public void checkSavedVoucher3InSalesINvoiceVATWithFormulaControl() throws Exception
	 {
		 VPO=new VoucherPrintOptionsPage(getDriver());
		 Assert.assertEquals(VPO.checkSavedVoucher3InSalesINvoiceVATWithFormulaControl(), true);
	 }
	 
	 @Test(priority=20018)
	 public void checkChangingRatePreloadBehaviourToRateSalesInvoiceVAT() throws Exception
	 {
		 VPO=new VoucherPrintOptionsPage(getDriver());
		 Assert.assertEquals(VPO.checkChangingRatePreloadBehaviourToRateSalesInvoiceVAT(), true);
	 }
	 
	 
	 
	 
	 // Credit Limit Validations
	 
	 @Test(priority=20019)
	 public void checkSavingSalesInvoiceVATVoucher4WithRateAsLessThanCreditLimit() throws Exception
	 {
		 VPO=new VoucherPrintOptionsPage(getDriver());
		 Assert.assertEquals(VPO.checkSavingSalesInvoiceVATVoucher4WithRateAsLessThanCreditLimit(), true);
	 }
	 
	 @Test(priority=20020)
	 public void checkSavedVoucher4AndTryingToUpdateVoucherRateGreaterThanCreditLimit() throws Exception
	 {
		 VPO=new VoucherPrintOptionsPage(getDriver());
		 Assert.assertEquals(VPO.checkSavedVoucher4AndTryingToUpdateVoucherRateGreaterThanCreditLimit(), true);
	 }
	 
	 
	 */
	
	
	 //@Test(priority=30000)
	 public void checkEraseAllTransactionsAfterCompletingAllValidations() throws Exception
	 {
		 VPO=new VoucherPrintOptionsPage(getDriver());
		 Assert.assertEquals(VPO.checkEraseAllTransactionsAfterCompletingAllValidations(), true);
	 }
	 
	 

}
