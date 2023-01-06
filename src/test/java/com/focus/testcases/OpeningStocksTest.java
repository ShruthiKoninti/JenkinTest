package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.OpeningStocksPage;
import com.focus.base.BaseEngine;

public class OpeningStocksTest extends BaseEngine
{
	OpeningStocksPage OSP;
	
	
	// @Test(priority=100)
	 public void checkRestoreAutomationCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkRestoreAutomationCompany(), true);
	 }
	
	
	 @Test(priority=220)
	 public void checkSavingOpeningStockVoucher1WithBatchItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkSavingOpeningStockVoucher1WithBatchItem(), true);
	 }
	 
	 @Test(priority=221)
	 public void checkEditingAndVerifingDataOfSavedOpeningStocksVoucher1AddingRow2WithBatchItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkEditingAndVerifingDataOfSavedOpeningStocksVoucher1AddingRow2WithBatchItem(), true);
	 }
	 
	 @Test(priority=222)
	 public void checkEditingAndVerifingDataOfSavedOpeningStocksVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkEditingAndVerifingDataOfSavedOpeningStocksVoucher1(), true);
	 }
	 
	 @Test(priority=223)
	 public void checkSavingOpeningStockVoucher2WithBatchItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkSavingOpeningStockVoucher2WithBatchItem(), true);
	 }
	 
	 @Test(priority=224)
	 public void checkVerifingDataOfSavedOpeningStocksVoucher2AndDeleting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkVerifingDataOfSavedOpeningStocksVoucher2AndDeleting(), true);
	 }
	 
	 



	 
	 @Test(priority=225)
	 public void checkSavingOpeningStockVoucher2WithStockItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkSavingOpeningStockVoucher2WithStockItem(), true);
	 }
	 
	 @Test(priority=226)
	 public void checkEditingAndVerifingDataOfSavedOpeningStocksVoucher2AddingRow2WithStockItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkEditingAndVerifingDataOfSavedOpeningStocksVoucher2AddingRow2WithStockItem(), true);
	 }
	 
	 @Test(priority=227)
	 public void checkEditingAndVerifingDataOfSavedOpeningStocksVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkEditingAndVerifingDataOfSavedOpeningStocksVoucher2(), true);
	 }
	 
	 @Test(priority=228)
	 public void checkSavingOpeningStockVoucher3WithStockItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkSavingOpeningStockVoucher3WithStockItem(), true);
	 }
	 
	 @Test(priority=229)
	 public void checkVerifingDataOfSavedOpeningStocksVoucher3AndDeleting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkVerifingDataOfSavedOpeningStocksVoucher3AndDeleting(), true);
	 }
	 
	 
	 
	 
	 
	 @Test(priority=230) // Bin Popup//balance issue
	 public void checkOpeningStockVoucher3Row1WithBinItemWithAutoAllocationOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkOpeningStockVoucher3Row1WithBinItemWithAutoAllocationOption(), true);
	 }
	 
	 @Test(priority=231)// Bin Popup//balance issue
	 public void checkSavingOpeningStockVoucher3With2RowsAndBinItemDoubleClickInIndex() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkSavingOpeningStockVoucher3With2RowsAndBinItemDoubleClickInIndex(), true);
	 }
	 
	 @Test(priority=232) // Bin Popup//balance issue
	 public void checkEditingAndVerifingDataOfSavedOpeningStocksVoucher3AddingRow3WithBinItemByAutoAllocation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkEditingAndVerifingDataOfSavedOpeningStocksVoucher3AddingRow3WithBinItemByAutoAllocation(), true);
	 }
	 
	 @Test(priority=233)
	 public void checkVerifingDataOfSavedOpeningStocksVoucher3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkVerifingDataOfSavedOpeningStocksVoucher3(), true);
	 }
	 
	 @Test(priority=234)
	 public void checkSavingOpeningStockVoucher4WithBinItemByManuallyEnteringQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkSavingOpeningStockVoucher4WithBinItemByManuallyEnteringQuantity(), true);
	 }
	 
	 @Test(priority=235)
	 public void checkVerifingDataOfSavedOpeningStocksVoucher4AndDeleting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkVerifingDataOfSavedOpeningStocksVoucher4AndDeleting(), true);
	 }
	 
	
	 
	 
	 
	 
	 @Test(priority=236)
	 public void checkSavingOpeningStockVoucher4WithRMAItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkSavingOpeningStockVoucher4WithRMAItem(), true);
	 }
	 
	 @Test(priority=237)
	 public void checkEditingAndVerifingDataOfSavedOpeningStocksVoucher4AddingRow2WithRMAItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkEditingAndVerifingDataOfSavedOpeningStocksVoucher4AddingRow2WithRMAItem(), true);
	 }
	 
	 @Test(priority=238)
	 public void checkEditingAndVerifingDataOfSavedOpeningStocksVoucher4() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkEditingAndVerifingDataOfSavedOpeningStocksVoucher4(), true);
	 }
	 
	 @Test(priority=239)
	 public void checkSavingOpeningStockVoucher5WithRMAItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkSavingOpeningStockVoucher5WithRMAItem(), true);
	 }
	 
	 @Test(priority=240)
	 public void checkVerifingDataOfSavedOpeningStocksVoucher5AndDeleting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkVerifingDataOfSavedOpeningStocksVoucher5AndDeleting(), true);
	 }
	
	
	 
	 
	 @Test(priority=241)
	 public void checkVerifingDetailsOfSavedOpeningStocksVouchersInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkVerifingDetailsOfSavedOpeningStocksVouchersInHomepage(), true);
	 }
	
	 @Test(priority=242)
	 public void checkSaveInCreateViewOfHomePageOfOpeningBalance() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkSaveInCreateViewOfHomePageOfOpeningBalance(), true);
	 }
	 
	 @Test(priority=300)
	 public void checkSavingAndSuspendingOpeningStockVoucher5WithStockItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkSavingAndSuspendingOpeningStockVoucher5WithStockItem(), true);
	 }
	 
	 @Test(priority=301)
	 public void checkVerifingDetailsOfSuspendedOpeningStocksVouchersInCreatedView() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkVerifingDetailsOfSuspendedOpeningStocksVouchersInCreatedView(), true);
	 }
	 
	 @Test(priority=302)
	 public void checkResavingOpeningStockVoucher5AndSuspendingFromHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkResavingOpeningStockVoucher5AndSuspendingFromHomepage(), true);
	 }
	 
	 @Test(priority=303)
	 public void checkDeleteOptionInOpeningStocksHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkDeleteOptionInOpeningStocksHomePage(), true);
	 }
	 
	 
	 
	 
	 
	/* @Test(priority=243)
	 public void checkSaveCreateViewOfHomePageOfOpeningBalance() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkSaveCreateViewOfHomePageOfOpeningBalance(), true);
	 }
	 
	 //@Test(priority=244)
	 public void checkEditViewofHomePageOpeningStocks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkEditViewofHomePageOpeningStocks(), true);
	 }
	
	 
	 @Test(priority=245)
	 public void checkValidationofSavedViewAfterRefreshtheBrowser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkValidationofSavedViewAfterRefreshtheBrowser(), true);
	 }
	
	 @Test(priority=246)
	 public void checkValidationofDeletingViewinHomePageofOpeningStocks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkValidationofDeletingViewinHomePageofOpeningStocks(), true);
	 }
	 
	 
	 @Test(priority=247)
	 public void checkDeletedViewinHomePageofOpeningStocks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 OSP=new OpeningStocksPage(getDriver());
		 Assert.assertEquals(OSP.checkDeletedViewinHomePageofOpeningStocks(), true);
	 }*/
}