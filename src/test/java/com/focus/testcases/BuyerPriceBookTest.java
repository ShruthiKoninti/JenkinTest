package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.BuyerPriceBookPage;
import com.focus.base.BaseEngine;

public class BuyerPriceBookTest extends BaseEngine
{
	BuyerPriceBookPage BPB;
	
	 
     @Test(priority=173)
	 public void checkSavingBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPB=new BuyerPriceBookPage(getDriver());
		 Assert.assertEquals(BPB.checkSavingBuyerPriceBook(), true);
	 }
     
     @Test(priority=174)
	 public void checkUpdatingBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPB=new BuyerPriceBookPage(getDriver());
		 Assert.assertEquals(BPB.checkUpdatingBuyerPriceBook(), true);
	 }
	
	 @Test(priority=175)
	 public void checkCopyPasteAndSavingBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPB=new BuyerPriceBookPage(getDriver());
		 Assert.assertEquals(BPB.checkCopyPasteAndSavingBuyerPriceBook(), true);
	 }
	
	 @Test(priority=176)
	 public void checkDeletingBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPB=new BuyerPriceBookPage(getDriver());
		 Assert.assertEquals(BPB.checkDeletingBuyerPriceBook(), true);
	 }
	 
	 
	 @Test(priority=177)
	 public void checkExportingBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPB=new BuyerPriceBookPage(getDriver());
		 Assert.assertEquals(BPB.checkExportingBuyerPriceBook(), true);
	 }
	 
	 @Test(priority=178)
	 public void checkImportingBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPB=new BuyerPriceBookPage(getDriver());
		 Assert.assertEquals(BPB.checkImportingBuyerPriceBook(), true);
	 }
	 
	
	/////////////////////////////////////////////////Validating Footer Fields///////////////////////////////////////
	 
	 @Test(priority=179)
	 public void checkAdvExcelImportingBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPB=new BuyerPriceBookPage(getDriver());
		 Assert.assertEquals(BPB.checkAdvExcelImportingBuyerPriceBook(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=180)
	 public void checkExportingXMLBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPB=new BuyerPriceBookPage(getDriver());
		 Assert.assertEquals(BPB.checkExportingXMLBuyerPriceBook(), true);
	 }
	 
	 
	 @Test(priority=181)
	 public void checkImportingXMLBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPB=new BuyerPriceBookPage(getDriver());
		 Assert.assertEquals(BPB.checkImportingXMLBuyerPriceBook(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=182)
	 public void checkPrintRateHistoryBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPB=new BuyerPriceBookPage(getDriver());
		 Assert.assertEquals(BPB.checkPrintRateHistoryBuyerPriceBook(), true);
	 }

	 
	 
	 
	 @Test(priority=183)
	 public void checkPreviewandClearBtninPriceRateHistory() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPB=new BuyerPriceBookPage(getDriver());
		 Assert.assertEquals(BPB.checkPreviewandClearBtninPriceRateHistory(), true);
	 }
	 
		 
	 
	 @Test(priority=184)
	 public void checkViewPriceCustomizationBtnBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPB=new BuyerPriceBookPage(getDriver());
		 Assert.assertEquals(BPB.checkViewPriceCustomizationBtnBuyerPriceBook(), true);
	 }
	
	 
	 @Test(priority=185)
		 public void checkViewPriceBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 BPB=new BuyerPriceBookPage(getDriver());
			 Assert.assertEquals(BPB.checkViewPriceBuyerPriceBook(), true);
		 }
	 
	 
	 
	 @Test(priority=186)
	 public void checkExportingtoExcelinViewPriceBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BPB=new BuyerPriceBookPage(getDriver());
		 Assert.assertEquals(BPB.checkExportingtoExcelinViewPriceBuyerPriceBook(), true);
	 }
	 
	 
	 
	 
	 
	 
	

}
