package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.BinsPage;
import com.focus.base.BaseEngine;

public class BinsTest extends BaseEngine
{
	BinsPage BP;
	

	 @Test(priority=160)
	 public void checkBinsMasterInMasterMenu() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BP=new BinsPage(getDriver());
		 Assert.assertEquals(BP.checkBinsMasterInMasterMenu(), true);
	 }
	
	 @Test(priority=161)
	 public void checkSavingBinsInBinsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BP=new BinsPage(getDriver());
		 Assert.assertEquals(BP.checkSavingBinsInBinsMaster(), true);
	 }
	
	 @Test(priority=162)
	 public void checkEditingInSavedBinsINBinsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BP=new BinsPage(getDriver());
		 Assert.assertEquals(BP.checkEditingInSavedBinsINBinsMaster(), true);
	 }
	
	 @Test(priority=163)
	 public void checkUpdatedBinsMasterDisplayInBins() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BP=new BinsPage(getDriver());
		 Assert.assertEquals(BP.checkUpdatedBinsMasterDisplayInBins(), true);
	 }
	
	 @Test(priority=164)
	 public void checkDeleteInBinsInBinsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BP=new BinsPage(getDriver());
		 Assert.assertEquals(BP.checkDeleteInBinsInBinsMaster(), true);
	 }
	 
	 
	 
	 @Test(priority=165)
	 public void checkBinsListBeforeImportingDataFromExcel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BP=new BinsPage(getDriver());
		 Assert.assertEquals(BP.checkBinsListBeforeImportingDataFromExcel(), true);
	 }
	 
	 
	 @Test(priority=166)
	 public void checkAdvanceMasterImportExportBins() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BP=new BinsPage(getDriver());
		 Assert.assertEquals(BP.checkAdvanceMasterImportExportBins(), true);
	 }
	 
	 
	 @Test(priority=167)
	 public void checkImportDatainAdvanceMasterImportExportBins() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BP=new BinsPage(getDriver());
		 Assert.assertEquals(BP.checkImportDatainAdvanceMasterImportExportBins(), true);
	 }
	 
	 
	 @Test(priority=168)//not importing
	 public void checkBinsListAfterImportingDataFromExcel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BP=new BinsPage(getDriver());
		 Assert.assertEquals(BP.checkBinsListAfterImportingDataFromExcel(), true);
	 }
	 
	 
	 
	 @Test(priority=169)//not importing
	 public void checkImportedBinDelete() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 BP=new BinsPage(getDriver());
		 Assert.assertEquals(BP.checkImportedBinDelete(), true);
	 }
	
}
