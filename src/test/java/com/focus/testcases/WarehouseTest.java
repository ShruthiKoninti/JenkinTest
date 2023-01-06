package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.WarehousePage;
import com.focus.base.BaseEngine;

public class WarehouseTest extends BaseEngine
{
	WarehousePage WP;
	

	 @Test(priority=154)
	 public void checkWarehouseMasterInMAsterMenu() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkWarehouseMasterInMAsterMenu(), true);
	 }
	
	 @Test(priority=155)
	 public void checkSavingWarehouseInWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkSavingWarehouseInWarehouseMaster(), true);
	 }
	
	 @Test(priority=156)
	 public void checkEditingInSavedWarehouseINWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkEditingInSavedWarehouseINWarehouseMaster(), true);
	 }
	
	 @Test(priority=157)
	 public void checkUpdatedWarehouseMasterDisplayInWarehouse() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkUpdatedWarehouseMasterDisplayInWarehouse(), true);
	 }
	
	 @Test(priority=158)
	 public void checkDeleteInWarehouseInWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkDeleteInWarehouseInWarehouseMaster(), true);
	 }
	
	 @Test(priority=159)
	 public void checkWarehouseListBeforeImportingDataFromExcel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkWarehouseListBeforeImportingDataFromExcel(), true);
	 }
	 
	 
	 @Test(priority=160)
	 public void checkAdvanceMasterImportExportWarehouse() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkAdvanceMasterImportExportWarehouse(), true);
	 }
	 
	 
	 @Test(priority=161)
	 public void checkImportDatainAdvanceMasterImportExportWarehouse() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkImportDatainAdvanceMasterImportExportWarehouse(), true);
	 }
	 
	 
	 @Test(priority=162)
	 public void checkWarehouseListAfterImportingDataFromExcel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkWarehouseListAfterImportingDataFromExcel(), true);
	 }
	 
	 @Test(priority=163)
	 public void checkImportedWarehouseDelete() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkImportedWarehouseDelete(), true);
	 }

	 
	 
	 
	 ////////////////////////Checking New Groups and SubGroups added/////////////////////////////
	 
		
	 @Test(priority=164)
	 public void checkSavingWarehouseGroupinWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkSavingWarehouseGroupinWarehouseMaster(), true);
	 }
	 
	 
	 @Test(priority=165)
	 public void checkSavingWarehouseSubGroupinWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkSavingWarehouseSubGroupinWarehouseMaster(), true);
	 }
	 

	 @Test(priority=166)
	 public void checkCreatedGroupDisplayedonLeftPanel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkCreatedGroupDisplayedonLeftPanel(), true);
	 }
	 
	
	 @Test(priority=167)
	 public void checkSavingWarehousesinSubGroupsinWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkSavingWarehousesinSubGroupsinWarehouseMaster(), true);
	 }
	
	 @Test(priority=168)
	 public void checkWarehousesSavedinGroupsandSubGroupsinWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkWarehousesSavedinGroupsandSubGroupsinWarehouseMaster(), true);
	 }
	 
	 
	 
	 @Test(priority=169)
	 public void checkDeleteGroupWarehouseinWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 WP=new WarehousePage(getDriver());
		 Assert.assertEquals(WP.checkDeleteGroupWarehouseinWarehouseMaster(), true);
	 }
	 
	 
	 
	 
}
