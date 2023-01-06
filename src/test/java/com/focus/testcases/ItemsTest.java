package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.ItemsPage;
import com.focus.base.BaseEngine;

public class ItemsTest extends BaseEngine
{
	ItemsPage IP;
	
	
	
	 @Test(priority=122)
	 public void checkItemMasterInMasterMenu() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkItemMasterInMasterMenu(), true);
	 }
	
	 @Test(priority=123)
	 public void checkSavingItemInItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkSavingItemInItemMaster(), true);
	 }
	
	 @Test(priority=124)
	 public void checkEditingInSavedItemINItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkEditingInSavedItemINItemMaster(), true);
	 }
	
	 @Test(priority=125)
	 public void checkUpdatedItemMasterDisplayInItemGridBeforeCustomization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkUpdatedItemMasterDisplayInItemGridBeforeCustomization(), true);
	 }
	
	 @Test(priority=126)
	 public void checkDeleteItemInItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkDeleteItemInItemMaster(), true);
	 }
	 
	 //////////////////Importing Data from Excel//////////////////////////
	 @Test(priority=127)
	 public void checkItemListBeforeImportingDataFromExcel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkItemListBeforeImportingDataFromExcel(), true);
	 }
	 
	 
	 @Test(priority=128)
	 public void checkAdvanceMasterImportExportItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkAdvanceMasterImportExportItem(), true);
	 }
	 
	 
	 @Test(priority=129)
	 public void checkImportDatainAdvanceMasterImportExportAccount() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkImportDatainAdvanceMasterImportExportAccount(), true);
	 }
	 
	 
	 @Test(priority=130)
	 public void checkItemListAfterImportingDataFromExcel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkItemListAfterImportingDataFromExcel(), true);
	 }
	 
	  
	 
	 @Test(priority=131)
	 public void checkImportedItemDelete() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkImportedItemDelete(), true);
	 }
	 
	 
	 
	 ////////////////////////Checking New Groups and SubGroups added/////////////////////////////
	 
		
	 @Test(priority=132)
	 public void checkSavingItemGroupinItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkSavingItemGroupinItemMaster(), true);
	 }
	 
	 
	 @Test(priority=133)
	 public void checkSavingItemSubGroupinItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkSavingItemSubGroupinItemMaster(), true);
	 }
	 

	 @Test(priority=134)
	 public void checkCreatedGroupDisplayedonLeftPanel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkCreatedGroupDisplayedonLeftPanel(), true);
	 }
	 
	
	 @Test(priority=135)
	 public void checkSavingItemsinSubGroupsinItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkSavingItemsinSubGroupsinItemMaster(), true);
	 }
	
	 @Test(priority=136)
	 public void checkItemsSavedinGroupsandSubGroupsinItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkItemsSavedinGroupsandSubGroupsinItemMaster(), true);
	 }
 
	 @Test(priority=137)
	 public void checkDeleteGroupIteminItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.checkDeleteGroupIteminItemMaster(), true);
	 }
	 
	 
	 
	 
	 
	 
}
