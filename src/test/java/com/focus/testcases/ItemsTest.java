package com.focus.testcases;

import java.awt.AWTException;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.AccountsPage;
import com.focus.Pages.ItemsPage;
import com.focus.base.BaseEngine;

public class ItemsTest extends BaseEngine
{
	ItemsPage IP;
	
	
	
	 @Test(priority=120)
	 public void CheckLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IP=new ItemsPage(getDriver());
		 Assert.assertEquals(IP.CheckLogin(), true);
	 }
/*
	 
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
	 
	 
	
	////Copy and Paste Item 
	 
		 @Test(priority=140)
		 public void checkCopyandPasteofAnItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 IP=new ItemsPage(getDriver());
			 Assert.assertEquals(IP.checkCopyandPasteofAnItem(), true);
		 }
		 
		 
		
		 
		 @Test(priority=141)
		 public void checkUnitsTabDataAfterPasteItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 IP=new ItemsPage(getDriver());
			 Assert.assertEquals(IP.checkUnitsTabDataAfterPasteItem(), true);
		 }
		 
		 
		 
		 @Test(priority=142)
		 public void checkSettingsTabDataAfterPasteItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 IP=new ItemsPage(getDriver());
			 Assert.assertEquals(IP.checkSettingsTabDataAfterPasteItem(), true);
		 } 
			 
			 
		@Test(priority=143)
		 public void checkClasifictationTabDataAfterPasteItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkClasifictationTabDataAfterPasteItem(), true);
		}
		 
		 
		@Test(priority=144)
		public void checkOtherDetailsTabDataAfterPasteItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkOtherDetailsTabDataAfterPasteItem(), true);
		}
			 
		 
		@Test(priority=145)
		public void checkReplenishmentTabDataAfterPasteItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
				IP=new ItemsPage(getDriver());
				Assert.assertEquals(IP.checkReplenishmentTabDataAfterPasteItem(), true);
		}
			
		
		@Test(priority=146)
		public void CheckSavingItemAfterPaste() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
				IP=new ItemsPage(getDriver());
				Assert.assertEquals(IP.CheckSavingItemAfterPaste(), true);
		}
	 
		*/
		
		///Clone Item
		
		

		@Test(priority=147)
		public void checkItemCloneOptionFromHome() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
				IP=new ItemsPage(getDriver());
				Assert.assertEquals(IP.checkItemCloneOptionFromHome(), true);
		}
		

		@Test(priority=148)
		public void checkCloneIteminDetailsTabAfterSave() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
				IP=new ItemsPage(getDriver());
				Assert.assertEquals(IP.checkCloneIteminDetailsTabAfterSave(), true);
		}
		
		
		
		 @Test(priority=149)
		 public void checkUnitsTabDataAfterCloneItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 IP=new ItemsPage(getDriver());
			 Assert.assertEquals(IP.checkUnitsTabDataAfterPasteItem(), true);
		 }
		 
		 
		 
		 @Test(priority=150)
		 public void checkSettingsTabDataAfterCloneItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 IP=new ItemsPage(getDriver());
			 Assert.assertEquals(IP.checkSettingsTabDataAfterPasteItem(), true);
		 } 
			 
			 
		@Test(priority=151)
		 public void checkClasifictationTabDataAfterCloneItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkClasifictationTabDataAfterPasteItem(), true);
		}
		 
		 
		@Test(priority=152)
		public void checkOtherDetailsTabDataAfterCloneItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkOtherDetailsTabDataAfterPasteItem(), true);
		}
			 
		 
		@Test(priority=153)
		public void checkReplenishmentTabDataAfterCloneItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
				IP=new ItemsPage(getDriver());
				Assert.assertEquals(IP.checkReplenishmentTabDataAfterCloneItem(), true);
		}
	 
	 
///////////////Creating Extra Fields//////////////////////
	 
	 
	
		@Test(priority=154)
		public void checkCreatingExtraFieldsinItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		IP=new ItemsPage(getDriver());
		Assert.assertEquals(IP.checkCreatingExtraFieldsinItemMaster(), true);
		}
		
		
		@Test(priority=155)
		public void checkAddedExtraFieldinGeneralTabofItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		IP=new ItemsPage(getDriver());
		Assert.assertEquals(IP.checkAddedExtraFieldinGeneralTabofItemMaster(), true);
		}
		
		
		
		
		@Test(priority=156)
		public void checkEditandUpdateExtraFieldinCustomizeMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		IP=new ItemsPage(getDriver());
		Assert.assertEquals(IP.checkEditandUpdateExtraFieldinCustomizeMaster(), true);
		}
		
		
		@Test(priority=157)
		public void checkLogoutAfterExtraFieldUpdated() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			IP=new ItemsPage(getDriver());
			Assert.assertEquals(IP.checkLogoutItemsPage(), true);
		}
		
		
		 @Test(priority=158)
		 public void CheckLoginAfterExtraFieldUpdated() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 IP=new ItemsPage(getDriver());
			 Assert.assertEquals(IP.CheckLogin(), true);
		 }
		

		@Test(priority=159)
		public void checkUpdatedExtraFieldinGeneralTabofItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		IP=new ItemsPage(getDriver());
		Assert.assertEquals(IP.checkUpdatedExtraFieldinGeneralTabofItemMaster(), true);
		}
		
		
		
		
		@Test(priority=160)
		public void checkDeleteExtraFieldinCustomizeMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		IP=new ItemsPage(getDriver());
		Assert.assertEquals(IP.checkDeleteExtraFieldinCustomizeMaster(), true);
		}
		
		
		
		
		@Test(priority=161)
		public void checkLogoutAfterExtraFieldDeleted() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			IP=new ItemsPage(getDriver());
			Assert.assertEquals(IP.checkLogoutItemsPage(), true);
		}
		
		
		 @Test(priority=162)
		 public void CheckLoginAfterExtraFieldDeleted() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 IP=new ItemsPage(getDriver());
			 Assert.assertEquals(IP.CheckLogin(), true);
		 }
		
		
		
		
		@Test(priority=163)
		public void checkDeletedExtraFieldinGeneralTabofItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		IP=new ItemsPage(getDriver());
		Assert.assertEquals(IP.checkDeletedExtraFieldinGeneralTabofItemMaster(), true);
		}

		@Test(priority=164)
		public void checkEditedItemInCustomizeDisplayColumns() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		IP=new ItemsPage(getDriver());
		Assert.assertEquals(IP.checkEditedItemInCustomizeDisplayColumns(), true);
		}

		 //////////////////Importing Data from Excel//////////////////////////
		 @Test(priority=166)
		 public void checkItemListBeforeImportingDataFromExcel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 IP=new ItemsPage(getDriver());
			 Assert.assertEquals(IP.checkItemListBeforeImportingDataFromExcel(), true);
		 }
		 
		 
		 @Test(priority=167)
		 public void checkAdvanceMasterImportExportItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 IP=new ItemsPage(getDriver());
			 Assert.assertEquals(IP.checkAdvanceMasterImportExportItem(), true);
		 }
		 
		 
		 @Test(priority=168)
		 public void checkImportDatainAdvanceMasterImportExportAccount() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 IP=new ItemsPage(getDriver());
			 Assert.assertEquals(IP.checkImportDatainAdvanceMasterImportExportAccount(), true);
		 }
		 
		 
		 @Test(priority=169)
		 public void checkItemListAfterImportingDataFromExcel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 IP=new ItemsPage(getDriver());
			 Assert.assertEquals(IP.checkItemListAfterImportingDataFromExcel(), true);
		 }
		 
		
		 
		 @Test(priority=170)
		 public void checkImportedItemDelete() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 IP=new ItemsPage(getDriver());
			 Assert.assertEquals(IP.checkImportedItemDelete(), true);
		 }
		
		 
		///Creating Item Set
		
			
			 @Test(priority=175)
			 public void checkSavingItemSetinItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkSavingItemSetinItemMaster(), true);
			 }
			 
			 
				 
			 @Test(priority=176)
			 public void checkVerifyingtheItemSetinItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkVerifyingtheItemSetinItemMaster(), true);
			 }
			 
			 
			 
			 @Test(priority=177)
			 public void checkDeletingItemSetinItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkDeletingItemSetinItemMaster(), true);
			 }
			
			
			 //Backup
			 
			 
			///Importing Bulk and Image
				
			 
			 @Test(priority=178)
			 public void checkBulkExportAndImportinItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkBulkExportAndImportinItemMaster(), true);
			 }
			
			
			
			@Test(priority=179)
			 public void checkItemsAfterBulkImportinItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkItemsAfterBulkImportinItemMaster(), true);
			 }
			
			
			//Details Tab
			
			@Test(priority=180)
			 public void checkUpdatingDetilsTabDataforExistingItemthroughBulkImport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkUpdatingDetilsTabDataforExistingItemthroughBulkImport(), true);
			 }
			
			
			
			@Test(priority=181)//
			 public void checkImportedDatainDetailsTabAfterUpdatingExistingItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkImportedDatainDetailsTabAfterUpdatingExistingItem(), true);
			 }
			
			
			
			
			@Test(priority=182)//
			 public void checkImportedDatainDetailsTabAfterUpdatingExistingItem1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkImportedDatainDetailsTabAfterUpdatingExistingItem1(), true);
			 }
			
			
			
			@Test(priority=183)
			 public void checkUpdatingDetilsTabDataforNewItemthroughBulkImport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkUpdatingDetilsTabDataforNewItemthroughBulkImport(), true);
			 }
			
			
			
			@Test(priority=184)
			 public void checkImportedDatainDetailsTabAfterUpdatingNewItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkImportedDatainDetailsTabAfterUpdatingNewItem(), true);
			 }
			
			//Units Tab
			
			
			@Test(priority=185)
			 public void checkUpdatingUnitsTabDataforExistingItemthroughBulkImport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkUpdatingUnitsTabDataforExistingItemthroughBulkImport(), true);
			 }
			
			
			@Test(priority=186)
			 public void checkImportedDatainUnitsTabAfterUpdatingExistingItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkImportedDatainUnitsTabAfterUpdatingExistingItem(), true);
			 }
			
			
			
			@Test(priority=187)
			 public void checkImportedDatainUnitsTabAfterUpdatingExistingItem1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkImportedDatainUnitsTabAfterUpdatingExistingItem1(), true);
			 }
			
			
			@Test(priority=188)
			 public void checkUpdatingUnitsTabDataforNewItemthroughBulkImport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkUpdatingUnitsTabDataforNewItemthroughBulkImport(), true);
			 }
			
			
			@Test(priority=189)
			 public void checkImportedDatainUnitsTabAfterUpdatingNewItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkImportedDatainUnitsTabAfterUpdatingNewItem(), true);
			 }
			
			
			//Other Details Tab
			
			
			
			@Test(priority=190)
			 public void checkUpdatingOtherDetailsTabforExistingItemthroughBulkImport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkUpdatingOtherDetailsTabforExistingItemthroughBulkImport(), true);
			 }
			
			
			
			@Test(priority=191)//
			 public void checkImportedDatainOtherDetailsTabAfterUpdatingExistingItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkImportedDatainOtherDetailsTabAfterUpdatingExistingItem(), true);
			 }
			
			
			@Test(priority=192)
			 public void checkUpdatingOtherDetailsTabDataforNewItemthroughBulkImport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkUpdatingOtherDetailsTabDataforNewItemthroughBulkImport(), true);
			 }
			
			@Test(priority=193)//
			 public void checkImportedDatainOtherDetailsTabAfterUpdatingNewItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkImportedDatainOtherDetailsTabAfterUpdatingNewItem(), true);
			 }
			
			
			
			//GST Tab
			
			
			
			@Test(priority=194)
			 public void checkUpdatingGSTTabDataforNewItemThroughbulkImport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkUpdatingGSTTabDataforNewItemThroughbulkImport(), true);
			 }
			
			
			
			@Test(priority=195)
			 public void checkImportedDatainGSTTabAfterUpdatingNewItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkImportedDatainGSTTabAfterUpdatingNewItem(), true);
			 }
			
			@Test(priority=196)
			 public void checkUpdatingGSTTabDataforExistingItemThroughbulkImport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkUpdatingGSTTabDataforExistingItemThroughbulkImport(), true);
			 }
			
			
			@Test(priority=197)
			 public void checkImportedDatainGSTTabAfterUpdatingExistingItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkImportedDatainGSTTabAfterUpdatingExistingItem(), true);
			 }
		
			
			//Importing with Image
			
			@Test(priority=198)
			 public void checkImportingItemAlongWithImageinItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 IP=new ItemsPage(getDriver());
				 Assert.assertEquals(IP.checkImportingItemAlongWithImageinItemMaster(), true);
			 }
		
				
			
		
		
		@Test(priority=199)
		 public void checkImportedItemWithImage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 IP=new ItemsPage(getDriver());
			 Assert.assertEquals(IP.checkImportedItemWithImage(), true);
		 }
		
		
		
		
		@Test(priority=200)
		 public void checkSavingImageofItemImported() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 IP=new ItemsPage(getDriver());
			 Assert.assertEquals(IP.checkSavingImageofItemImported(), true);
		 }
		
		
		@Test(priority=201)
		 public void checkDeletingtheImageinImportedItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 IP=new ItemsPage(getDriver());
			 Assert.assertEquals(IP.checkDeletingtheImageinImportedItem(), true);
		 }
		
		@Test(priority=210)
		public void checkLogoutItemsPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		IP=new ItemsPage(getDriver());
		Assert.assertEquals(IP.checkLogoutItemsPage(), true);
		}
	 
	 
}
