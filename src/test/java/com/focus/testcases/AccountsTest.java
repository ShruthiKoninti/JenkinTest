package com.focus.testcases;

import java.awt.AWTException;


import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.AccountsPage;
import com.focus.Pages.AccountsPage;
import com.focus.base.BaseEngine;

public class AccountsTest extends BaseEngine
{
	AccountsPage AP;
	
	
	
	
	@Test(priority=100)
		 public void CheckLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 AP=new AccountsPage(getDriver());
			 Assert.assertEquals(AP.CheckLogin(), true);
		 }

	@Test(priority=110)
	 public void checkRestoreAutomationCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkRestoreAutomationCompany(), true);
	 }
	
	
	@Test(priority=117)
	 public void checkSavingAccountINAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkSavingAccountINAccountMaster(), true);
	 }
	
	@Test(priority=118)
	 public void checkEditingInSavedAccountINAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkEditingInSavedAccountINAccountMaster(), true);
	 }
	
	@Test(priority=119)
	 public void checkUpdatedAccountDisplayInAccountGridBeforeCustomization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkUpdatedAccountDisplayInAccountGridBeforeCustomization(), true);
	 }
	 
	@Test(priority=120)
	 public void checkDeleteAccountInAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkDeleteAccountInAccountMaster(), true);
	 }
	 
	 ////////////////////////Checking New Groups and SubGroups added/////////////////////////////
	 
		
	@Test(priority=122)
	 public void checkSavingAccountGroupinAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkSavingAccountGroupinAccountMaster(), true);
	 }
 
	@Test(priority=123)
	 public void checkSavingAccountSubGroupinAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkSavingAccountSubGroupinAccountMaster(), true);
	 }
	 
	 
	 
	
	@Test(priority=124)
	 public void checkCreatedGroupDisplayedonLeftPanel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkCreatedGroupDisplayedonLeftPanel(), true);
	 }


	@Test(priority=125)
	 public void checkSavingAccountsinSubGroupsinAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkSavingAccountsinSubGroupsinAccountMaster(), true);
	 }
	 


	 @Test(priority=126)
	 public void checkAccountsSavedinGroupsandSubGroupsinAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkAccountsSavedinGroupsandSubGroupsinAccountMaster(), true);
	 }
	
	 
	 @Test(priority=127)
	 public void checkDeleteGroupAccountinAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkDeleteGroupAccountinAccountMaster(), true);
	 }
	 
	 
	 
 ///////////////////Checking the Group Account//////////////////////
	 
	 
	 
	 
	 @Test(priority=128)
	 public void checkOpenGroupAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkOpenGroupAccountMaster(), true);
	 }
	 
	 
	 
	 
	 
	 @Test(priority=129)
	 public void checkSearchinganAccountUnderSubGroup() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkSearchinganAccountUnderSubGroup(), true);
	 }
	 
	 
	
///////////////Creating Extra Fields//////////////////////
	 
	 
	 
	 @Test(priority=132)
	 public void checkCreatingExtraFieldsinAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkCreatingExtraFieldsinAccountMaster(), true);
	 }
	
	
	 @Test(priority=135)
	 public void checkAddedExtraFieldinGeneralTabofAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkAddedExtraFieldinGeneralTabofAccountMaster(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=136)
	 public void checkEditandUpdateExtraFieldinCustomizeMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkEditandUpdateExtraFieldinCustomizeMaster(), true);
	 }
	 
	 
	 
		@Test(priority=137)
		public void checkLogoutAfterExtraFieldUpdated() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			AP=new AccountsPage(getDriver());
			Assert.assertEquals(AP.checkLogoutAccountsPage(), true);
		}
		
		
		 @Test(priority=138)
		 public void CheckLoginAfterExtraFieldUpdated() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 AP=new AccountsPage(getDriver());
			 Assert.assertEquals(AP.CheckLogin(), true);
		 }
	 
	 
	 @Test(priority=139)
	 public void checkUpdatedExtraFieldinGeneralTabofAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkUpdatedExtraFieldinGeneralTabofAccountMaster(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=140)
	 public void checkDeleteExtraFieldinCustomizeMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkDeleteExtraFieldinCustomizeMaster(), true);
	 }
	 
	 
	 @Test(priority=141)
		public void checkLogoutAfterExtraFieldDeleted() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			AP=new AccountsPage(getDriver());
			Assert.assertEquals(AP.checkLogoutAccountsPage(), true);
		}
		
		
		 @Test(priority=142)
		 public void CheckLoginAfterExtraFieldDeleted() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 AP=new AccountsPage(getDriver());
			 Assert.assertEquals(AP.CheckLogin(), true);
		 }
	 
	 @Test(priority=143)
	 public void checkDeletedExtraFieldinGeneralTabofAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkDeletedExtraFieldinGeneralTabofAccountMaster(), true);
	 }
	 
	 
	 
 ///Creating Tree Tab
	 
	 @Test(priority=144)
	 public void checkCreatingaTreeTabinAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkCreatingaTreeTabinAccountMaster(), true);
	 }
	
	 @Test(priority=145)
	 public void checkEditingNewlyCreatedTreeTabDisplay() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkEditingNewlyCreatedTreeTabDisplay(), true);
	 }
	 
	 
	 @Test(priority=146)
	 public void checkDeleteTabOptionForCreateTreeTabDisplayInTabs() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkDeleteTabOptionForCreateTreeTabDisplayInTabs(), true);
	 }
	 
	 
	
	 
	 
	 ///Document Tab
	 
	 @Test(priority=147)
	 public void checkCreateDocumentTabOptionsInCustomizeMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkCreateDocumentTabOptionsInCustomizeMaster(), true);
	 }
 
	// @Test(priority=148)
	 public void checkCreateDocumentTabCreationWithExistingDocumentTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkCreateDocumentTabCreationWithExistingDocumentTab(), true);
	 }
	 
	 
	 
	 
	// @Test(priority=143)//Getting Error Message
	 public void checkCreateCreateTabInCustomizeMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkCreateCreateTabInCustomizeMaster(), true);
	 }
	 
	 
	 
	// @Test(priority=144)
	 public void checkAddExtraFieldOptionInCreateCreateTabDisplayInCustomizeTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkAddExtraFieldOptionInCreateCreateTabDisplayInCustomizeTab(), true);
	 }
	 
	 
	 
	 @Test(priority=149)
	 public void checkLogoutAccountPageAfterTabCreation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkLogoutAccountsPage(), true);
	 }
	 
	 
	 
	 
	 
	 @Test(priority=150)
	 public void checkLoginAfterLogout() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.CheckLogin(), true);
	 }

	 @Test(priority=151)
	 public void checkAfterTabCreationinAccountMasterPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkAfterTabCreationinAccountMasterPage(), true);
	 }
	 
	 

	 
 ///////////////////Importing Data from Excel///////////////////////
	/* 
	 @Test(priority=153)
	 public void checkAccountsListBeforeImportingDataFromExcel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkAccountsListBeforeImportingDataFromExcel(), true);
	 }
	 
	 @Test(priority=154)
	 public void checkAdvanceMasterImportExportAccount() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkAdvanceMasterImportExportAccount(), true);
	 }
	 
	 @Test(priority=155)
	 public void checkImportDatainAdvanceMasterImportExportAccount() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkImportDatainAdvanceMasterImportExportAccount(), true);
	 }
	 
	 @Test(priority=156)
	 public void checkAccountsListAfterImportingDataFromExcel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkAccountsListAfterImportingDataFromExcel(), true);
	 }
	 
	 
	 @Test(priority=157)
	 public void checkImportedAccountDelete() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkImportedAccountDelete(), true);
	 }
	 
	 */
	 
	 
	 @Test(priority=180)
	 public void checkLogoutAccountsPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkLogoutAccountsPage(), true);
	 }
	 
	 
}
