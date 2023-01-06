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
	
	// @Test(priority=110)
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
	 
	 
	 
 ///////////////////Importing Data from Excel///////////////////////
	 
	/* @Test(priority=121)
	 public void checkAccountsListBeforeImportingDataFromExcel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkAccountsListBeforeImportingDataFromExcel(), true);
	 }
	 
	 @Test(priority=122)
	 public void checkAdvanceMasterImportExportAccount() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkAdvanceMasterImportExportAccount(), true);
	 }
	 
	 @Test(priority=123)
	 public void checkImportDatainAdvanceMasterImportExportAccount() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkImportDatainAdvanceMasterImportExportAccount(), true);
	 }
	 
	 @Test(priority=124)
	 public void checkAccountsListAfterImportingDataFromExcel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkAccountsListAfterImportingDataFromExcel(), true);
	 }
	 
	 
	 @Test(priority=125)
	 public void checkImportedAccountDelete() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkImportedAccountDelete(), true);
	 }
	 
	 */
	 
	

	 
	 ////////////////////////Checking New Groups and SubGroups added/////////////////////////////
	 
		
	 @Test(priority=128)
	 public void checkSavingAccountGroupinAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkSavingAccountGroupinAccountMaster(), true);
	 }
 
	 @Test(priority=129)
	 public void checkSavingAccountSubGroupinAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkSavingAccountSubGroupinAccountMaster(), true);
	 }
	 
	 
	 
	 @Test(priority=130)
	 public void checkCreatedGroupDisplayedonLeftPanel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkCreatedGroupDisplayedonLeftPanel(), true);
	 }


	 @Test(priority=131)
	 public void checkSavingAccountsinSubGroupsinAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkSavingAccountsinSubGroupsinAccountMaster(), true);
	 }
	 


	 @Test(priority=132)
	 public void checkAccountsSavedinGroupsandSubGroupsinAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkAccountsSavedinGroupsandSubGroupsinAccountMaster(), true);
	 }
	
	 
	 @Test(priority=133)
	 public void checkDeleteGroupAccountinAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkDeleteGroupAccountinAccountMaster(), true);
	 }
	 
	 
	 
 ///////////////////Checking the Group Account//////////////////////
	 
	 
	 
	 
	 @Test(priority=134)
	 public void checkOpenGroupAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkOpenGroupAccountMaster(), true);
	 }
	 
	 
	 
	 
	 
	 @Test(priority=135)
	 public void checkSearchinganAccountUnderSubGroup() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 AP=new AccountsPage(getDriver());
		 Assert.assertEquals(AP.checkSearchinganAccountUnderSubGroup(), true);
	 }
	 
	 
	 
	 
	 
	 
}
