package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.MasterAuthorizationPage;
import com.focus.base.BaseEngine;

public class MasterAuthorizationTest extends BaseEngine
{
	MasterAuthorizationPage MP;

/*	 @Test(priority=211)
	 public void checkSavingMasterAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkSavingMasterAuthorization(), true);
	 }
	 
	 @Test(priority=212)
	 public void checkCreatingWarehouseFromVoucherLevelAndValidatingWarehouseList() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkCreatingWarehouseFromVoucherLevelAndValidatingWarehouseList(), true);
	 }
	 
	 @Test(priority=213)
	 public void checkShowAllUnAuthorizeOptionInAccountsMasterPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkShowAllUnAuthorizeOptionInAccountsMasterPage(), true);
	 }
	
	 @Test(priority=214)
	 public void checkLogoutAndLoginWithUserAllOptionsForWarehouseAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkLogoutAndLoginWithUserAllOptions(), true);
	 }
	 
	 @Test(priority=215)
	 public void checkValidatingAndAuthorizingWarehouseForAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkValidatingAndAuthorizingWarehouseForAuthorization(), true);
	 }
	 
	 @Test(priority=217)
	 public void checkLogoutAndLoginWithSUAfterWarehouseAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkLogoutAndLoginWithSU(), true);
	 }
	 
	 @Test(priority=218)
	 public void checkVerifyingWarehouseInPurchaseVoucherAfterAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkVerifyingWarehouseInPurchaseVoucherAfterAuthorization(), true);
	 }
	 */
	 @Test(priority=219)
	 public void checkEditingMasterAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkEditingMasterAuthorization(), true);
	 }
	
	 @Test(priority=220)
	 public void checkCreatingDemoWarehouseAndTrialWarehouse() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkCreatingDemoWarehouseAndTrialWarehouse(), true);
	 }
 
	 @Test(priority=221)
	 public void checkVerifyingWarehouseCreated() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkVerifyingWarehouseCreated(), true);
	 }
	
	 @Test(priority=222)
	 public void checkLogoutAndLoginWithUserAllOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkLogoutAndLoginWithUserAllOptions(), true);
	 }
	 
	 @Test(priority=223)
	 public void checkValidatingAndAuthorizingTrialWarehouse() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkValidatingAndAuthorizingTrialWarehouse(), true);
	 }
	 
	 @Test(priority=224)
	 public void checkLogoutAndLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkLogoutAndLoginWithSU(), true);
	 }
	 
	 @Test(priority=225)
	 public void checkVerifyingWarehouseCreatedAfterAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkVerifyingWarehouseCreatedAfterAuthorization(), true);
	 }
	 
	 @Test(priority=226)
	 public void checkDeleteInWarehouseInWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkDeleteInWarehouseInWarehouseMaster(), true);
	 }
	 
	 @Test(priority=227)
	 public void checkDeletingMasterAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MasterAuthorizationPage(getDriver());
		 Assert.assertEquals(MP.checkDeletingMasterAuthorization(), true);
	 }
	 
}
