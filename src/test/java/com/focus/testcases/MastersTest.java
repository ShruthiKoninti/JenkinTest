package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.MastersPage;
import com.focus.Pages.MastersPage;
import com.focus.base.BaseEngine;

public class MastersTest extends BaseEngine
{

	MastersPage MP;
	
	
	// Account
	
	 @Test(priority=201)
	 public void checkSavingAccountINAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkSavingAccountINAccountMaster(), true);
	 }
	
	 @Test(priority=202)
	 public void checkEditingInSavedAccountINAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkEditingInSavedAccountINAccountMaster(), true);
	 }
	
	 @Test(priority=203)
	 public void checkDeleteAccountInAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkDeleteAccountInAccountMaster(), true);
	 }
	 
	 
	 
	 // Item
	 
	 
	 @Test(priority=208)
	 public void checkSavingItemInItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkSavingItemInItemMaster(), true);
	 }
	
	 @Test(priority=209)
	 public void checkEditingInSavedItemINItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkEditingInSavedItemINItemMaster(), true);
	 }
	
	 @Test(priority=210)
	 public void checkDeleteItemInItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkDeleteItemInItemMaster(), true);
	 }
	 
	 
	 
	 // Units
	 
	 @Test(priority=215)
	 public void checkSavingUnitsInUnitsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkSavingUnitsInUnitsMaster(), true);
	 }
	
	 @Test(priority=216)
	 public void checkEditingInSavedUnitsINUnitsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkEditingInSavedUnitsINUnitsMaster(), true);
	 }
	
	 @Test(priority=217)
	 public void checkDeleteUnitsInUnitsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkDeleteUnitsInUnitsMaster(), true);
	 }
	 
	 
	 
	 // Department
	 
	 @Test(priority=222)
	 public void checkSavingDepartmentInDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkSavingDepartmentInDepartmentMaster(), true);
	 }
	
	 @Test(priority=223)
	 public void checkEditingInSavedDepartmentINDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkEditingInSavedDepartmentINDepartmentMaster(), true);
	 }
	
	 @Test(priority=224)
	 public void checkDeleteInDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkDeleteInDepartmentMaster(), true);
	 }
	 
	 
	 
	 
	 
	 
	 // Bin
	 
	 @Test(priority=229)
	 public void checkSavingBinsInBinsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkSavingBinsInBinsMaster(), true);
	 }
	
	 @Test(priority=230)
	 public void checkEditingInSavedBinsINBinsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkEditingInSavedBinsINBinsMaster(), true);
	 }
	
	 @Test(priority=231)
	 public void checkDeleteInBinsInBinsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkDeleteInBinsInBinsMaster(), true);
	 }
	 
	 
	 
	 
	 
	 // Warehouse
	 
	 @Test(priority=236)
	 public void checkSavingWarehouseInWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkSavingWarehouseInWarehouseMaster(), true);
	 }
	 
	 @Test(priority=237)
	 public void checkEditingInSavedWarehouseINWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkEditingInSavedWarehouseINWarehouseMaster(), true);
	 }
	 
	 @Test(priority=238)
	 public void checkDeleteInWarehouseInWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkDeleteInWarehouseInWarehouseMaster(), true);
	 }
	 
	 
	 
	 
	 
	 // Currency
	 
	 @Test(priority=243)
	 public void checkUpdatingISOCurrencyCodeSLLCoinName() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkUpdatingISOCurrencyCodeSLLCoinName(), true);
	 }
	 
	 @Test(priority=244)
	 public void checkUpdatingAgainISOCurrencyCodeSLLCoinName() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkUpdatingAgainISOCurrencyCodeSLLCoinName(), true);
	 }
	 
	 
	 
	 
	// Exchange Rate & Exchange History
	 
	 @Test(priority=250)
	 public void checkDeleteOptionInExchangeRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkDeleteOptionInExchangeRate(), true);
	 }
	 
	 @Test(priority=251)
	 public void checkSavingCurrencyExchnageRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkSavingCurrencyExchnageRate(), true);
	 }
	 
	 @Test(priority=252)
	 public void checkExchangeRateHistoryLoadButtonWithBaseCurrency() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkExchangeRateHistoryLoadButtonWithBaseCurrency(), true);
	 }
	 
	 @Test(priority=253)
	 public void checkUpdateExchangeRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkUpdateExchangeRate(), true);
	 }
	 
	 @Test(priority=254)
	 public void checkExchangeRateHistoryAfterUpdating() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkExchangeRateHistoryAfterUpdating(), true);
	 }
	 
	 
	 
	 
	 
	 
	 // Unit Conversion
	 
	 @Test(priority=259)
	 public void checkDeletingUnitConversion() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkDeletingUnitConversion(), true);
	 }
	
	 @Test(priority=260)
	 public void checkSavingUnitConversion() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkSavingUnitConversion(), true);
	 }

	 @Test(priority=261)
	 public void checkUpdatingUnitConversion() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkUpdatingUnitConversion(), true);
	 }
	 
	
	 
	
	// Seller Price Book
	 
	@Test(priority=266)
	public void checkSavingSellerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkSavingSellerPriceBook(), true);
	}
	
	@Test(priority=267)
	public void checkUpdatingSellerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkUpdatingSellerPriceBook(), true);
	}
	
	@Test(priority=268)
	public void checkDeletingSellerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkDeletingSellerPriceBook(), true);
	}
	
	
	
	
	// Buyer Price Book
	
	@Test(priority=274)
	public void checkSavingBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkSavingBuyerPriceBook(), true);
	}
	
	@Test(priority=275)
	public void checkUpdatingBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkUpdatingBuyerPriceBook(), true);
	}
	
	@Test(priority=276)
	public void checkDeletingBuyerPriceBook() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkDeletingBuyerPriceBook(), true);
	}
	
	
	
	// Transaction Authorization
	
	@Test(priority=280)
	public void checkSavingTransactionAuthorizationInPurchaseReturns() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkSavingTransactionAuthorizationInPurchaseReturns(), true);
	}
 
	@Test(priority=281)
	public void checkEditingTransactionAuthorizationInPurchaseReturns() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkEditingTransactionAuthorizationInPurchaseReturns(), true);
	}
 
	@Test(priority=282)
	public void checkDeletingTransactionAuthorizationInPurchaseReturns() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkDeletingTransactionAuthorizationInPurchaseReturns(), true);
	} 

	
	
	
	
	// Master Authorization
	
	@Test(priority=287)
	public void checkSavingMasterAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkSavingMasterAuthorization(), true);
	} 
	
	@Test(priority=288)
	public void checkEditingMasterAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkEditingMasterAuthorization(), true);
	} 
	
	@Test(priority=289)
	public void checkDeletingMasterAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkDeletingMasterAuthorization(), true);
	} 
	
	
	
	// Design Workflow
	
	@Test(priority=295)
	public void checkSavingTestWorkFlow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkSavingTestWorkFlow(), true);
	}
 
	@Test(priority=296)
	public void checkEditingSavedTestWorkFlow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkEditingSavedTestWorkFlow(), true);
	}
 
	@Test(priority=297)
	public void checkDeletingTestWorkFlow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkDeletingTestWorkFlow(), true);
	}
	
	
	
	
	// Security
	
	 @Test(priority=1001)
	 public void verifyPasswordPolicyTestSave() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkPasswordPolicyTestSave(), true);		
	 }

	 
	 @Test(priority=1002)
	 public void verifyUpdateSimplePasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkUpdateTestPasswordPolicy(), true);		
	 }
	 
	 @Test(priority=1003)
	 public void verifyAlertOnDeletePasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 MP=new MastersPage(getDriver());
		 Assert.assertEquals(MP.checkAlertOnDeletePasswordPolicy(), true);
	 }
	 
	 
	 
	 
	 // Profile
	 
	 @Test(priority=1004)
	 public void CheckSavingAllProfileSTwithAlloptionsEnabled() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	 	MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.CheckSavingAllProfileSTwithAlloptionsEnabled(), true);
  	 }
	 
	 @Test(priority=1005)
	 public void checkUpdatingThreSavedProfile() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	 	MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkUpdatingThreSavedProfile(), true);
  	 }
	 
	
	 
	 @Test(priority=1006)
	 public void checkDeleteUpdatedProfile() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	 	MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkDeleteUpdatedProfile(), true);
  	 }
	 
	 
	 
	 //Role
	 
	 @Test(priority=1007)
	 public void checkSaveingRoleWithAllfieldsInAllTabs() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkSaveingRoleWithAllfieldsInAllTabs(), true);
	 }
	 
	 @Test(priority=1008)
	 public void checkEditAndUpdateSavedRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkEditAndUpdateSavedRole(), true);
	 }
	 
	 @Test(priority=1009)
	 public void checkCompareAndDeleteUpdatedRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkCompareAndDeleteUpdatedRole(), true);
	 }
	 
	 
	 
	 //User
	 
	 @Test(priority=1010)
	 public void checkCreateUserWithRestrictionsForAccount() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkCreateUserWithRestrictionsForAccount(), true);
	 }
	 
	 @Test(priority=1011)
	 public void checkUpdatingTheCreatedUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkUpdatingTheCreatedUser(), true);
	 }
	 
	 @Test(priority=1012)
	 public void checkDeletingUpdatedUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    MP=new MastersPage(getDriver());
		Assert.assertEquals(MP.checkDeletingUpdatedUser(), true);
	 }
	  
	
	 
	 
	 
	 
	 
	 
	 
}
