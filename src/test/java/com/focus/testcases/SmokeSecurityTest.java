package com.focus.testcases;

import com.focus.base.BaseEngine;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.SmokeSecurityPage;

public class SmokeSecurityTest extends BaseEngine
{
	
	static SmokeSecurityPage ssp;
	 
	 // PassWord Policy
	 
	 @Test(priority=1008)
	 public void verifyPasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 ssp=new SmokeSecurityPage(getDriver());
		 Assert.assertEquals(ssp.checkPasswordPolicy(), true);
	 }
	 
	 @Test(priority=1013)
	 public void verifyPasswordPolicyTestSave() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 ssp=new SmokeSecurityPage(getDriver());
		 Assert.assertEquals(ssp.checkPasswordPolicyTestSave(), true);		
	 }
	
	 @Test(priority=1014)
	 public void verifyPasswordPolicyLists() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 ssp=new SmokeSecurityPage(getDriver());
		 Assert.assertEquals(ssp.checkPasswordPolicyLists(), true);
	 }
	 
	 @Test(priority=1015)
	 public void verifyGetTheSimplePasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 ssp=new SmokeSecurityPage(getDriver());
		 Assert.assertEquals(ssp.checkGetTheTestPasswordPolicy(), true);
	 }
	 
	 @Test(priority=1016)
	 public void verifyUpdateSimplePasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 ssp=new SmokeSecurityPage(getDriver());
		 Assert.assertEquals(ssp.checkUpdateTestPasswordPolicy(), true);		
	 }
	 
	 @Test(priority=1017)
	 public void verifyGetTheUpdatedPasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 ssp=new SmokeSecurityPage(getDriver());
		 Assert.assertEquals(ssp.checkGetTheUpdatedPasswordPolicy(), true);
	 }
	 
	 @Test(priority=1018)
	 public void verifyAlertOnDeletePasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 ssp=new SmokeSecurityPage(getDriver());
		 Assert.assertEquals(ssp.checkAlertOnDeletePasswordPolicy(), true);
	 }
	 
	
	
	 //create profile
	 
	 @Test(priority=1055)
	 public void CheckSavingAllProfileSTwithAlloptionsEnabled() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	 	ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.CheckSavingAllProfileSTwithAlloptionsEnabled(), true);
  	 }
	 
	 @Test(priority=1056)
	 public void checkCreateProfileLoadFromAllProfileAndValidatePurchasesVouchers() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkCreateProfileLoadFromAllProfileAndValidatePurchasesVouchers(), true);
	 }
	 
	 @Test(priority=1057)
	 public void checkCreateProfileSavingValidatingAndDeletingLoadedProfile() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkCreateProfileSavingValidatingAndDeletingLoadedProfile(), true);
	 }
	 
	
	 

	 
	 //Role
	
	 @Test(priority=1079)
	 public void checkSaveingRoleWithAllfieldsInAllTabs() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkSaveingRoleWithAllfieldsInAllTabs(), true);
	 }
	 
	 @Test(priority=1080)
	 public void checkEditAndUpdateSavedRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkEditAndUpdateSavedRole(), true);
	 }
	 
	 @Test(priority=1082)
	 public void checkCompareAndDeleteUpdatedRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkCompareAndDeleteUpdatedRole(), true);
	 }
	 
	   
	 
	 	  
	 
	 // Create User

	 @Test(priority=1107)
	 public void checkCreateUserWithAllMandatoryFileds() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkCreateUserWithAllMandatoryFileds(), true);
	 } 
	
	@Test(priority=1108)
	 public void checkLogoutAndLoginWithCreatedUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkLogoutAndLoginWithCreatedUser(), true);
	 } 
	
	
	@Test(priority=1109)
	 public void checkLogoutLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkLogoutAndLoginWithSU(), true);
	 } 
	
	
	@Test(priority=1110)
	 public void checkUpdatingTheCreatedUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkUpdatingTheCreatedUser(), true);
	 } 
	
	
	 @Test(priority=1111)
	 public void checkDeletingUpdatedUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkDeletingUpdatedUser(), true);
	 }
	
	 @Test(priority=1112)
	 public void checInputMandatoryFieldsInCreateUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checInputMandatoryFieldsInCreateUser(), true);
	 } 
	 
	  @Test(priority=1113)
	 public void checkCreateUserWithRestrictionsForAccount() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkCreateUserWithRestrictionsForAccount(), true);
	 }
	
	 @Test(priority=1121)
	 public void checkLogoutAndLoginWithNewlySavedUser_Testing() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkLogoutAndLoginWithNewlySavedUser(), true);
	 }
	 
	 @Test(priority=1122)
	 public void checkLogoutAndLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkLogoutAndLoginWithSU(), true);
	 }
	 
	 
	 @Test(priority=1123)
	 public void checkEditingAndUpdatingSavedUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkEditingAndUpdatingSavedUser(), true);
	 }
	 
	
	 @Test(priority=1125)
	 public void checkLogoutAndLoginWithUpdatedUserWithPreviousPassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkLogoutAndLoginWithUpdatedUserWithPreviousPassword(), true);
	 }
	 
	 
	 
	 @Test(priority=1126)
	 public void checkLoginWithUpdatedPassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkLoginWithUpdatedPassword(), true);
	 }
	 
	 
	 
	 @Test(priority=1127)//issue with pronghorn
	 public void checkPurchasesVoucherForRestrictedAccountsInUserThroughRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkPurchasesVoucherForRestrictedAccountsInUserThroughRole(), true);
	 }
	 
	 @Test(priority=1128)//issue with Pronghorn
	 public void checkSalesInvoicesVATForRestrictedAccountsInUserThroughRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkSalesInvoicesVATForRestrictedAccountsInUserThroughRole(), true);
	 }
	 
	 @Test(priority=1129)
	 public void checkLedgerForRestrictedAccountsInUserThroughRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkLedgerForRestrictedAccountsInUserThroughRole(), true);
	 }
	 
	 @Test(priority=1130)
	 public void checkLogoutAndLoginWithCreatedRestrictedUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkLogoutAndLoginWithCreatedRestrictedUser(), true);
	 }
	 

	 @Test(priority=1131)//issue with Pronghorn
	 public void checkPurchasesVocherForRestrictedAccountThroughUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkPurchasesVocherForRestrictedAccountThroughUser(), true);
	 }
	 
	 
	 @Test(priority=1132)//issue with Pronghorn
	 public void checkSalesInvoiceVATForRestrictedAccountsThroughUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkSalesInvoiceVATForRestrictedAccountsThroughUser(), true);
	 }
	 
	 @Test(priority=1133)
	 public void checkLedgerForRestrictedAccountThroughUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkLedgerForRestrictedAccountThroughUser(), true);
	 }
	 
	
	 @Test(priority=1134)
	 public void checkLogoutAndLoginWithSUToUpdatedUserDeleteUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkLogoutAndLoginWithSUToUpdatedUserDeleteUser(), true);
	 }
	 
		
	 
	 
	 
	/* @Test(priority=1135)
	 public void checkChangePasswordatUserLevelwithNewPassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkChangePasswordatUserLevelwithNewPassword(), true);
	 }
	 
	 
	 
	 @Test(priority=1136)
	 public void  checkLoginangLogoutSUwithChangePassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLoginangLogoutSUwithChangePassword(), true);
	 }
	
	 @Test(priority=1137)
	 public void checkChangePasswordatUserLevelwithOldPassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkChangePasswordatUserLevelwithOldPassword(), true);
	 }
	 
	 
	 
	 @Test(priority=1138)
	 public void  checkLoginangLogoutSUwithChangePasswordUpdate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLoginangLogoutSUwithChangePasswordUpdate(), true);
	 }*/
}
