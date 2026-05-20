package com.focus.testcases;

import com.focus.base.BaseEngine;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.SmokeSecurityPage;

public class SmokeSecurityTest extends BaseEngine
{
	
	static SmokeSecurityPage ssp;
	
	
	 @Test(priority=1006)
	 public void CheckLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 ssp=new SmokeSecurityPage(getDriver());
		 Assert.assertEquals(ssp.CheckLogin(), true);
	 }

	 
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
	
	 @Test(priority=1105)
	 public void checkDetailsofSUUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkDetailsofSUUser(), true);
	 } 
  
	//@Test(priority=1106)
	 public void checkLogoutAndLoginWithSUBeforeCreateUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkLogoutAndLoginWithSU(), true);
	 }
	 
	 @Test(priority=1107)
	 public void checkCreateUserWithAllMandatoryFileds() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkCreateUserWithAllMandatoryFileds(), true);
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
	 
	 @Test(priority=1127)//checks Vendor B
	 public void checkPurchasesVoucherForRestrictedAccountsInUserThroughRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkPurchasesVoucherForRestrictedAccountsInUserThroughRole(), true);
	 }
	 
	 @Test(priority=1128)//checks Customer C
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
	 
	 @Test(priority=1131)//checks Vendor B
	 public void checkPurchasesVocherForRestrictedAccountThroughUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkPurchasesVocherForRestrictedAccountThroughUser(), true);
	 }
	 
	 @Test(priority=1132)//checks Customer C
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
	 
	//Do not allow Copy/Paste 
	 
	 @Test(priority=1135)
	 public void checkDonotallowcopypasteinprofileinuserlevel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkDonotallowcopypasteinprofileinuserlevel(), true);
	 }
	 
	 @Test(priority=1136)
	 public void checkLogoutAndLoginTestUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkLogoutAndLoginTestUser(), true);
	 }
	 
	 @Test(priority=1137)
	 public void VerifyDonotallowcopypasteDisplayedInPurchasevoucherInuserlevel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.VerifyDonotallowcopypasteDisplayedInPurchasevoucherInuserlevel(), true);
	 }
	 
	 @Test(priority=1138)
	 public void EnableDonotallowcopypasteinprofileInSuUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.EnableDonotallowcopypasteinprofileInSuUser(), true);
	 }
	 
	 @Test(priority=1139)
	 public void VerifyDonotallowcopypasteInPurchasevoucherInuserlevel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.VerifyDonotallowcopypasteInPurchasevoucherInuserlevel(), true);
	 }
	 
	 @Test(priority=1140)
	 public void CheckDonotallowcopypastewithRoleinUserlevel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.checkDonotallowcopypastewithRoleinUserlevel(), true);
	 }
	 
	 @Test(priority=1141)
	 public void VerifyDonotallowcopypastewithRoleinUserlevel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.VerifyDonotallowcopypasteDisplayedInPurchasevoucherInuserlevel(), true);
	 }
	 
	 /*
	 
	 //Password Policy 
	 
	 @Test(priority=1140)
	 public void  checkComplexityofPasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkComplexityofPasswordPolicy(), true);
	 }
	
	 
	 
	 @Test(priority=1141)
	 public void  checkCreateaNewUserforPasswordasAlphaNumeric() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkCreateaNewUserforPasswordasAlphaNumeric(), true);
	 }
	
	 
	 
	 @Test(priority=1142)
	 public void  checkComplexityofPasswordPolicyasAlphaNumericwithRestriction() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkComplexityofPasswordPolicyasAlphaNumericwithRestriction(), true);
	 }
	 
	 @Test(priority=1143)
	 public void  checkChangePasswordwhileChangePasswordPolicyComplexityAsAplhaNumericwithRestriction() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkChangePasswordwhileChangePasswordPolicyComplexityAsAplhaNumericwithRestriction(), true);
	 }
	 
	 @Test(priority=1144)
	 public void   checkComplexityofPasswordPolicyasAlphaNumericandSpecialCharcter() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.  checkComplexityofPasswordPolicyasAlphaNumericandSpecialCharcter(), true);
	 }
	 
	 @Test(priority=1145)
	 public void  checkChangePasswordwhileChangePasswordPolicyComplexityAsAplhaNumericandSpecialCharacter() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkChangePasswordwhileChangePasswordPolicyComplexityAsAplhaNumericandSpecialCharacter(), true);
	 }
	 
	 @Test(priority=1146)
	 public void   checkComplexityofPasswordPolicyasAlphaNumericwandSpecialCharacterithRestriction() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.  checkComplexityofPasswordPolicyasAlphaNumericwandSpecialCharacterithRestriction(), true);
	 }
	 
	 @Test(priority=1147)
	 public void  checkChangePasswordwhileChangePasswordPolicyComplexityAsAplhaNumericandSpecialCharacterWithRestriction() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkChangePasswordwhileChangePasswordPolicyComplexityAsAplhaNumericandSpecialCharacterWithRestriction(), true);
	 }
	 
	 
	 
	 @Test(priority=1148)
	 public void  checkDoNotAllowPreviousPasswordinPasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkDoNotAllowPreviousPasswordinPasswordPolicy(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=1149)
	 public void  checkSavingNewUserafterDoNotAllowPreviousPasswordinPasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkSavingNewUserafterDoNotAllowPreviousPasswordinPasswordPolicy(), true);
	 }
	 
	 
	 
	 @Test(priority=1150)
	 public void  checkLogoutAndLoginWithPwdUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithPwdUser(), true);
	 }
	 
	 
	 @Test(priority=1151)
	 public void  checkChangePasswordforNewUsertoCheckDoNotAllowPreviousPassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkChangePasswordforNewUsertoCheckDoNotAllowPreviousPassword(), true);
	 }
	 
	 
	 
	 @Test(priority=1152)
	 public void  checkLogoutAndLoginWithSUAgain() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithSUAgain(), true);
	 }
	 
	 
	 
	 @Test(priority=1153)
	 public void  checkDeletingPwdUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkDeletingPwdUser(), true);
	 }
	 
	 
	 @Test(priority=1154)
	 public void  checkInvalidAtetmptsandLockOutPeriodinPasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkInvalidAtetmptsandLockOutPeriodinPasswordPolicy(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=1155)
	 public void  checkLogoutAndLoginWithNewUserforInvalidAttempts() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithNewUserforInvalidAttempts(), true);
	 }
	
	 @Test(priority=1156)
	 public void  checkLogoutAndLoginWithSUAgainAfterInvalidAttempts() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithSUAgain(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=1157)
	 public void  checkSendEMailonLoginSuccessinPasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkSendEMailonLoginSuccessinPasswordPolicy(), true);
	 }
	 
	 
	 @Test(priority=1158)
	 public void  checkLogoutAndLoginWithNewUserAgain() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithNewUser(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=1159)
	 public void  checkEmailAfterSuccessfulLoginofNewUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkEmailAfterSuccessfulLoginofNewUser(), true);
	 }
	 
	 
	 @Test(priority=1160)
	 public void  checkLogoutAndLoginWithSUAgainafterEmailSuccess() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithSUAgain(), true);
	 }
	 
	 
	 
	 @Test(priority=1161)
	 public void  checkSendEMailonLoginFailureinPasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkSendEMailonLoginFailureinPasswordPolicy(), true);
	 }
	 
	 
	 @Test(priority=1162)
	 public void  checkLogoutAndLoginWithNewUserforIncorrectPwd() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithNewUserforIncorrectPwd(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=1163)
	 public void  checkEmailAfterFailureLoginofNewUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkEmailAfterFailureLoginofNewUser(), true);
	 }
	 
	 
	 
	 @Test(priority=1164)
	 public void  checkLoginWithSUAfterEmail() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLoginWithSU(), true);
	 }
	 

	 
	 @Test(priority=1165)
	 public void   checkChangePasswordafterFirstLogininPasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp.  checkChangePasswordafterFirstLogininPasswordPolicy(), true);
	 }
	 
	 
	 @Test(priority=1166)
	 public void  checkCreateaNewUserforPasswordChangeafterFirstLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkCreateaNewUserforPasswordChangeafterFirstLogin(), true);
	 }
	 
	 
	 @Test(priority=1167)
	 public void  checkLogoutAndLoginWithDemoUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithDemoUser(), true);
	 }
	 
	 
	 @Test(priority=1168)
	 public void  checkChangePasswordTitleaftertLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkChangePasswordTitleaftertLogin(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=1169)
	 public void  checkLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLoginWithSU(), true);
	 }
	 
	 
	 
	 
	 
	 @Test(priority=1170)
	 public void  checkCannotChangePasswordinPasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkCannotChangePasswordinPasswordPolicy(), true);
	 }
	 
	 
	 
	 @Test(priority=1171)
	 public void  checkLogoutAndLoginWithNewUserAgainforPwd() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithNewUser(), true);
	 }
	 
	
	 
	 
	 @Test(priority=1172)
	 public void  checkCannotChangePasswordforNewUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkCannotChangePasswordforNewUser(), true);
	 }
	 
	 
	 
	 @Test(priority=1173)
	 public void  checkLogoutAndLoginWithSUAgainafterCannotChangePassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithSUAgain(), true);
	 }
	 
	 
	 
	 @Test(priority=1174)
	 public void  checkOTPBasedLogininPasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkOTPBasedLogininPasswordPolicy(), true);
	 }
	 
	 
	 
	 
	 
	@Test(priority=1175)
	 public void  checkLogoutAndLoginWithNewUserforOTPLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithNewUserforOTPLogin(), true);
	 }
	 
	 	 
	 
	@Test(priority=1176)
	 public void  checkEmailafterOTPSent() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkEmailafterOTPSent(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=1177)
	 public void  checkInvalidAtetmptsandLockOutPeriodinPasswordPolicyAgain() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkInvalidAtetmptsandLockOutPeriodinPasswordPolicyAgain(), true);
	 }
	 
	 
	 
	 
	 
	 @Test(priority=1178)
	 public void  checkLogoutAndLoginWithNewUserforInvalidAttemptsAgain() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithNewUserforInvalidAttemptsAgain(), true);
	 }
	 
	 
	 
	 
	 
	 @Test(priority=1179)
	 public void  checkDoNotLockAccountinNewUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkDoNotLockAccountinNewUser(), true);
	 }
	 
	 
	 
	 @Test(priority=1180)
	 public void  checkLogoutAndLoginWithNewUserAfterDonotLockCheckbox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithNewUser(), true);
	 }
	 
	 @Test(priority=1181)
	 public void  checkLogoutAndLoginWithSUAfterDoNotLockCheckBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithSUAgain(), true);
	 }
	 
	 
	 
	 @Test(priority=1182)
	 public void  checkAccountDisabledinNewUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkEnableAccountDisabledinNewUser(), true);
	 }
	 
	 
	 @Test(priority=1184)
	 public void  checkLogoutandLoginwithNewUserAfterAccountDisabledChkBoxEnabled() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutandLoginwithNewUserAfterAccountDisabledChkBoxEnabled(), true);
	 }
	 
	 
	 @Test(priority=1185)
	 public void  checkLoginWithSUAfterAccountDisabledChkBoxEnabled() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLoginWithSUAfterAccountDisabledChkBoxEnabled(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=1186)
	 public void  checkDisableAccountDisabledinNewUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkDisableAccountDisabledinNewUser(), true);
	 }
	 
	 @Test(priority=1187)
	 public void  checkLogoutAndLoginWithNewUserAfterDisableAccountDisableChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithNewUserAfterDisableAccountDisableChkBox(), true);
	 }
	 
	 
	 @Test(priority=1188)
	 public void  checkLogoutAndLoginWithSUAfterDisableAccountDisableChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithSUAgain(), true);
	 }
	 
	
	 
	
	 
	 
	 @Test(priority=1189)
	 public void  checkEnableAllowMultiLoginCheckBoxinNewUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkEnableAllowMultiLoginCheckBoxinNewUser(), true);
	 }
	 
	 
	 
	 @Test(priority=1200)
	 public void  checkLogoutAndLoginWithNewUserAfterAllowMultiLoginCheckbox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithNewUser(), true);
	 }
	 
	 
	 
	 @Test(priority=1201)
	 public void  checkNewUsertoAllowMultiLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkNewUsertoAllowMultiLogin(), true);
	 }
	 
	 @Test(priority=1202)
	 public void  checkLogoutAndLoginWithSUAfterEnableAllowMultiLoginCheckBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithSUAgain(), true);
	 }
	 
	 
	 @Test(priority=1203)
	 public void  checkDisableAllowMultiLoginCheckBoxinNewUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkDisableAllowMultiLoginCheckBoxinNewUser(), true);
	 }
	 
	 @Test(priority=1204)
	 public void  checkLogoutAndLoginWithNewUserAfterDisableAllowMultiLoginCheckbox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithNewUser(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=1205)
	 public void  checkNewUserDonotAllowMultiLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkNewUserDonotAllowMultiLogin(), true);
	 }
	 
	 @Test(priority=1206)
	 public void  checkLoginSUAgain() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. CheckLogin(), true);
	 }
	 
	 
	 @Test(priority=1207)
	 public void  checkSendEmailNotificationinNewUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkSendEmailNotificationinNewUser(), true);
	 }
	 
	 @Test(priority=1208)
	 public void  checkLogoutAndLoginWithNewUserAfterSendEmailNotificationsCheckbox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutAndLoginWithNewUser(), true);
	 }
	 
	  
	 
	 @Test(priority=1209)
	 public void  checkEmailAfterSuccessfulLoginofUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkEmailAfterSuccessfulLoginofUser(), true);
	 }
	 
	 
	 
	 @Test(priority=1210)
	 public void  checkLogoutandLoginNewUserforLoginFailure() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutandLoginNewUserforLoginFailure(), true);
	 }
	 
	 @Test(priority=1211)
	 public void  checkEmailAfterLoginFailureofUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkEmailAfterLoginFailureofUser(), true);
	 }
	 
	 
	 @Test(priority=1212)
	 public void  checkLoginSUAgainAfterEmail() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. CheckLogin(), true);
	 }
	 
	 
	 
	 
	 
 //Unlock User
	 
	 @Test(priority=1213)
	 public void  checkLogoutandLoginforUserstoLocktheUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutandLoginforUserstoLocktheUser(), true);
	 }
	 
	 
	@Test(priority=1214)
	 public void  checkLoginSUAgainAfterLockedAccounts() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. CheckLogin(), true);
	 }
	 
	 
	 @Test(priority=1215)
	 public void  checkUnlockUserOptioninCreateUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkUnlockUserOptioninCreateUser(), true);
	 }
	 
	 
	 ///Group user
	 
	 
	 
	 @Test(priority=1216)
	 public void  checkCreatingGroupUserinCreateUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkCreatingGroupUserinCreateUser(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=1217)
	 public void  checkCreateUserFromLoadFromOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkCreateUserFromLoadFromOption(), true);
	 }
	 
	 
	 
	 
	 
	 @Test(priority=1218)
	 public void  checkCreatingNewUserFromLoadFromSearchOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkCreatingNewUserFromLoadFromSearchOption(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=1219)
	 public void  checkResetOptioninCreateUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkResetOptioninCreateUser(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=1220)
	 public void  checkSearchOptioninCreateUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkSearchOptioninCreateUser(), true);
	 }
	 
	 */
	 @Test(priority=2000)
	 public void  checkLogoutSmokeSecurityPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
	    ssp=new SmokeSecurityPage(getDriver());
		Assert.assertEquals(ssp. checkLogoutSmokeSecurityPage(), true);
	 }
	 
	 
	 
	 
	 

	 /*
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 ////////////////////////////////////No Need///////////////////////////////////////////////////////////////
	 
	/* 
	 @Test(priority=1135)
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
	 }
	
	 
	*/
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
}