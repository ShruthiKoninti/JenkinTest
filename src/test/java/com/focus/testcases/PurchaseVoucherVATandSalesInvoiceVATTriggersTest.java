package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.PurchaseVoucherVATandSalesInvoiceVATTriggersPage;
import com.focus.Pages.SanityDocCustViewPage;
import com.focus.base.BaseEngine;

public class PurchaseVoucherVATandSalesInvoiceVATTriggersTest extends BaseEngine{
	PurchaseVoucherVATandSalesInvoiceVATTriggersPage PST;
	
	
	
	@Test(priority=99)
	public void checkLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.CheckLogin(), true);
	}
	

	
	@Test(priority=100)
	public void checkEraseAllTransactions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkEraseAllTransactions(), true);
	}
	
	@Test(priority=101)
	public void checkDeleteVATFieldinEditScreenofPurchaseVosucherVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkDeleteVATFieldinEditScreenofPurchaseVoucherVAT(), true);
	}
	
	
	
	@Test(priority=102)
	public void checkSaveOptionInTriggerTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSaveTriggerwithDiscountasMapFieldinPVVAT(), true);
	}
	
	
	
	
	@Test(priority=103)
	public void checkSavingTriggeronWarehouseRaiseSameVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingTriggeronWarehouseRaiseSameVoucher(), true);
	}
	
	
	@Test(priority=104)
	public void checkSavingPurchaseVoucherVATVoucher1withDiscountasMapField() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingPurchaseVoucherVATVoucher1withDiscountasMapField(), true);
	}
	
	
	
	
	
	@Test(priority=105)
	public void checkSavingPurchaseVoucherVATforWarehouseTrigger() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingPurchaseVoucherVATforWarehouseTrigger(), true);
	}
	
	
	
	
	@Test(priority=106)
	public void checkRaisedVoucherAfterSavingTriggerVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkRaisedVoucherAfterSavingTriggerVoucher1(), true);
	}
	
	
	@Test(priority=107)
	public void checkRaisedVoucherAfterSavingTriggerVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkRaisedVoucherAfterSavingTriggerVoucher2(), true);
	}
	
	
	
	
	
	@Test(priority=108)
	public void checkSavingMultipleConditionTriggerinPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingMultipleConditionTriggerinPurchaseVoucherVAT(), true);
	}
	
	
	@Test(priority=109)
	public void checkSavingPurchaseVoucherVATVoucherforMultipleConditionTrigger() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingPurchaseVoucherVATVoucherforMultipleConditionTrigger(), true);
	}
	
	@Test(priority=110)
	public void checkRaisedVoucherforMultipleConditionTrigger() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkRaisedVoucherforMultipleConditionTrigger(), true);
	}
	
	@Test(priority=111)
	public void checkSavingPurchaseVoucherVATVoucher2forMultipleConditionTrigger() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingPurchaseVoucherVATVoucher2forMultipleConditionTrigger(), true);
	}
	
	@Test(priority=112)
	public void checkRaisedVoucherforMultipleConditionTriggerisNotTriggered() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkRaisedVoucherforMultipleConditionTriggerisNotTriggered(), true);
	}
	
	
	
	@Test(priority=113)
	public void checkEnablingDoNotModifyDocumentIfTriggerisRaised() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkEnablingDoNotModifyDocumentIfTriggerisRaised(), true);
	}
	
	
	
	@Test(priority=114)
	public void checkSavingPurchaseVoucherVATVoucherandEditingVoucherafterVoucherTriggered() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingPurchaseVoucherVATVoucherandEditingVoucherafterVoucherTriggered(), true);
	}
	
	
	
	@Test(priority=115)
	public void checkDisablingDoNotModifyDocumentIfTriggerisRaised() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkDisablingDoNotModifyDocumentIfTriggerisRaised(), true);
	}
	
	
	@Test(priority=116)
	public void checkSavingTriggerWithoutEnablingLineWiseCheckBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingTriggerWithoutEnablingLineWiseCheckBox(), true);
	}
	
	@Test(priority=117)
	public void checkSavingPurchaseVoucherVATforDisableLinewiseCheckBoxinTrigger() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingPurchaseVoucherVATforLinewiseCheckBoxinTrigger(), true);
	}
	
	@Test(priority=118)
	public void checkRaisedVoucherafterSavingPurchaseVoucherVATWithoutLineWiseCheckBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkRaisedVoucherafterSavingPurchaseVoucherVATWithoutLineWiseCheckBox(), true);
	}
	
	
	@Test(priority=119)
	public void checkEditingTriggertoEnableLinewiseCheckingCheckBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkEditingTriggertoEnableLinewiseCheckingCheckBox(), true);
	}
	
	
	
	@Test(priority=120)
	public void checkSavingPurchaseVoucherVATforEnableLinewiseCheckBoxinTrigger() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingPurchaseVoucherVATforLinewiseCheckBoxinTrigger(), true);
	}
	
	
	
	@Test(priority=121)
	public void checkRaisedVoucherafterSavingPurchaseVoucherVATWithLineWiseCheckBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkRaisedVoucherafterSavingPurchaseVoucherVATWithLineWiseCheckBox(), true);
	}
	
	
	@Test(priority=122)
	public void checkDeleteTrigger() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkDeleteTrigger(), true);
	}
	
	@Test(priority=123)
	public void checkSavingTriggertoRaiseaDocumentOnSaveandSuspend() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingTriggertoRaiseaDocumentOnSaveandSuspend(), true);
	}
	
	
	@Test(priority=124)
	public void checkSavingPurchaseVoucherVATtoCheckSuspendStatus() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingPurchaseVoucherVATtoCheckSuspendStatus(), true);
	}
	
	
	
	@Test(priority=125)
	public void checkEditingTriggertoEnableLineWiseCheckBoxtoSuspend() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkEditingTriggertoEnableLineWiseCheckBoxtoSuspend(), true);
	}
	
	
	
	
	
	@Test(priority=126)
	public void checkSavingPuchaseVoucherVATtoCheckStausofSuspendwhenLinewiseChkBoxEnabled() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingPuchaseVoucherVATtoCheckStausofSuspendwhenLinewiseChkBoxEnabled(), true);
	}
	
	
	
	
	
//Alert
	
	@Test(priority=127)
	public void checkCreatingAlertinTriggersPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkCreatingAlertinTriggersPurchaseVoucherVAT(), true);
	}
	
	
	
	
	@Test(priority=128)
	public void checkCreatedAlertandUserSelectinPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkCreatedAlertandUserSelectinPurchaseVoucherVAT(), true);
	}
	
	//@Test(priority=129)
	public void checkLogoutAndLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkLogoutAndLoginWithSU(), true);
	}
	
	
	@Test(priority=130)
	public void checkSavingVoucher2inPVVATforAlerts() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingVoucher2inPVVATforAlerts(), true);
	}
	
	
	
	
	@Test(priority=131)
	public void checkLogoutandLoginwithUserAllOptionsSTforAlertMessage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkLogoutandLoginwithUserAllOptionsSTforAlertMessage(), true);
	}
	
	
	
	
	@Test(priority=132)
	public void checkTriggerAlertMessageinUserAllOptionsST() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkTriggerAlertMessageinUserAllOptionsST(), true);
	}
	
	@Test(priority=133)
	public void checkAgainLogoutAndLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkLogoutAndLoginWithSU(), true);
	}
	
	
	
	@Test(priority=134)
	public void checkCreateUserforTriggers() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkCreateUserforTriggers(), true);
	}
	
	
	
	
	@Test(priority=135)
	public void checkEmailTemplateCreationinPurchaseVoucherVATTriggers() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkEmailTemplateCreationinPurchaseVoucherVATTriggers(), true);
	}
	
	
	
	
	@Test(priority=136)
	public void checkCreatedEmailTemplateinPVVATView() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkCreatedEmailTemplateinPVVATView(), true);
	}
	
	
	//@Test(priority=137)
	public void checkAgainLogoutAndLoginWithSUforEmailTemplateValidation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkLogoutAndLoginWithSU(), true);
	}
	
	
	@Test(priority=138)
	public void checkSavingVoucher3inPVVATforEmailValidation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingVoucher3inPVVATforEmailValidation(), true);
	}
	
	
	
	@Test(priority=139)
	public void checkEmailValidationforLogininCreatedTriggerTemplate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkEmailValidationforLogininCreatedTriggerTemplate(), true);
		 
	}
	
	
	@Test(priority=140)
	public void checkCreateEmailTemplateforVendor() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkCreateEmailTemplateforVendor(), true);
	}
	
	@Test(priority=141)
	public void checkSavingPurchaseVoucherVATforVendorEmail() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingPurchaseVoucherVATforVendorEmail(), true);
	}
	
	
	@Test(priority=142)
	public void checkEmailValidationforVendorinCreatedTriggerTemplate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkEmailValidationforVendorinCreatedTriggerTemplate(), true);
	}
	
	
	
	//Role
	 
		@Test(priority=143)
		public void checkSavingTriggertoSendEmailbyRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavingTriggertoSendEmailbyRole(), true);
		}
		
		
		
		
		@Test(priority=144)
		public void checkSavingPurchaseVoucherVATVoucherforSendtoRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavingPurchaseVoucherVATVoucherforSendtoRole(), true);
		}
		
		
		
		@Test(priority=145)
		public void checkEmailValidationforUser1inRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailValidationforUser1inRole(), true);
		}
		
		
		@Test(priority=146)
		public void checkEmailValidationforUser2inRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailValidationforUser2inRole(), true);
		}
		
		
		@Test(priority=147)
		public void checkEmailValidationforUser3inRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailValidationforUser3inRole(), true);
		}
		
		
	
	
	
	
	
	
	@Test(priority=148)
	public void checkCreatingSMSTemplateforTriggersinPVVATVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkCreatingSMSTemplateforTriggersinPVVATVoucher(), true);
	}
	
	@Test(priority=149)
	public void checkCreatedSMSAlertinPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkCreatedSMSAlertinPurchaseVoucherVAT(), true);
	}
	
	
	
	
	///Authorization
	
	@Test(priority=167)
	public void checkSaveTriggerOptionasOnAuthorizationinPurchaseOrder() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSaveTriggerOptionasOnAuthorizationinPurchaseOrder(), true);
	}
	
	
	@Test(priority=168)
	public void checkSavingPurchaseOrdersVouchertoTrigger() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingPurchaseOrdersVouchertoTrigger(), true);
	}
	
	
	
	@Test(priority=169)
	public void checkVerifyingTriggerOptioninPurchaseVoucherBeforeAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkVerifyingTriggerOptioninPurchaseVoucherBeforeAuthorization(), true);
	}
	
	
	@Test(priority=170)
	public void checkLogoutandLoginwithUserAllOptionsSTforAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkLogoutandLoginwithUserAllOptionsSTforAlertMessage(), true);
	}
	
	
	
	
	
	@Test(priority=171)
	public void checkAuthorizationInPurchaseOrderVoucherWithUserAllOptionsST() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkAuthorizationInPurchaseOrderVoucherWithUserAllOptionsSTFromEntryPage(), true);
	}
	
	
	
	
	@Test(priority=172)
	public void checkLogoutAndLoginWithSUAfterAuthorizingPurchaseOrderVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkLogoutAndLoginWithSU(), true);
	}
	
	
	
	@Test(priority=173)
	public void checkVerifingDetailsOfSavedPurchaseVouchersInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkVerifingDetailsOfSavedPurchaseVouchersInHomepage(), true);
	}
	
	
	
	
	@Test(priority=174)
	public void checkCreateEmailTemplateforPurchaseOrderVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkCreateEmailTemplateforPurchaseOrderVoucher(), true);
	}
	
	
	@Test(priority=175)
	public void checkSavingPurchaseOrderVoucherforEmailValidation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingPurchaseOrderVouchertoAuthorize(), true);
	}
	
	
	@Test(priority=176)
	public void checkLogoutandLoginwithUserAllOptionsSTforEmailValidation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkLogoutandLoginwithUserAllOptionsSTforAlertMessage(), true);
	}
	
	
	
	@Test(priority=177)
	public void checkAuthorizationInPurchaseOrderVoucherWithUserAllOptionsSTFromEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkAuthorizationInPurchaseOrderVoucherWithUserAllOptionsSTFromEntryPage(), true);
	}
	
	
	@Test(priority=178)
	public void checkLogoutandLogintoSUAgain() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkLogoutAndLoginWithSU(), true);
	}
	
	
	@Test(priority=179)
	public void checkVerifingDetailsOfSavedPurchaseVouchersInHomepageafterAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkVerifingDetailsOfSavedPurchaseVouchers2InHomepage(), true);
	}
	
	
	
	
	@Test(priority=180)
	public void checkEmailAfterAuthorizationofPurchaseOrderVoucherFromEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkEmailAfterAuthorizationofPurchaseOrderVoucherFromEntryPage(), true);
	}
	
	
	
	
	@Test(priority=181)
	public void SavingPurchaseOrderVouchertoAuthorizeFromHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingPurchaseOrderVouchertoAuthorize(), true);
	}
	
	@Test(priority=182)
	public void checkLogoutandLoginwithUserAllOptionsSTforEmailValidationthroughHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkLogoutandLoginwithUserAllOptionsSTforAlertMessage(), true);
	}
	
	@Test(priority=183)
	public void checkAuthorizationofPurchaseOrderinHomePageatUserALLOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkAuthorizationofPurchaseOrderinHomePageatUserALLOptions(), true);
	}
	
	@Test(priority=184)
	public void checkLogoutandLogintoSUAgainforAthorizationFromHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkLogoutAndLoginWithSU(), true);
	}
	
	@Test(priority=185)
	public void checkVerifingDetailsOfSavedPurchaseVouchers3InHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkVerifingDetailsOfSavedPurchaseVouchers3InHomepage(), true);
	}
	
	@Test(priority=186)
	public void checkEmailAfterAuthorizationofPurchaseOrderVoucherfromHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkEmailAfterAuthorizationofPurchaseOrderVoucherfromHomePage(), true);
	}
	
	
	
	
	///Authorize from email
	
		@Test(priority=187)
		public void checkSavingPurchaseOrderVouchertoAuthorizefromEmail() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavingPurchaseOrderVouchertoAuthorize(), true);
		}
		
		
		@Test(priority=188)
		public void checkAuthorizingVoucherFromEmail() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkAuthorizingVoucherFromEmail(), true);
		}
		
		
		
		
		
		@Test(priority=189)
		public void checkEmailAfterAuthorizationofPurchaseOrderVoucherThroughEmail() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailAfterAuthorizationofPurchaseOrderVoucherThroughEmail(), true);
		}
		
	
	
	
	///Multiple Vouchers for authorization to check Email
	
		@Test(priority=190)
		public void checkSavingMultipleRowsPurchaseOrderVouchertoAuthorizeFromHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavingMultipleRowsPurchaseOrderVouchertoAuthorizeFromHomePage(), true);
		}
		
		
		
		
		@Test(priority=191)
		public void checkSavingMultipleVouchersinPurchaseOrderVouchertoAuthorizeFromHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavingMultipleVouchersinPurchaseOrderVouchertoAuthorizeFromHomePage(), true);
		}
		
		
		
		@Test(priority=192)
		public void checkLogoutandLoginwithUserAllOptionsSTForAutnorizationinHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkLogoutandLoginwithUserAllOptionsSTforAlertMessage(), true);
		}
		
		@Test(priority=193)
		public void checkAuthorizingTwoVouchersinPOVoucherfromHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkAuthorizingTwoVouchersinPOVoucherfromHomePage(), true);
		}
		
		
		@Test(priority=194)
		public void checkEmailAttachementsAfterAuthorizingTwoVouchers() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailAttachementsAfterAuthorizingTwoVouchers(), true);
		}
		
		
		
		@Test(priority=195)
		public void checkAuthorizingThreeVouchersFromHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkAuthorizingThreeVouchersFromHomePage(), true);
		}
		
		
		
		
		
		@Test(priority=196)//some times mails getting with two attachments so here 2 mails only
		public void checkEmailAttachementsAfterAuthorizingThreeVouchers() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailAttachementsAfterAuthorizingThreeVouchers(), true);
		}
		
		
		
		
		
		@Test(priority=197)
		public void checkAuthorizingFiveVouchersinPOVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkAuthorizingFiveVouchersinPOVoucher(), true);
		}
		
		
		@Test(priority=198)//some times mails getting with two attachments so here 4 mails only
		public void checkEmailAttachementsAfterAuthorizingFiveVouchers() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailAttachementsAfterAuthorizingFiveVouchers(), true);
		}
		
		
		
		
		@Test(priority=199)
		public void checkAuthorizingTenVouchersInPOVoucherFromHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkAuthorizingTenVouchersInPOVoucherFromHomePage(), true);
		}
		
		
		
		
		@Test(priority=200)//some times mails getting with two attachments so here 9 mails only
		public void checkEmailAttachementsAfterAuthorizingTenVouchers() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailAttachementsAfterAuthorizingTenVouchers(), true);
		}
		
		
		@Test(priority=201)
		public void checkLogoutandLogintoSUAfterMultipleVouchersAthorizationFromHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkLogoutAndLoginWithSU(), true);
		}
	
		
		//Email for Login Option
		@Test(priority=202)
		public void checkEmailTemplateofLoginforPurchaseOrderVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailTemplateofLoginforPurchaseOrderVoucher(), true);
		}
		
		
		
		@Test(priority=203)
		public void checkSavingPurchaseOrderVouchertoAuthorizeFromHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavingPurchaseOrderVouchertoAuthorize(), true);
		}
		
	
		@Test(priority=204)
		public void checkLogoutandLoginwithUserAllOptionsSTforEmailLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkLogoutandLoginwithUserAllOptionsSTforAlertMessage(), true);
		}
		
		
		
		
		@Test(priority=205)
		public void checkAuthorizationofPurchaseOrderVoucherinHomePageatUserALLOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkAuthorizationofPurchaseOrderinHomePageatUserALLOptions(), true);
		}
		

		@Test(priority=206)
		public void checkLogoutAndLoginWithSUAfterAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkLogoutAndLoginWithSU(), true);
		}
		
		
		
		@Test(priority=207)
		public void checkVerifingDetailsOfSavedPurchaseVouchers4InHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkVerifingDetailsOfSavedPurchaseVouchers4InHomepage(), true);
		}
		
		
		@Test(priority=208)
		public void checkEmailLoginAfterAuthorizationofPurchaseOrderVoucherfromHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailAfterEnablingLoginOptionAfterAuthorizationofPurchaseOrderVoucherfromHomePage(), true);
		}
		
		//Email from Extra Field
		
		@Test(priority=209)
		public void checkEnablingPickEmailFromNarrationinPOVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEnablingPickEmailFromNarrationinPOVoucher(), true);
		}
		
		
		
		
		
		@Test(priority=210)
		public void checkSavingPOVoucher1toSendEmailfromExtraField() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavingPOVoucher1toSendEmailfromExtraField(), true);
		}
		
		
		@Test(priority=211)
		public void checkSavingPOVoucher2toSendEmailfromExtraField() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavingPOVoucher2toSendEmailfromExtraField(), true);
		}
		
		@Test(priority=212)
		public void checkSavingPOVoucher3toSendEmailfromExtraField() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavingPOVoucher3toSendEmailfromExtraField(), true);
		}
		
		
		
		@Test(priority=213)
		public void checkLogoutandLoginwithUserAllOptionsSTforEmailExtraField() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkLogoutandLoginwithUserAllOptionsSTforAlertMessage(), true);
		}
		
		
			
		
		@Test(priority=214)
		public void checkAuthorizingVouchersFromHomePageforExtraFieldEmail() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkAuthorizingVouchersFromHomePageforExtraFieldEmail(), true);
		}
		
		
		
		@Test(priority=215)
		public void checkEmailfromExtraFieldAfterauthorizationofVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailfromExtraFieldAfterauthorizationofVoucher1(), true);
		}
		
		
		@Test(priority=216)
		public void checkEmailfromExtraFieldAfterauthorizationofVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailfromExtraFieldAfterauthorizationofVoucher2(), true);
		}
		
		@Test(priority=217)
		public void checkEmailfromExtraFieldAfterauthorizationofVoucher3() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailfromExtraFieldAfterauthorizationofVoucher3(), true);
		}
		
		
		@Test(priority=218)
		public void checkLogoutAndLoginWithSUAfterAuthorizationforExtraField() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkLogoutAndLoginWithSU(), true);
		}
		
		
	//Send Email by Role	
	
		
		@Test(priority=219)
		public void checkSavingTriggerinPOVouchertoSendMailtoRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavingTriggerinPOVouchertoSendMailtoRole(), true);
		}
		

		
		@Test(priority=220)
		public void checkSavingPOVoucherforRoleEmailValidation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavingPOVoucherforRoleEmailValidation(), true);
		}
		
		
		
		@Test(priority=221)
		public void checkLogoutandLoginwithUserAllOptionsSTforEmailRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkLogoutandLoginwithUserAllOptionsSTforAlertMessage(), true);
		}
		
		
	
		
		@Test(priority=222)
		public void checkAuthorizationofPurchaseOrderinHomePageatUserALLOptionsforRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkAuthorizationofPurchaseOrderinHomePageatUserALLOptions(), true);
		}
		
		
		
		@Test(priority=223)
		public void checkEmailValidationforUser1inRoleAfterAuth() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailValidationforUser1inRoleAfterAuth(), true);
		}
		
		
		
		@Test(priority=224)
		public void checkEmailValidationforUser2inRoleAfterAuth() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailValidationforUser2inRoleAfterAuth(), true);
		}
		
		
		
		@Test(priority=225)
		public void checkEmailValidationforUser3inRoleAfterAuth() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEmailValidationforUser3inRoleAfterAuth(), true);
		}
		
		
		@Test(priority=226)
		public void checkLogoutAndLoginWithSUAfterAuthorizationforRole() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkLogoutAndLoginWithSU(), true);
		}
		
		
		
	
	////Receipts Voucher Email
	@Test(priority=227)
	public void checkSavingEmailTemplateinTriggerinReceiptsVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingEmailTemplateinTriggerinReceiptsVoucher(), true);
	}
	
	
	
	
	
	@Test(priority=228)
	public void checkSaveReceiptsVoucherafterEmailInvoiceCreated() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSaveReceiptsVoucherafterEmailInvoiceCreated(), true);
	}
	
	
	
	@Test(priority=229)
	public void checkEmailValidationAfterReceiptsVoucherSaved() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkEmailValidationAfterReceiptsVoucherSaved(), true);
	}
	
	
	
	///Triggers on Petty Cash 
	
			@Test(priority=230)
			public void checkSavingOpeningBalanceVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkSavingOpeningBalanceVoucher(), true);
			}
			
			
			
			@Test(priority=231)
			public void checkSavingTriggersinPettyCashVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkSavingTriggersinPettyCashVoucher(), true);
			}
			
			
			
			@Test(priority=232)
			public void checkSavingPettyCashVoucherforTriggers() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkSavingPettyCashVoucherforTriggers(), true);
			}
			
			@Test(priority=233)//not getting new reference 
			public void checkRaisedPettyCashNewVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkRaisedPettyCashNewVoucher(), true);
			}
			
		
				
				
				
		////PO Voucher--Purchase Returns for DocNo
			
			
				
				
			@Test(priority=234)
			public void checkSavingTriggerinPOVoucherBasedOnSaveforDocumentNumberSeries() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkSavingTriggerinPOVoucherBasedOnSaveforDocumentNumberSeries(), true);
			}
				
			
			
			@Test(priority=235)
			public void checkSavingPOVoucherforDocumentNumber() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkSavingPOVoucherforDocumentNumber(), true);
			}
				
			
				
				
			@Test(priority=236)
			public void checkRaisedVoucherPurchseReturnforDocNumber() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkRaisedVoucherPurchseReturnforDocNumber(), true);
			}
				
					
				
			
				
			@Test(priority=237)
			public void checkSavingTriggerinPOVoucherBasedonOnAuthorizeforDocNumber() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkSavingTriggerinPOVoucherBasedonOnAuthorizeforDocNumber(), true);
			}
				
			
			
			@Test(priority=238)
			public void checkSavingPOVoucherforAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkSavingPOVoucherforAuthorization(), true);
			}
				
				
			@Test(priority=239)
			public void checkLogoutandLoginwithUserAllOptionsSTforAuthorizePOVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkLogoutandLoginwithUserAllOptionsSTforAlertMessage(), true);
			}
			
			
			
			@Test(priority=240)
			public void checkAuthorizingthePOVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkAuthorizingthePOVoucher(), true);
			}
			
			@Test(priority=241)
			public void checkLogoutandLogintoSUAfterAthorizationMRNVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkLogoutAndLoginWithSU(), true);
			}
				
				
				
				
			@Test(priority=242)
			public void checkRaisedVoucherPurchaseREturnsAfterAuthorizePOVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkRaisedVoucherPurchaseREturnsAfterAuthorizePOVoucher(), true);
			}	
				
			
			
			
			///MRN Trigger to Sales Orders
			
			
			
			@Test(priority=243)
			public void checkUpdatingSOVouchertoAutoReservation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkUpdatingSOVouchertoAutoReservation(), true);
			}	
			
			
			@Test(priority=244)
			public void checkSavingTriggerinMRNVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkSavingTriggerinMRNVoucher(), true);
			}	
			
			
			
			

			@Test(priority=245)
			public void checkSavingMRNVoucherAfterTriggerSaved() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkSavingMRNVoucherAfterTriggerSaved(), true);
			}	
			
			
			
			@Test(priority=246)
			public void checkRaisedSalesOrderVoucherforAutoReservation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			{
				  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
				  Assert.assertEquals(PST.checkRaisedSalesOrderVoucherforAutoReservation(), true);
			}	
			
	
	/*
	
		
	/////Material Receipts Note
		@Test(priority=242)
		public void checkSaveTriggerinMRNVoucherbasedonBPAAmountwithEnableRaiseandOnSaveOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSaveTriggerinMRNVoucherbasedonBPAAmountwithEnableRaiseandOnSaveOption(), true);
		}
		
		
		@Test(priority=243)
		public void checkSaveTriggerinMRNVoucherbasedonBPAAmountwithEnableRaiseandOnAuthorizeOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSaveTriggerinMRNVoucherbasedonBPAAmountwithEnableRaiseandOnAuthorizeOption(), true);
		}
		
		
		
		@Test(priority=244)
		public void checkSaveTriggerinMRNVoucherbasedonBPADAmountwithEnableRaiseandOnSaveOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSaveTriggerinMRNVoucherbasedonBPADAmountwithEnableRaiseandOnSaveOption(), true);
		}
		
		
		
		
		@Test(priority=245)
		public void checkSaveTriggerinMRNVoucherbasedonBPADAmountwithEnableRaiseandOnAuthorizeOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSaveTriggerinMRNVoucherbasedonBPADAmountwithEnableRaiseandOnAuthorizeOption(), true);
		}
		
		
		
		@Test(priority=246)
		public void checkSaveTriggerinMRNVoucherbasedonDebitNoteAmountwithEnableRaiseandOnSaveOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSaveTriggerinMRNVoucherbasedonDebitNoteAmountwithEnableRaiseandOnSaveOption(), true);
		}
		
		
		
		
		@Test(priority=247)
		public void checkSaveTriggerinMRNVoucherbasedonDebitNoteAmountwithEnableRaiseandOnAuthorizeOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSaveTriggerinMRNVoucherbasedonDebitNoteAmountwithEnableRaiseandOnAuthorizeOption(), true);
		}
		
		@Test(priority=248)
		public void checkSavingMRNVouchertoRaiseaTriggerVoucherBasedonBAPAmountandOnSaveOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavingMRNVouchertoRaiseaTriggerVoucherBasedonBAPAmount(), true);
		}
		
		
		
		@Test(priority=249)
		public void checkJournalEntriesRaisedVoucherafterSavingTriggerMRNVOucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkJournalEntriesRaisedVoucherafterSavingTriggerMRNVOucher(), true);
		}
		
		
		@Test(priority=250)
		public void checkEditingMRNVoucherAfterRaiseDocumentOnSave() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkEditingMRNVoucherAfterRaiseDocumentOnSave(), true);
		}
		
		
		
		
		@Test(priority=251)
		public void checkJournalEntriesVoucherAfterEditingMRNVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkJournalEntriesVoucherAfterEditingMRNVoucher(), true);
		}
		
		
		
		@Test(priority=252)
		public void checkLogoutandLoginwithUserAllOptionsSTforAuthorizeMRNVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkLogoutandLoginwithUserAllOptionsSTforAlertMessage(), true);
		}
		
		
		
		@Test(priority=253)
		public void checkAuthorizationInMRNVoucherWithUserAllOptionsST() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkAuthorizationInMRNVoucherWithUserAllOptionsST(), true);
		}
		
		@Test(priority=254)
		public void checkLogoutandLogintoSUAfterAthorizationMRNVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkLogoutAndLoginWithSU(), true);
		}
		
		
		
		@Test(priority=255)
		public void checkJournalEntriesVoucherAfterAuthorizationMRNVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkJournalEntriesVoucherAfterAuthorizationMRNVoucher(), true);
		}

		

		//Views Creation in PVVAT
	/*	
		
		
		@Test(priority=256)
		public void checkSavingViewinPurchaseVoucherVATVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavingViewinPurchaseVoucherVATVoucher(), true);
		}
		
		
		
		@Test(priority=257)
		public void checkCreatedDepartmentViewinPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkCreatedDepartmentViewinPurchaseVoucherVAT(), true);
		}
		
		
		
		
		@Test(priority=258)
		public void checkSavingVendorViewinPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavingVendorViewinPurchaseVoucherVAT(), true);
		}
		
		
		
		
		
		@Test(priority=259)
		public void checkSavedViewofVendorinPVVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavedViewofVendorinPVVAT(), true);
		}
		
		
		
		
		
		@Test(priority=260)
		public void checkSavingWarehouseViewinPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavingWarehouseViewinPurchaseVoucherVAT(), true);
		}
		
		
		@Test(priority=261)
		public void checkSavedViewofWarehouseinPVVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSavedViewofWarehouseinPVVAT(), true);
		}
		
		
		
		
		@Test(priority=262)
		public void checkSaveFilterPermenantlycheckBoxinHomeFilterOptioninPVVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkSaveFilterPermenantlycheckBoxinHomeFilterOptioninPVVAT(), true);
		}
		
		
		
		@Test(priority=263)
		public void checkDisablingSaveFilterPermenantlyChkBoxinPVVATVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
			  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
			  Assert.assertEquals(PST.checkDisablingSaveFilterPermenantlyChkBoxinPVVATVoucher(), true);
		}
		*/
		
		
		
		
		
	@Test(priority=270)
	public void checkLogoutPurchaseandSalesInvoiceVoucherVATPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkLogoutPurchaseandSalesInvoiceVoucherVATPage(), true);
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	///////////////////////////////////////////////////////No need-------------------------------------------------------------
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
////Sales Invoice VAT
	/*
	
	@Test(priority=125)
	public void checkDeleteVATFieldinEditScreenofSalesInvoiceVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkDeleteVATFieldinEditScreenofSalesInvoiceVAT(), true);
	}
	
	
	@Test(priority=126)
	public void checkSaveTriggerwithDiscountasMapFieldinSalesInvoiceVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSaveTriggerwithDiscountasMapFieldinSalesInvoiceVAT(), true);
	}
	
	@Test(priority=127)
	public void checkSavingSalesInvoiceVATVoucher1withDiscountasMapField() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingSalesInvoiceVATVoucher1withDiscountasMapField(), true);
	}
	
	@Test(priority=128)
	public void checkCashSalesVoucherafterTriggerOptionsSavewithMapField() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkCashSalesVoucherafterTriggerOptionsSavewithMapField(), true);
	}
	
	@Test(priority=129)
	public void checkCreatingAlertinTriggersSalesInvoiceVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkCreatingAlertinTriggersSalesInvoiceVAT(), true);
	}
	
	@Test(priority=130)
	public void checkCreatedAlertandUserSelectinSalesInvoiceVATVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkCreatedAlertandUserSelectinSalesInvoiceVATVoucher(), true);
	}
	
	@Test(priority=131)
	public void checkSavingVoucher2inSalesInvoiceVATforAlertMessage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingVoucher2inSalesInvoiceVATforAlertMessage(), true);
	}
	
	
	@Test(priority=132)
	public void checkLogoutandLoginwithUserAllOptionsSTforAlertMessageofSalesInvoiceVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkLogoutandLoginwithUserAllOptionsSTforAlertMessage(), true);
	}
	
	@Test(priority=133)
	public void checkTriggerAlertMessageinUserAllOptionsSTofSalesInvoiceVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkTriggerAlertMessageinUserAllOptionsSTofSalesInvoiceVAT(), true);
	}
	
	@Test(priority=134)
	public void checkLogoutAndLoginWithSUforSalesInvoiceVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkLogoutAndLoginWithSU(), true);
	}
	
	
	
	
	
	
	@Test(priority=135)
	public void checkEmailTemplateCreationinSalesInvoiceVATTriggers() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkEmailTemplateCreationinSalesInvoiceVATTriggers(), true);
	}
	
	
	@Test(priority=136)
	public void checkCreatedEmailTemplateinSalesInvoiceVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkCreatedEmailTemplateinSalesInvoiceVAT(), true);
	}
	
	
	
	@Test(priority=137)
	public void checkSavingVoucher3inSalesInvoiceVATforEmailValidation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingVoucher3inSalesInvoiceVATforEmailValidation(), true);
	}
	
	
	@Test(priority=138)
	public void checkEmailValidationUserforCreatedSalesTriggerTemplate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkEmailValidationUserforCreatedSalesTriggerTemplate(), true);
	}
	
	@Test(priority=139)
	public void checkCreateEmailTemplateforCustomer() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkCreateEmailTemplateforCustomer(), true);
	}
	
	@Test(priority=140)
	public void checkSavingVoucher4inSalesInvoiceVATforCustomerEmailValidation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkSavingVoucher4inSalesInvoiceVATforCustomerEmailValidation(), true);
	}
	
	@Test(priority=141)
	public void checkEmailValidationUserforCustomerCreatedSalesTriggerTemplate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkEmailValidationUserforCreatedSalesTriggerTemplate(), true);
	}
	
	
	@Test(priority=142)
	public void checkCreatingSMSTemplateforTriggersinSalesInvoiceVATVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkCreatingSMSTemplateforTriggersinSalesInvoiceVATVoucher(), true);
	}
	
	@Test(priority=143)
	public void checkCreatedSMSAlertinSalesInvoiceVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		  PST=new PurchaseVoucherVATandSalesInvoiceVATTriggersPage(getDriver());	
		  Assert.assertEquals(PST.checkCreatedSMSAlertinSalesInvoiceVAT(), true);
	}
	
	
	
	*/
	
}
