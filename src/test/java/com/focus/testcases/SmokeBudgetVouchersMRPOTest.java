package com.focus.testcases;

import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;
import com.focus.Pages.SmokeBudgetVouchersMRPOPage;
import com.focus.base.BaseEngine;

public class SmokeBudgetVouchersMRPOTest extends BaseEngine
{
	static SmokeBudgetVouchersMRPOPage sbvp;

    //@Test(priority=30001)
  	public void checkSignInToVerifyBudgetsMRPOVouchers() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkSignInToVerifyBudgetsMRPOVouchers(), true);
  	}
    
    @Test(priority=30002)
  	public void checkToOpenMaterialRequisitionVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkToOpenMaterialRequisitionVoucher(), true);
  	}
    
    @Test(priority=30003)
  	public void checkMaterialRequisitionVoucherNewEntryPageHeaderToggleOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkMaterialRequisitionVoucherNewEntryPageHeaderToggleOptions(), true);
  	}
    
    @Test(priority=30004)
  	public void checkBudgetAlertOnMoreQtyInputValue() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkBudgetAlertOnMoreQtyInputValue(), true);
  	}

    @Test(priority=30005)
  	public void checkBudgetAlertOnLessQtyInputValue() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkBudgetAlertOnLessQtyInputValue(), true);
  	}
    
    @Test(priority=30006)
  	public void checkSaveMaterialRequisitionVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkSaveMaterialRequisitionVoucher(), true);
  	}
    
    @Test(priority=30007)
  	public void checkMaterailRequsitionHomePageStatus() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkMaterailRequsitionHomePageStatus(), true);
  	}
    
    @Test(priority=30008)
  	public void checkMaterialRequsitionVoucherCloseButtonInHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkMaterialRequsitionVoucherCloseButtonInHomePage(), true);
  	}
    
    
    // Purcase Orders
    
    @Test(priority=30009)
  	public void checkOpenPurchasesOrderVoucherToDisplayLinksWithoutAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkOpenPurchasesOrderVoucherToDisplayLinksWithoutAuthorization(), true);
  	}
    
    @Test(priority=30010)
  	public void checkPurchasesOrderVoucherPendingLinksOptionsInHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkPurchasesOrderVoucherPendingLinksOptionsInHomePage(), true);
  	}
    
    @Test(priority=30011)
  	public void checkPurchaseOrderVoucherNewEntryPageHeaderToggleOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkPurchaseOrderVoucherNewEntryPageHeaderToggleOptions(), true);
  	}
    
    
    @Test(priority=30012)
  	public void checkPurchasesOrderVoucherPendingLinksOptionsInEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkPurchasesOrderVoucherPendingLinksOptionsInEntryPage(), true);
  	}
    
    
    @Test(priority=30013)
  	public void checkPurchaseOrderVoucherCloseButtonInEntryAndHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkPurchaseOrderVoucherCloseButtonInEntryAndHomePage(), true);
  	}
    
    
    
  
    
   // Logout From Su and Login to Authorize user
    
    @Test(priority=30014)
  	public void checkAlertsOfTransactionAuthorizationLogoutAndLoginCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkAlertsOfTransactionAuthorizationLogoutAndLoginCompany(), true);
  	}
    
   
    @Test(priority=30015)
  	public void checkClickOnAlertsInWrapperLayoutAuthorizeUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkClickOnAlertsInWrapperLayoutAuthorizeUser(), true);
  	}
    
    @Test(priority=30016)
  	public void checkMaterialRequisitionVoucherOpensOnAlertClick() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkMaterialRequisitionVoucherOpensOnAlertClick(), true);
  	}
    
    @Test(priority=30017)
  	public void checInputMoreQtyInAuthorizeUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checInputMoreQtyInAuthorizeUser(), true);
  	}
    
    @Test(priority=30018)
  	public void checkAuthorizeBtnWithInputMoreQty() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkAuthorizeBtnWithInputMoreQty(), true);
  	}
    
    @Test(priority=30019)
  	public void checInputLessQtyInAuthorizeUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checInputLessQtyInAuthorizeUser(), true);
  	}
    
    @Test(priority=30020)
  	public void checkMaterailRequsitionHomePageStatusInAuthUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkMaterailRequsitionHomePageStatusInAuthUser(), true);
  	}
    
    @Test(priority=30021)
  	public void checkMaterialRequsitionVoucherCloseButtonInAuthUserHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkMaterialRequsitionVoucherCloseButtonInAuthUserHomePage(), true);
  	}
    
    
    
    
    
    
    
    @Test(priority=30022)
  	public void checLogoutAndLoginCompanyToSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checLogoutAndLoginCompanyToSU(), true);
  	}
    
    @Test(priority=30023)
  	public void checkOpenPurchasesOrderVoucherToDisplayLinksOnAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkOpenPurchasesOrderVoucherToDisplayLinksOnAuthorization(), true);
  	}
    
    @Test(priority=30024)
  	public void checkPurchasesOrderVoucherPendingLinksOptionsInHomePageOnAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkPurchasesOrderVoucherPendingLinksOptionsInHomePageOnAuthorization(), true);
  	}
    
    @Test(priority=30025)
  	public void checkPurchasesOrderVoucherOpenInHomePageOnAuthorization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkPurchasesOrderVoucherOpenInHomePageOnAuthorization(), true);
  	}
       
    @Test(priority=30026)
  	public void checkPurchaseVoucherInputQtyMoreOnAthorizeVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkPurchaseVoucherInputQtyMoreOnAthorizeVoucher(), true);
  	}
    
    @Test(priority=30027)
  	public void checkLinksPopUpScreenUserInputMoreQty() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkLinksPopUpScreenUserInputMoreQty(), true);
  	}
    
    
    
    @Test(priority=30028)
  	public void checkSavePurchaseOrdersVoucherInUnAuthorizeUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkSavePurchaseOrdersVoucherInUnAuthorizeUser(), true);
  	}
    
    
    
    @Test(priority=30029) //
  	public void checkPurchasesOrderHomePageStatusInUnAuthorizeUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkPurchasesOrderHomePageStatusInUnAuthorizeUser(), true);
  	}
    
    
    @Test(priority=30030)
  	public void checkPurchasesOrdersVoucherCloseButtonInHomePageInUnAuthorize() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkPurchasesOrdersVoucherCloseButtonInHomePageInUnAuthorize(), true);
  	}
    
    
    // Check Authorization in Authorize User for voucher Purchase order
    
    @Test(priority=30031)
  	public void checLogoutAndLoginCompanyToAuthorizeUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checLogoutAndLoginCompanyToAuthorizeUser(), true);
  	}
    
    @Test(priority=30032)
  	public void checkOpenPurchasesOrderVoucherToDisplayPendingAuthorizationInAuthorizeHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkOpenPurchasesOrderVoucherToDisplayPendingAuthorizationInAuthorizeHomePage(), true);
  	}
    
   
    @Test(priority=30033) //
  	public void checkOpenPurchasesOrderVoucherAuthorizeFromHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checkOpenPurchasesOrderVoucherAuthorizeFromHomePage(), true);
  	}
    
    
    @Test(priority=30034)
  	public void checLogoutAndLoginCompanyToUnAuthorizeUserToCheckPVPendingLinks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
  	{
  		sbvp=new SmokeBudgetVouchersMRPOPage(getDriver());
  		Assert.assertEquals(sbvp.checLogoutAndLoginCompanyToUnAuthorizeUserToCheckPVPendingLinks(), true);
  	}
    
   
    
    
}
