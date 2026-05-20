package com.focus.testcases;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.CompanyCreatePage;
import com.focus.base.BaseEngine;

public class CompanyCreateTest extends BaseEngine{
	
	CompanyCreatePage ccp;
	
	
	
	
	@Test(priority=1)
	public void checkLoginScreen() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkLoginScreen(),true);	
	}

	@Test(priority=2)
	public void checkCreateCompanyInLoginPage() throws InterruptedException, IOException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkCreateCompanyInLoginPage(),true);	
	}
	
	@Test(priority=3)
	public void checkLogintoCreatedCompany() throws InterruptedException, IOException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkLogintoCreatedCompany(),true);	
	}
	
	@Test(priority=4)
	public void checkEditCompanyAfterCompanyCreatedWithSecurityasNormal() throws InterruptedException, IOException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkEditCompanyAfterCompanyCreatedWithSecurityasNormal(),true);	
	}
	
	@Test(priority=5)
	public void checkLogintoCreatedCompanyAfterEdit() throws InterruptedException, IOException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkLogintoCreatedCompany(),true);	
	}
	
	@Test(priority=6)
	public void checkSavingItemsinItemMaster() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkSavingItemsinItemMaster(),true);	
	}
	
	@Test(priority=7)
	public void checkSavingPasswordPolicy() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkSavingPasswordPolicy(),true);	
	}
	
	@Test(priority=8)
	public void CheckSavingProfilewithAlloptionsEnabled() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.CheckSavingProfilewithAlloptionsEnabled(),true);	
	}
	
	@Test(priority=9)
	public void checkSavingRoleWithAllFieldsinAllTabs() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkSavingRoleWithAllFieldsinAllTabs(),true);	
	}
	
	@Test(priority=10)
	public void checkInputMandatoryFieldsInCreateUser() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkInputMandatoryFieldsInCreateUser(),true);	
	}
	
	@Test(priority=11)
	public void checkSavingVoucherinPurchseVoucher() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkSavingVoucherinPurchseVoucher(),true);	
	}
	
	@Test(priority=12)
	public void checkEditingVoucher1inPurchaseVouchers() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkEditingVoucher1inPurchaseVouchers(),true);	
	}
	
	@Test(priority=13)
	public void checkSuspendtheVoucherinPurchaseVouchers() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkSuspendtheVoucherinPurchaseVouchers(),true);	
	}
	
	@Test(priority=14)
	public void checkLogoutAndLoginWithNewlySavedUser() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkLogoutAndLoginWithNewlySavedUser(),true);	
	}
	
	@Test(priority=15)
	public void checkEditingVoucherinPurchaseVouchersatUser() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkEditingVoucherinPurchaseVouchersatUser(),true);	
	}
	
	@Test(priority=16)
	public void checkSupendtheVoucherinPurchaseVouchersfromEntryPageatUser() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkSupendtheVoucherinPurchaseVouchersfromEntryPageatUser(),true);	
	}
	
	@Test(priority=17)
	public void checkLogoutAndLoginWithSU() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkLogoutAndLoginWithSU(),true);	
	}
	
	//Security AS Low
	
	@Test(priority=18)
	public void checkEditCompanyWithSecurityasLow() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkEditCompanyWithSecurityasLow(),true);	
	}
	
	@Test(priority=19)
	public void checkLogintoCreatedCompanyAfterEditSecuritylevelasLow() throws InterruptedException, IOException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkLogintoCreatedCompany(),true);	
	}
	
	@Test(priority=20)
	public void checkSavingVoucher2inPurchaseVoucher() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkSavingVoucher2inPurchaseVoucher(),true);	
	}
	
	@Test(priority=21)
	public void checkEditingandSavingVoucher2inPurchaseVoucher() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkEditingandSavingVoucher2inPurchaseVoucher(),true);	
	}
	
	@Test(priority=22)
	public void checkSuspendVoucher2FromEntryPageinPurchaseVoucher() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkSuspendVoucher2FromEntryPageinPurchaseVoucher(),true);	
	}
	
	@Test(priority=23)
	public void checkLogoutAndLoginWithNewlySavedUserAfterEditSecurityLow() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkLogoutAndLoginWithNewlySavedUser(),true);	
	}
	
	@Test(priority=24)
	public void checkEditingVoucher2inPurchaseVoucheratUserLevel() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkEditingVoucher2inPurchaseVoucheratUserLevel(),true);	
	}
	
	@Test(priority=25)
	public void checkSuspendtheVoucher2inPurchaseVouchers() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkSuspendtheVoucher2inPurchaseVouchers(),true);	
	}
	
	@Test(priority=26)
	public void checkLogoutAndLoginWithSUAgain() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkLogoutAndLoginWithSU(),true);	
	}
	
	///Security as High
	
	@Test(priority=27)
	public void checkEditCompanyWithSecurityasHigh() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkEditCompanyWithSecurityasHigh(),true);	
	}
	
	@Test(priority=28)
	public void checkLogintoCreatedCompanyAfterEditSecuritylevelasHigh() throws InterruptedException, IOException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkLogintoCreatedCompany(),true);	
	}
	
	@Test(priority=29)
	public void checkSavingVoucher3inPurchaseVoucher() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkSavingVoucher3inPurchaseVoucher(),true);	
	}
	
	@Test(priority=30)
	public void checkEditingandSavingVoucher3inPurchaseVoucher() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkEditingandSavingVoucher3inPurchaseVoucher(),true);	
	}
	
	@Test(priority=31)
	public void checkSuspendVoucher3FromEntryPageinPurchaseVoucher() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkSuspendVoucher3FromEntryPageinPurchaseVoucher(),true);	
	}
	
	@Test(priority=32)
	public void checkLogoutAndLoginWithNewlySavedUserAfterEditSecurityHigh() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkLogoutAndLoginWithNewlySavedUser(),true);	
	}
	
	@Test(priority=33)
	public void checkEditingVoucher3inPurchaseVoucheratUserLevel() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkEditingVoucher3inPurchaseVoucheratUserLevel(),true);	
	}
	
	@Test(priority=34)
	public void checkSuspendtheVoucher3inPurchaseVouchers() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkSuspendtheVoucher3inPurchaseVouchers(),true);	
	}
	
	@Test(priority=35)
	public void checkLogoutCompanyCreationPage() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		ccp = new CompanyCreatePage(getDriver());
		
		 Assert.assertEquals(ccp.checkLogoutCompanyCreationPage(),true);	
	}
	
}
